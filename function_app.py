import azure.functions as func
import datetime
import json
import logging
import os
import re
from typing import Dict, Iterable, List, Optional

import requests

from azure.storage.blob import BlobServiceClient

app = func.FunctionApp()


def _parse_last_sync(last_sync_raw: Optional[str]) -> datetime.datetime:
    if not last_sync_raw:
        return datetime.datetime(1970, 1, 1, tzinfo=datetime.timezone.utc)

    raw_value = last_sync_raw.strip()
    if not raw_value:
        return datetime.datetime(1970, 1, 1, tzinfo=datetime.timezone.utc)

    # Support the previous format and ISO-8601 values for backward compatibility.
    formats = ["%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%SZ"]
    for fmt in formats:
        try:
            parsed = datetime.datetime.strptime(raw_value, fmt)
            if parsed.tzinfo is None:
                parsed = parsed.replace(tzinfo=datetime.timezone.utc)
            return parsed.astimezone(datetime.timezone.utc)
        except ValueError:
            continue

    try:
        normalized = raw_value.replace("Z", "+00:00")
        parsed = datetime.datetime.fromisoformat(normalized)
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=datetime.timezone.utc)
        return parsed.astimezone(datetime.timezone.utc)
    except ValueError:
        logging.warning("Unrecognized last-sync format '%s'; defaulting to epoch", raw_value)
        return datetime.datetime(1970, 1, 1, tzinfo=datetime.timezone.utc)


def _get_graph_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_response = requests.post(
        token_url,
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=30,
    )
    logging.info(f"Token response status: {token_response.status_code}")
    if token_response.status_code != 200:
        logging.error(f"Token request failed: {token_response.text}")
    token_response.raise_for_status()
    token_json = token_response.json()
    access_token = token_json.get("access_token")
    if not access_token:
        raise RuntimeError("Graph token response did not include access_token")
    logging.info("Successfully obtained Graph access token")
    return access_token


def _graph_get(url: str, headers: Dict[str, str]) -> Dict:
    response = requests.get(url, headers=headers, timeout=60)
    response.raise_for_status()
    return response.json()


def _get_site_id(hostname: str, site_path: str, headers: Dict[str, str]) -> str:
    site_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"
    site_json = _graph_get(site_url, headers)
    site_id = site_json.get("id")
    if not site_id:
        raise RuntimeError("Unable to resolve SharePoint site id")
    return site_id


def _resolve_site_id(hostname: str, site_path: str, headers: Dict[str, str], site_id_override: Optional[str] = None) -> str:
    if site_id_override:
        logging.info("Using configured SharePoint site id override: %s", site_id_override)
        return site_id_override
    return _get_site_id(hostname, site_path, headers)


def _normalize_drive_name(drive_name: str) -> str:
    return re.sub(r"\s+", " ", (drive_name or "").strip()).lower()


def _get_drive_id(site_id: str, drive_name: str, headers: Dict[str, str]) -> str:
    next_link = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives?$top=200"
    requested_drive_name = _normalize_drive_name(drive_name)
    candidate_drive: Optional[Dict] = None
    while next_link:
        drives_response = _graph_get(next_link, headers)
        for drive in drives_response.get("value", []):
            current_drive_name = _normalize_drive_name(drive.get("name", ""))
            if current_drive_name == requested_drive_name:
                drive_id = drive.get("id")
                if drive_id:
                    return drive_id

            if candidate_drive is None and drive.get("id"):
                candidate_drive = drive
        next_link = drives_response.get("@odata.nextLink")

    if candidate_drive and candidate_drive.get("id"):
        return candidate_drive["id"]

    raise RuntimeError(f"Drive '{drive_name}' not found in site '{site_id}'")


def _list_all_items(drive_id: str, headers: Dict[str, str]) -> Iterable[Dict]:
    # Breadth-first listing of all items in a drive to support nested folders.
    queue: List[str] = ["root"]
    while queue:
        parent = queue.pop(0)
        next_link = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{parent}/children?$top=200"
        while next_link:
            children_page = _graph_get(next_link, headers)
            for item in children_page.get("value", []):
                yield item
                if "folder" in item:
                    child_id = item.get("id")
                    if child_id:
                        queue.append(child_id)
            next_link = children_page.get("@odata.nextLink")


def _to_blob_name(file_name: str) -> str:
    base_name = os.path.basename((file_name or "").strip())
    normalized = re.sub(r"[^0-9A-Za-z._-]+", "_", base_name)
    normalized = re.sub(r"_+", "_", normalized).strip("._-")
    if not normalized:
        normalized = "file"
    return normalized


def _to_blob_metadata_value(raw_value: object) -> str:
    if raw_value is None:
        return ""
    if isinstance(raw_value, list):
        # Multi-value lookup columns are returned as a list of
        # {"LookupId": ..., "LookupValue": ...} objects.  Extract just
        # the display values and serialise as a JSON array.
        values: List[str] = []
        for item in raw_value:
            if isinstance(item, dict):
                lookup_value = item.get("LookupValue") or item.get("lookupValue")
                values.append(str(lookup_value) if lookup_value is not None else str(item))
            else:
                values.append(str(item))
        text = json.dumps(values, ensure_ascii=True)
    elif isinstance(raw_value, dict):
        # Single-value lookup columns are returned as {"LookupId": ..., "LookupValue": ...}.
        lookup_value = raw_value.get("LookupValue") or raw_value.get("lookupValue")
        text = str(lookup_value) if lookup_value is not None else str(raw_value)
    else:
        text = str(raw_value)
    # Azure Blob metadata values are ASCII-only; drop unsupported chars.
    return text.encode("ascii", errors="ignore").decode("ascii")


def _get_drive_list_id(drive_id: str, headers: Dict[str, str]) -> str:
    """Return the SharePoint list GUID for the document library behind a drive."""
    drive_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}?$select=id,sharePointIds"
    drive_json = _graph_get(drive_url, headers)
    list_id = ((drive_json.get("sharePointIds") or {}).get("listId"))
    if not list_id:
        raise RuntimeError(f"Could not resolve SharePoint list ID for drive '{drive_id}'")
    return list_id


def _get_lookup_column_info(site_id: str, list_id: str, column_name: str, headers: Dict[str, str]) -> Optional[Dict[str, str]]:
    """Return lookup metadata (target list + target column) for a list column."""
    columns_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/columns?$select=name,displayName,lookup"
    columns_json = _graph_get(columns_url, headers)
    target = (column_name or "").strip().lower()

    for column in columns_json.get("value", []):
        name = (column.get("name") or "").strip().lower()
        display_name = (column.get("displayName") or "").strip().lower()
        if target not in {name, display_name}:
            continue

        lookup = column.get("lookup") or {}
        lookup_list_id = lookup.get("listId")
        if not lookup_list_id:
            return None

        # The lookup source column is often "Title" when the list displays Name.
        # Keep a robust fallback chain when reading the lookup item fields.
        lookup_column = lookup.get("columnName") or "Title"
        return {
            "lookup_list_id": str(lookup_list_id),
            "lookup_column": str(lookup_column),
        }

    return None


def _get_lookup_item_display_value(
    site_id: str,
    lookup_list_id: str,
    lookup_item_id: object,
    lookup_column: str,
    headers: Dict[str, str],
) -> Optional[str]:
    """Resolve a lookup item ID to human-readable text from the lookup list."""
    if lookup_item_id is None:
        return None

    lookup_item_id_text = str(lookup_item_id).strip()
    if not lookup_item_id_text:
        return None

    select = f"{lookup_column},Title,Name"
    lookup_url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{lookup_list_id}"
        f"/items/{lookup_item_id_text}?$expand=fields($select={select})"
    )

    try:
        lookup_json = _graph_get(lookup_url, headers)
    except Exception as exc:
        logging.warning("Failed to resolve lookup item %s from list %s: %s", lookup_item_id_text, lookup_list_id, exc)
        return None

    fields = (lookup_json.get("fields") or {})
    for candidate in [lookup_column, "Title", "Name"]:
        value = fields.get(candidate)
        if value is not None and str(value).strip():
            return str(value)
    return None


def _fetch_item_fields(
    drive_id: str,
    item_id: str,
    site_id: str,
    list_id: str,
    headers: Dict[str, str],
) -> Dict:
    """Return the SharePoint list-item fields dict for a drive item.

    Uses the sites/lists endpoint which reliably returns lookup display values
    (LookupValue) that the drive-based endpoint sometimes omits.
    """
    # Step 1: resolve the SharePoint list item integer ID from the drive item.
    sp_id_url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"?$expand=listItem($select=id)"
    )
    sp_id_json = _graph_get(sp_id_url, headers)
    sp_item_id = ((sp_id_json.get("listItem") or {}).get("id"))
    if not sp_item_id:
        logging.warning("Could not resolve SharePoint item ID for drive item %s", item_id)
        return {}

    # Step 2: fetch fields via the sites/lists endpoint.  This path is more
    # feature-complete and returns LookupValue for lookup columns.
    select = "PrefixLookupId,PrefixLookupValue,HODSContentType"
    fields_url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}"
        f"/items/{sp_item_id}?$expand=fields($select={select})"
    )
    fields_json = _graph_get(fields_url, headers)
    fields = (fields_json.get("fields") or {})
    logging.info("Item %s available field keys: %s", item_id, sorted(fields.keys()))
    return fields


def _upload_changed_files(
    blob_service_client: BlobServiceClient,
    container_name: str,
    drive_id: str,
    site_id: str,
    last_sync: datetime.datetime,
    headers: Dict[str, str],
    max_files: int = 5,
) -> int:
    # Resolve the SharePoint list ID once for the whole run.
    list_id = _get_drive_list_id(drive_id, headers)
    prefix_lookup_info = _get_lookup_column_info(site_id, list_id, "Prefix", headers)
    uploaded = 0
    for item in _list_all_items(drive_id, headers):
        if uploaded >= max_files:
            break

        if "file" not in item:
            continue

        modified_raw = item.get("lastModifiedDateTime")
        if not modified_raw:
            continue

        modified_at = _parse_last_sync(modified_raw)
        if modified_at <= last_sync:
            continue

        item_id = item.get("id")
        file_name = item.get("name")
        if not item_id or not file_name:
            continue

        content_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
        content_response = requests.get(content_url, headers=headers, timeout=120)
        content_response.raise_for_status()

        blob_name = _to_blob_name(file_name)
        blob_client = blob_service_client.get_blob_client(container=container_name, blob=blob_name)

        # Fetch SharePoint list-item fields once for all metadata columns.
        fields = _fetch_item_fields(drive_id, item_id, site_id, list_id, headers)

        metadata: Dict[str, str] = {}

        # Always capture the last-modified timestamp.
        metadata["Modified"] = _to_blob_metadata_value(modified_raw)

        # Always capture the single-value "Prefix" lookup column as text.
        prefix_raw = fields.get("PrefixLookupValue")
        if prefix_raw is None:
            prefix_lookup_id = fields.get("PrefixLookupId")
            if prefix_lookup_id is not None and prefix_lookup_info:
                prefix_raw = _get_lookup_item_display_value(
                    site_id=site_id,
                    lookup_list_id=prefix_lookup_info["lookup_list_id"],
                    lookup_item_id=prefix_lookup_id,
                    lookup_column=prefix_lookup_info["lookup_column"],
                    headers=headers,
                )

        if prefix_raw is not None:
            metadata["Prefix"] = _to_blob_metadata_value(prefix_raw)
        else:
            logging.warning(
                "Could not resolve Prefix display value for item %s. Available field keys: %s",
                item_id, sorted(fields.keys()),
            )

        # Always capture the multi-value "HODS Content Type" lookup column.
        hods_content_type_raw = fields.get("HODSContentType")
        if hods_content_type_raw is not None:
            metadata["ContentType"] = _to_blob_metadata_value(hods_content_type_raw)
        else:
            logging.warning(
                "'HODSContentType' field not found for item %s. Available field keys: %s",
                item_id, sorted(fields.keys()),
            )

        blob_client.upload_blob(content_response.content, overwrite=True, metadata=metadata or None)
        uploaded += 1

    return uploaded

@app.timer_trigger(schedule="0 */1 * * * *", arg_name="myTimer", run_on_startup=False,
              use_monitor=False) 
def Ingest(myTimer: func.TimerRequest) -> None:
    
    if myTimer.past_due:
        logging.info('The timer is past due!')

    now_utc = datetime.datetime.now(datetime.timezone.utc)
    blob_connection_string = os.getenv("BLOB_STORAGE_CONNECTION_STRING")
    container_name = os.getenv("BLOB_CONTAINER_NAME", "ingest-output")
    tenant_id = os.getenv("SHAREPOINT_TENANT_ID")
    client_id = os.getenv("SHAREPOINT_CLIENT_ID")
    client_secret = os.getenv("SHAREPOINT_CLIENT_SECRET")
    site_hostname = os.getenv("SHAREPOINT_SITE_HOSTNAME")
    site_path = os.getenv("SHAREPOINT_SITE_PATH")
    site_id_override = os.getenv("SHAREPOINT_SITE_ID")
    drive_name = os.getenv("SHAREPOINT_LIBRARY_DRIVE_NAME", "Documents")

    if not blob_connection_string:
        logging.error("Missing app setting: BLOB_STORAGE_CONNECTION_STRING")
        return

    required_sharepoint_settings = {
        "SHAREPOINT_TENANT_ID": tenant_id,
        "SHAREPOINT_CLIENT_ID": client_id,
        "SHAREPOINT_CLIENT_SECRET": client_secret,
        "SHAREPOINT_SITE_HOSTNAME": site_hostname,
        "SHAREPOINT_SITE_PATH": site_path,
    }
    missing = [k for k, v in required_sharepoint_settings.items() if not v]
    if missing:
        logging.error("Missing SharePoint app settings: %s", ", ".join(missing))
        return

    blob_service_client = BlobServiceClient.from_connection_string(blob_connection_string)

    # Read last-sync value from blob storage
    last_sync_raw = None
    try:
        last_sync_blob_client = blob_service_client.get_blob_client(container=container_name, blob="last-sync")
        last_sync_raw = last_sync_blob_client.download_blob().readall().decode("utf-8")
        logging.info("Last sync raw value: %s", last_sync_raw)
    except Exception:
        logging.info("No last-sync blob found, this may be the first run")

    last_sync = _parse_last_sync(last_sync_raw)
    logging.info("Using last-sync timestamp (UTC): %s", last_sync.isoformat())

    try:
        container_client = blob_service_client.get_container_client(container_name)
        container_client.create_container()
    except Exception:
        # Container may already exist, so continue with upload attempt.
        pass

    try:
        token = _get_graph_token(tenant_id, client_id, client_secret)
        headers = {"Authorization": f"Bearer {token}"}
        site_id = _resolve_site_id(site_hostname, site_path, headers, site_id_override)
        drive_id = _get_drive_id(site_id, drive_name, headers)

        uploaded_count = _upload_changed_files(
            blob_service_client=blob_service_client,
            container_name=container_name,
            drive_id=drive_id,
            site_id=site_id,
            last_sync=last_sync,
            headers=headers,
            max_files=5,
        )
        logging.info("Completed SharePoint sync. Files uploaded: %s", uploaded_count)
    except Exception as exc:
        logging.exception("Failed to sync SharePoint files: %s", exc)
        return
    
    # Update last-sync value in blob storage
    try:
        last_sync_time = now_utc.isoformat()
        config_blob_client = blob_service_client.get_blob_client(container=container_name, blob="last-sync")
        config_blob_client.upload_blob(last_sync_time, overwrite=True)
        logging.info("Updated last-sync to: %s", last_sync_time)
    except Exception as exc:
        logging.exception("Failed to update last-sync blob: %s", exc)
