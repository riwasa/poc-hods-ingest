import azure.functions as func
import datetime
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


def _to_blob_metadata_key(column_name: str) -> str:
    # Blob metadata keys are restricted to a subset of ASCII characters.
    normalized = re.sub(r"[^0-9A-Za-z_]", "_", (column_name or "").strip())
    if not normalized:
        return "sharepoint_column"
    if not re.match(r"^[A-Za-z_]", normalized):
        normalized = f"m_{normalized}"
    return normalized.lower()


def _to_blob_metadata_value(raw_value: object) -> str:
    if raw_value is None:
        return ""
    if isinstance(raw_value, (dict, list)):
        text = str(raw_value)
    else:
        text = str(raw_value)
    # Azure Blob metadata values are ASCII-only; drop unsupported chars.
    return text.encode("ascii", errors="ignore").decode("ascii")


def _get_item_field_value(
    drive_id: str,
    item_id: str,
    column_name: str,
    headers: Dict[str, str],
) -> Optional[str]:
    if not column_name:
        return None

    item_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}?$expand=listItem($expand=fields)"
    item_json = _graph_get(item_url, headers)

    normalized_column_name = column_name.strip().lower()
    if normalized_column_name in {"name", "filename", "fileleafref"}:
        item_name = item_json.get("name")
        if item_name is not None:
            return _to_blob_metadata_value(item_name)

    fields = ((item_json.get("listItem") or {}).get("fields") or {})
    raw_value = fields.get(column_name)
    if raw_value is None:
        return None
    return _to_blob_metadata_value(raw_value)


def _upload_changed_files(
    blob_service_client: BlobServiceClient,
    container_name: str,
    drive_id: str,
    last_sync: datetime.datetime,
    headers: Dict[str, str],
    sharepoint_metadata_column: Optional[str] = None,
    blob_metadata_key: Optional[str] = None,
    max_files: int = 5,
) -> int:
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

        metadata = None
        if sharepoint_metadata_column:
            field_value = _get_item_field_value(drive_id, item_id, sharepoint_metadata_column, headers)
            if field_value is not None:
                metadata_name = blob_metadata_key or _to_blob_metadata_key(sharepoint_metadata_column)
                metadata = {metadata_name: field_value}

        blob_client.upload_blob(content_response.content, overwrite=True, metadata=metadata)
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
    sharepoint_metadata_column = os.getenv("SHAREPOINT_METADATA_COLUMN")
    blob_metadata_key = os.getenv("BLOB_METADATA_KEY")

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
            last_sync=last_sync,
            headers=headers,
            sharepoint_metadata_column=sharepoint_metadata_column,
            blob_metadata_key=blob_metadata_key,
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
