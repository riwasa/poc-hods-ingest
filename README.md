# poc-hods-ingest

## Settings

These settings must be updated in the local.settings.json file when running locally. Some settings, those with a value of
Yes in the 'Needs to be added to Azure Function App' column below, must be added to the Function App in Azure under 
Settings / Environment variables / App settings.

| Setting name | Setting value | Needs to be added to Azure Function App |
| ------------ | ------------- | --------------------------------------- |
| FUNCTIONS_WORKER_RUNTIME | python | No |
| AzureWebJobsStorage | UseDevelopmentStorage=true when running locally | No; use the pre-defined value in Azure |
| BLOB_STORAGE_CONNECTION_STRING | DefaultEndpointsProtocol=https;AccountName=<account-name>;AccountKey=<account-key>;EndpointSuffix=core.windows.net | Yes. Whether running locally or in Azure, replace <account-name> and <account-key>. |
| BLOB_CONTAINER_NAME | Name of container where files should be written, e.g. ingest-output | Yes | 
| SHAREPOINT_TENANT_ID | Tenant id of the service principal used to connect to SharePoint | Yes |
| SHAREPOINT_CLIENT_ID | Client id of the service principal used to connect to SharePoint | Yes |
| SHAREPOINT_CLIENT_SECRET | Secret for the service principal used to connect to SharePoint | Yes |
| SHAREPOINT_SITE_HOSTNAME | Hostname of the SharePoint site (e.g. contoso.sharepoint.com) | Yes | 
| SHAREPOINT_SITE_PATH | /sites/YourSiteName | Yes | 
| SHAREPOINT_LIBRARY_DRIVE_NAME | Documents | Yes |
| SHAREPOINT_METADATA_COLUMN | Name of a single SharePoint column to copy as a metadata value on the blob | Yes | 
| BLOB_METADATA_KEY | Name of the blob metadata key to hold the SharePoint column value. If not provided, the SHAREPOINT_METADATA_COLUMN name will be used | Yes |

## Description

This Azure Function App pulls changed SharePoint files (using the lastModifiedDateTime) since the value contained in a blob named 'last-sync' in the Storage Account. A max of 
5 files are then copied from SharePoint to the Storage Account and stored in BLOB_CONTAINER_NAME. The app then updates the last-sync blob with
the current date/time.
	
Fill in the SharePoint app settings in local.settings.json and ensure your Entra app has Graph application permissions (typically Sites.Read.All, 
or permissions set at a more restrictive level), then run the function host.

## Requirements

- Python 3.13
- Visual Studio Code
- Azure Function Core Tools (https://learn.microsoft.com/en-us/azure/azure-functions/functions-run-local?pivots=programming-language-python&tabs=windows%2Cisolated-process%2Cnode-v4%2Cpython-v2%2Chttp-trigger%2Ccontainer-apps#install-the-azure-functions-core-tools)
- Azurite storage emulator (https://learn.microsoft.com/en-us/azure/storage/common/storage-install-azurite?toc=%2Fazure%2Fstorage%2Fblobs%2Ftoc.json&bc=%2Fazure%2Fstorage%2Fblobs%2Fbreadcrumb%2Ftoc.json&tabs=visual-studio%2Cblob-storage)

# Setup

- Setup and activate a Python virtual environment.
- Install requirements
  - pip install -r requirements.txt
- Change to the directory with the code.
- On Windows in a PowerShell terminal
  - if (!(Test-Path .azurite)) { New-Item -ItemType Directory .azurite | Out-Null }; $env:NODE_OPTIONS=''; npx -y azurite --location .azurite --silent
	- func start
  - If the above command says port 7071 is busy, then use
  - func start --port 7072