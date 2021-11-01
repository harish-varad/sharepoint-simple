# sharepoint_simple

**sharepoint_simple** is a python library to upload and download the files from SharePoint.\
It uses SharePoint REST services at the backend to perform CRUD operations and runs on app-only principals.

## Pre-Requisites
**sharepoint_simple** uses App-Only model for OAuth authentication.\
Detailed SharePoint app creation is demonstrated here.\
Client ID, Client Secret and Tenant ID is required to communicate with your SharePoint.


Read this [SharePoint documentation](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs/) to generate Client ID, Client Secret and Tenant ID.
You must have Owner access to your SharePoint site to create an app.
## Methods
There are four methods in **sharepoint_simple**
- connect
- download
- upload
- create_folder
