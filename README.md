# sharepointsimple

**sharepointsimple** is a python library to upload and download the files from SharePoint.\
It uses SharePoint REST services to perform CRUD operations and runs on SharePoint's app-only principals using App-Only model for OAuth authentication.


## Pre-Requisites

*Client ID*, *Client Secret* and *Tenant ID* is required to communicate with your SharePoint.

Read this [documentation](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs/) to generate Client ID, Client Secret and Tenant ID.

>*Note:*\
> *You must have Owner access to a SharePoint site to generate Client ID and Client Secret*  


## Installation
```sh
pip install sharepointsimple
```


## Methods

**sharepointsimple** has four methods
- [connect](#connect)
- [download](#download)
- [upload](#upload)
- [create_folder](#create_folder)



## Usage
### Step 1 ::
### **connect**
```sh
import sharepointsimple as ss

ss.connect(clientid = 'Your-ClientID',
           clientsecret = 'Your-ClientSecret',
           tenantid = "Your-TenantID", # A method to get TenantID is at end of the page
           domain = 'yourCompany',     #Eg: As in "yourCompany.sharepoint.com"
           SP_sitename = 'yourSharePointSite'
           )
```

> - You must use ***'connect'*** method first, before using ***download*** or ***upload*** method
> - [Know how to generate Tenant ID](#how-to-get-tenant-id)

### Step 2 (1) :: 
### **download**
```sh
ss.download(local_path = "/Users/Folder",
            SP_path = "Folder/SubFolder",           #SP path starts from the root folder directly inside a SP Sites
            files_to_download = "file1.xlsx,file2.txt" #(Optional) Remove this to download all the files in SP folder
           )
```

> - Multiple file names should be given in comma seperated as a single string (list is also accepted)
> - Remove argument ***filename*** if you want to download all the files in the folder i.e: SP_path
> - You need not include *'Shared Documents'*, a default folder, under which all the files are present in

### Step 2 (2) :: 
### **upload**
```sh
ss.upload(SP_path = "Folder/SubFolder",
          local_path = "/Users/Folder",
          files_to_download = "file1.xlsx,file2.txt", #(Optional) Remove this to upload all the files in local system
         )
```

*Hola! You are done*

---

### Optional Step ::
### **create_folder**

***create_folder*** is an *optional* method if you want to create a new folder\
It is ***NOT*** necessary to call this method before uploading a file to a non-existing folder in SP\
***upload*** method already has a built in functionality to create a folder if it doesn't exist in SP site

```sh
ss.create_folder(SP_path = "Folder/SubFolder")
```
> Folder will be created in the SharePoint path\
> New folder will not be created if the folder is already present in SharePoint.


## License
*MIT*\
*Hit it, it costs nothing!* :blush:

---

## How to get Tenant ID
1) Open any browser
2) Type in the address bar, https://login.microsoftonline.com/[YourCompany].onmicrosoft.com/.well-known/openid-configuration\
     Replace "[YourCompany]" with your organisation name
3) In the response, look at the key named *"token_endpoint"*
3) Your Tenant ID is, after  ***https://login.microsoftonline.com/*** and before ***/oauth2/token***\
>  Example:\
>  In the below response, Tenant ID is, ***05200b0x-xxx2-4xx4-x25x-020ab5865xx3***\
>      ***"token_endpoint":"https://login.microsoftonline.com/05200b0x-xxx2-4xx4-x25x-020ab5865xx3/oauth2/token"***\
> 


