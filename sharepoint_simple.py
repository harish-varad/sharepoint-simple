try:import requests
except:raise ModuleNotFoundError("No module named 'requests' found. Please install 'requests' module.")

import os

#Connect
def connect(**kwargs):
    global access_token,domain_name,spsitename
    
    clientid = kwargs.get('clientid',None)
    clientsecret= kwargs.get('clientsecret',None)
    tenantid= kwargs.get('tenantid',None)
    spurl= kwargs.get('SP_url',None)
    domain=kwargs.get('domain',None)
    spsitename=kwargs.get('SP_sitename',None)
    
    try:
        spurl=spurl.replace(".sharepoint.com","")
    except:pass
    try:
        spurl=spurl.replace("https://","")
    except:pass
    
    if domain==None:
        domain_name=spurl
    else:domain_name=domain
    
    url = "https://accounts.accesscontrol.windows.net/"+tenantid+"/tokens/OAuth/2"
    payload = {'grant_type': 'client_credentials',
            'client_id': clientid+'@'+tenantid,
            'client_secret': clientsecret,
            'resource': '00000003-0000-0ff1-ce00-000000000000/'+domain_name+".sharepoint.com@"+tenantid
            }
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    response = requests.request("POST", url, headers=headers, data = payload)
    
    if "[200]" not in str(response):
        raise RuntimeError("Unable to generate the access token. "+str(response)+". Possible reasons are, Wrong ClientID/ClientSecret/TenantID/Domain Name \n\n Usually it looks like \n 1) ClientID: 36f9167x-09xe-4da6-xxxx-a339823b238x \n 2) Client Secret: XX+c7qSTKnE3ZF0yLnOC5x1mx3xJXXlN7xxxmauMhCY= \n 3) TenantID: 05273b5d-xxx2-4fe4-x00x-020ab2222xx3 \n If you are unaware of what ClientID/ClientSecret/TenantID is, visit 'https' to generate one.")
    print("Access token generated successfully. ",response)
    
    access_token = response.json().get('access_token')
    

def download(**kwargs):
    try: access_token
    except:raise RuntimeError("Connect to the SharePoint APP using : sharepoint_simple.connect() \n For more details please refer the documentation at 'https' ")

    sharepointfolder= kwargs.get('SP_path',None)
    localPath=kwargs.get('local_path',None)
    filenamesp=kwargs.get('filename',None)
    spsitename=kwargs.get('SP_sitename',None)
    
    sharepointfolder="Shared Documents/"+sharepointfolder
    
    requestGetURL = 'https://'+domain_name+'.sharepoint.com/sites/'+spsitename+'/_api/web/GetFolderByServerRelativeUrl(\'' + sharepointfolder + '\')/files'
    update_headers = {
                'Accept':'application/json; odata=verbose',
                'Content-Type':'application/json; odata=verbose',
                'Authorization': 'Bearer {}'.format(access_token)
            } 
    data = requests.get(requestGetURL, headers=update_headers)
    data = data.json()
    fileNameData = data['d']['results']
    
    if filenamesp!=None:    
        if "." not in filenamesp:
            raise RuntimeError("Filename with extension is missing. Files with only valid extension can be downloaded.")
        else:
            if type(filenamesp) is not list:
                try:filenamesp=filenamesp.split(",")
                except:pass
            for onefile in filenamesp:
                for key in fileNameData:
                    fileName = key['Name']
                    if str(fileName) == onefile:
                        requestFileNameGetUrl = 'https://'+domain_name+'.sharepoint.com/sites/'+spsitename+'/_api/web/GetFolderByServerRelativeUrl(\'' + sharepointfolder + '\')/Files(\''+ fileName + '\')/$value'
                        savedpathtofile = os.path.join(localPath,fileName)
                        getFile = requests.get(requestFileNameGetUrl, headers=update_headers)
                        getFilestatus = str(getFile.status_code)
                        with open(savedpathtofile, 'wb+') as f:
                            for chunk in getFile.iter_content(chunk_size=1024): 
                                if chunk:
                                    f.write(chunk)
                                    f.flush()
                        print ("File Downloaded: ",fileName)
    else:     
        for key in fileNameData:
            fileName = key['Name']
            requestFileNameGetUrl = 'https://'+domain_name+'.sharepoint.com/sites/'+spsitename+'/_api/web/GetFolderByServerRelativeUrl(\'' + sharepointfolder + '\')/Files(\''+ fileName + '\')/$value'
            savedpathtofile = os.path.join(localPath,fileName)
            getFile = requests.get(requestFileNameGetUrl, headers=update_headers)
            getFilestatus = str(getFile.status_code)
            with open(savedpathtofile, 'wb+') as f:
                for chunk in getFile.iter_content(chunk_size=1024): 
                    if chunk:
                        f.write(chunk)
                        f.flush()
            print ("File Downloaded: ",fileName)

def create_folder(**kwargs):
    
    newfolder_with_entire_path=kwargs.get('SP_path',None)    
    
    url = 'https://'+domain_name+'.sharepoint.com/sites/'+spsitename+'/_api/web/folders'
    update_headers = {
            'Accept': 'application/json; odata=verbose',
            'Content-Type': 'application/json; odata=verbose',
            'Authorization': 'Bearer {}'.format(access_token)
        }
    
    if newfolder_with_entire_path[0:17]=="Shared Documents/":
        newfolder_with_entire_path=newfolder_with_entire_path.replace("Shared Documents/","")
    newfolder_with_entire_path=newfolder_with_entire_path.replace("\\","/")
    if "/" in newfolder_with_entire_path:
        newfolder_with_entire_path=newfolder_with_entire_path.split("/")
        folders_to_create=[]
        temp_sp=""
        for folder in newfolder_with_entire_path:    
            temp_sp+="/"+folder
            folders_to_create.append(temp_sp)
        for eachfolder in folders_to_create:            
            json = {"__metadata": {"type": "SP.Folder"}, "ServerRelativeUrl": "Shared Documents"  + eachfolder}
            response = requests.post(url, headers=update_headers, json=json)
        if "[501]" in str(response):
            raise RuntimeError("Unable to Create folder. "+str(response)+". Please ensure that the format is correct. Please refer the below examples.\n 1) To Create one folder      : Folder1\n 2) To Create multiple folders: Folder1\Subfolder2\Subfolder3")
            
        elif "[201]" in str(response):
            print("Folder created successfully => Shared Documents"+ eachfolder+" "+str(response))
        else:pass
    else:
        json = {"__metadata": {"type": "SP.Folder"}, "ServerRelativeUrl": "Shared Documents" + '/' + newfolder_with_entire_path}
        response = requests.post(url, headers=update_headers, json=json)
        if "[501]" in str(response):
            raise RuntimeError("Unable to Create folder. "+str(response)+". Please ensure that the format is correct. Please refer the below examples.\n 1) To Create one folder      : Folder1\n 2) To Create multiple folders: Folder1\Subfolder2\Subfolder3")            
        elif "[201]" in str(response):
            print("Folder created successfully => Shared Documents/"+ newfolder_with_entire_path+" "+str(response))
        else:pass

def upload(**kwargs):

    sharepointfolder=kwargs.get('SP_path',None)
    localPath=kwargs.get('local_path',None)
    filenamesp=kwargs.get('files_to_upload',None)
    
    create_folder(SP_path=sharepointfolder)
    sharepointfolder="Shared Documents/"+sharepointfolder
    
    if filenamesp!=None:    
        if type(filenamesp) is not list:
            try:filenamesp=filenamesp.split(",")
            except:pass

        for x in filenamesp:
            if "." not in x:
                raise RuntimeError("Extension missing for filename: "+ x +". Files with only valid extension can be uploaded.")

            for onefile in filenamesp:
                
                with open(localPath+"/"+onefile, 'rb') as f: 
                    fileBufferRead = f.read()
                requestGetURL = 'https://'+domain_name+'.sharepoint.com/sites/'+spsitename+'/_api/web/getfolderbyserverrelativeurl(\'' + sharepointfolder + '\')/Files/add(url=\'' + onefile + '\',overwrite=true)'
                update_headers = {
                            'Accept':'application/json; odata=verbose',
                            'Content-Type':'application/json; odata=verbose',
                            'Authorization': 'Bearer {}'.format(access_token)
                        } 
                data = requests.post(requestGetURL, headers=update_headers, data = fileBufferRead )
                datastatus = str(data.status_code)
                print("Uploaded "+onefile+" to SharePoint successfully!")
    else:
        all_files_to_upload=os.listdir(localPath)
        for onefile in all_files_to_upload:
                
                with open(localPath+"/"+onefile, 'rb') as f: 
                    fileBufferRead = f.read()
                requestGetURL = 'https://'+domain_name+'.sharepoint.com/sites/'+spsitename+'/_api/web/getfolderbyserverrelativeurl(\'' + sharepointfolder + '\')/Files/add(url=\'' + onefile + '\',overwrite=true)'
                update_headers = {
                            'Accept':'application/json; odata=verbose',
                            'Content-Type':'application/json; odata=verbose',
                            'Authorization': 'Bearer {}'.format(access_token)
                        } 
                data = requests.post(requestGetURL, headers=update_headers, data = fileBufferRead )
                datastatus = str(data.status_code)
                print("Uploaded "+onefile+" to SharePoint successfully!")

