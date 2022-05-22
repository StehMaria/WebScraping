
from googleapiclient.http import MediaFileUpload
from Google import Create_Service



def drive():
    CLIENT_SERVICE_FILE = 'client_secret.json'
    
    API_NAME = 'drive'
    
    API_VERSION = 'v3'
    
    SCOPES = ['https://www.googleapis.com/auth/drive']
    
    
    
    service = Create_Service(CLIENT_SERVICE_FILE,API_NAME,API_VERSION,SCOPES)
    
    
    
    file_name = 'Vulnerabilidades.xlsx'
    
    mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    
    
    
    
    file_metadados = {
        'title' : "Teste",
    
    }
    
    media = MediaFileUpload(file_name,mimetype=mime_type)
    service.files().create(
        body=file_metadados,
        media_body=media,
        fields='id').execute()
    
    request = service.files().get(fileId='id') 
    print(request)
                