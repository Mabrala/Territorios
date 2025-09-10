from django.contrib import messages
from django.shortcuts import render, redirect
import os
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from .models import *
                
#Revisa las credenciales, si no las tiene devuelve None, si las tine devuelve lista
def check_creds(request):
    return request.session.get('credentials')
    
def clasify_items(items):
    for item in items:
        if item["mimeType"] == "application/vnd.google-apps.folder":
            item["type"] = "folder"
        else:
            # Clasificar archivos por tipo
            if item["mimeType"] in [
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "application/msword"
            ]:
                item["type"] = "docx"
            elif item["mimeType"] == "application/pdf":
                item["type"] = "pdf"
            elif item["mimeType"] in [
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.ms-excel"
            ]:
                item["type"] = "excel"
            elif item["mimeType"].startswith("image/"):
                item["type"] = "image"
            else:
                item["type"] = "other"
    return items

def index(request):
    creds_info = check_creds(request)
    if creds_info:
        drive_folder = Folder.objects.first()
        if drive_folder != None:
            creds = Credentials(**creds_info)
            service = build('drive', 'v3', credentials=creds)

            results = service.files().list(
                q=f"'{drive_folder.id_folder}' in parents",
                pageSize=20,
                fields="files(id, name, mimeType)",
                orderBy="folder,name",  # primero carpetas, luego archivos
             ).execute()
            
            items = results.get("files", [])
            
            items = clasify_items(items)
                    
            return render(request, 'index.html', {"files":items})
    else:
        return redirect("drive_auth_init")

#Lista el contenido de cualquier carpeta
def list_folder_content(request,folder_id):
    creds_info = check_creds(request)
    if creds_info:
        creds = Credentials(**creds_info)
        service = build('drive', 'v3', credentials=creds)
        folder = service.files().get(
            fileId=folder_id,
            fields="name"
        ).execute()
        results = service.files().list(
                q=f"'{folder_id}' in parents",
                pageSize=20,
                fields="files(id, name, mimeType)",
                orderBy="folder,name",  # primero carpetas, luego archivos
             ).execute()
        
        items = results.get("files", [])
        
        items = clasify_items(items)

        return render(request, 'folder/list.html', {"files":items,"folder_name":folder})   
    
    else:
        return redirect("drive_auth_init")

def search_in_folder(request, query):
    creds_info = check_creds(request)
    
    if not creds_info:
        return redirect("drive_auth_init")

    drive_folder = Folder.objects.first()
    folder_id = drive_folder.id_folder
    creds = Credentials(**creds_info)
    service = build('drive', 'v3', credentials=creds)

    # Obtener el nombre de la carpeta principal
    folder = service.files().get(fileId=folder_id, fields="name").execute()

    # Función recursiva para buscar en carpeta y subcarpetas
    def recursive_search(folder_id, query):
        results = []
        # Buscar archivos y carpetas en la carpeta actual
        response = service.files().list(
            q=f"'{folder_id}' in parents and name contains '{query}'",
            pageSize=100,
            fields="files(id, name, mimeType)"
        ).execute()
        results.extend(response.get("files", []))

        # Buscar subcarpetas para recursión
        subfolders = service.files().list(
            q=f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.folder'",
            fields="files(id, name)"
        ).execute().get("files", [])

        for subfolder in subfolders:
            results.extend(recursive_search(subfolder["id"], query))

        return results

    items = recursive_search(folder_id, query)
    items = clasify_items(items)  # Clasifica por tipo: carpeta, docx, pdf, imagen, excel

    return render(request, 'folder/list.html', {"files": items, "folder_name": folder})
                
#Esta vista solo se utiliza para listar las carpetas de Drive cuando 
# se va a seleccionar la carpeta de trabajo
def list_drive_files(request):
    creds_info = check_creds(request)
    if creds_info:
        creds = Credentials(**creds_info)
        service = build('drive', 'v3', credentials=creds)

        # Pedimos nombre, id y mimeType para saber si es archivo o carpeta
        results = service.files().list(
            q = "mimeType='application/vnd.google-apps.folder'",
            pageSize=20,
            fields="files(id, name, mimeType)"
        ).execute()

        items = results.get('files', [])
        
        return render(request, "choose_drive_file/choose_drive_file.html", {"files": items})
    else:
        return redirect("drive_auth_init")

#Selecciona la carpeta sobre la que se trabajara
def select_drive_folder(request, id_folder):
    creds_info = check_creds(request)
    if creds_info:
        creds = Credentials(**creds_info)
        service = build('drive', 'v3', credentials=creds)
        folder = service.files().get(
            fileId=id_folder,
            fields="id,name,mimeType"
        ).execute()
        if folder["mimeType"] != "application/vnd.google-apps.folder":
            messages.error(request, "El ID no corresponde a una carpeta.")
            return redirect("list_drive_files")
                
        drive_folder = Folder.objects.update_or_create(
            id=1,
            defaults={
                "id_folder": folder["id"],
                "name": folder["name"]
                }
        )
        messages.success(request, f"Carpeta '{folder['name']}' seleccionada como activa.")
        return redirect("index")

    else:
        return redirect("drive_auth_init")

SCOPES = ['https://www.googleapis.com/auth/drive']

from django.shortcuts import redirect
from google_auth_oauthlib.flow import Flow

def drive_auth_init(request):
    flow = Flow.from_client_config(
        {
            "web": {
                "client_id": os.environ["DRIVE_CLIENT"],
                "client_secret": os.environ["DRIVE_SECRET"],
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": [os.environ["DRIVE_REDIRECT_URI"]],
            }
        },
        scopes=SCOPES,
    )
    flow.redirect_uri = os.environ["DRIVE_REDIRECT_URI"]

    authorization_url, state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true"
    )
    request.session['oauth_state'] = state
    return redirect(authorization_url)

from google.oauth2.credentials import Credentials

def drive_auth_callback(request):
    # Recupera el estado guardado al iniciar OAuth
    state = request.session.get('oauth_state')
    
    flow = Flow.from_client_config(
        {
            "web": {
                "client_id": os.environ["DRIVE_CLIENT"],
                "client_secret": os.environ["DRIVE_SECRET"],
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": [os.environ["DRIVE_REDIRECT_URI"]],
            }
        },
        scopes=SCOPES,
        state=state
    )
    flow.redirect_uri = os.environ["DRIVE_REDIRECT_URI"]

    # Construye la URL completa con parámetros que envió Google
    authorization_response = request.build_absolute_uri()
    flow.fetch_token(authorization_response=authorization_response)

    # Credenciales que permiten acceder a Google Drive
    credentials = flow.credentials

    # Guardarlas en la sesión de Django
    request.session['credentials'] = {
        'token': credentials.token,
        'refresh_token': credentials.refresh_token,
        'token_uri': credentials.token_uri,
        'client_id': credentials.client_id,
        'client_secret': credentials.client_secret,
        'scopes': credentials.scopes
    }

    # Redirigir a la vista que lista archivos o a otra página
    return redirect('index')

