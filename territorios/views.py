from django.contrib import messages
from django.shortcuts import render, redirect
import os
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from .models import *
from google_auth_oauthlib.flow import Flow
from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger

#para imagenes
from django.http import HttpResponse, Http404
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from googleapiclient.errors import HttpError
import io
import mimetypes

#escribit docx y subirlo
from docx import Document
import os
                
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

def paginate_items(request, items, per_page=10):
    # Ordena primero carpetas, luego archivos por nombre
    items_sorted = sorted(
        items,
        key=lambda x: (x.get("type") != "folder", x.get("name", "").lower())
    )
    page = request.GET.get('page', 1)
    paginator = Paginator(items_sorted, per_page)
    try:
        files = paginator.page(page)
    except PageNotAnInteger:
        files = paginator.page(1)
    except EmptyPage:
        files = paginator.page(paginator.num_pages)
    return files

def index(request):
    creds_info = check_creds(request)
    if creds_info:
        drive_folder = Folder.objects.first()
        if drive_folder != None:
            creds = Credentials(**creds_info)
            service = build('drive', 'v3', credentials=creds)

            results = service.files().list(
                q=f"'{drive_folder.id_folder}' in parents",
                pageSize=100,
                fields="files(id, name, mimeType)",
                orderBy="folder,name",
             ).execute()
            
            items = results.get("files", [])
            items = clasify_items(items)

            files = paginate_items(request, items)
            return render(request, 'index.html', {"files": files})
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
                fields="files(id, name, mimeType)",
                orderBy="folder,name",
             ).execute()
        
        items = results.get("files", [])
        items = clasify_items(items)

        files = paginate_items(request, items)
        return render(request, 'folder/list.html', {"files": files, "folder_name": folder})   
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

    folder = service.files().get(fileId=folder_id, fields="name").execute()

    def recursive_search(folder_id, query):
        results = []
        response = service.files().list(
            q=f"'{folder_id}' in parents and name contains '{query}'",
            pageSize=100,
            fields="files(id, name, mimeType)"
        ).execute()
        results.extend(response.get("files", []))
        subfolders = service.files().list(
            q=f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.folder'",
            fields="files(id, name)"
        ).execute().get("files", [])
        for subfolder in subfolders:
            results.extend(recursive_search(subfolder["id"], query))
        return results

    items = recursive_search(folder_id, query)
    items = clasify_items(items)

    files = paginate_items(request, items)
    return render(request, 'folder/list.html', {"files": files, "folder_name": folder})


                
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

#Asigna el territorio en el docx
def update_docx(request, file_id, territory_code, col_idx, new_value):
    creds_info = check_creds(request)
    if not creds_info:
        return redirect("drive_auth_init")

    creds = Credentials(**creds_info)
    service = build("drive", "v3", credentials=creds)

    # 1. Descargar archivo desde Drive a memoria
    request_file = service.files().get_media(fileId=file_id)
    downloaded_file = io.BytesIO()
    downloader = MediaIoBaseDownload(downloaded_file, request_file)

    done = False
    while not done:
        status, done = downloader.next_chunk()
    downloaded_file.seek(0)

    # 2. Abrir con python-docx desde memoria
    doc = Document(downloaded_file)

    # 3. Buscar la fila y modificar la celda
    for table in doc.tables:
        for row in table.rows:
            if row.cells[0].text.strip() == territory_code:
                row.cells[col_idx].text = str(new_value)
                break

    # 4. Guardar cambios en un nuevo buffer
    updated_file = io.BytesIO()
    doc.save(updated_file)
    updated_file.seek(0)

    # 5. Subir archivo actualizado a Drive (sobrescribe el original)
    media_body = MediaIoBaseUpload(
        updated_file,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    updated = service.files().update(
        fileId=file_id,
        media_body=media_body
    ).execute()

    return f"Archivo actualizado en Drive: {updated['name']}"

#visualizar imagenes
def view_file(request, file_id):
    creds_info = check_creds(request)
    if not creds_info:
        return redirect("drive_auth_init")

    creds = Credentials(**creds_info)
    service = build("drive", "v3", credentials=creds)

    try:
        # Obtener metadata del archivo
        file = service.files().get(
            fileId=file_id,
            fields="id, name, mimeType"
        ).execute()

        mime_type = file.get("mimeType", "")
        if not (mime_type.startswith("image/") or mime_type == "application/pdf"):
            raise Http404("El archivo no es una imagen ni un PDF")

        # Descargar el archivo
        request_file = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request_file)

        done = False
        while not done:
            status, done = downloader.next_chunk()

        fh.seek(0)  # reiniciar puntero

        # Determinar content type
        content_type = mime_type or mimetypes.guess_type(file["name"])[0] or "application/octet-stream"

        return HttpResponse(fh.read(), content_type=content_type)

    except Exception as e:
        raise Http404(f"No se pudo mostrar el archivo: {str(e)}")



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
DRIVE_CLIENT = os.environ.get("DRIVE_CLIENT")
DRIVE_SECRET = os.environ.get("DRIVE_SECRET")
DRIVE_REDIRECT_URI = os.environ.get("DRIVE_REDIRECT_URI")

def drive_auth_init(request):
    client = DRIVE_CLIENT
    secret = DRIVE_SECRET
    redirect_uri = DRIVE_REDIRECT_URI
    flow = Flow.from_client_config(
        {
            "web": {
                "client_id": client,
                "client_secret": secret,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": redirect_uri,
            }
        },
        scopes=SCOPES,
    )
    flow.redirect_uri = redirect_uri

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
    client = DRIVE_CLIENT
    secret = DRIVE_SECRET
    redirect_uri = DRIVE_REDIRECT_URI    
    flow = Flow.from_client_config(
        {
            "web": {
                "client_id": client,
                "client_secret": secret,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": redirect_uri,
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

