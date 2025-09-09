from django.shortcuts import render, redirect
import os
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

def index(request):
    return render(request, 'index.html')

def check_creds(request):
    return request.session.get('credentials')
    

def list_drive_files(request):
    creds_info = check_creds(request)
    if creds_info:
        creds = Credentials(**creds_info)
        service = build('drive', 'v3', credentials=creds)

        results = service.files().list(pageSize=10).execute()
        items = results.get('files', [])

        return render(request, 'index.html', {'files': items})
    else:
        return redirect('drive_auth_init')

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
    return redirect('listar')

