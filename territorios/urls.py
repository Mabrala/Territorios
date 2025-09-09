from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('listar/', views.list_drive_files, name='listar'),
    #identidicaciones
    path('drive/auth/', views.drive_auth_init, name='drive_auth_init'),
    path('oauth2/callback', views.drive_auth_callback, name='drive_auth_callback')

]
