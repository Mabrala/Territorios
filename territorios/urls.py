from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('list_drive_files/', views.list_drive_files, name='list_drive_files'),
    path("list_content/<str:folder_id>/", views.list_folder_content, name="list_content"),
    path("drive/select/<str:id_folder>/", views.select_drive_folder, name="select_drive_folder"),


   #imagenes
    path("drive/file/<str:file_id>/", views.view_file, name="view_file"),

    #identidicaciones
    path('drive/auth/', views.drive_auth_init, name='drive_auth_init'),
    path('oauth2/callback', views.drive_auth_callback, name='drive_auth_callback')

]
