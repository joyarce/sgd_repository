from django.urls import path
from . import views

app_name = "usuario"

urlpatterns = [
    path('', views.inicio, name='inicio'),  # <-- aquí la raíz de usuario/
    path('usuarios/', views.lista_usuarios, name='lista_usuarios'),  # ← Asegúrate que exista
    path('proyectos/', views.lista_proyectos, name='lista_proyectos'),  # ← Asegúrate que exista
    path('files/', views.list_files, name='list_files'),
    path('files/upload/', views.upload_file, name='upload_file'),
    path('files/download/<path:file_id>/', views.download_file, name='download_file'),
    path('files/delete/<path:file_id>/', views.delete_file, name='delete_file'),
    path('files/new_folder/', views.new_folder, name='new_folder'),
    path("proyectos/", views.lista_proyectos, name="lista_proyectos"),
    path("proyectos/<int:proyecto_id>/", views.detalle_proyecto, name="detalle_proyecto"),
]

