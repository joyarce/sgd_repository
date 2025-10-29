from django.urls import path
from . import views
from .views import crear_proyecto

app_name = "usuario"

urlpatterns = [
    path('', views.inicio, name='inicio'), 
    path('usuarios/', views.lista_usuarios, name='lista_usuarios'), 
    path('proyectos/', views.lista_proyectos, name='lista_proyectos'), 
    path('files/', views.list_files, name='list_files'),
    path('files/upload/', views.upload_file, name='upload_file'),
    path('files/download/<path:file_id>/', views.download_file, name='download_file'),
    path('files/delete/<path:file_id>/', views.delete_file, name='delete_file'),
    path('files/new_folder/', views.new_folder, name='new_folder'),
    path("proyectos/<int:proyecto_id>/", views.detalle_proyecto, name="detalle_proyecto"),
    path('proyectos/validar_orden/', views.validar_orden_ajax, name='validar_orden'),
    path('usuario/documento/<int:documento_id>/', views.detalle_documento, name='detalle_documento'),
    path("usuario/proyecto/<int:proyecto_id>/nuevo-requerimiento/", views.nuevo_requerimiento, name="nuevo_requerimiento"),
    path('requerimiento/<int:requerimiento_id>/editar/', views.editar_requerimiento, name='editar_requerimiento'),
    path('requerimiento/<int:requerimiento_id>/eliminar/', views.eliminar_requerimiento, name='eliminar_requerimiento'),
    path('proyectos/crear/', views.crear_proyecto, name='crear_proyecto'),


    path('leer_excel/', views.leer_excel_numero_orden, name='leer_excel_numero_orden'),
    path('guardar_paso_proyecto/', views.guardar_paso_proyecto, name='guardar_paso_proyecto'),

    path('usuarios/<int:user_id>/editar/', views.editar_usuario, name='editar_usuario'),
    path('usuarios/<int:user_id>/estadisticas/', views.ver_estadisticas_usuario, name='ver_estadisticas_usuario'),


    path('editar/proyecto/<int:proyecto_id>/', views.editar_proyecto, name='editar_proyecto'),
    path('editar/contrato/<int:contrato_id>/', views.editar_contrato, name='editar_contrato'),
    path('editar/cliente/<int:cliente_id>/', views.editar_cliente, name='editar_cliente'),
    path('editar/maquina/<int:maquina_id>/', views.editar_maquina, name='editar_maquina'),


]