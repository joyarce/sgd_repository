from django.urls import path
from . import views

app_name = "plantillas"

urlpatterns = [

    # LISTADO
    path("", views.lista_plantillas, name="lista_plantillas"),

    # CATEGORÍAS / TIPOS
    path("categoria/<int:categoria_id>/", views.categoria_detalle, name="detalle_categoria"),
    path("categoria/<int:categoria_id>/editar/", views.editar_categoria, name="editar_categoria"),

    path("tipos/<int:tipo_id>/", views.tipo_detalle, name="detalle_tipo"),
    path("tipos/<int:tipo_id>/editar/", views.editar_tipo_documento, name="editar_tipo_documento"),

    # TIPOS - SUBIR PLANTILLA (CUERPO)
    path("tipo/<int:tipo_id>/subir/", views.subir_plantilla, name="subir_plantilla"),
    path("subir-plantilla-tipo/<int:tipo_id>/", views.subir_plantilla, name="subir_plantilla"),

    # CREAR
    path("crear/categoria/", views.crear_categoria, name="crear_categoria"),
    path("crear/tipo/", views.crear_tipo_documento, name="crear_tipo_documento"),

    # DESCARGA / ELIMINACIÓN
    path("descargar/<path:path>/", views.descargar_gcs, name="descargar_gcs"),
    path("eliminar-plantilla/<int:tipo_id>/", views.eliminar_plantilla, name="eliminar_plantilla"),
    path("version/<int:version_id>/eliminar/", views.eliminar_version, name="eliminar_version"),
    path("ajax/detectar-controles/", views.detectar_controles_ajax, name="ajax_detectar_controles"),

]
