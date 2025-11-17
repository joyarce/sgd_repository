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

    # TIPOS - SUBIR PLANTILLA
    path("tipo/<int:tipo_id>/subir/", views.subir_plantilla, name="subir_plantilla"),
    path("subir-plantilla-tipo/<int:tipo_id>/",
         views.subir_plantilla_tipo_doc,
         name="subir_plantilla_tipo_doc"),

    # CREAR
    path("crear/categoria/", views.crear_categoria, name="crear_categoria"),
    path("crear/tipo/", views.crear_tipo_documento, name="crear_tipo_documento"),

    # PORTADA WORD
    path("portada/word/", views.portada_word_detalle, name="portada_word_detalle"),
    path("portada/word/subir/", views.subir_portada_word, name="subir_portada_word"),

    # PORTADA EXCEL
    path("portada/excel/", views.portada_excel_detalle, name="portada_excel_detalle"),
    path("portada/excel/subir/", views.subir_portada_excel, name="subir_portada_excel"),

    # DESCARGA / ELIMINACIÓN
    path("descargar/<path:path>/", views.descargar_gcs, name="descargar_gcs"),
    path("eliminar-plantilla/<int:tipo_id>/", views.eliminar_plantilla, name="eliminar_plantilla"),
    path("version/<int:version_id>/eliminar/", views.eliminar_version, name="eliminar_version"),
]
