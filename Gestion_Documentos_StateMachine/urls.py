# Gestion_Documentos_StateMachine/urls.py
from django.urls import path
from . import views

from .views import (
    detalle_documento,
    validar_controles_doc_ajax,
)

app_name = "documentos"

urlpatterns = [
    path("mis-documentos/", views.lista_documentos_asignados, name="lista_documentos_asignados"),

    # Vista principal (GET para mostrar, POST para ejecutar eventos)
    path("<int:requerimiento_id>/", views.detalle_documento, name="detalle_documento"),

    # AJAX comparación controles (PREVIEW)
    path("documentos/validar-controles/<int:requerimiento_id>/",
        validar_controles_doc_ajax,
        name="validar_controles_doc_ajax"),

    # Descargar plantilla copiada al RQ
    path(
       "descargar-plantilla/<int:requerimiento_id>/",
        views.descargar_plantilla_rq,
        name="descargar_plantilla_rq"
    ),
]
