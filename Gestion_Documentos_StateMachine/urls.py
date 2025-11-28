#C:\Users\jonat\Documents\gestion_docs\Gestion_Documentos_StateMachine\urls.py
from django.urls import path
from . import views

app_name = "documentos"

urlpatterns = [
    path("mis-documentos/", views.lista_documentos_asignados, name="lista_documentos_asignados"),
    path("<int:requerimiento_id>/", views.detalle_documento, name="detalle_documento"),
    path("<int:requerimiento_id>/subir/", views.subir_archivo_documento, name="subir_archivo_documento"),

    # --- RUTA CORRECTA AJAX ---
    path(
        "validar-controles/<int:requerimiento_id>/",
        views.validar_controles_doc_ajax,
        name="validar_controles_doc_ajax"
    ),

    path("descargar-plantilla/<int:requerimiento_id>/", 
     views.descargar_plantilla_rq, 
     name="descargar_plantilla_rq"),

]
