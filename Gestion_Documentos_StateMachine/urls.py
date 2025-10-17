from django.urls import path
from . import views

urlpatterns = [
    path("mis-documentos/", views.lista_documentos_asignados, name="lista_documentos_asignados"),
    path("simulador/", views.simulador_estado, name="simulador_estado"),
]
