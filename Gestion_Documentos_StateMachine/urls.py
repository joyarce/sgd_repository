from django.urls import path
from . import views
from .views import detalle_documento

urlpatterns = [
    path("mis-documentos/", views.lista_documentos_asignados, name="lista_documentos_asignados"),
    path("simulador/", views.simulador_estado, name="simulador_estado"),
    path('documentos/<int:requerimiento_id>/', views.detalle_documento, name='detalle_documento'),
    path('documentos/documentos/<int:requerimiento_id>/', detalle_documento, name='detalle_documento'),

]

