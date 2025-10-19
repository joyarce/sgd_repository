from django.contrib import admin
from django.urls import path, include
from Gestion_Documentos_StateMachine import views

urlpatterns = [
    path("admin/", admin.site.urls),
    path("", include("microsoft_auth.urls")),
    path("usuario/", include("Usuario.urls")),
    path("documentos/", include("Gestion_Documentos_StateMachine.urls")),
    path("simulador/", views.simulador_estado, name="simulador_estado"),
]