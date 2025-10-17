from django.contrib import admin
from django.urls import path, include

urlpatterns = [
    path("admin/", admin.site.urls),
    path("", include("microsoft_auth.urls")),
    path("usuario/", include("Usuario.urls")),
    path("documentos/", include("Gestion_Documentos_StateMachine.urls")),  # ✅ correcto
]
