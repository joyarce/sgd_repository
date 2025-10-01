from django.contrib import admin
from django.urls import path, include
from Usuario.views import list_files

# Gestion_Doc>urls.py:
urlpatterns = [
    path("admin/", admin.site.urls),
    path("", include("microsoft_auth.urls")),   # sin prefijo
    path("usuario/", include("Usuario.urls")),
]