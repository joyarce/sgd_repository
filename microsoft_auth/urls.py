from django.urls import path
from . import views

app_name = "microsoft_auth"   # ✅ Esto registra el namespace

urlpatterns = [
    path("", views.inicio, name="inicio"),
    path("login/", views.login, name="login"),
    path("callback/", views.callback, name="callback"),
    path("logout/", views.logout, name="logout"),
]
