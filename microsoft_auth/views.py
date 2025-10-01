# microsoft_auth/views.py
from django.shortcuts import render, redirect
from django.conf import settings
from msal import ConfidentialClientApplication
import requests
from django.contrib.auth import login as django_login, get_user_model
from django.contrib.auth.decorators import login_required


def inicio(request):
    """Página de inicio pública."""
    if request.user.is_authenticated:
        return redirect("/usuario/")  # Redirige al dashboard si ya está logueado
    return render(request, "inicio.html")


def login(request):
    """Redirige a Microsoft para autenticación."""
    app = ConfidentialClientApplication(
        client_id=settings.MICROSOFT_CLIENT_ID,
        authority=settings.MICROSOFT_AUTHORITY,
        client_credential=settings.MICROSOFT_CLIENT_SECRET
    )
    auth_url = app.get_authorization_request_url(
        scopes=["User.Read"],
        redirect_uri=settings.MICROSOFT_REDIRECT_URI
    )
    return redirect(auth_url)


def callback(request):
    """Callback de Microsoft: obtiene token, registra y loguea usuario."""
    code = request.GET.get("code")
    if not code:
        return redirect("/")  # No se recibió código

    app = ConfidentialClientApplication(
        client_id=settings.MICROSOFT_CLIENT_ID,
        authority=settings.MICROSOFT_AUTHORITY,
        client_credential=settings.MICROSOFT_CLIENT_SECRET
    )

    result = app.acquire_token_by_authorization_code(
        code,
        scopes=["User.Read"],
        redirect_uri=settings.MICROSOFT_REDIRECT_URI
    )

    if "access_token" not in result:
        if settings.DEBUG:
            print("Error al obtener token:", result)
        return redirect("/")

    user_data = requests.get(
        "https://graph.microsoft.com/v1.0/me",
        headers={"Authorization": f"Bearer {result['access_token']}"}
    ).json()

    nombre = user_data.get("displayName")
    email = user_data.get("mail") or user_data.get("userPrincipalName")
    microsoft_id = user_data.get("id")

    # Registrar usuario en PostgreSQL interno
    registrar_usuario_postgres(nombre, email, microsoft_id)

    # Crear o recuperar usuario Django y hacer login
    User = get_user_model()
    user, created = User.objects.get_or_create(
        username=microsoft_id,
        defaults={'first_name': nombre, 'email': email}
    )
    django_login(request, user)

    return redirect("/usuario/")  # Ir al dashboard


@login_required
def logout(request):
    """Cierra sesión del usuario."""
    from django.contrib.auth import logout as django_logout
    django_logout(request)
    return redirect("/")


def registrar_usuario_postgres(nombre, email, microsoft_id):
    """Registra usuario en tabla externa de PostgreSQL."""
    import psycopg2
    from django.conf import settings

    try:
        conn = psycopg2.connect(
            host=settings.DATABASES['default']['HOST'],
            database=settings.DATABASES['default']['NAME'],
            user=settings.DATABASES['default']['USER'],
            password=settings.DATABASES['default']['PASSWORD'],
            port=settings.DATABASES['default']['PORT'],
        )
        cur = conn.cursor()
        cur.execute("SELECT * FROM usuarios_microsoft WHERE microsoft_id=%s", (microsoft_id,))
        usuario = cur.fetchone()

        if usuario is None:
            cur.execute(
                "INSERT INTO usuarios_microsoft (nombre, email, microsoft_id) VALUES (%s, %s, %s)",
                (nombre, email, microsoft_id)
            )
            conn.commit()

        cur.close()
        conn.close()
    except Exception as e:
        if settings.DEBUG:
            print("Error al registrar usuario:", e)
