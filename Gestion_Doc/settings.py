import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

SECRET_KEY = 'tu-secret-key'
DEBUG = True
ALLOWED_HOSTS = []

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'Usuario',  
    'microsoft_auth', 
    "Gestion_Documentos_StateMachine",
    "plantillas_documentos_tecnicos"
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

ROOT_URLCONF = 'Gestion_Doc.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [os.path.join(BASE_DIR, 'templates')], 
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    }
]

WSGI_APPLICATION = 'Gestion_Doc.wsgi.application'
LOGIN_URL = '/login/'
LOGIN_REDIRECT_URL = '/usuario/inicio'
DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': BASE_DIR / 'db.sqlite3',
    }
}

STATIC_URL = 'static/'

# ------------------------------
# Configuración de Google Cloud Storage
# ------------------------------

# Ruta al JSON de la Service Account
GCP_SERVICE_ACCOUNT_JSON = os.path.join(BASE_DIR, 'Usuario', 'service_account.json')

# Nombre de tu bucket en GCS
GCP_BUCKET_NAME = "sgdmtso_jova"


# Credenciales del service account
GOOGLE_DRIVE_SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, 'Usuario', 'service_account.json')
GOOGLE_DRIVE_FOLDER_ID = "1pu5aVJtVbsaGqKrTb1avUoO4m2G83ipM"

# Microsoft OAuth
MICROSOFT_CLIENT_ID = "c4601061-a49f-478a-ad5c-654e3cca3868"
MICROSOFT_CLIENT_SECRET = "fgy8Q~KXD3foDHVQunLEKODHYG4jNB3QxKNCKaW4"
MICROSOFT_AUTHORITY = "https://login.microsoftonline.com/common"
MICROSOFT_REDIRECT_URI = "http://localhost:8000/callback/"

EMAIL_BACKEND = 'django.core.mail.backends.smtp.EmailBackend'
EMAIL_HOST = 'smtp.office365.com'
EMAIL_PORT = 587
EMAIL_USE_TLS = True
EMAIL_HOST_USER = 'alertas_sgd@outlook.com'
EMAIL_HOST_PASSWORD = 'tu_contraseña_o_app_password'
DEFAULT_FROM_EMAIL = EMAIL_HOST_USER




DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.postgresql',
        'NAME': 'django_db',
        'USER': 'django_user',
        'PASSWORD': 'django_pass',
        'HOST': '127.0.0.1',
        'PORT': '5432',
    }
}



