#Gestion_Doc\settings.py
import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

SECRET_KEY = '#-#-#'
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

ROOT_URLCONF = '#.urls'

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

WSGI_APPLICATION = '#.wsgi.application'
LOGIN_URL = '/#/'
LOGIN_REDIRECT_URL = '/#/#'


STATIC_URL = 'static/'

STATICFILES_DIRS = [
    BASE_DIR / "static",
]
# ------------------------------
# Configuración de Google Cloud Storage
# ------------------------------

# Ruta al JSON de la Service Account
GCP_SERVICE_ACCOUNT_JSON = os.path.join(BASE_DIR, 'Usuario', 'service_account.json')

# Nombre de tu bucket en GCS
GCP_BUCKET_NAME = "#"


# Credenciales del service account
GOOGLE_DRIVE_SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, 'Usuario', 'service_account.json')
GOOGLE_DRIVE_FOLDER_ID = "#"

# Microsoft OAuth
MICROSOFT_CLIENT_ID = "#"
MICROSOFT_CLIENT_SECRET = "#"
MICROSOFT_AUTHORITY = "#"
MICROSOFT_REDIRECT_URI = "#"

EMAIL_BACKEND = '#'
EMAIL_HOST = '#'
EMAIL_PORT = #
EMAIL_USE_TLS = #
EMAIL_HOST_USER = '#'
EMAIL_HOST_PASSWORD = #'
DEFAULT_FROM_EMAIL = EMAIL_HOST_USER




DATABASES = {
    'default': {
        'ENGINE': '#',
        'NAME': '#',
        'USER': '#',
        'PASSWORD': '#',
        'HOST': '#',
        'PORT': '#',
    }
}



