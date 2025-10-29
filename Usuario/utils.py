import os
from google.cloud import storage
from django.conf import settings

from openpyxl import load_workbook
# Configurar la ruta del JSON de la Service Account
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = settings.GCP_SERVICE_ACCOUNT_JSON

# Inicializar cliente de Cloud Storage
client = storage.Client()
bucket = client.get_bucket(settings.GCP_BUCKET_NAME)






