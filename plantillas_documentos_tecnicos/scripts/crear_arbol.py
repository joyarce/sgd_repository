# C:\Users\jonat\Documents\gestion_docs\plantillas_documentos_tecnicos\scripts\crear_arbol.py

import os
import django
import sys
from unidecode import unidecode
from google.cloud import storage
from django.db import connection

# ---------------------------------------------------------------------
# CONFIGURACIÓN DJANGO
# ---------------------------------------------------------------------
RUTA_PROYECTO = r"C:\Users\jonat\Documents\gestion_docs"
sys.path.append(RUTA_PROYECTO)
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'gestion_docs.settings')
django.setup()

# ---------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------
def clean(texto):
    return unidecode(texto).replace(" ", "_").replace("/", "_")

# ---------------------------------------------------------------------
# CREAR CARPETA EN GCS
# ---------------------------------------------------------------------
def mkdir_gcs(bucket, path):
    if not path.endswith("/"):
        path += "/"
    blob = bucket.blob(path)
    blob.upload_from_string("")  # carpeta vacía
    return path

# ---------------------------------------------------------------------
# PROCESO PRINCIPAL
# ---------------------------------------------------------------------
def crear_arbol_bucket():
    print("\n=== CREANDO ÁRBOL DE PLANTILLAS EN GCS ===")

    # Cliente GCS
    client = storage.Client()
    bucket = client.bucket("sgdmtso_jova")   # <-- tu bucket real

    # Raíz principal
    base_root = "Plantillas/"
    utilidad_root = base_root + "Utilidad/"
    portada_root = utilidad_root + "Portada/"
    docs_root = base_root + "Documentos_Tecnicos/"

    # Crear estructura fija
    mkdir_gcs(bucket, base_root)
    mkdir_gcs(bucket, utilidad_root)
    mkdir_gcs(bucket, portada_root)
    mkdir_gcs(bucket, docs_root)

    print("✔ Carpetas raíz creadas.")

    # -----------------------------------------------------------------
    # 1. Obtener categorías
    # -----------------------------------------------------------------
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id, nombre, abreviatura
            FROM categoria_documentos_tecnicos
            ORDER BY id
        """)
        categorias = cursor.fetchall()

    for cat_id, cat_nombre, cat_abrev in categorias:

        cat_clean = clean(cat_nombre)
        cat_path = f"{docs_root}{cat_clean}/"
        mkdir_gcs(bucket, cat_path)

        print(f"✔ Categoría creada: {cat_path}")

        # -------------------------------------------------------------
        # 2. Obtener tipos asociados
        # -------------------------------------------------------------
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT id, nombre, abreviatura
                FROM tipo_documentos_tecnicos
                WHERE categoria_id = %s
                ORDER BY id
            """, [cat_id])
            tipos = cursor.fetchall()

        for tipo_id, tipo_nombre, tipo_abrev in tipos:
            tipo_clean = clean(tipo_nombre)
            tipo_path = f"{cat_path}{tipo_clean}/"

            mkdir_gcs(bucket, tipo_path)

            print(f"   ✔ Tipo creado: {tipo_path}")

            # ---------------------------------------------------------
            # 3. Registrar path en la tabla plantillas_documentos_tecnicos
            # (solo si no existe)
            # ---------------------------------------------------------
            with connection.cursor() as cursor:
                cursor.execute("""
                    SELECT COUNT(*) 
                    FROM plantillas_documentos_tecnicos
                    WHERE tipo_documento_id = %s AND gcs_path = %s
                """, [tipo_id, tipo_path])
                existe = cursor.fetchone()[0]

            if existe == 0:
                with connection.cursor() as cursor:
                    cursor.execute("""
                        INSERT INTO plantillas_documentos_tecnicos
                        (tipo_documento_id, gcs_path, version)
                        VALUES (%s, %s, 1)
                    """, [tipo_id, tipo_path])
                print(f"      ✔ Path registrado en BD: {tipo_path}")
            else:
                print(f"      (i) Ya registrado en BD")

    print("\n=== PROCESO COMPLETADO ===\n")


# ---------------------------------------------------------------------
# EJECUCIÓN DIRECTA
# ---------------------------------------------------------------------
if __name__ == "__main__":
    crear_arbol_bucket()
