# C:\Users\jonat\Documents\gestion_docs\plantillas_documentos_tecnicos\utils_documentos.py

from datetime import timedelta
import urllib.parse
from django.db import connection
import json
def obtener_plantilla_usada(requerimiento_id):
    """
    Devuelve la versión de plantilla realmente usada por el RQ.
    Busca por coincidencia en ruta_gcs LIKE %RQ-{id}/%.
    Si no hay copia, retorna fallback versión_actual.
    """

    with connection.cursor() as cursor:
        # 1) Buscar la plantilla copiada al RQ
        cursor.execute("""
            SELECT 
                DG.plantilla_documento_tecnico_version_id,
                V.gcs_path
            FROM documentos_generados DG
            LEFT JOIN plantilla_tipo_doc_versiones V
                ON V.id = DG.plantilla_documento_tecnico_version_id
            WHERE DG.ruta_gcs LIKE %s
            ORDER BY DG.fecha_generacion ASC, DG.id ASC
            LIMIT 1;
        """, [f"%RQ-{requerimiento_id}/%"])

        row = cursor.fetchone()

    if row and row[0]:
        return {
            "version_id": row[0],
            "ruta": row[1],
            "tipo": "copiada"
        }

    # 2) Fallback: versión_actual del tipo_documento
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT V.id, V.gcs_path
            FROM requerimiento_documento_tecnico R
            JOIN tipo_documentos_tecnicos TDT
                ON R.tipo_documento_id = TDT.id
            JOIN plantilla_tipo_doc P
                ON P.tipo_documento_id = TDT.id
            JOIN plantilla_tipo_doc_versiones V
                ON V.id = P.version_actual_id
            WHERE R.id = %s
            ORDER BY V.creado_en DESC, V.id DESC
            LIMIT 1;
        """, [requerimiento_id])

        row = cursor.fetchone()

    if row:
        return {
            "version_id": row[0],
            "ruta": row[1],
            "tipo": "fallback"
        }

    return None





def obtener_blob_plantilla_usada(cursor, bucket, requerimiento_id):
    """
    Obtiene el blob y signed_url de la plantilla usada en el RQ.
    """

    cursor.execute("""
        SELECT ruta_gcs, plantilla_documento_tecnico_version_id
        FROM documentos_generados
        WHERE requerimiento_id = %s
        ORDER BY fecha_generacion ASC
        LIMIT 1;
    """, [requerimiento_id])

    row = cursor.fetchone()
    if not row:
        return None

    ruta_gcs, version_id = row
    blob = bucket.blob(ruta_gcs)

    url = blob.generate_signed_url(
        version="v4",
        expiration=timedelta(hours=24),
        method="GET"
    )

    return {
        "ruta": ruta_gcs,
        "version_id": version_id,
        "signed_url": url
    }

def obtener_estructura_plantilla_usada(requerimiento_id):
    """
    Devuelve (estructura_json_dict, version_id) de la plantilla REAL
    usada cuando se inicializó el RQ.
    """

    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT plantilla_documento_tecnico_version_id
            FROM documentos_generados
            WHERE ruta_gcs LIKE %s
            ORDER BY fecha_generacion DESC, id DESC
            LIMIT 1
        """, [f"%RQ-{requerimiento_id}/%"])

        row = cursor.fetchone()

    if not row or not row[0]:
        raise ValueError("No se encontró una plantilla usada para este requerimiento.")

    version_id = row[0]

    # Obtener estructura JSON asociada a esa versión
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT estructura_json
            FROM plantilla_estructura_version
            WHERE version_id = %s
            LIMIT 1
        """, [version_id])

        row = cursor.fetchone()

    if not row:
        raise ValueError(f"No existe estructura registrada para version_id={version_id}")

    estructura_json = row[0]

    if isinstance(estructura_json, str):
        estructura_json = json.loads(estructura_json)

    return estructura_json, version_id





def extract_blob_name_from_signed_url(signed_url: str) -> str:
    """
    Extrae la ruta real del blob GCS desde una SignedURL generada por Google.

    Ejemplo de SignedURL:
    https://storage.googleapis.com/bucket-name/Carpeta1/Doc1.docx?X-Goog-...

    Retorna:
        Carpeta1/Doc1.docx
    """
    if not signed_url:
        return None

    try:
        # Quitar parámetros (?X-Goog-...)
        clean_url = signed_url.split("?", 1)[0]

        # Particionar por el dominio
        if "storage.googleapis.com" in clean_url:
            # https://storage.googleapis.com/bucket/...
            parts = clean_url.split("/", 4)
            if len(parts) >= 5:
                return parts[4]  # Carpeta/.../archivo.docx

        # Firmas con domain personalizado
        if "https://" in clean_url:
            # Última parte del dominio hacia adelante
            after_domain = clean_url.split(".com/", 1)[-1]
            return after_domain

        return clean_url

    except Exception:
        return None

# ============================================================
# EXTRACCIÓN DE BLOBNAME DESDE gcs_path (PLANTILLAS)
# ============================================================

def extract_blob_name_from_gcs_path(path: str) -> str:
    """
    Convierte un gcs_path almacenado tal cual en la BD
    a un blob_name usable por GCS.

    Ejemplo:
        'Plantillas/TipoA/v1.docx'
    """
    if not path:
        return ""
    return urllib.parse.unquote(path)


# ============================================================
# INSERTAR EN documentos_generados
# ============================================================

def insertar_documento_generado(
    cursor,
    proyecto_id: int,
    tipo_documento_id: int,
    ruta_gcs: str,
    plantilla_version_id: int,
    formato_id: int,
):
    """
    Inserta un documento generado en la tabla documentos_generados.
    """
    cursor.execute(
        """
        INSERT INTO documentos_generados (
            proyecto_id,
            tipo_documento_id,
            ruta_gcs,
            fecha_generacion,
            plantilla_documento_tecnico_version_id,
            formato_id
        )
        VALUES (%s, %s, %s, NOW(), %s, %s)
        RETURNING id;
        """,
        [
            proyecto_id,
            tipo_documento_id,
            ruta_gcs,
            plantilla_version_id,
            formato_id,
        ],
    )
    return cursor.fetchone()[0]


# ============================================================
# CREAR VERSION INICIAL v0.0.1 DEL REQUERIMIENTO
# ============================================================

def inicializar_version_inicial(
    cursor,
    bucket,
    requerimiento_id: int,
    ruta_plantilla: str,
    codigo_documento: str,
):
    """
    Crea la versión inicial v0.0.1 para un RQ dado,
    copiando la última plantilla disponible del tipo de documento.

    - Obtiene proyecto, tipo_documento, formato y plantilla actual
    - Copia la plantilla a:
        {ruta_plantilla}/"v0.0.1 - {codigo_documento}.docx"
    - Inserta en version_documento_tecnico
    - Inserta en documentos_generados

    Devuelve:
        version_id (id en version_documento_tecnico)
    """

    # -----------------------------------------------------------
    # 1) Obtener datos del requerimiento + tipo_doc + plantilla
    # -----------------------------------------------------------
    cursor.execute(
        """
        SELECT
            R.proyecto_id,
            R.tipo_documento_id,
            T.formato_id,
            P.version_actual_id,
            V.gcs_path
        FROM requerimiento_documento_tecnico R
        JOIN tipo_documentos_tecnicos T
          ON R.tipo_documento_id = T.id
        JOIN plantilla_tipo_doc P
          ON P.tipo_documento_id = T.id
        JOIN plantilla_tipo_doc_versiones V
          ON V.id = P.version_actual_id
        WHERE R.id = %s
        LIMIT 1;
        """,
        [requerimiento_id],
    )
    row = cursor.fetchone()
    if not row:
        raise Exception(
            "⚠ No se pudo obtener datos de RQ + tipo_documento + plantilla."
        )

    proyecto_id, tipo_documento_id, formato_id, plantilla_version_id, gcs_path = row

    if not plantilla_version_id or not gcs_path:
        raise Exception("⚠ No hay versión_actual de plantilla configurada.")

    # -----------------------------------------------------------
    # 2) Descargar plantilla base
    # -----------------------------------------------------------
    blob_origen = bucket.blob(extract_blob_name_from_gcs_path(gcs_path))
    contenido = blob_origen.download_as_bytes()

    # -----------------------------------------------------------
    # 3) Subir a carpeta del requerimiento (Plantilla/)
    # -----------------------------------------------------------
    if not ruta_plantilla.endswith("/"):
        ruta_plantilla += "/"

    nombre_final = f"v0.0.1 - {codigo_documento}.docx"
    blob_destino_name = f"{ruta_plantilla}{nombre_final}"

    blob_destino = bucket.blob(blob_destino_name)
    blob_destino.upload_from_string(
        contenido,
        content_type=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
    )

    # -----------------------------------------------------------
    # 4) signed_url temporal para vista previa
    # -----------------------------------------------------------
    signed_url = blob_destino.generate_signed_url(
        version="v4",
        expiration=timedelta(hours=24),
        method="GET",
    )

    # -----------------------------------------------------------
    # 5) Estado inicial
    # -----------------------------------------------------------
    cursor.execute(
        """
        SELECT id
        FROM estado_documento
        WHERE nombre ILIKE 'Pendiente de Inicio'
        LIMIT 1;
        """
    )
    row = cursor.fetchone()
    estado_id = row[0] if row else None

    # -----------------------------------------------------------
    # 6) Insertar versión v0.0.1
    # -----------------------------------------------------------
    cursor.execute(
        """
        INSERT INTO version_documento_tecnico
            (requerimiento_documento_id,
             version,
             estado_id,
             fecha,
             comentario,
             usuario_id,
             signed_url)
        VALUES (%s, %s, %s, NOW(), %s, NULL, %s)
        RETURNING id;
        """,
        [
            requerimiento_id,
            "v0.0.1",
            estado_id,
            "Versión inicial generada automáticamente.",
            signed_url,
        ],
    )
    version_id = cursor.fetchone()[0]

    # -----------------------------------------------------------
    # 7) Registrar en documentos_generados
    # -----------------------------------------------------------
    insertar_documento_generado(
        cursor=cursor,
        proyecto_id=proyecto_id,
        tipo_documento_id=tipo_documento_id,
        ruta_gcs=blob_destino_name,
        plantilla_version_id=plantilla_version_id,
        formato_id=formato_id,
    )

    return version_id
