# C:\Users\jonat\Documents\gestion_docs\plantillas_documentos_tecnicos\views.py

# IMPORTS
# ============================
from django.http import HttpResponse, JsonResponse, Http404
from django.contrib.auth.decorators import login_required
from django.shortcuts import render, redirect, get_object_or_404
from django.db import connection
from django.contrib import messages
from django.conf import settings
import urllib.parse
import re
import os
import tempfile
from zipfile import ZipFile, BadZipFile
from datetime import timedelta
import json

from google.cloud import storage
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table
from openpyxl.utils import range_boundaries
from unidecode import unidecode

from lxml import etree


# =============================================================================
# UTILIDADES
# =============================================================================

def dictfetchall(cursor):
    columns = [col[0] for col in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def dictfetchone(cursor):
    row = cursor.fetchone()
    if row is None:
        return None
    columns = [col[0] for col in cursor.description]
    return dict(zip(columns, row))


def clean(texto):
    """Limpia texto para usar en rutas del bucket."""
    return unidecode(texto).replace(" ", "_").replace("/", "_")


def siguiente_version(version_actual):
    """
    Recibe una cadena tipo '1.0' y devuelve la siguiente versión: '1.1'.
    Si no hay versión o es inválida -> '1.0'.
    """
    if not version_actual:
        return "1.0"
    try:
        partes = str(version_actual).split(".")
        if len(partes) == 2:
            major = int(partes[0])
            minor = int(partes[1])
            return f"{major}.{minor + 1}"
        # fallback sencillo
        v = float(version_actual)
        return f"{v + 0.1:.1f}"
    except Exception:
        return "1.0"


# =============================================================================
# LISTADO GENERAL
# =============================================================================

def lista_plantillas(request):
    categorias = []
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id, nombre, descripcion, abreviatura
            FROM categoria_documentos_tecnicos
            ORDER BY id ASC
        """)
        categorias_data = cursor.fetchall()

        for cat_id, cat_nombre, cat_desc, cat_abrev in categorias_data:
            cursor.execute("""
                SELECT id, nombre, abreviatura, descripcion
                FROM tipo_documentos_tecnicos
                WHERE categoria_id = %s
                ORDER BY id ASC
            """, [cat_id])
            tipos = cursor.fetchall()

            categorias.append({
                "id": cat_id,
                "nombre": cat_nombre,
                "abreviatura": cat_abrev,
                "descripcion": cat_desc,
                "tipos": [{
                    "id": t[0],
                    "nombre": t[1],
                    "abreviatura": t[2],
                    "descripcion": t[3],
                } for t in tipos]
            })

    return render(request, "lista_plantillas.html", {"categorias": categorias})


# =============================================================================
# DETALLES DE CATEGORÍA Y TIPO
# =============================================================================

def categoria_detalle(request, categoria_id):
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id, nombre, descripcion, abreviatura
            FROM categoria_documentos_tecnicos
            WHERE id = %s
        """, [categoria_id])
        categoria = dictfetchone(cursor)

        cursor.execute("""
            SELECT id, nombre, descripcion, abreviatura
            FROM tipo_documentos_tecnicos
            WHERE categoria_id = %s
            ORDER BY nombre
        """, [categoria_id])
        tipos = dictfetchall(cursor)

    return render(request, "categoria_detalle.html", {
        "categoria": categoria,
        "tipos": tipos,
    })


@login_required
def editar_categoria(request, categoria_id):

    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id, nombre, descripcion, abreviatura
            FROM categoria_documentos_tecnicos
            WHERE id = %s
        """, [categoria_id])
        categoria = dictfetchone(cursor)

    if not categoria:
        messages.error(request, "Categoría no encontrada.")
        return redirect("plantillas:lista_plantillas")

    # Valores originales (para paths)
    nombre_original = categoria["nombre"]
    nombre_original_clean = clean(nombre_original)

    if request.method == "POST":
        nombre = request.POST.get("nombre", "").strip()
        descripcion = request.POST.get("descripcion", "").strip()
        abreviatura = request.POST.get("abreviatura", "").strip().upper()

        if not nombre:
            messages.error(request, "El nombre de la categoría es obligatorio.")
            return redirect(request.path)

        nombre_nuevo_clean = clean(nombre)

        # Si cambió el nombre → mover carpetas en GCS + actualizar rutas en BD
        if nombre_nuevo_clean != nombre_original_clean:
            old_prefix = f"Plantillas/Documentos_Tecnicos/{nombre_original_clean}/"
            new_prefix = f"Plantillas/Documentos_Tecnicos/{nombre_nuevo_clean}/"

            try:
                mover_carpeta_gcs(old_prefix, new_prefix)

                # Actualizar rutas en BD (nueva tabla de versiones)
                with connection.cursor() as cursor:
                    cursor.execute("""
                        UPDATE plantilla_tipo_doc_versiones
                        SET gcs_path = REPLACE(gcs_path, %s, %s)
                        WHERE gcs_path LIKE %s
                    """, [old_prefix, new_prefix, old_prefix + "%"])
            except Exception as e:
                messages.error(request, f"Error al renombrar en GCS: {e}")
                return redirect("plantillas:detalle_categoria", categoria_id=categoria_id)

        # Actualizar la categoría en BD
        with connection.cursor() as cursor:
            cursor.execute("""
                UPDATE categoria_documentos_tecnicos
                SET nombre = %s, descripcion = %s, abreviatura = %s
                WHERE id = %s
            """, [nombre, descripcion, abreviatura, categoria_id])

        messages.success(request, "✔ Categoría actualizada.")
        return redirect("plantillas:detalle_categoria", categoria_id=categoria_id)

    return render(request, "editar_categoria.html", {
        "categoria": categoria
    })


def tipo_detalle(request, tipo_id):

    # ============================
    # 1) Datos del tipo de documento
    # ============================
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT 
                t.id, 
                t.categoria_id,
                t.nombre, 
                t.descripcion, 
                t.abreviatura,
                t.formato_id,
                c.nombre AS categoria_nombre
            FROM tipo_documentos_tecnicos t
            JOIN categoria_documentos_tecnicos c ON c.id = t.categoria_id
            WHERE t.id = %s
        """, [tipo_id])
        tipo = dictfetchone(cursor)

        if not tipo:
            return render(request, "404.html", status=404)

        # Última versión (actual) desde las tablas nuevas
        cursor.execute("""
            SELECT 
                v.id,
                v.plantilla_id,
                v.gcs_path,
                v.version,
                v.creado_en
            FROM plantilla_tipo_doc p
            JOIN plantilla_tipo_doc_versiones v
                ON v.id = p.version_actual_id
            WHERE p.tipo_documento_id = %s
            LIMIT 1
        """, [tipo_id])
        plantilla = dictfetchone(cursor)

        # Todas las versiones de esa plantilla (si existe)
        if plantilla:
            cursor.execute("""
                SELECT 
                    id,
                    gcs_path,
                    version,
                    creado_en
                FROM plantilla_tipo_doc_versiones
                WHERE plantilla_id = %s
                ORDER BY creado_en DESC, id DESC
            """, [plantilla["plantilla_id"]])
            versiones = dictfetchall(cursor)
        else:
            versiones = []

    archivo_existe = False
    preview_url = None
    controles = []

    # ============================
    # 2) Validación de plantilla REAL en GCS
    # ============================
    if plantilla:

        ruta = plantilla["gcs_path"]

        # Ruta vacía → plantilla no válida
        if not ruta:
            plantilla = None

        else:
            client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
            bucket = client.bucket(settings.GCP_BUCKET_NAME)

            # Caso: carpeta
            if ruta.endswith("/"):
                blobs = list(bucket.list_blobs(prefix=ruta))
                archivos = [b for b in blobs if not b.name.endswith("/")]

                if archivos:
                    archivo_existe = True
                    archivo_real = archivos[0].name
                    plantilla["gcs_path"] = archivo_real

                    preview_url = generar_url_previa(archivo_real)
                    controles = extraer_controles_contenido_desde_gcs(archivo_real)
                else:
                    plantilla = None  # carpeta vacía
                    archivo_existe = False

            else:
                # Caso: archivo directo
                blob = bucket.blob(ruta)

                if blob.exists():
                    archivo_existe = True
                    preview_url = generar_url_previa(ruta)
                    controles = extraer_controles_contenido_desde_gcs(ruta)
                else:
                    plantilla = None
                    archivo_existe = False

    office_url = office_or_download_url(preview_url) if preview_url else None

    # ============================
    # 3) Procesar historial de versiones
    # ============================
    versiones_procesadas = []

    for v in versiones:
        entrada = v.copy()

        gpath = v.get("gcs_path")

        if gpath:
            try:
                prev = generar_url_previa(gpath)
                entrada["office_url"] = office_or_download_url(prev) if prev else None
            except:
                entrada["office_url"] = None
        else:
            entrada["office_url"] = None

        versiones_procesadas.append(entrada)

    # ============================
    # 4) Render final
    # ============================
    return render(request, "tipo_detalle.html", {
        "tipo": tipo,
        "plantilla": plantilla,
        "archivo_existe": archivo_existe,
        "preview_url": preview_url,
        "office_url": office_url,
        "controles": controles,
        "versiones": versiones_procesadas,
    })


# =============================================================================
# ABREVIATURAS
# =============================================================================

PALABRAS_IGNORAR = {
    "y", "de", "del", "la", "los", "las", "el", "en", "por",
    "para", "ltlda", "ltda", "sa", "empresa", "asociacion",
    "compania", "hermanos"
}


def generar_abreviatura(nombre, tipo="categoria"):

    if not nombre:
        return ""

    limpio = re.sub(r"[^A-Za-zÁÉÍÓÚáéíóúÑñ\s]", "", nombre)
    palabras = [p for p in limpio.split() if p.lower() not in PALABRAS_IGNORAR]

    if not palabras:
        return ""

    if len(palabras) == 1:
        return palabras[0][:3].upper() if tipo == "categoria" else palabras[0][:4].upper()

    return "".join(p[0].upper() for p in palabras[:4])


# =============================================================================
# CREAR CATEGORÍA
# =============================================================================

def crear_categoria(request):
    abreviatura_generada = ""

    if request.method == "POST":
        nombre = request.POST.get("nombre", "").strip()
        descripcion = request.POST.get("descripcion", "").strip()

        abreviatura = generar_abreviatura(nombre, tipo="categoria")
        abreviatura_generada = abreviatura

        if not abreviatura:
            messages.error(request, "El nombre no genera abreviatura válida.")
        else:
            with connection.cursor() as cursor:
                cursor.execute("""
                    SELECT COUNT(*) 
                    FROM categoria_documentos_tecnicos
                    WHERE UPPER(abreviatura) = UPPER(%s)
                """, [abreviatura])
                existe = cursor.fetchone()[0]

            if existe > 0:
                messages.warning(request, f"La abreviatura '{abreviatura}' ya existe.")
            else:
                messages.success(request, f"✔ Abreviatura: {abreviatura}")

    return render(request, "crear_categoria_documento.html", {
        "abreviatura": abreviatura_generada
    })


# =============================================================================
# CREAR TIPO DOCUMENTO
# =============================================================================

def crear_tipo_documento(request):

    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id, nombre, abreviatura 
            FROM categoria_documentos_tecnicos
            ORDER BY nombre
        """)
        categorias = [
            {"id": r[0], "nombre": r[1], "abreviatura": r[2]}
            for r in cursor.fetchall()
        ]

    if request.method == "POST":

        nombre = request.POST.get("nombre", "").strip()
        descripcion = request.POST.get("descripcion", "").strip()
        categoria_id = request.POST.get("categoria_id")
        abreviatura_manual = request.POST.get("abreviatura", "").strip().upper()

        if not categoria_id or not nombre:
            messages.error(request, "Todos los campos son obligatorios.")
            return redirect(request.path)

        abreviatura_generada = generar_abreviatura(nombre, tipo="tipo_documento")
        abreviatura_final = abreviatura_manual or abreviatura_generada

        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT COUNT(*)
                FROM tipo_documentos_tecnicos
                WHERE categoria_id = %s AND UPPER(abreviatura) = UPPER(%s)
            """, [categoria_id, abreviatura_final])
            existe = cursor.fetchone()[0]

        if existe > 0:
            messages.error(request, "Abreviatura duplicada.")
            return redirect(request.path)

        with connection.cursor() as cursor:
            cursor.execute("""
                INSERT INTO tipo_documentos_tecnicos
                (categoria_id, nombre, descripcion, abreviatura)
                VALUES (%s, %s, %s, %s)
                RETURNING id
            """, [categoria_id, nombre, descripcion, abreviatura_final])
            nuevo_id = cursor.fetchone()[0]

        messages.success(request, "✔ Tipo creado.")
        return redirect("plantillas:detalle_tipo", tipo_id=nuevo_id)

    return render(request, "crear_tipo_documento.html", {"categorias": categorias})


# =============================================================================
# SUBIR PLANTILLA DE DOCUMENTO TÉCNICO (GENÉRICA)
# =============================================================================

@login_required
def subir_plantilla(request, tipo_id):
    """
    Versión genérica. Ahora usa las tablas:
      - plantilla_tipo_doc
      - plantilla_tipo_doc_versiones
    """

    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT t.id, t.nombre, c.nombre as categoria
            FROM tipo_documentos_tecnicos t
            JOIN categoria_documentos_tecnicos c ON c.id = t.categoria_id
            WHERE t.id = %s
        """, [tipo_id])
        tipo = dictfetchone(cursor)

    if not tipo:
        messages.error(request, "Tipo de documento no encontrado.")
        return redirect("plantillas:lista_plantillas")

    if request.method == "POST":

        archivo = request.FILES.get("plantilla")
        if not archivo:
            messages.error(request, "Selecciona un archivo DOCX.")
            return redirect(request.path)

        categoria = clean(tipo["categoria"])
        tipo_nom = clean(tipo["nombre"])

        base_path = f"Plantillas/Documentos_Tecnicos/{categoria}/{tipo_nom}/"
        client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
        bucket = client.bucket(settings.GCP_BUCKET_NAME)

        # ===== 1) Buscar/crear registro maestro en plantilla_tipo_doc =====
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT p.id AS plantilla_id, v.version, v.gcs_path
                FROM plantilla_tipo_doc p
                LEFT JOIN plantilla_tipo_doc_versiones v
                    ON v.id = p.version_actual_id
                WHERE p.tipo_documento_id = %s
                LIMIT 1
            """, [tipo_id])
            anterior = dictfetchone(cursor)

        if anterior:
            plantilla_id = anterior["plantilla_id"]
            version_anterior = anterior["version"]
            ruta_anterior = anterior["gcs_path"]
            controles_antes = extraer_controles_contenido_desde_gcs(ruta_anterior) if ruta_anterior else []
        else:
            # crear registro maestro
            with connection.cursor() as cursor:
                cursor.execute("""
                    INSERT INTO plantilla_tipo_doc (tipo_documento_id)
                    VALUES (%s)
                    RETURNING id
                """, [tipo_id])
                plantilla_id = cursor.fetchone()[0]
            version_anterior = None
            controles_antes = []

        # ===== 2) Guardar archivo temporal =====
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            for chunk in archivo.chunks():
                tmp.write(chunk)
            nuevo_local = tmp.name

        controles_despues = extraer_controles_contenido_desde_file(nuevo_local)

        # ===== 3) Calcular versión =====
        nueva_version = versionar_plantilla(version_anterior, controles_antes, controles_despues)

        # ===== 4) Crear carpeta V{version} =====
        version_folder = f"{base_path}V{nueva_version}/"
        bucket.blob(version_folder).upload_from_string("")

        filename = archivo.name
        blob_name = f"{version_folder}{filename}"

        blob = bucket.blob(blob_name)
        blob.upload_from_filename(nuevo_local)

        # ===== 5) Insertar nueva versión en plantilla_tipo_doc_versiones =====
        with connection.cursor() as cursor:
            cursor.execute("""
                INSERT INTO plantilla_tipo_doc_versiones (plantilla_id, version, gcs_path, controles)
                VALUES (%s, %s, %s, %s)
                RETURNING id
            """, [plantilla_id, nueva_version, blob_name, json.dumps(controles_despues)])
            version_id = cursor.fetchone()[0]

            cursor.execute("""
                UPDATE plantilla_tipo_doc
                SET version_actual_id = %s,
                    actualizado_en = NOW()
                WHERE id = %s
            """, [version_id, plantilla_id])

        messages.success(request, f"✔ Plantilla subida (versión {nueva_version}).")
        return redirect("plantillas:detalle_tipo", tipo_id=tipo_id)

    return render(request, "subir_plantilla.html", {"tipo": tipo})


# =============================================================================
# EXTRACCIÓN DE ETIQUETAS DOCX (sin BD)
# =============================================================================

def extraer_etiquetas_word(blob):
    """
    Recibe un blob de GCS, descarga el DOCX y extrae {{etiquetas}} del cuerpo.
    No guarda nada en BD. Se usará bajo demanda (auditorías, etc.).
    """
    with tempfile.NamedTemporaryFile(delete=False) as tmp:
        blob.download_to_file(tmp)
        tmp_path = tmp.name

    with ZipFile(tmp_path) as docx:
        xml = docx.read("word/document.xml").decode("utf-8")

    encontrados = re.findall(r"\{\{(.*?)\}\}", xml)
    return sorted(set(e.strip() for e in encontrados))


# =============================================================================
# ======================   PORTADA   WORD   =======================
# =============================================================================

@login_required
def portada_word_detalle(request):
    """
    Usa plantilla_portada + plantilla_portada_versiones
    y filtra por ruta que contenga /Portada/Word/
    """
    with connection.cursor() as cursor:
        # Versión actual (Word)
        cursor.execute("""
            SELECT 
                pv.id,
                pv.plantilla_id,
                pv.gcs_path,
                pv.version,
                pv.creado_en
            FROM plantilla_portada p
            JOIN plantilla_portada_versiones pv
                ON pv.id = p.version_actual_id
            WHERE p.utilidad_id = 1
              AND pv.gcs_path LIKE 'Plantillas/Utilidad/Portada/Word/%'
            ORDER BY pv.creado_en DESC, pv.id DESC
            LIMIT 1
        """)
        plantilla = dictfetchone(cursor)

        versiones = []
        if plantilla:
            cursor.execute("""
                SELECT 
                    id,
                    gcs_path,
                    version,
                    creado_en
                FROM plantilla_portada_versiones
                WHERE plantilla_id = %s
                ORDER BY creado_en DESC, id DESC
            """, [plantilla["plantilla_id"]])
            versiones = dictfetchall(cursor)

    preview_url = None
    controles = []

    if plantilla:
        preview_url = generar_url_previa(plantilla["gcs_path"])
        controles = extraer_controles_contenido_desde_gcs(plantilla["gcs_path"])

    office_url = office_or_download_url(preview_url) if preview_url else None

    versiones_procesadas = []
    for v in versiones:
        entrada = v.copy()
        if v["gcs_path"]:
            try:
                prev = generar_url_previa(v["gcs_path"])
                entrada["office_url"] = office_or_download_url(prev)
            except:
                entrada["office_url"] = None
        else:
            entrada["office_url"] = None
        versiones_procesadas.append(entrada)

    return render(request, "portada_word_detalle.html", {
        "plantilla": plantilla,
        "preview_url": preview_url,
        "office_url": office_url,
        "controles": controles,
        "versiones": versiones_procesadas,
    })


@login_required
def subir_portada_word(request):

    if request.method == "POST":
        archivo = request.FILES.get("plantilla")

        if not archivo:
            messages.error(request, "Debes seleccionar un archivo DOCX.")
            return redirect(request.path)

        client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
        bucket = client.bucket(settings.GCP_BUCKET_NAME)

        # ===========================
        # 1) Buscar/crear portada (utilidad_id=1) asociada a Word
        # ===========================
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT 
                    p.id AS plantilla_id,
                    pv.id AS version_id,
                    pv.gcs_path,
                    pv.version
                FROM plantilla_portada p
                LEFT JOIN plantilla_portada_versiones pv
                    ON pv.id = p.version_actual_id
                WHERE p.utilidad_id = 1
                  AND (pv.gcs_path LIKE 'Plantillas/Utilidad/Portada/Word/%'
                       OR pv.gcs_path IS NULL)
                ORDER BY p.id
                LIMIT 1
            """)
            existente = dictfetchone(cursor)

        controles_antes = []
        version_anterior = None

        if existente and existente.get("version") is not None:
            plantilla_id = existente["plantilla_id"]
            version_anterior = existente["version"]
            if existente["gcs_path"]:
                controles_antes = extraer_controles_contenido_desde_gcs(existente["gcs_path"])
        else:
            # Crear registro maestro para utilidad Portada (id=1)
            with connection.cursor() as cursor:
                cursor.execute("""
                    INSERT INTO plantilla_portada (utilidad_id)
                    VALUES (1)
                    RETURNING id
                """)
                plantilla_id = cursor.fetchone()[0]

        # ===========================
        # 2) Guardar archivo temporal
        # ===========================
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            for chunk in archivo.chunks():
                tmp.write(chunk)
            nuevo_path_local = tmp.name

        controles_despues = extraer_controles_contenido_desde_file(nuevo_path_local)

        # ===========================
        # 3) Calcular nueva versión
        # ===========================
        nueva_version = versionar_plantilla(version_anterior, controles_antes, controles_despues)

        # ===========================
        # 4) Crear carpeta V{version}
        # ===========================
        base_path = "Plantillas/Utilidad/Portada/Word/"
        version_folder = f"{base_path}V{nueva_version}/"
        bucket.blob(version_folder).upload_from_string("")

        # ===========================
        # 5) Nombre final del archivo
        # ===========================
        filename = archivo.name
        blob_name = f"{version_folder}{filename}"

        # ===========================
        # 6) Subir archivo a la carpeta
        # ===========================
        blob = bucket.blob(blob_name)
        blob.upload_from_filename(nuevo_path_local)

        # ===========================
        # 7) Insertar nueva versión en plantilla_portada_versiones
        # ===========================
        with connection.cursor() as cursor:
            cursor.execute("""
                INSERT INTO plantilla_portada_versiones (plantilla_id, version, gcs_path, controles)
                VALUES (%s, %s, %s, %s)
                RETURNING id
            """, [plantilla_id, nueva_version, blob_name, json.dumps(controles_despues)])
            version_id = cursor.fetchone()[0]

            cursor.execute("""
                UPDATE plantilla_portada
                SET version_actual_id = %s,
                    actualizado_en = NOW()
                WHERE id = %s
            """, [version_id, plantilla_id])

        messages.success(request, f"✔ Portada Word actualizada (versión {nueva_version})")
        return redirect("plantillas:portada_word_detalle")

    return render(request, "subir_portada_word.html")


# =============================================================================
# ======================   PORTADA   EXCEL   ======================
# =============================================================================

@login_required
def portada_excel_detalle(request):
    """
    Usa plantilla_portada + plantilla_portada_versiones,
    filtrando por rutas con /Portada/Excel/
    """
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT 
                pv.id,
                pv.plantilla_id,
                pv.gcs_path,
                pv.version,
                pv.creado_en
            FROM plantilla_portada p
            JOIN plantilla_portada_versiones pv
                ON pv.id = p.version_actual_id
            WHERE p.utilidad_id = 1
              AND pv.gcs_path LIKE 'Plantillas/Utilidad/Portada/Excel/%'
            ORDER BY pv.creado_en DESC, pv.id DESC
            LIMIT 1
        """)
        plantilla = dictfetchone(cursor)

    etiquetas = []  # Por ahora, sin análisis de etiquetas en Excel

    return render(request, "portada_excel_detalle.html", {
        "plantilla": plantilla,
        "etiquetas": etiquetas,
    })


@login_required
def subir_portada_excel(request):

    if request.method == "POST":
        archivo = request.FILES.get("plantilla")

        if not archivo:
            messages.error(request, "Debes seleccionar un archivo Excel.")
            return redirect(request.path)

        client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
        bucket = client.bucket(settings.GCP_BUCKET_NAME)

        # Buscar/crear portada Excel (utilidad_id=1, ruta Excel)
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT 
                    p.id AS plantilla_id,
                    pv.id AS version_id,
                    pv.gcs_path,
                    pv.version
                FROM plantilla_portada p
                LEFT JOIN plantilla_portada_versiones pv
                    ON pv.id = p.version_actual_id
                WHERE p.utilidad_id = 1
                  AND (pv.gcs_path LIKE 'Plantillas/Utilidad/Portada/Excel/%'
                       OR pv.gcs_path IS NULL)
                ORDER BY p.id
                LIMIT 1
            """)
            existente = dictfetchone(cursor)

        if existente and existente.get("version") is not None:
            plantilla_id = existente["plantilla_id"]
            version_anterior = existente["version"]
        else:
            with connection.cursor() as cursor:
                cursor.execute("""
                    INSERT INTO plantilla_portada (utilidad_id)
                    VALUES (1)
                    RETURNING id
                """)
                plantilla_id = cursor.fetchone()[0]
            version_anterior = None

        nueva_version = siguiente_version(version_anterior)

        base_path = "Plantillas/Utilidad/Portada/Excel/"
        version_folder = f"{base_path}V{nueva_version}/"
        bucket.blob(version_folder).upload_from_string("")

        blob_name = f"{version_folder}{archivo.name}"
        blob = bucket.blob(blob_name)
        blob.upload_from_file(archivo)

        with connection.cursor() as cursor:
            cursor.execute("""
                INSERT INTO plantilla_portada_versiones (plantilla_id, version, gcs_path)
                VALUES (%s, %s, %s)
                RETURNING id
            """, [plantilla_id, nueva_version, blob_name])
            version_id = cursor.fetchone()[0]

            cursor.execute("""
                UPDATE plantilla_portada
                SET version_actual_id = %s,
                    actualizado_en = NOW()
                WHERE id = %s
            """, [version_id, plantilla_id])

        messages.success(request, f"✔ Portada Excel subida/actualizada (versión {nueva_version}).")
        return redirect("plantillas:portada_excel_detalle")

    return render(request, "subir_portada_excel.html")


# =============================================================================
# URL FIRMADA GCS
# =============================================================================

def generar_url_previa(blob_path):
    """
    Genera URL temporal para el archivo (3 horas).
    """
    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
    bucket = client.bucket(settings.GCP_BUCKET_NAME)
    blob = bucket.blob(blob_path)

    try:
        url = blob.generate_signed_url(
            version="v4",
            expiration=timedelta(hours=3),
            method="GET",
        )
        return url
    except Exception:
        return None




# =============================================================================
# SUBIR PLANTILLA (ESPECÍFICA PARA TIPO DE DOCUMENTO)
# =============================================================================

@login_required
def subir_plantilla_tipo_doc(request, tipo_id):

    # ===== 1) Datos del tipo de documento =====
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT t.id, t.nombre, c.nombre as categoria
            FROM tipo_documentos_tecnicos t
            JOIN categoria_documentos_tecnicos c ON c.id = t.categoria_id
            WHERE t.id = %s
        """, [tipo_id])
        tipo = dictfetchone(cursor)

    if not tipo:
        messages.error(request, "Tipo de documento no encontrado.")
        return redirect("plantillas:lista_plantillas")

    if request.method == "POST":

        archivo = request.FILES.get("plantilla")
        if not archivo:
            messages.error(request, "Debes seleccionar un archivo DOCX.")
            return redirect(request.path)

        categoria = clean(tipo["categoria"])
        tipo_nom = clean(tipo["nombre"])

        # Base path sin versión
        base_path = f"Plantillas/Documentos_Tecnicos/{categoria}/{tipo_nom}/"

        # ===== 2) Buscar registro maestro y versión anterior =====
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT 
                    p.id AS plantilla_id,
                    v.version,
                    v.gcs_path
                FROM plantilla_tipo_doc p
                LEFT JOIN plantilla_tipo_doc_versiones v
                    ON v.id = p.version_actual_id
                WHERE p.tipo_documento_id = %s
                LIMIT 1
            """, [tipo_id])
            anterior = dictfetchone(cursor)

        if anterior:
            plantilla_id = anterior["plantilla_id"]
            version_anterior = anterior["version"]
            ruta_anterior = anterior["gcs_path"]
            controles_antes = extraer_controles_contenido_desde_gcs(ruta_anterior) if ruta_anterior else []
        else:
            with connection.cursor() as cursor:
                cursor.execute("""
                    INSERT INTO plantilla_tipo_doc (tipo_documento_id)
                    VALUES (%s)
                    RETURNING id
                """, [tipo_id])
                plantilla_id = cursor.fetchone()[0]
            version_anterior = None
            controles_antes = []

        # ===== 3) Guardar archivo temporal =====
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            for chunk in archivo.chunks():
                tmp.write(chunk)
            nuevo_local = tmp.name

        controles_despues = extraer_controles_contenido_desde_file(nuevo_local)

        # ===== 4) Calcular versión =====
        nueva_version = versionar_plantilla(version_anterior, controles_antes, controles_despues)

        # ===== 5) Crear carpeta V{version} =====
        client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
        bucket = client.bucket(settings.GCP_BUCKET_NAME)

        version_folder = f"{base_path}V{nueva_version}/"
        bucket.blob(version_folder).upload_from_string("")  # Crear carpeta vacía

        # ===== 6) Subir archivo dentro de carpeta versión =====
        filename = archivo.name
        blob_name = f"{version_folder}{filename}"

        blob = bucket.blob(blob_name)
        blob.upload_from_filename(nuevo_local)

        # ===== 7) Insertar registro de versión =====
        with connection.cursor() as cursor:
            cursor.execute("""
                INSERT INTO plantilla_tipo_doc_versiones (plantilla_id, version, gcs_path, controles)
                VALUES (%s, %s, %s, %s)
                RETURNING id
            """, [plantilla_id, nueva_version, blob_name, json.dumps(controles_despues)])
            version_id = cursor.fetchone()[0]

            cursor.execute("""
                UPDATE plantilla_tipo_doc
                SET version_actual_id = %s,
                    actualizado_en = NOW()
                WHERE id = %s
            """, [version_id, plantilla_id])

        messages.success(
            request,
            f"✔ Plantilla subida correctamente (versión {nueva_version})."
        )

        return redirect("plantillas:detalle_tipo", tipo_id=tipo_id)

    return render(request, "subir_plantilla_tipo_doc.html", {"tipo": tipo})


@login_required
def descargar_gcs(request, path):
    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
    bucket = client.bucket(settings.GCP_BUCKET_NAME)
    blob = bucket.blob(path)

    if not blob.exists():
        raise Http404("Archivo no encontrado.")

    contenido = blob.download_as_bytes()

    response = HttpResponse(
        contenido,
        content_type="application/octet-stream"
    )
    response['Content-Disposition'] = f'attachment; filename="{os.path.basename(path)}"'
    return response


def office_or_download_url(preview_url):
    """
    Intenta abrir en Office Web Apps.
    Si Office no puede abrirlo, Office lo descarga.
    """
    encoded = urllib.parse.quote(preview_url, safe='')
    return f"https://view.officeapps.live.com/op/view.aspx?src={encoded}"


def versionar_plantilla(version_actual, controles_antes, controles_despues):
    """
    Sistema de versionado:
    - La primera versión SIEMPRE es 1.0
    - Cambio mayor si los controles cambian
    - Cambio menor si solo cambia el archivo sin afectar controles
    """

    # 1) Primera vez → siempre 1.0
    if not version_actual:
        return "1.0"

    # 2) Validar formato major.minor
    try:
        partes = str(version_actual).split(".")
        major = int(partes[0])
        minor = int(partes[1])
    except:
        major, minor = 1, 0

    # 3) Comparar controles
    set_antes = set(controles_antes or [])
    set_despues = set(controles_despues or [])

    # CAMBIO MAYOR → si cambia algún control
    if set_antes != set_despues:
        return f"{major + 1}.0"

    # CAMBIO MENOR → si no cambian controles
    return f"{major}.{minor + 1}"


def crear_carpeta_version(bucket, base_path, version):
    """
    Crea la carpeta de versión, ej:
    base_path = Plantillas/Documentos_Tecnicos/Mecanica/Informe/
    version = 1.0
    Resultado → Plantillas/Documentos_Tecnicos/Mecanica/Informe/V1.0/
    """
    version_path = f"{base_path}V{version}/"
    bucket.blob(version_path).upload_from_string("")  # crear carpeta vacía
    return version_path




@login_required
def eliminar_plantilla(request, tipo_id):

    # ============================
    # 1) Obtener datos del tipo
    # ============================
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT t.id, t.nombre, c.nombre AS categoria
            FROM tipo_documentos_tecnicos t
            JOIN categoria_documentos_tecnicos c ON c.id = t.categoria_id
            WHERE t.id = %s
        """, [tipo_id])
        tipo = dictfetchone(cursor)

    if not tipo:
        messages.error(request, "Tipo de documento no encontrado.")
        return redirect("plantillas:lista_plantillas")

    categoria = clean(tipo["categoria"])
    tipo_nom = clean(tipo["nombre"])

    # Carpeta base del tipo de documento
    base_path = f"Plantillas/Documentos_Tecnicos/{categoria}/{tipo_nom}/"

    # ============================
    # 2) Cliente GCS
    # ============================
    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
    bucket = client.bucket(settings.GCP_BUCKET_NAME)

    # ============================
    # 3) Eliminar todos los blobs en la carpeta
    # ============================
    blobs = list(bucket.list_blobs(prefix=base_path))

    for b in blobs:
        try:
            b.delete()
        except Exception:
            pass  # si ya no existe, continuar

    # ============================
    # 4) Eliminar registros en BD (nuevas tablas)
    # ============================
    with connection.cursor() as cursor:
        # Primero eliminar versiones
        cursor.execute("""
            DELETE FROM plantilla_tipo_doc_versiones
            WHERE plantilla_id IN (
                SELECT id
                FROM plantilla_tipo_doc
                WHERE tipo_documento_id = %s
            )
        """, [tipo_id])

        # Luego eliminar registros maestro
        cursor.execute("""
            DELETE FROM plantilla_tipo_doc
            WHERE tipo_documento_id = %s
        """, [tipo_id])

    messages.success(request, "️Todas las versiones de la plantilla fueron eliminadas correctamente.")

    return redirect("plantillas:detalle_tipo", tipo_id=tipo_id)


@login_required
def editar_tipo_documento(request, tipo_id):

    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT 
                t.id, 
                t.categoria_id, 
                t.nombre, 
                t.descripcion, 
                t.abreviatura,
                t.formato_id,
                c.nombre AS categoria_nombre
            FROM tipo_documentos_tecnicos t
            JOIN categoria_documentos_tecnicos c ON c.id = t.categoria_id
            WHERE t.id = %s
        """, [tipo_id])
        tipo = dictfetchone(cursor)

        cursor.execute("""
            SELECT id, nombre
            FROM categoria_documentos_tecnicos
            ORDER BY nombre
        """)
        categorias = dictfetchall(cursor)

        cursor.execute("""
            SELECT id, nombre, extension
            FROM formato_archivo
            ORDER BY id
        """)
        formatos = dictfetchall(cursor)

    if not tipo:
        messages.error(request, "Tipo no encontrado.")
        return redirect("plantillas:lista_plantillas")

    # Valores originales para paths
    nombre_original = tipo["nombre"]
    categoria_original_nombre = tipo["categoria_nombre"]
    categoria_original_clean = clean(categoria_original_nombre)
    tipo_original_clean = clean(nombre_original)

    if request.method == "POST":
        nombre = request.POST.get("nombre", "").strip()
        descripcion = request.POST.get("descripcion", "").strip()
        abreviatura = request.POST.get("abreviatura", "").strip()
        categoria_id = request.POST.get("categoria_id")
        formato_id = request.POST.get("formato_id")

        if not nombre or not categoria_id:
            messages.error(request, "Nombre y categoría son obligatorios.")
            return redirect(request.path)

        # Obtener nombre de la nueva categoría
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT nombre
                FROM categoria_documentos_tecnicos
                WHERE id = %s
            """, [categoria_id])
            row = cursor.fetchone()

        if not row:
            messages.error(request, "Categoría seleccionada no existe.")
            return redirect(request.path)

        categoria_nueva_nombre = row[0]
        categoria_nueva_clean = clean(categoria_nueva_nombre)
        tipo_nuevo_clean = clean(nombre)

        old_prefix = f"Plantillas/Documentos_Tecnicos/{categoria_original_clean}/{tipo_original_clean}/"
        new_prefix = f"Plantillas/Documentos_Tecnicos/{categoria_nueva_clean}/{tipo_nuevo_clean}/"

        # Si cambió nombre y/o categoría → mover carpeta y actualizar paths
        if new_prefix != old_prefix:
            try:
                mover_carpeta_gcs(old_prefix, new_prefix)

                with connection.cursor() as cursor:
                    cursor.execute("""
                        UPDATE plantilla_tipo_doc_versiones
                        SET gcs_path = REPLACE(gcs_path, %s, %s)
                        WHERE plantilla_id IN (
                            SELECT id
                            FROM plantilla_tipo_doc
                            WHERE tipo_documento_id = %s
                        )
                          AND gcs_path LIKE %s
                    """, [old_prefix, new_prefix, tipo_id, old_prefix + "%"])
            except Exception as e:
                messages.error(request, f"Error al renombrar en GCS: {e}")
                return redirect("plantillas:detalle_tipo", tipo_id=tipo_id)

        # Actualizar datos del tipo
        with connection.cursor() as cursor:
            cursor.execute("""
                UPDATE tipo_documentos_tecnicos
                SET categoria_id = %s,
                    nombre = %s,
                    descripcion = %s,
                    abreviatura = %s,
                    formato_id = %s
                WHERE id = %s
            """, [categoria_id, nombre, descripcion, abreviatura, formato_id, tipo_id])

        messages.success(request, "✔ Tipo de documento actualizado.")
        return redirect("plantillas:detalle_tipo", tipo_id=tipo_id)

    return render(request, "editar_tipo_documento.html", {
        "tipo": tipo,
        "categorias": categorias,
        "formatos": formatos
    })


@login_required
def eliminar_version(request, version_id):

    # 1) Obtener registro de la versión (nueva tabla)
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT 
                v.id,
                v.gcs_path,
                v.plantilla_id,
                p.tipo_documento_id,
                p.version_actual_id
            FROM plantilla_tipo_doc_versiones v
            JOIN plantilla_tipo_doc p ON p.id = v.plantilla_id
            WHERE v.id = %s
        """, [version_id])
        reg = dictfetchone(cursor)

    if not reg:
        messages.error(request, "Versión no encontrada.")
        return redirect("plantillas:lista_plantillas")

    tipo_id = reg["tipo_documento_id"]
    gcs_path = reg["gcs_path"]
    plantilla_id = reg["plantilla_id"]
    version_actual_id = reg["version_actual_id"]

    # 2) Extraer carpetas desde gcs_path
    version_folder = "/".join(gcs_path.split("/")[:-1]) + "/"
    base_folder = "/".join(version_folder.split("/")[:-2]) + "/"

    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
    bucket = client.bucket(settings.GCP_BUCKET_NAME)

    # 3) Eliminar carpeta de la versión
    blobs_version = list(bucket.list_blobs(prefix=version_folder))
    for blob in blobs_version:
        try:
            blob.delete()
        except:
            pass

    # 4) Eliminar registro en BD
    with connection.cursor() as cursor:
        cursor.execute("DELETE FROM plantilla_tipo_doc_versiones WHERE id = %s", [version_id])

        # Si era la versión actual, buscar nueva versión actual
        if version_id == version_actual_id:
            cursor.execute("""
                SELECT id
                FROM plantilla_tipo_doc_versiones
                WHERE plantilla_id = %s
                ORDER BY version DESC, creado_en DESC, id DESC
                LIMIT 1
            """, [plantilla_id])
            nueva = cursor.fetchone()
            nuevo_version_actual_id = nueva[0] if nueva else None

            cursor.execute("""
                UPDATE plantilla_tipo_doc
                SET version_actual_id = %s
                WHERE id = %s
            """, [nuevo_version_actual_id, plantilla_id])

    # 5) Verificar si quedan otras versiones dentro de la carpeta base del tipo
    blobs_base = list(bucket.list_blobs(prefix=base_folder))

    carpetas_versiones = set()
    for b in blobs_base:
        partes = b.name.split("/")
        for p in partes:
            if p.startswith("V") and "." in p:
                carpetas_versiones.add(p)

    # 6) Si NO quedan carpetas → eliminar carpeta base del tipo
    if len(carpetas_versiones) == 0:

        carpetas_base_restantes = list(bucket.list_blobs(prefix=base_folder))
        for blob in carpetas_base_restantes:
            try:
                blob.delete()
            except:
                pass

        messages.success(request, "✔ Versión y carpeta del tipo eliminadas completamente.")
        return redirect("plantillas:lista_plantillas")

    # Si sí quedan versiones → dejar carpeta base intacta
    messages.success(request, "✔ Versión eliminada exitosamente.")
    return redirect("plantillas:detalle_tipo", tipo_id=tipo_id)


def renombrar_archivo_gcs(ruta_actual, nueva_ruta):
    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
    bucket = client.bucket(settings.GCP_BUCKET_NAME)

    blob_old = bucket.blob(ruta_actual)
    blob_new = bucket.blob(nueva_ruta)

    bucket.copy_blob(blob_old, bucket, nueva_ruta)
    blob_old.delete()

    return nueva_ruta


def mover_carpeta_gcs(old_prefix, new_prefix):
    """
    Mueve TODOS los blobs cuya ruta comienza con old_prefix a new_prefix,
    manteniendo el resto del path igual.
    """
    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
    bucket = client.bucket(settings.GCP_BUCKET_NAME)

    blobs = list(bucket.list_blobs(prefix=old_prefix))

    for b in blobs:
        old_name = b.name
        new_name = old_name.replace(old_prefix, new_prefix, 1)
        bucket.copy_blob(b, bucket, new_name)
        b.delete()

    return True


@login_required
def eliminar_version_portada(request, version_id):

    # 1. Buscar versión en BD nueva
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT 
                v.id,
                v.gcs_path,
                v.plantilla_id,
                p.utilidad_id,
                p.version_actual_id
            FROM plantilla_portada_versiones v
            JOIN plantilla_portada p ON p.id = v.plantilla_id
            WHERE v.id = %s
        """, [version_id])
        reg = dictfetchone(cursor)

    if not reg:
        messages.error(request, "Versión no encontrada.")
        return redirect("plantillas:portada_word_detalle")

    gcs_path = reg["gcs_path"]
    plantilla_id = reg["plantilla_id"]
    version_actual_id = reg["version_actual_id"]

    # 2. Extraer carpeta de la versión
    version_folder = "/".join(gcs_path.split("/")[:-1]) + "/"
    base_folder = "/".join(version_folder.split("/")[:-2]) + "/"

    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
    bucket = client.bucket(settings.GCP_BUCKET_NAME)

    # 3. Borrar todos los archivos dentro de la carpeta de versión
    blobs_version = list(bucket.list_blobs(prefix=version_folder))
    for b in blobs_version:
        try:
            b.delete()
        except:
            pass

    # 4. Eliminar registro en BD
    with connection.cursor() as cursor:
        cursor.execute("DELETE FROM plantilla_portada_versiones WHERE id = %s", [version_id])

        # Si era la versión actual, recalcular
        if version_id == version_actual_id:
            cursor.execute("""
                SELECT id
                FROM plantilla_portada_versiones
                WHERE plantilla_id = %s
                ORDER BY version DESC, creado_en DESC, id DESC
                LIMIT 1
            """, [plantilla_id])
            nueva = cursor.fetchone()
            nuevo_version_actual_id = nueva[0] if nueva else None

            cursor.execute("""
                UPDATE plantilla_portada
                SET version_actual_id = %s
                WHERE id = %s
            """, [nuevo_version_actual_id, plantilla_id])

    # 5. Revisar si quedan otras versiones
    blobs_base = list(bucket.list_blobs(prefix=base_folder))

    carpetas_versiones = set()
    for b in blobs_base:
        partes = b.name.split("/")
        for p in partes:
            if p.startswith("V") and "." in p:
                carpetas_versiones.add(p)

    # 6. Si no hay versiones → eliminar carpeta base
    if len(carpetas_versiones) == 0:
        for b in blobs_base:
            try:
                b.delete()
            except:
                pass
        
        messages.success(request, "✔ Se eliminaron todas las versiones de la portada.")
        return redirect("plantillas:portada_word_detalle")

    messages.success(request, "✔ Versión eliminada.")
    return redirect("plantillas:portada_word_detalle")


# =============================================================================
# EXTRACCIÓN DE CONTROLES (w:sdt) DESDE DOCX (GCS / LOCAL) - MODO PRO
# =============================================================================

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def _normalizar_control(nombre):
    """
    Normaliza el nombre de un control para comparar de forma robusta:
    - quita acentos
    - quita espacios raros
    - reemplaza espacios por _
    - pasa a minúsculas
    - deja solo [a-z0-9_.], el resto lo convierte en _
    - colapsa múltiples _ consecutivos
    """
    if not nombre:
        return None

    # A texto simple
    s = str(nombre)

    # Quitar espacios no separables y similares
    s = s.replace("\u00A0", " ")

    # Quitar acentos
    s = unidecode(s)

    # Strip y espacios → _
    s = s.strip()
    s = re.sub(r"\s+", "_", s)

    # Minúsculas para comparar sin importar mayúsculas
    s = s.lower()

    # Solo letras, números, punto y guión bajo
    s = re.sub(r"[^a-z0-9_.]+", "_", s)

    # Colapsar guiones bajos
    s = re.sub(r"_+", "_", s)

    return s.strip("_") or None


def _extraer_tag_o_alias(sdt_node):
    """
    Devuelve primero el TAG del control si existe (w:tag @w:val),
    si no, devuelve el ALIAS (w:alias @w:val).
    """
    # 1) TAG
    tag_node = sdt_node.xpath(".//w:tag", namespaces=NS)
    if tag_node:
        val = tag_node[0].get(f"{{{NS['w']}}}val")
        if val and str(val).strip():
            return str(val).strip()

    # 2) ALIAS
    alias_node = sdt_node.xpath(".//w:alias", namespaces=NS)
    if alias_node:
        val = alias_node[0].get(f"{{{NS['w']}}}val")
        if val and str(val).strip():
            return str(val).strip()

    return None


def _agregar_control(mapa_controles, nombre_crudo):
    """
    Usa _normalizar_control para evitar duplicados "visualmente iguales".
    mapa_controles: dict { nombre_normalizado: nombre_original_primero }
    """
    norm = _normalizar_control(nombre_crudo)
    if not norm:
        return

    if norm not in mapa_controles:
        mapa_controles[norm] = nombre_crudo.strip()


def _procesar_docx_zip(docx, mapa_controles, log_prefix=""):
    """
    Procesa un ZipFile de DOCX y agrega controles encontrados en:
    - word/document.xml
    - word/header*.xml
    - word/footer*.xml
    """
    print(f"{log_prefix}📂 Archivos dentro del DOCX:")
    for n in docx.namelist():
        print(f"{log_prefix}   - {n}")

    # ---------------------------
    # Helper interno para escanear una parte
    # ---------------------------
    def _scan_part(part_name, etiqueta="document"):
        try:
            xml = etree.fromstring(docx.read(part_name))
            encontrados = xml.xpath("//w:sdt", namespaces=NS)
            print(f"{log_prefix}🧩 Controles en {etiqueta} ({part_name}): {len(encontrados)}")
            for sdt in encontrados:
                raw = _extraer_tag_o_alias(sdt)
                print(f"{log_prefix}   - raw: {raw}")
                if raw:
                    _agregar_control(mapa_controles, raw)
        except Exception as e:
            print(f"{log_prefix}⚠ Error procesando parte {part_name}: {e}")

    # ---------------------------
    # DOCUMENTO PRINCIPAL
    # ---------------------------
    if "word/document.xml" in docx.namelist():
        _scan_part("word/document.xml", etiqueta="document.xml")

    # ---------------------------
    # ENCABEZADOS
    # ---------------------------
    for name in docx.namelist():
        if name.startswith("word/header") and name.endswith(".xml"):
            _scan_part(name, etiqueta="header")

    # ---------------------------
    # PIE DE PÁGINA
    # ---------------------------
    for name in docx.namelist():
        if name.startswith("word/footer") and name.endswith(".xml"):
            _scan_part(name, etiqueta="footer")


def extraer_controles_contenido_desde_gcs(gcs_path):
    """
    Descarga un DOCX desde GCS y extrae todos los controles de contenido (w:sdt).
    Devuelve lista ordenada (sin duplicados, con preferencia TAG > ALIAS).
    """

    print("\n" + "=" * 80)
    print("🔍 EXTRAER CONTROLES DESDE GCS (MODO PRO)")
    print("📌 Ruta solicitada:", gcs_path)

    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
    bucket = client.bucket(settings.GCP_BUCKET_NAME)
    blob = bucket.blob(gcs_path)

    exists = blob.exists()
    print("📦 ¿Blob existe en GCS?:", exists)

    if not exists:
        print("❌ Blob no existe en GCS")
        print("=" * 80)
        return []

    # Descargar archivo local temp
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        blob.download_to_filename(tmp.name)
        local_path = tmp.name

    print("📄 Archivo DOCX descargado en:", local_path)
    size = os.path.getsize(local_path)
    print("📏 Tamaño del archivo:", size, "bytes")

    if size < 50:
        print("⚠ Archivo demasiado pequeño, parece inválido")
        print("=" * 80)
        return []

    mapa_controles = {}

    try:
        with ZipFile(local_path, "r") as docx:
            _procesar_docx_zip(docx, mapa_controles, log_prefix="[GCS] ")
    except Exception as e:
        print("⚠ ERROR procesando DOCX (GCS):", e)
        import traceback
        traceback.print_exc()

    # Convertir dict a lista de nombres "bonitos"
    controles = sorted(mapa_controles.values(), key=lambda x: x.lower())

    print("[GCS] 🔎 Total controles detectados (normalizados):", len(mapa_controles))
    print("[GCS] ➡ Controles devueltos:", controles)
    print("=" * 80 + "\n")

    return controles


def extraer_controles_contenido_desde_file(local_path):
    """
    Extrae controles desde un DOCX local (ruta en disco).
    Misma lógica que extraer_controles_contenido_desde_gcs, pero sin GCS.
    """

    print("\n" + "=" * 80)
    print("🔍 EXTRAER CONTROLES DESDE ARCHIVO LOCAL (MODO PRO)")
    print("📄 Archivo:", local_path)

    if not os.path.exists(local_path):
        print("❌ El archivo local no existe.")
        print("=" * 80)
        return []

    size = os.path.getsize(local_path)
    print("📏 Tamaño:", size, "bytes")

    if size < 50:
        print("⚠ Archivo demasiado pequeño, parece inválido")
        print("=" * 80)
        return []

    mapa_controles = {}

    try:
        with ZipFile(local_path, "r") as docx:
            _procesar_docx_zip(docx, mapa_controles, log_prefix="[LOCAL] ")
    except Exception as e:
        print("⚠ ERROR leyendo archivo local:", e)
        import traceback
        traceback.print_exc()

    controles = sorted(mapa_controles.values(), key=lambda x: x.lower())

    print("[LOCAL] 🔎 Total controles detectados (normalizados):", len(mapa_controles))
    print("[LOCAL] ➡ Controles devueltos:", controles)
    print("=" * 80 + "\n")

    return controles
