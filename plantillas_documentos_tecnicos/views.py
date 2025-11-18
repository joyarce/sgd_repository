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

                # Actualizar rutas en BD
                with connection.cursor() as cursor:
                    cursor.execute("""
                        UPDATE plantillas_documentos_tecnicos
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

        # Última versión (actual)
        cursor.execute("""
            SELECT id, gcs_path, version, creado_en
            FROM plantillas_documentos_tecnicos
            WHERE tipo_documento_id = %s
            ORDER BY id DESC LIMIT 1
        """, [tipo_id])
        plantilla = dictfetchone(cursor)

        # Todas las versiones
        cursor.execute("""
            SELECT id, gcs_path, version, creado_en
            FROM plantillas_documentos_tecnicos
            WHERE tipo_documento_id = %s
            ORDER BY id DESC
        """, [tipo_id])
        versiones = dictfetchall(cursor)

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
    Versión genérica que podrías seguir usando, pero ya tienes
    `subir_plantilla_tipo_doc`. Puedes dejar una sola en uso si quieres.
    """

    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT t.id, t.nombre, c.nombre as categoria
            FROM tipo_documentos_tecnicos t
            JOIN categoria_documentos_tecnicos c ON c.id = t.categoria_id
            WHERE t.id = %s
        """, [tipo_id])
        tipo = dictfetchone(cursor)

    if request.method == "POST":

        archivo = request.FILES.get("plantilla")
        if not archivo:
            messages.error(request, "Selecciona un archivo DOCX.")
            return redirect(request.path)

        categoria = clean(tipo["categoria"])
        tipo_nom = clean(tipo["nombre"])

        path = f"Plantillas/Documentos_Tecnicos/{categoria}/{tipo_nom}/"
        blob_name = f"{path}{archivo.name}"

        client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
        bucket = client.bucket(settings.GCP_BUCKET_NAME)
        blob = bucket.blob(blob_name)
        blob.upload_from_file(archivo)

        # Obtener última versión previa, si existe
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT version
                FROM plantillas_documentos_tecnicos
                WHERE tipo_documento_id = %s
                ORDER BY id DESC LIMIT 1
            """, [tipo_id])
            row = cursor.fetchone()

        last_version = row[0] if row else None
        nueva_version = siguiente_version(last_version)

        with connection.cursor() as cursor:
            cursor.execute("""
                INSERT INTO plantillas_documentos_tecnicos (tipo_documento_id, gcs_path, version)
                VALUES (%s, %s, %s)
            """, [tipo_id, blob_name, nueva_version])

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
    # 1. Obtener TODAS las versiones
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT 
                pu.id,
                pu.nombre,
                pu.gcs_path,
                pu.version,
                pu.creado_en,
                fa.extension,
                fa.mime_type
            FROM plantillas_utilidad pu
            LEFT JOIN formato_archivo fa ON fa.id = pu.formato_id
            WHERE pu.tipo_id = 1 AND pu.formato_id = 1
            ORDER BY pu.id DESC
        """)
        versiones = dictfetchall(cursor)

    plantilla = versiones[0] if versiones else None

    preview_url = None
    controles = []

    # 2. Generar controls de contenido para la versión ACTUAL
    if plantilla:
        preview_url = generar_url_previa(plantilla["gcs_path"])
        controles = extraer_controles_contenido_desde_gcs(plantilla["gcs_path"])

    office_url = office_or_download_url(preview_url) if preview_url else None

    # 3. Generar office_url para cada versión historica
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
        # 1) Buscar portada existente
        # ===========================
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT id, gcs_path, version
                FROM plantillas_utilidad
                WHERE tipo_id = 1 AND formato_id = 1
                ORDER BY id DESC
                LIMIT 1
            """)
            existente = dictfetchone(cursor)

        controles_antes = []
        version_anterior = None

        if existente:
            version_anterior = existente["version"]
            # extraer controles anteriores
            controles_antes = extraer_controles_contenido_desde_gcs(existente["gcs_path"])

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
        bucket.blob(version_folder).upload_from_string("")  # crea carpeta vacía

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
        # 7) Insertar o actualizar BD
        # ===========================
        if existente:
            with connection.cursor() as cursor:
                cursor.execute("""
                    UPDATE plantillas_utilidad
                    SET gcs_path = %s,
                        version = %s,
                        creado_en = NOW()
                    WHERE id = %s
                """, [blob_name, nueva_version, existente["id"]])

        else:
            with connection.cursor() as cursor:
                cursor.execute("""
                    INSERT INTO plantillas_utilidad (tipo_id, formato_id, nombre, gcs_path, version)
                    VALUES (1, 1, 'Portada Word', %s, %s)
                """, [blob_name, nueva_version])

        messages.success(request, f"✔ Portada Word actualizada (versión {nueva_version})")
        return redirect("plantillas:portada_word_detalle")

    return render(request, "subir_portada_word.html")



# =============================================================================
# ======================   PORTADA   EXCEL   ======================
# =============================================================================

@login_required
def portada_excel_detalle(request):

    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT 
                pu.id,
                pu.nombre,
                pu.gcs_path,
                pu.version,
                pu.creado_en,
                fa.extension,
                fa.mime_type,
                tpu.nombre AS tipo_nombre
            FROM plantillas_utilidad pu
            LEFT JOIN formato_archivo fa ON fa.id = pu.formato_id
            LEFT JOIN tipo_plantilla_utilidad tpu ON tpu.id = pu.tipo_id
            WHERE pu.tipo_id = 1 AND pu.formato_id = 2
            ORDER BY pu.id DESC
            LIMIT 1
        """)
        plantilla = dictfetchone(cursor)

    etiquetas = []

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

        path = "Plantillas/Utilidad/Portada/Excel/"
        blob_name = f"{path}{archivo.name}"

        client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
        bucket = client.bucket(settings.GCP_BUCKET_NAME)
        blob = bucket.blob(blob_name)
        blob.upload_from_file(archivo)

        # Buscar si ya existe una portada Excel
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT id, version
                FROM plantillas_utilidad
                WHERE tipo_id = 1 AND formato_id = 2
                ORDER BY id DESC
                LIMIT 1
            """)
            existente = dictfetchone(cursor)

        if existente:
            nuevo_id = existente["id"]
            nueva_version = siguiente_version(existente["version"])

            with connection.cursor() as cursor:
                cursor.execute("""
                    UPDATE plantillas_utilidad
                    SET gcs_path = %s,
                        version = %s,
                        creado_en = NOW()
                    WHERE id = %s
                """, [blob_name, nueva_version, nuevo_id])

        else:
            nueva_version = "1.0"
            with connection.cursor() as cursor:
                cursor.execute("""
                INSERT INTO plantillas_utilidad (tipo_id, formato_id, nombre, gcs_path, version)
                VALUES (1, 2, 'Portada Excel', %s, %s)
                    RETURNING id
                """, [blob_name, nueva_version])
                nuevo_id = cursor.fetchone()[0]

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
# EXTRACCIÓN DE CONTROLES (w:sdt) DESDE GCS
# =============================================================================

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def extraer_controles_contenido_desde_gcs(gcs_path):
    """
    Descarga un DOCX desde GCS y extrae TODOS los controles de contenido (w:sdt),
    devolviendo una lista única ordenada.
    """

    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
    bucket = client.bucket(settings.GCP_BUCKET_NAME)
    blob = bucket.blob(gcs_path)

    # 1) Validar que el archivo exista en GCS
    if not blob.exists():
        return []

    # 2) Descargar archivo temporal
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        blob.download_to_filename(tmp.name)
        local_path = tmp.name

    # 3) Validar tamaño mínimo (archivo vacío → plantilla rota o no subida)
    if os.path.getsize(local_path) < 50:     # puedes bajar o subir el umbral
        return []

    controles = set()

    # 4) Intentar abrir como zip
    try:
        with ZipFile(local_path, "r") as docx:

            # --- Cuerpo ---
            if "word/document.xml" in docx.namelist():
                xml_doc = etree.fromstring(docx.read("word/document.xml"))
                for sdt in xml_doc.xpath("//w:sdt", namespaces=NS):
                    alias = _extraer_alias(sdt)
                    if alias:
                        controles.add(alias)

            # --- Encabezados ---
            for name in docx.namelist():
                if name.startswith("word/header") and name.endswith(".xml"):
                    xml_h = etree.fromstring(docx.read(name))
                    for sdt in xml_h.xpath("//w:sdt", namespaces=NS):
                        alias = _extraer_alias(sdt)
                        if alias:
                            controles.add(alias)

            # --- Pies ---
            for name in docx.namelist():
                if name.startswith("word/footer") and name.endswith(".xml"):
                    xml_f = etree.fromstring(docx.read(name))
                    for sdt in xml_f.xpath("//w:sdt", namespaces=NS):
                        alias = _extraer_alias(sdt)
                        if alias:
                            controles.add(alias)

    except BadZipFile:
        # archivo no válido → devolver vacío en vez de crashear
        return []
    except Exception:
        return []

    return sorted(controles)



def _extraer_alias(sdt_node):
    """Obtiene el alias de un control de contenido w:sdt."""
    alias_node = sdt_node.xpath(".//w:alias", namespaces=NS)
    if alias_node:
        return alias_node[0].get("{%s}val" % NS["w"])
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

        # ===== 2) Buscar plantilla anterior =====
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT id, gcs_path, version
                FROM plantillas_documentos_tecnicos
                WHERE tipo_documento_id = %s
                ORDER BY id DESC LIMIT 1
            """, [tipo_id])
            anterior = dictfetchone(cursor)

        # ==== protección contra NULL en gcs_path ====
        controles_antes = []
        version_anterior = None

        if anterior:
            version_anterior = anterior["version"]
            ruta_anterior = anterior["gcs_path"]

            if ruta_anterior:
                controles_antes = extraer_controles_contenido_desde_gcs(ruta_anterior)
            else:
                controles_antes = []

        # ===== 3) Guardar archivo temporal =====
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            for chunk in archivo.chunks():
                tmp.write(chunk)
            nuevo_local = tmp.name

        # Controles nuevos
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
                INSERT INTO plantillas_documentos_tecnicos (tipo_documento_id, gcs_path, version)
                VALUES (%s, %s, %s)
            """, [tipo_id, blob_name, nueva_version])

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
    Si Office no puede abrirlo (Chrome, Brave, etc), Office lo descarga.
    """
    encoded = urllib.parse.quote(preview_url, safe='')
    return f"https://view.officeapps.live.com/op/view.aspx?src={encoded}"



def versionar_plantilla(version_actual, controles_antes, controles_despues):
    """
    Sistema de versionado corregido:
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



def extraer_controles_contenido_desde_file(local_path):
    controles = set()
    try:
        with ZipFile(local_path, "r") as docx:

            if "word/document.xml" in docx.namelist():
                xml_doc = etree.fromstring(docx.read("word/document.xml"))
                for sdt in xml_doc.xpath("//w:sdt", namespaces=NS):
                    alias = _extraer_alias(sdt)
                    if alias:
                        controles.add(alias)

            for name in docx.namelist():
                if name.startswith("word/header") and name.endswith(".xml"):
                    xml_h = etree.fromstring(docx.read(name))
                    for sdt in xml_h.xpath("//w:sdt", namespaces=NS):
                        alias = _extraer_alias(sdt)
                        if alias:
                            controles.add(alias)

            for name in docx.namelist():
                if name.startswith("word/footer") and name.endswith(".xml"):
                    xml_f = etree.fromstring(docx.read(name))
                    for sdt in xml_f.xpath("//w:sdt", namespaces=NS):
                        alias = _extraer_alias(sdt)
                        if alias:
                            controles.add(alias)

    except:
        return []

    return sorted(controles)



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
    # 4) Eliminar registros en BD
    # ============================
    with connection.cursor() as cursor:
        cursor.execute("""
            DELETE FROM plantillas_documentos_tecnicos
            WHERE tipo_documento_id = %s
        """, [tipo_id])

    messages.success(request, "🗑️ Todas las versiones de la plantilla fueron eliminadas correctamente.")

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
                        UPDATE plantillas_documentos_tecnicos
                        SET gcs_path = REPLACE(gcs_path, %s, %s)
                        WHERE tipo_documento_id = %s
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

    # 1) Obtener registro de la versión
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id, tipo_documento_id, gcs_path
            FROM plantillas_documentos_tecnicos
            WHERE id = %s
        """, [version_id])
        reg = dictfetchone(cursor)

    if not reg:
        messages.error(request, "Versión no encontrada.")
        return redirect("plantillas:lista_plantillas")

    tipo_id = reg["tipo_documento_id"]
    gcs_path = reg["gcs_path"]

    # 2) Extraer carpetas desde gcs_path
    # Ej: Plantillas/.../TipoDocumento/V1.0/archivo.docx
    version_folder = "/".join(gcs_path.split("/")[:-1]) + "/"
    # Ej: Plantillas/.../TipoDocumento/V1.0/

    base_folder = "/".join(version_folder.split("/")[:-2]) + "/"
    # Ej: Plantillas/.../TipoDocumento/

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
        cursor.execute("DELETE FROM plantillas_documentos_tecnicos WHERE id = %s", [version_id])

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

        # Lista ahora vacía
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

    Ej:
      old_prefix = Plantillas/Documentos_Tecnicos/MECANICA/INFORME/
      new_prefix = Plantillas/Documentos_Tecnicos/MECÁNICA_NUEVA/INFORME_NUEVO/
    """
    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
    bucket = client.bucket(settings.GCP_BUCKET_NAME)

    blobs = list(bucket.list_blobs(prefix=old_prefix))

    for b in blobs:
        old_name = b.name
        new_name = old_name.replace(old_prefix, new_prefix, 1)
        # copiar en la nueva ruta
        bucket.copy_blob(b, bucket, new_name)
        # eliminar antiguo
        b.delete()

    return True




@login_required
def eliminar_version_portada(request, version_id):

    # 1. Buscar versión en BD
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id, gcs_path
            FROM plantillas_utilidad
            WHERE id = %s
        """, [version_id])
        reg = dictfetchone(cursor)

    if not reg:
        messages.error(request, "Versión no encontrada.")
        return redirect("plantillas:portada_word_detalle")

    gcs_path = reg["gcs_path"]

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
        cursor.execute("DELETE FROM plantillas_utilidad WHERE id = %s", [version_id])

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
