# C:\Users\jonat\Documents\gestion_docs\plantillas_documentos_tecnicos\views.py
from django.http import HttpResponse, JsonResponse, Http404
from django.contrib.auth.decorators import login_required
from django.shortcuts import render, redirect
from django.db import connection
from django.contrib import messages
from django.conf import settings

import urllib.parse
import re
import os
import tempfile
from zipfile import ZipFile
from datetime import timedelta
import json

from google.cloud import storage
from unidecode import unidecode
from lxml import etree


# =============================================================================
# UTILIDADES
# =============================================================================
# --- AJAX: detectar controles sin subir la plantilla ---

from django.views.decorators.http import require_POST
from django.views.decorators.csrf import csrf_exempt
import tempfile, os

@csrf_exempt
@require_POST
def detectar_controles_ajax(request):
    file = request.FILES.get("archivo")
    if not file:
        return JsonResponse({"ok": False, "error": "No se recibió archivo"})

    if not file.name.lower().endswith(".docx"):
        return JsonResponse({"ok": False, "error": "Formato inválido"})

    # Guardar temporalmente
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        for chunk in file.chunks():
            tmp.write(chunk)
        local_path = tmp.name

    # EXTRAER CONTROLES (ya tienes esta función)
    controles = extraer_controles_contenido_desde_file(local_path)

    os.remove(local_path)

    return JsonResponse({
        "ok": True,
        "controles": controles
    })


def gcs_exists(path):
    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
    bucket = client.bucket(settings.GCP_BUCKET_NAME)
    blob = bucket.blob(path)
    return blob.exists()

def calcular_stats_versiones(versiones):
    """
    versiones = lista de dicts: {id, gcs_path, version, creado_en}
    """
    total = len(versiones)
    disponibles = 0
    rotas = 0

    for v in versiones:
        if gcs_exists(v["gcs_path"]):
            disponibles += 1
        else:
            rotas += 1

    return {
        "total_versiones": total,
        "versiones_ok": disponibles,
        "versiones_rotas": rotas,
    }
def calcular_stats_controles(lista_controles):
    """
    lista_controles = controles detectados en versión ACTUAL (lista simple)
    """
    total = len(lista_controles)
    unicos = len(set(lista_controles))

    return {
        "controles_total": total,
        "controles_unicos": unicos,
    }
def evaluar_calidad_controles(unicos):
    if unicos >= 10:
        return "alta"
    if unicos >= 3:
        return "media"
    return "baja"
def comparar_controles(controles_actual, controles_anterior):
    set_actual = set(controles_actual or [])
    set_anterior = set(controles_anterior or [])

    nuevos = list(set_actual - set_anterior)
    eliminados = list(set_anterior - set_actual)

    return {
        "controles_nuevos": sorted(nuevos),
        "controles_eliminados": sorted(eliminados),
    }

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
    return unidecode(texto).replace(" ", "_").replace("/", "_")


def siguiente_version(version_actual):
    if not version_actual:
        return "1.0"
    try:
        partes = str(version_actual).split(".")
        if len(partes) == 2:
            major = int(partes[0])
            minor = int(partes[1])
            return f"{major}.{minor + 1}"
        v = float(version_actual)
        return f"{v + 0.1:.1f}"
    except Exception:
        return "1.0"


# =============================================================================
# LISTADO GENERAL
# =============================================================================

def lista_plantillas(request):
    categorias = []

    # ======================================================
    #      ARMAR LISTA COMPLETA DE CATEGORÍAS Y TIPOS
    # ======================================================
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id, nombre, descripcion, abreviatura
            FROM categoria_documentos_tecnicos
            ORDER BY id ASC
        """)
        categorias_data = cursor.fetchall()

        for cat_id, cat_nombre, cat_desc, cat_abrev in categorias_data:

            cursor.execute("""
                SELECT 
                    t.id,
                    t.nombre,
                    t.abreviatura,
                    t.descripcion,

                    EXISTS (
                        SELECT 1
                        FROM plantilla_tipo_doc p
                        WHERE p.tipo_documento_id = t.id
                    ) AS tiene_plantilla

                FROM tipo_documentos_tecnicos t
                WHERE t.categoria_id = %s
                ORDER BY t.nombre
            """, [cat_id])

            tipos_raw = dictfetchall(cursor)

            categorias.append({
                "id": cat_id,
                "nombre": cat_nombre,
                "abreviatura": cat_abrev,
                "descripcion": cat_desc,
                "tipos": tipos_raw
            })

    # ======================================================
    #      ESTADÍSTICAS GLOBALES DEL SISTEMA
    # ======================================================

    total_tipos = 0
    con_plantilla = 0
    sin_plantilla = 0
    rotas = 0

    with connection.cursor() as cursor:

        # N° total de tipos
        cursor.execute("SELECT id FROM tipo_documentos_tecnicos")
        tipos_all = [r[0] for r in cursor.fetchall()]
        total_tipos = len(tipos_all)

        # Tipos con plantilla
        cursor.execute("SELECT tipo_documento_id FROM plantilla_tipo_doc")
        p = [r[0] for r in cursor.fetchall()]
        con_plantilla = len(p)

        sin_plantilla = total_tipos - con_plantilla

        # Plantillas rotas
        cursor.execute("""
            SELECT v.gcs_path
            FROM plantilla_tipo_doc_versiones v
        """)
        paths = [r[0] for r in cursor.fetchall()]

        for ruta in paths:
            if not gcs_exists(ruta):
                rotas += 1

    stats_globales = {
        "total_tipos": total_tipos,
        "con_plantilla": con_plantilla,
        "sin_plantilla": sin_plantilla,
        "plantillas_rotas": rotas,
    }

    # ======================================================
    # RENDER
    # ======================================================
    return render(request, "lista_plantillas.html", {
        "categorias": categorias,
        "stats_globales": stats_globales,
    })

# =============================================================================
# DETALLES DE CATEGORÍA
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

    nombre_original = categoria["nombre"]
    nombre_original_clean = clean(nombre_original)

    if request.method == "POST":
        nombre = request.POST.get("nombre", "").strip()
        descripcion = request.POST.get("descripcion", "").strip()
        abreviatura = request.POST.get("abreviatura", "").strip().upper()

        if not nombre:
            messages.error(request, "El nombre es obligatorio.")
            return redirect(request.path)

        nuevo_clean = clean(nombre)

        # mover carpetas si cambia nombre
        if nuevo_clean != nombre_original_clean:
            old_prefix = f"Plantillas/Documentos_Tecnicos/{nombre_original_clean}/"
            new_prefix = f"Plantillas/Documentos_Tecnicos/{nuevo_clean}/"

            try:
                mover_carpeta_gcs(old_prefix, new_prefix)
                with connection.cursor() as cursor:
                    cursor.execute("""
                        UPDATE plantilla_tipo_doc_versiones
                        SET gcs_path = REPLACE(gcs_path, %s, %s)
                        WHERE gcs_path LIKE %s
                    """, [old_prefix, new_prefix, old_prefix + "%"])
            except Exception as e:
                messages.error(request, f"Error renombrando carpeta en GCS: {e}")
                return redirect("plantillas:detalle_categoria", categoria_id=categoria_id)

        with connection.cursor() as cursor:
            cursor.execute("""
                UPDATE categoria_documentos_tecnicos
                SET nombre = %s, descripcion = %s, abreviatura = %s
                WHERE id = %s
            """, [nombre, descripcion, abreviatura, categoria_id])

        messages.success(request, "✔ Categoría actualizada.")
        return redirect("plantillas:detalle_categoria", categoria_id=categoria_id)

    return render(request, "editar_categoria.html", {"categoria": categoria})


# =============================================================================
# DETALLE TIPO
# =============================================================================

def tipo_detalle(request, tipo_id):

    with connection.cursor() as cursor:
        # Datos del tipo
        cursor.execute("""
            SELECT 
                t.id, t.categoria_id, t.nombre, t.descripcion,
                t.abreviatura, t.formato_id,
                c.nombre AS categoria_nombre
            FROM tipo_documentos_tecnicos t
            JOIN categoria_documentos_tecnicos c ON c.id = t.categoria_id
            WHERE t.id = %s
        """, [tipo_id])
        tipo = dictfetchone(cursor)

        if not tipo:
            return render(request, "404.html", status=404)

        # Plantilla actual (si existe)
        cursor.execute("""
            SELECT 
                v.id, v.plantilla_id, v.gcs_path, v.version, v.creado_en
            FROM plantilla_tipo_doc p
            JOIN plantilla_tipo_doc_versiones v
                ON v.id = p.version_actual_id
            WHERE p.tipo_documento_id = %s
            LIMIT 1
        """, [tipo_id])
        plantilla = dictfetchone(cursor)

        # Historial completo
        if plantilla:
            cursor.execute("""
                SELECT id, gcs_path, version, creado_en
                FROM plantilla_tipo_doc_versiones
                WHERE plantilla_id = %s
                ORDER BY creado_en DESC, id DESC
            """, [plantilla["plantilla_id"]])
            versiones = dictfetchall(cursor)
        else:
            versiones = []

    # ======================================================
    #      EXISTENCIA ARCHIVO + CONTROLES DE CONTENIDO
    # ======================================================

    archivo_existe = False
    preview_url = None
    controles = []

    if plantilla:
        ruta = plantilla["gcs_path"]
        if gcs_exists(ruta):
            archivo_existe = True
            preview_url = generar_url_previa(ruta)
            controles = extraer_controles_contenido_desde_gcs(ruta)
        else:
            plantilla = None  # Archivo perdido → eliminar visual

    office_url = office_or_download_url(preview_url) if preview_url else None

    # ======================================================
    #      PROCESAR VERSIONES PARA TABLA
    # ======================================================
    versiones_proc = []
    for v in versiones:
        entrada = v.copy()
        gp = v.get("gcs_path")

        if gp and gcs_exists(gp):
            prev = generar_url_previa(gp)
            entrada["office_url"] = office_or_download_url(prev) if prev else None
        else:
            entrada["office_url"] = None

        # 🔥 NUEVO: saber si está en uso
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT COUNT(*)
                FROM documentos_generados
                WHERE plantilla_documento_tecnico_version_id = %s
            """, [v["id"]])
            entrada["en_uso"] = cursor.fetchone()[0] > 0

        versiones_proc.append(entrada)

    # ======================================================
    #      ESTADÍSTICAS
    # ======================================================

    # A) Stats de versiones
    if versiones:
        stats_versiones = calcular_stats_versiones(versiones)
    else:
        stats_versiones = {
            "total_versiones": 0,
            "versiones_ok": 0,
            "versiones_rotas": 0
        }

    # B) Stats de controles
    stats_controles = calcular_stats_controles(controles)

    # C) Diferencias con penúltima versión
    controles_anterior = []
    if len(versiones) >= 2:
        penult = versiones[1]
        if penult["gcs_path"] and gcs_exists(penult["gcs_path"]):
            controles_anterior = extraer_controles_contenido_desde_gcs(penult["gcs_path"])

    stats_diferencias = comparar_controles(controles, controles_anterior)

    # D) Calidad estructural
    calidad = evaluar_calidad_controles(stats_controles["controles_unicos"])

    # ======================================================
    # RENDER
    # ======================================================
    return render(request, "tipo_detalle.html", {
        "tipo": tipo,
        "plantilla": plantilla,
        "archivo_existe": archivo_existe,
        "preview_url": preview_url,
        "office_url": office_url,

        "controles": controles,
        "versiones": versiones_proc,

        # --- NUEVAS ESTADÍSTICAS ---
        "stats_versiones": stats_versiones,
        "stats_controles": stats_controles,
        "stats_diferencias": stats_diferencias,
        "calidad": calidad,
    })

# =============================================================================
# CREAR CATEGORÍA
# =============================================================================

def generar_abreviatura(nombre, tipo="categoria"):
    if not nombre:
        return ""
    limpio = re.sub(r"[^A-Za-zÁÉÍÓÚáéíóúÑñ\s]", "", nombre)
    palabras = [p for p in limpio.split() if p.lower() not in {
        "y", "de", "del", "la", "los", "las", "el", "en",
        "por", "para", "ltlda", "ltda", "sa", "empresa",
        "asociacion", "compania", "hermanos"
    }]
    if not palabras:
        return ""
    if len(palabras) == 1:
        return palabras[0][:3].upper() if tipo == "categoria" else palabras[0][:4].upper()
    return "".join(p[0].upper() for p in palabras[:4])


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
                    SELECT COUNT(*) FROM categoria_documentos_tecnicos
                    WHERE UPPER(abreviatura) = UPPER(%s)
                """, [abreviatura])
                existe = cursor.fetchone()[0]

            if existe > 0:
                messages.warning(request, "La abreviatura ya existe.")
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
        categorias = dictfetchall(cursor)

    if request.method == "POST":

        nombre = request.POST.get("nombre", "").strip()
        descripcion = request.POST.get("descripcion", "").strip()
        categoria_id = request.POST.get("categoria_id")
        abreviatura_manual = request.POST.get("abreviatura", "").strip().upper()

        if not categoria_id or not nombre:
            messages.error(request, "Nombre y categoría obligatorios.")
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
# SUBIR PLANTILLA (CUERPO)
# =============================================================================

@login_required
def subir_plantilla(request, tipo_id):

    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT t.id, t.nombre, c.nombre as categoria
            FROM tipo_documentos_tecnicos t
            JOIN categoria_documentos_tecnicos c ON c.id = t.categoria_id
            WHERE t.id = %s
        """, [tipo_id])
        tipo = dictfetchone(cursor)

    if not tipo:
        messages.error(request, "Tipo no encontrado.")
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

        # buscar registro maestro
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
            with connection.cursor() as cursor:
                cursor.execute("""
                    INSERT INTO plantilla_tipo_doc (tipo_documento_id)
                    VALUES (%s)
                    RETURNING id
                """, [tipo_id])
                plantilla_id = cursor.fetchone()[0]
            version_anterior = None
            controles_antes = []

        # guardar archivo temporal
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            for chunk in archivo.chunks():
                tmp.write(chunk)
            nuevo_local = tmp.name

        controles_despues = extraer_controles_contenido_desde_file(nuevo_local)

        # versionado
        nueva_version = versionar_plantilla(version_anterior, controles_antes, controles_despues)

        version_folder = f"{base_path}V{nueva_version}/"
        bucket.blob(version_folder).upload_from_string("")

        filename = archivo.name
        blob_name = f"{version_folder}{filename}"

        bucket.blob(blob_name).upload_from_filename(nuevo_local)

        # insertar versión
        with connection.cursor() as cursor:
            cursor.execute("""
                INSERT INTO plantilla_tipo_doc_versiones (plantilla_id, version, gcs_path, controles)
                VALUES (%s, %s, %s, %s)
                RETURNING id
            """, [plantilla_id, nueva_version, blob_name, json.dumps(controles_despues)])
            new_version_id = cursor.fetchone()[0]

            cursor.execute("""
                UPDATE plantilla_tipo_doc
                SET version_actual_id = %s, actualizado_en = NOW()
                WHERE id = %s
            """, [new_version_id, plantilla_id])

        messages.success(request, f"✔ Plantilla subida (versión {nueva_version}).")
        return redirect("plantillas:detalle_tipo", tipo_id=tipo_id)

    return render(request, "subir_plantilla.html", {"tipo": tipo})


# =============================================================================
# DESCARGAR GCS
# =============================================================================

@login_required
def descargar_gcs(request, path):
    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
    bucket = client.bucket(settings.GCP_BUCKET_NAME)
    blob = bucket.blob(path)

    if not blob.exists():
        raise Http404("Archivo no encontrado.")

    contenido = blob.download_as_bytes()
    response = HttpResponse(contenido, content_type="application/octet-stream")
    response['Content-Disposition'] = f'attachment; filename="{os.path.basename(path)}"'
    return response


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


def office_or_download_url(preview_url):
    encoded = urllib.parse.quote(preview_url, safe='')
    return f"https://view.officeapps.live.com/op/view.aspx?src={encoded}"


# =============================================================================
# VERSIONAMIENTO
# =============================================================================

def versionar_plantilla(version_actual, controles_antes, controles_despues):

    if not version_actual:
        return "1.0"

    try:
        partes = str(version_actual).split(".")
        major = int(partes[0])
        minor = int(partes[1])
    except:
        major, minor = 1, 0

    set_antes = set(controles_antes or [])
    set_despues = set(controles_despues or [])

    if set_antes != set_despues:
        return f"{major + 1}.0"

    return f"{major}.{minor + 1}"



# =============================================================================
# EDITAR TIPO DOCUMENTO
# =============================================================================

@login_required
def editar_tipo_documento(request, tipo_id):

    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT 
                t.id, t.categoria_id, t.nombre, t.descripcion,
                t.abreviatura, t.formato_id,
                c.nombre AS categoria_nombre
            FROM tipo_documentos_tecnicos t
            JOIN categoria_documentos_tecnicos c ON c.id = t.categoria_id
            WHERE t.id = %s
        """, [tipo_id])
        tipo = dictfetchone(cursor)

        cursor.execute("SELECT id, nombre FROM categoria_documentos_tecnicos ORDER BY nombre")
        categorias = dictfetchall(cursor)

        cursor.execute("SELECT id, nombre, extension FROM formato_archivo ORDER BY id")
        formatos = dictfetchall(cursor)

    if not tipo:
        messages.error(request, "Tipo no encontrado.")
        return redirect("plantillas:lista_plantillas")

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
            messages.error(request, "Nombre y categoría obligatorios.")
            return redirect(request.path)

        with connection.cursor() as cursor:
            cursor.execute("SELECT nombre FROM categoria_documentos_tecnicos WHERE id = %s", [categoria_id])
            row = cursor.fetchone()

        if not row:
            messages.error(request, "Categoría inválida.")
            return redirect(request.path)

        categoria_nueva_nombre = row[0]
        categoria_nueva_clean = clean(categoria_nueva_nombre)
        tipo_nuevo_clean = clean(nombre)

        old_prefix = f"Plantillas/Documentos_Tecnicos/{categoria_original_clean}/{tipo_original_clean}/"
        new_prefix = f"Plantillas/Documentos_Tecnicos/{categoria_nueva_clean}/{tipo_nuevo_clean}/"

        if new_prefix != old_prefix:
            try:
                mover_carpeta_gcs(old_prefix, new_prefix)

                with connection.cursor() as cursor:
                    cursor.execute("""
                        UPDATE plantilla_tipo_doc_versiones
                        SET gcs_path = REPLACE(gcs_path, %s, %s)
                        WHERE plantilla_id IN (
                            SELECT id FROM plantilla_tipo_doc WHERE tipo_documento_id = %s
                        )
                        AND gcs_path LIKE %s
                    """, [old_prefix, new_prefix, tipo_id, old_prefix + "%"])
            except Exception as e:
                messages.error(request, f"Error al renombrar en GCS: {e}")
                return redirect("plantillas:detalle_tipo", tipo_id=tipo_id)

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


# =============================================================================
# ELIMINAR VERSIÓN INDIVIDUAL
# =============================================================================
@login_required
def eliminar_version(request, version_id):

    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT 
                v.id, v.gcs_path, v.plantilla_id,
                p.tipo_documento_id, p.version_actual_id
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

    # Carpetas GCS
    version_folder = "/".join(gcs_path.split("/")[:-1]) + "/"
    base_folder = "/".join(version_folder.split("/")[:-2]) + "/"

    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)

    # 0) Verificar si esta versión está siendo usada en algún requerimiento técnico
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT COUNT(*) 
            FROM documentos_generados 
            WHERE plantilla_documento_tecnico_version_id = %s
        """, [version_id])
        usada_en_requerimientos = cursor.fetchone()[0]

    if usada_en_requerimientos > 0:
        messages.error(
            request,
            "❌ No se puede eliminar esta versión porque está siendo utilizada "
            "por uno o más requerimientos técnicos."
        )
        return redirect("plantillas:detalle_tipo", tipo_id=tipo_id)



    bucket = client.bucket(settings.GCP_BUCKET_NAME)

    # 1) Borrar archivos de la versión en GCS
    for blob in list(bucket.list_blobs(prefix=version_folder)):
        try:
            blob.delete()
        except:
            pass

    with connection.cursor() as cursor:

        # 2) Borrar versión en BD
        cursor.execute("DELETE FROM plantilla_tipo_doc_versiones WHERE id = %s", [version_id])

        # 3) Ver cuántas versiones quedan
        cursor.execute("""
            SELECT COUNT(*)
            FROM plantilla_tipo_doc_versiones
            WHERE plantilla_id = %s
        """, [plantilla_id])
        quedan = cursor.fetchone()[0]

        # ========================================================
        #     ⚠⚠ REGLA TUYA: SOLO eliminar plantilla_tipo_doc
        #       cuando NO QUEDA NINGUNA versión (quedan = 0)
        # ========================================================
        if quedan == 0:

            # Borrar carpeta padre (Vacía)
            for blob in list(bucket.list_blobs(prefix=base_folder)):
                try:
                    blob.delete()
                except:
                    pass

            # Eliminar registro maestro
            cursor.execute("""
                DELETE FROM plantilla_tipo_doc
                WHERE id = %s
            """, [plantilla_id])

            messages.success(
                request,
                "✔ No quedan versiones: plantilla eliminada completamente."
            )
            return redirect("plantillas:lista_plantillas")

        # 4) Si aún quedan versiones → recalcular versión actual
        if version_id == version_actual_id:
            cursor.execute("""
                SELECT id
                FROM plantilla_tipo_doc_versiones
                WHERE plantilla_id = %s
                ORDER BY version DESC, creado_en DESC, id DESC
                LIMIT 1
            """, [plantilla_id])
            nueva = cursor.fetchone()
            nuevo_id = nueva[0] if nueva else None

            cursor.execute("""
                UPDATE plantilla_tipo_doc
                SET version_actual_id = %s
                WHERE id = %s
            """, [nuevo_id, plantilla_id])

    messages.success(request, "✔ Versión eliminada.")
    return redirect("plantillas:detalle_tipo", tipo_id=tipo_id)


# =============================================================================
# MOVER CARPETA EN GCS
# =============================================================================

def mover_carpeta_gcs(old_prefix, new_prefix):
    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
    bucket = client.bucket(settings.GCP_BUCKET_NAME)

    blobs = list(bucket.list_blobs(prefix=old_prefix))
    for b in blobs:
        old_name = b.name
        new_name = old_name.replace(old_prefix, new_prefix, 1)
        bucket.copy_blob(b, bucket, new_name)
        b.delete()

    return True


# =============================================================================
# EXTRACCIÓN DE CONTROLES CONTENIDO DOCX (CUERPO)
# =============================================================================

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def _normalizar_control(nombre):
    if not nombre:
        return None
    s = str(nombre)
    s = s.replace("\u00A0", " ")
    s = unidecode(s)
    s = s.strip()
    s = re.sub(r"\s+", "_", s)
    s = s.lower()
    s = re.sub(r"[^a-z0-9_.]+", "_", s)
    s = re.sub(r"_+", "_", s)
    return s.strip("_") or None


def _extraer_tag_o_alias(sdt_node):
    tag_node = sdt_node.xpath(".//w:tag", namespaces=NS)
    if tag_node:
        val = tag_node[0].get(f"{{{NS['w']}}}val")
        if val:
            return str(val).strip()

    alias_node = sdt_node.xpath(".//w:alias", namespaces=NS)
    if alias_node:
        val = alias_node[0].get(f"{{{NS['w']}}}val")
        if val:
            return str(val).strip()

    return None


def _agregar_control(mapa, nombre_crudo):
    norm = _normalizar_control(nombre_crudo)
    if not norm:
        return
    if norm not in mapa:
        mapa[norm] = nombre_crudo.strip()


def _procesar_docx_zip(docx, mapa, log_prefix=""):

    def _scan(part, etiqueta="document"):
        try:
            xml = etree.fromstring(docx.read(part))
            encontrados = xml.xpath("//w:sdt", namespaces=NS)
            for sdt in encontrados:
                raw = _extraer_tag_o_alias(sdt)
                if raw:
                    _agregar_control(mapa, raw)
        except Exception:
            pass

    # Documento principal
    if "word/document.xml" in docx.namelist():
        _scan("word/document.xml", "document")

    # Encabezados
    for name in docx.namelist():
        if name.startswith("word/header") and name.endswith(".xml"):
            _scan(name, "header")

    # Pies de página
    for name in docx.namelist():
        if name.startswith("word/footer") and name.endswith(".xml"):
            _scan(name, "footer")


def extraer_controles_contenido_desde_gcs(gcs_path):
    """
    Descarga DOCX desde GCS y extrae controles w:sdt.
    Devuelve lista sin duplicados.
    """
    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
    bucket = client.bucket(settings.GCP_BUCKET_NAME)
    blob = bucket.blob(gcs_path)

    if not blob.exists():
        return []

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        blob.download_to_filename(tmp.name)
        local_path = tmp.name

    if not os.path.exists(local_path):
        return []

    mapa = {}

    try:
        with ZipFile(local_path, "r") as docx:
            _procesar_docx_zip(docx, mapa, log_prefix="[GCS] ")
    except Exception:
        return []

    return sorted(mapa.values(), key=lambda x: x.lower())


def extraer_controles_contenido_desde_file(local_path):
    """
    Extrae controles desde un DOCX local.
    """
    if not os.path.exists(local_path):
        return []

    mapa = {}

    try:
        with ZipFile(local_path, "r") as docx:
            _procesar_docx_zip(docx, mapa, log_prefix="[LOCAL] ")
    except Exception:
        return []

    return sorted(mapa.values(), key=lambda x: x.lower())
