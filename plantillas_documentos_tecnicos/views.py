# plantillas_documentos_tecnicos/views.py

from django.http import HttpResponse, JsonResponse, Http404
from django.contrib.auth.decorators import login_required
from django.shortcuts import render, redirect
from django.db import connection
from django.contrib import messages
from django.conf import settings

from psycopg.types.json import Json

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

from django.views.decorators.http import require_POST
from django.views.decorators.csrf import csrf_exempt


# =============================================================================
# UTILIDADES BÁSICAS
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
    """
    Normaliza nombres para usarlos como parte de rutas en GCS.
    """
    return unidecode(texto).replace(" ", "_").replace("/", "_")


def gcs_exists(path: str) -> bool:
    """
    Verifica existencia de un blob en GCS.
    """
    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
    bucket = client.bucket(settings.GCP_BUCKET_NAME)
    blob = bucket.blob(path)
    return blob.exists()


def generar_url_previa(blob_path: str) -> str | None:
    """
    Genera URL temporal firmada (3 horas) para visualizar/descargar el archivo.
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


def office_or_download_url(preview_url: str) -> str:
    """
    Envuelve una URL firmada de GCS dentro del visor de Office.
    """
    encoded = urllib.parse.quote(preview_url, safe='')
    return f"https://view.officeapps.live.com/op/view.aspx?src={encoded}"


def mover_carpeta_gcs(old_prefix: str, new_prefix: str) -> bool:
    """
    Mueve (copia + borra) todos los blobs cuyo nombre parte con old_prefix a new_prefix.
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


# =============================================================================
# UTILIDADES DE ESTADÍSTICAS / ESTRUCTURA
# =============================================================================

def calcular_stats_versiones(versiones):
    """
    versiones: lista de dicts {id, gcs_path, version, creado_en}
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
    lista_controles = lista simple de alias/tag de controles
    (Utilidad genérica, por ahora no usada en la UI principal)
    """
    total = len(lista_controles or [])
    unicos = len(set(lista_controles or []))

    return {
        "controles_total": total,
        "controles_unicos": unicos,
    }


def evaluar_calidad_estructura(estructura: dict | None) -> str:
    """
    Evalúa un JSON estructural completo de la plantilla Word.

    Criterio:
        - Alta: contiene controles, tablas Word y excels embebidos
        - Media: contiene solo una o dos de las anteriores
        - Baja: estructura mínima o inexistente
    """

    if not isinstance(estructura, dict):
        return "baja"

    bloques = 0

    # Claves coherentes con generar_estructura() en leer_estructura_plantilla_word.py
    if estructura.get("controles"):
        bloques += 1
    if estructura.get("tablas_word"):
        bloques += 1
    if estructura.get("excels"):
        bloques += 1

    if bloques >= 3:
        return "alta"
    if bloques == 2:
        return "media"
    return "baja"


def comparar_estructuras(estructura_actual, estructura_anterior):
    """
    Compara estructuras JSON profundas:
    - controles
    - tablas_word
    - excels
    - imagenes
    - firma estructural

    Devuelve una estructura ALINEADA con el template:
      stats_diferencias = {
        "controles": {"nuevos": [...], "eliminados": [...]},
        "tablas_word": {"nuevas": [...], "eliminadas": [...]},
        "excels": {"nuevos": [...], "eliminados": [...]},
        "imagenes": {"nuevas": [...], "eliminadas": [...]},
        "signature_cambios": True/False,
      }
    """

    # =============================
    # NORMALIZACIÓN ROBUSTA
    # =============================
    def normalize(e):
        if isinstance(e, dict):
            return e
        if isinstance(e, str):
            try:
                return json.loads(e)
            except Exception:
                return {}
        return {}

    ea = normalize(estructura_actual)
    eb = normalize(estructura_anterior)

    # ===================================================
    # FUNCIONES AUXILIARES SEGURO-ROBUSTAS
    # ===================================================

    def safe_list(value):
        """Siempre devuelve lista."""
        if isinstance(value, list):
            return value
        return []

    def hash_dict(d):
        """Convierte un dict → string hash estable."""
        if not isinstance(d, dict):
            return None
        try:
            return json.dumps(d, sort_keys=True)
        except Exception:
            return str(d)

    # ---------- CONTROLES ----------

    def get_aliases(e):
        """Extrae alias/tag de forma resiliente."""
        if not isinstance(e, dict):
            return []
        controles = e.get("controles", [])
        if isinstance(controles, str):
            return []
        if not isinstance(controles, list):
            return []

        out = []
        for item in controles:
            if isinstance(item, dict):
                alias = item.get("alias") or item.get("tag")
                if alias:
                    out.append(alias)
            elif isinstance(item, str):
                if item.strip():
                    out.append(item.strip())
        return out

    a_controles = set(get_aliases(ea))
    b_controles = set(get_aliases(eb))

    # ---------- TABLAS WORD ----------
    def key_tabla_word(t):
        """Hash para tabla_word (dict)."""
        return hash_dict(t)

    a_tablas = {
        key_tabla_word(t) for t in safe_list(ea.get("tablas_word"))
        if key_tabla_word(t)
    }
    b_tablas = {
        key_tabla_word(t) for t in safe_list(eb.get("tablas_word"))
        if key_tabla_word(t)
    }

    # ---------- EXCELS ----------
    def key_excel(e):
        return hash_dict(e)

    a_excels = {
        key_excel(x) for x in safe_list(ea.get("excels"))
        if key_excel(x)
    }
    b_excels = {
        key_excel(x) for x in safe_list(eb.get("excels"))
        if key_excel(x)
    }

    # ---------- IMÁGENES ----------
    def key_imagen(img):
        return hash_dict(img)

    a_imgs = {
        key_imagen(i) for i in safe_list(ea.get("imagenes"))
        if key_imagen(i)
    }
    b_imgs = {
        key_imagen(i) for i in safe_list(eb.get("imagenes"))
        if key_imagen(i)
    }

    # ---------- FIRMA ESTRUCTURAL ----------
    from plantillas_documentos_tecnicos.leer_estructura_plantilla_word import (
        extract_structural_signature
    )

    sig_a = extract_structural_signature(ea)
    sig_b = extract_structural_signature(eb)

    firma_cambiada = (
        json.dumps(sig_a, sort_keys=True) != json.dumps(sig_b, sort_keys=True)
    )

    # ===================================================
    # RESULTADO FINAL ALINEADO CON EL TEMPLATE
    # ===================================================
    return {
        "controles": {
            "nuevos": sorted(a_controles - b_controles),
            "eliminados": sorted(b_controles - a_controles),
        },
        "tablas_word": {
            "nuevas": sorted(a_tablas - b_tablas),
            "eliminadas": sorted(b_tablas - a_tablas),
        },
        "excels": {
            "nuevos": sorted(a_excels - b_excels),
            "eliminados": sorted(b_excels - a_excels),
        },
        "imagenes": {
            "nuevas": sorted(a_imgs - b_imgs),
            "eliminadas": sorted(b_imgs - a_imgs),
        },
        "signature_cambios": firma_cambiada,
    }


def extract_aliases_from_estructura(estructura) -> list[str]:
    """
    Extrae alias/tag de controles de una estructura JSON de plantilla.

    Soporta:
        - estructura = dict con clave "controles": [ {alias, tag, ...}, ... ]
        - estructuras antiguas donde "controles" pueda ser lista de strings
        - ignora valores corruptos sin lanzar excepción
    """
    aliases: list[str] = []

    if not isinstance(estructura, dict):
        return aliases

    lista = estructura.get("controles", [])
    if not isinstance(lista, list):
        return aliases

    for c in lista:
        if isinstance(c, dict):
            alias = c.get("alias") or c.get("tag")
            if alias:
                aliases.append(alias)
        elif isinstance(c, str):
            # Formato antiguo: ya es alias directamente
            if c:
                aliases.append(c)

    return aliases


# =============================================================================
# AJAX: DETECTAR CONTROLES SIN SUBIR PLANTILLA (preview)
# =============================================================================

@csrf_exempt
@require_POST
def detectar_controles_ajax(request):
    file = request.FILES.get("archivo")
    if not file:
        return JsonResponse({"ok": False, "error": "No se recibió archivo"})

    if not file.name.lower().endswith(".docx"):
        return JsonResponse({"ok": False, "error": "Formato inválido"})

    # Guardar temporal
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        for chunk in file.chunks():
            tmp.write(chunk)
        local_path = tmp.name

    try:
        from plantillas_documentos_tecnicos.leer_estructura_plantilla_word import generar_estructura
        estructura = generar_estructura(local_path)
    finally:
        if os.path.exists(local_path):
            os.remove(local_path)

    controles_alias = extract_aliases_from_estructura(estructura)

    return JsonResponse({
        "ok": True,
        "controles": sorted(controles_alias),
        "estructura_json": estructura,  # dict serializable
    })


# =============================================================================
# LISTADO GENERAL DE TIPOS
# =============================================================================

def lista_plantillas(request):
    categorias = []

    # -----------------------------
    # Lista de categorías + tipos
    # -----------------------------
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

    # -----------------------------
    # Estadísticas globales
    # -----------------------------
    total_tipos = 0
    con_plantilla = 0
    sin_plantilla = 0
    rotas = 0

    with connection.cursor() as cursor:

        cursor.execute("SELECT id FROM tipo_documentos_tecnicos")
        tipos_all = [r[0] for r in cursor.fetchall()]
        total_tipos = len(tipos_all)

        cursor.execute("SELECT tipo_documento_id FROM plantilla_tipo_doc")
        p = [r[0] for r in cursor.fetchall()]
        con_plantilla = len(p)

        sin_plantilla = total_tipos - con_plantilla

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

        # Renombrar carpeta en GCS si cambia el nombre
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
# DETALLE TIPO DOCUMENTO
# =============================================================================

def tipo_detalle(request, tipo_id):

    # ============================================================
    # 1) Obtener datos del tipo de documento y su plantilla actual
    # ============================================================
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

        if not tipo:
            return render(request, "404.html", status=404)

        # Plantilla actual
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

        # Historial de versiones (todas)
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

    # ============================================================
    # 2) Obtener archivo actual + estructura JSON
    # ============================================================
    archivo_existe = False
    preview_url = None
    estructura_actual = {}
    estructura_anterior = {}

    if plantilla:
        ruta = plantilla["gcs_path"]

        if gcs_exists(ruta):
            archivo_existe = True
            preview_url = generar_url_previa(ruta)

            # JSON estructural asociado a esta versión
            with connection.cursor() as cursor:
                cursor.execute("""
                    SELECT estructura_json
                    FROM plantilla_estructura_version
                    WHERE version_id = %s
                    ORDER BY id DESC LIMIT 1
                """, [plantilla["id"]])
                row = cursor.fetchone()

            estructura_actual = row[0] if row else {}
        else:
            # No existe el archivo en GCS
            plantilla = None

    office_url = office_or_download_url(preview_url) if preview_url else None

    # ============================================================
    # 3) Obtener estructura anterior (para comparar)
    # ============================================================
    if len(versiones) >= 2:
        penult = versiones[1]

        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT estructura_json
                FROM plantilla_estructura_version
                WHERE version_id = %s
                ORDER BY id DESC LIMIT 1
            """, [penult["id"]])
            row_prev = cursor.fetchone()

        estructura_anterior = row_prev[0] if row_prev else {}

    else:
        estructura_anterior = {}

    # ============================================================
    # 4) Comparación profunda de estructura (NUEVO)
    # ============================================================
    stats_diferencias = comparar_estructuras(
        estructura_actual,
        estructura_anterior
    )

    # ============================================================
    # 5) Estadísticas generales de versiones
    # ============================================================
    stats_versiones = (
        calcular_stats_versiones(versiones)
        if versiones else {"total_versiones": 0, "versiones_ok": 0, "versiones_rotas": 0}
    )

    # ============================================================
    # 6) Estadísticas de calidad de la estructura
    # ============================================================
    calidad = evaluar_calidad_estructura(estructura_actual)

    # ============================================================
    # 7) Procesar historial para UI
    # ============================================================
    versiones_proc = []
    for v in versiones:
        data = v.copy()
        gp = v["gcs_path"]

        if gp and gcs_exists(gp):
            prev = generar_url_previa(gp)
            data["office_url"] = office_or_download_url(prev) if prev else None
        else:
            data["office_url"] = None

        # Verificar si la versión está en uso
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT COUNT(*)
                FROM documentos_generados
                WHERE plantilla_documento_tecnico_version_id = %s
            """, [v["id"]])
            data["en_uso"] = cursor.fetchone()[0] > 0

        versiones_proc.append(data)

    # ============================================================
    # 8) Render final
    # ============================================================
    return render(request, "tipo_detalle.html", {
        "tipo": tipo,
        "plantilla": plantilla,

        "archivo_existe": archivo_existe,
        "preview_url": preview_url,
        "office_url": office_url,

        # Estructura profunda
        "estructura_json": estructura_actual,
        "estructura_anterior": estructura_anterior,

        # Comparación avanzada
        "stats_diferencias": stats_diferencias,
        "stats_versiones": stats_versiones,
        "calidad": calidad,

        # Historial
        "versiones": versiones_proc,
    })


# =============================================================================
# CREAR CATEGORÍA / TIPO DOCUMENTO
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
# SUBIR PLANTILLA (VERSIONADO REAL POR JSON)
# =============================================================================

def versionar_plantilla_json(version_actual, estructura_antes, estructura_despues):
    """
    Versionamiento real basado en diferencias profundas del JSON estructural.

    Usa extract_structural_signature() definido en leer_estructura_plantilla_word.py
    para comparar solo la “forma” de la plantilla.
    """
    from plantillas_documentos_tecnicos.leer_estructura_plantilla_word import (
        extract_structural_signature
    )

    # Primera versión
    if not version_actual:
        return "1.0"

    try:
        major, minor = map(int, str(version_actual).split("."))
    except Exception:
        major, minor = 1, 0

    sig1 = extract_structural_signature(estructura_antes or {})
    sig2 = extract_structural_signature(estructura_despues or {})

    if json.dumps(sig1, sort_keys=True) != json.dumps(sig2, sort_keys=True):
        # Cambios estructurales → major
        return f"{major + 1}.0"

    # Cambios menores → minor
    return f"{major}.{minor + 1}"


@login_required
def subir_plantilla(request, tipo_id):

    # 1) Datos del tipo
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT t.id, t.nombre, c.nombre AS categoria
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

        categoria_clean = clean(tipo["categoria"])
        tipo_clean = clean(tipo["nombre"])

        base_path = f"Plantillas/Documentos_Tecnicos/{categoria_clean}/{tipo_clean}/"

        client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
        bucket = client.bucket(settings.GCP_BUCKET_NAME)

        # 2) Buscar maestro existente
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT p.id AS plantilla_id, v.id AS version_id, v.version
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
            version_actual_id = anterior["version_id"]

            # Obtener JSON estructural anterior (dict)
            with connection.cursor() as cursor:
                cursor.execute("""
                    SELECT estructura_json
                    FROM plantilla_estructura_version
                    WHERE version_id = %s
                    ORDER BY id DESC LIMIT 1
                """, [version_actual_id])
                row = cursor.fetchone()

            estructura_antes = row[0] if row else None
        else:
            # Crear maestro si no existe
            with connection.cursor() as cursor:
                cursor.execute("""
                    INSERT INTO plantilla_tipo_doc (tipo_documento_id)
                    VALUES (%s)
                    RETURNING id
                """, [tipo_id])
                plantilla_id = cursor.fetchone()[0]

            version_anterior = None
            estructura_antes = None

        # 3) Guardar archivo local temporalmente
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            for chunk in archivo.chunks():
                tmp.write(chunk)
            nuevo_local = tmp.name

        # 4) Generar nueva estructura JSON (dict)
        from plantillas_documentos_tecnicos.leer_estructura_plantilla_word import generar_estructura
        estructura_despues = generar_estructura(nuevo_local)

        # 5) Calcular nueva versión (major/minor real)
        nueva_version = versionar_plantilla_json(
            version_anterior,
            estructura_antes,
            estructura_despues
        )

        version_folder = f"{base_path}V{nueva_version}/"
        bucket.blob(version_folder).upload_from_string("")

        filename = archivo.name
        blob_name = f"{version_folder}{filename}"

        bucket.blob(blob_name).upload_from_filename(nuevo_local)

        # Limpieza temp
        if os.path.exists(nuevo_local):
            os.remove(nuevo_local)

        # 6) Registrar nueva versión en BD (SIN COLUMNA controles)
        with connection.cursor() as cursor:

            cursor.execute("""
                INSERT INTO plantilla_tipo_doc_versiones (plantilla_id, version, gcs_path)
                VALUES (%s, %s, %s)
                RETURNING id
            """, [plantilla_id, nueva_version, blob_name])
            new_version_id = cursor.fetchone()[0]

            cursor.execute("""
                UPDATE plantilla_tipo_doc
                SET version_actual_id = %s, actualizado_en = NOW()
                WHERE id = %s
            """, [new_version_id, plantilla_id])

            # 7) Guardar JSON estructural profundo
            cursor.execute("""
                INSERT INTO plantilla_estructura_version
                (version_id, estructura_json)
                VALUES (%s, %s)
            """, [
                new_version_id,
                Json(estructura_despues)   # dict → JSONB
            ])

        messages.success(request, f"✔ Plantilla subida (versión {nueva_version}).")
        return redirect("plantillas:detalle_tipo", tipo_id=tipo_id)

    return render(request, "subir_plantilla.html", {"tipo": tipo})


# =============================================================================
# DESCARGAR ARCHIVO DESDE GCS
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

    version_folder = "/".join(gcs_path.split("/")[:-1]) + "/"
    base_folder = "/".join(version_folder.split("/")[:-2]) + "/"

    client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
    bucket = client.bucket(settings.GCP_BUCKET_NAME)

    # Verificar si esta versión está en uso por documentos_generados
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

    # 1) Borrar archivos de la versión en GCS
    for blob in list(bucket.list_blobs(prefix=version_folder)):
        try:
            blob.delete()
        except Exception:
            pass

    with connection.cursor() as cursor:

        # 2) Borrar versión en BD
        cursor.execute("DELETE FROM plantilla_tipo_doc_versiones WHERE id = %s", [version_id])

        # 3) Consultar cuántas versiones quedan
        cursor.execute("""
            SELECT COUNT(*)
            FROM plantilla_tipo_doc_versiones
            WHERE plantilla_id = %s
        """, [plantilla_id])
        quedan = cursor.fetchone()[0]

        # Si no queda ninguna versión, eliminar maestro + carpeta
        if quedan == 0:

            for blob in list(bucket.list_blobs(prefix=base_folder)):
                try:
                    blob.delete()
                except Exception:
                    pass

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
