# ============================================================
#   Gestión de Documentos Técnicos - VIEWS (Versión Final A)
# ============================================================
from plantillas_documentos_tecnicos.utils_documentos import (
    obtener_plantilla_usada,
    obtener_estructura_plantilla_usada,
    extract_blob_name_from_signed_url,
)
import uuid
from django.shortcuts import render, redirect
from django.contrib import messages
from django.http import JsonResponse, HttpResponse
from django.db import connection, transaction
from django.contrib.auth.decorators import login_required
from datetime import datetime, timedelta
from django.conf import settings
import json
import os
import tempfile
from plantillas_documentos_tecnicos.leer_estructura_plantilla_word import (
    generar_estructura
)
from google.cloud import storage

from .state_machine import DocumentoTecnicoStateMachine
from decimal import Decimal
from django.views.decorators.http import require_POST
import mimetypes
import re
# ============================================================
# IMPORTS DE PLANTILLAS
# ============================================================
from plantillas_documentos_tecnicos.leer_estructura_plantilla_word import (
    generar_estructura as extraer_controles_contenido_desde_file
)
from pathlib import Path
from plantillas_documentos_tecnicos.views import comparar_estructuras


# ============================================================
# HELPERS
# ============================================================



def obtener_estado_id(nombre):
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id FROM estado_documento WHERE nombre = %s LIMIT 1
        """, [nombre])
        row = cursor.fetchone()
    return row[0] if row else None


def clean(x):
    """Limpia nombres para rutas GCS."""
    if not x:
        return ""
    x = str(x)
    x = re.sub(r"[\/\\]+", " ", x)
    x = re.sub(r"\s+", "_", x)
    x = re.sub(r"[:*?\"<>|]+", "_", x)
    x = re.sub(r"_+", "_", x)
    return x.strip("_")


def to_json_safe(data):
    """Convierte Decimals, fechas y None en tipos seguros para json.dumps."""
    if isinstance(data, list):
        return [to_json_safe(x) for x in data]
    if isinstance(data, dict):
        return {k: to_json_safe(v) for k, v in data.items()}
    if isinstance(data, Decimal):
        return float(data)
    if isinstance(data, datetime):
        return data.isoformat()
    if data is None:
        return 0
    return data
class VersionManager:
    """
    Maneja el versionamiento de version_documento_tecnico sin crear
    carpetas dummy en GCS. Usa la última versión registrada como base.
    """

    def __init__(self, requerimiento_id, cursor):
        self.requerimiento_id = requerimiento_id
        self.cursor = cursor
        self.version_actual = self.obtener_ultima_version()

    def obtener_ultima_version(self):
        self.cursor.execute("""
            SELECT version
            FROM version_documento_tecnico
            WHERE requerimiento_documento_id = %s
            ORDER BY fecha DESC
            LIMIT 1
        """, [self.requerimiento_id])
        row = self.cursor.fetchone()
        # Si no hay versión previa, asumimos que la inicial ya fue creada por inicializar_version_inicial
        return row[0] if row else "v0.0.1"

    def _count_suffix(self, token):
        self.cursor.execute("""
            SELECT COUNT(*)
            FROM version_documento_tecnico
            WHERE requerimiento_documento_id = %s
              AND version LIKE %s
        """, [self.requerimiento_id, f"%-{token}%"])
        return self.cursor.fetchone()[0]

    def nueva_version(self, evento):
        base = self.version_actual.split("-")[0]  # Ej: v1.0.0 o v0.0.1
        c1, c2, c3 = map(int, base.replace("v", "").split("."))

        if evento in ["enviar_revision", "reenviar_revision"]:
            n = self._count_suffix("REV")
            c2 += 1
            c3 = 0
            return f"v{c1}.{c2}.{c3}-REV{n+1}"

        if evento == "rechazar_revision":
            n = self._count_suffix("REJREV")
            return f"v{c1}.{c2}.{c3}-REJREV{n+1}"

        if evento == "revision_aceptada":
            return f"v{c1+1}.0.0-APR1"

        if evento == "rechazar_aprobacion":
            n = self._count_suffix("REJAPR")
            return f"v{c1}.{c2}.{c3}-REJAPR{n+1}"

        if evento == "aprobar_documento":
            return f"v{c1}.{c2}.{c3+1}-APROBADO"

        # publicar_documento podrías tratarlo como base-PUB si quisieras
        if evento == "publicar_documento":
            return f"{base}-PUB"

        return self.version_actual

    def registrar_version(self, evento, estado_nombre, usuario_id, comentario):
        nueva_version = self.nueva_version(evento)

        self.cursor.execute("""
            INSERT INTO version_documento_tecnico
                (requerimiento_documento_id, version, estado_id, fecha, comentario, usuario_id, signed_url)
            VALUES (
                %s,
                %s,
                (SELECT id FROM estado_documento WHERE nombre = %s),
                NOW(),
                %s,
                %s,
                NULL
            )
        """, [
            self.requerimiento_id,
            nueva_version,
            estado_nombre,
            comentario,
            usuario_id,
        ])

        self.version_actual = nueva_version
        return nueva_version


def obtener_estado_id(nombre):
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id FROM estado_documento WHERE nombre = %s LIMIT 1
        """, [nombre])
        row = cursor.fetchone()
    return row[0] if row else None

def clean(x):
    """Limpia nombres para rutas GCS."""
    import re
    if not x:
        return ""
    x = str(x)
    x = re.sub(r"[\/\\]+", " ", x)
    x = re.sub(r"\s+", "_", x)
    x = re.sub(r"[:*?\"<>|]+", "_", x)
    x = re.sub(r"_+", "_", x)
    return x.strip("_")


def to_json_safe(data):
    """Convierte decimales, fechas, None → tipos válidos para JSON."""
    from decimal import Decimal
    if isinstance(data, list):
        return [to_json_safe(x) for x in data]
    if isinstance(data, dict):
        return {k: to_json_safe(v) for k, v in data.items()}
    if isinstance(data, Decimal):
        return float(data)
    if isinstance(data, datetime):
        return data.isoformat()
    if data is None:
        return 0
    return data


def extraer_controles_contenido_desde_gcs(gcs_path):
    """Descarga archivo desde GCS y extrae controles con generar_estructura()."""
    client = storage.Client()
    bucket = client.bucket(settings.GCP_BUCKET_NAME)
    blob = bucket.blob(gcs_path)

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        blob.download_to_filename(tmp.name)
        tmp_path = tmp.name

    estructura = extraer_controles_contenido_desde_file(tmp_path)

    os.remove(tmp_path)
    return estructura


def obtener_tipo_documento_por_rq(requerimiento_id):
    """
    SELECT TDT.id, TDT.nombre
    FROM requerimiento_documento_tecnico R
    JOIN tipo_documentos_tecnicos TDT ON R.tipo_documento_id = TDT.id
    WHERE R.id = %s
    """
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT TDT.id, TDT.nombre
            FROM requerimiento_documento_tecnico R
            JOIN tipo_documentos_tecnicos TDT ON R.tipo_documento_id = TDT.id
            WHERE R.id = %s
        """, [requerimiento_id])
        row = cursor.fetchone()

    return {"id": row[0], "nombre": row[1]} if row else None


def obtener_estructura_plantilla_referencia(tipo_id):
    """Obtiene la estructura JSON de plantilla_tipo_doc versión_actual."""
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT pe.estructura_json
            FROM plantilla_tipo_doc pt
            JOIN plantilla_estructura_version pe ON pe.id = pt.version_actual_id
            WHERE pt.tipo_documento_id = %s
        """, [tipo_id])
        row = cursor.fetchone()

    if not row:
        return None

    estructura = row[0]
    return json.loads(estructura) if isinstance(estructura, str) else estructura


def extraer_controles_archivo_temporal(archivo_fileobj):
    """Guarda temporalmente y extrae controles de archivo subido."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp.write(archivo_fileobj.read())
        tmp_path = tmp.name

    estructura = extraer_controles_contenido_desde_file(tmp_path)
    os.remove(tmp_path)
    return estructura


def validar_contra_plantilla(estructura_subida, estructura_base):
    """
    Valida estructura del archivo subido contra la versión exacta usada
    en el requerimiento, comparando:
      - tablas Word
      - excels embebidos
      - imágenes
    Es compatible con la NUEVA estructura 2025 (filas, n_filas, n_columnas).
    """

    dif = {
        "tablas_word": {"faltantes": [], "sobrantes": []},
        "excels": {"faltantes": [], "sobrantes": []},
        "imagenes": {"faltantes": [], "sobrantes": []},
        "detalles": [],
        "status": "OK"
    }

    # ============================================================
    # 🔵 TABLAS WORD
    # ============================================================
    base_tab = {(t.get("n_filas", 0), t.get("n_columnas", 0)) for t in estructura_base.get("tablas_word", [])}
    sub_tab  = {(t.get("n_filas", 0), t.get("n_columnas", 0)) for t in estructura_subida.get("tablas_word", [])}

    dif["tablas_word"]["faltantes"] = list(base_tab - sub_tab)
    dif["tablas_word"]["sobrantes"] = list(sub_tab - base_tab)

    # ============================================================
    # 🔵 EXCELS embebidos
    # ============================================================
    base_exc = {e.get("excel"): e for e in estructura_base.get("excels", [])}
    sub_exc  = {e.get("excel"): e for e in estructura_subida.get("excels", [])}

    dif["excels"]["faltantes"] = list(set(base_exc.keys()) - set(sub_exc.keys()))
    dif["excels"]["sobrantes"] = list(set(sub_exc.keys()) - set(base_exc.keys()))

    # === Comparación por Excel ===
    for excel_name in base_exc:

        if excel_name not in sub_exc:
            continue

        b = base_exc[excel_name]
        s = sub_exc[excel_name]

        # Diccionarios tabla → tabla
        b_tab = {t.get("tabla"): t for t in b.get("tablas", [])}
        s_tab = {t.get("tabla"): t for t in s.get("tablas", [])}

        # Tablas faltantes / extra
        for falt in (set(b_tab.keys()) - set(s_tab.keys())):
            dif["excels"]["faltantes"].append(f"{excel_name}:{falt}")

        for sob in (set(s_tab.keys()) - set(b_tab.keys())):
            dif["excels"]["sobrantes"].append(f"{excel_name}:{sob}")

        # Comparación interna de tablas
        for tname in b_tab:

            if tname not in s_tab:
                continue

            bt = b_tab[tname]
            st = s_tab[tname]

            # === Nueva estructura ===
            # columnas = primera fila
            # registros = resto
            bt_cols = bt.get("filas", [[]])[0] if bt.get("filas") else []
            st_cols = st.get("filas", [[]])[0] if st.get("filas") else []

            if len(bt_cols) != len(st_cols):
                dif["detalles"].append(
                    f"Excel {excel_name} > Tabla {tname}: distinta cantidad de columnas."
                )

            if bt.get("n_filas") != st.get("n_filas"):
                dif["detalles"].append(
                    f"Excel {excel_name} > Tabla {tname}: distinta cantidad de filas."
                )

    # ============================================================
    # 🔵 IMÁGENES
    # ============================================================
    base_imgs = {i.get("nombre") for i in estructura_base.get("imagenes", [])}
    sub_imgs  = {i.get("nombre") for i in estructura_subida.get("imagenes", [])}

    dif["imagenes"]["faltantes"] = list(base_imgs - sub_imgs)
    dif["imagenes"]["sobrantes"] = list(sub_imgs - base_imgs)

    # ============================================================
    # 🔵 STATUS GLOBAL
    # ============================================================
    if any([
        dif["tablas_word"]["faltantes"],
        dif["tablas_word"]["sobrantes"],
        dif["excels"]["faltantes"],
        dif["excels"]["sobrantes"],
        dif["imagenes"]["faltantes"],
        dif["imagenes"]["sobrantes"],
        dif["detalles"],
    ]):
        dif["status"] = "ERROR"

    return dif







def subir_archivo_version(requerimiento_id, archivo, sufijo):
    """Sube archivo REV / APR / PUB a GCS."""
    extension = os.path.splitext(archivo.name)[1]
    stamp = datetime.now().strftime("%Y%m%d-%H%M")
    filename = f"docs/RQ-{requerimiento_id}/RQ-{requerimiento_id}-{sufijo}-{stamp}{extension}"

    client = storage.Client()
    bucket = client.bucket(settings.GCP_BUCKET_NAME)
    blob = bucket.blob(filename)

    blob.upload_from_file(archivo)

    return blob.public_url, filename


def registrar_estado(requerimiento_id, usuario_id, estado_id, observaciones):
    with connection.cursor() as cursor:
        cursor.execute("""
            INSERT INTO log_estado_requerimiento_documento
            (requerimiento_id, usuario_id, estado_destino_id, observaciones, fecha_cambio)
            VALUES (%s, %s, %s, %s, NOW())
        """, [requerimiento_id, usuario_id, estado_id, observaciones])



def crear_version(requerimiento_id, usuario_id, estado_id, comentario, archivo_url, sufijo):
    version_str = f"{sufijo}-{datetime.now().strftime('%Y.%m.%d-%H%M')}"

    with connection.cursor() as cursor:
        cursor.execute("""
            INSERT INTO version_documento_tecnico
                (requerimiento_documento_id, usuario_id, estado_id, comentario, version, signed_url, fecha)
            VALUES (%s, %s, %s, %s, %s, %s, NOW())
        """, [
            requerimiento_id, usuario_id, estado_id,
            comentario, version_str, archivo_url
        ])

    return version_str


# ============================================================
# 1) LISTA DOCUMENTOS ASIGNADOS (VERSIÓN COMPLETA)
# ============================================================

@login_required
def lista_documentos_asignados(request):
    """
    Dashboard real: muestra estadísticas y gráficos reales 
    de los documentos asignados al usuario logueado.
    """
    user_id = request.user.id

    sql = """
    WITH UltimoEstado AS (
        SELECT
            requerimiento_id,
            estado_destino_id,
            ROW_NUMBER() OVER (
                PARTITION BY requerimiento_id
                ORDER BY fecha_cambio DESC, id DESC
            ) AS rn
        FROM public.log_estado_requerimiento_documento
    ),
    EstadoActual AS (
        SELECT
            RDT.id AS requerimiento_id,
            COALESCE(E.id, 0) AS estado_id,
            COALESCE(E.nombre, 'Pendiente de Inicio') AS estado_actual
        FROM public.requerimiento_documento_tecnico RDT
        LEFT JOIN UltimoEstado UE
            ON UE.requerimiento_id = RDT.id AND UE.rn = 1
        LEFT JOIN public.estado_documento E
            ON UE.estado_destino_id = E.id
    )
    SELECT
        RDT.id AS requerimiento_id,
        RDT.fecha_registro,
        RDT.observaciones,
        TDT.nombre AS tipo_documento,
        CDT.nombre AS categoria_documento,
        EA.estado_actual,
        RR.nombre AS rol_asignado,
        RDT.proyecto_id,
        P.nombre AS nombre_proyecto
    FROM public.requerimiento_documento_tecnico RDT
    INNER JOIN EstadoActual EA ON EA.requerimiento_id = RDT.id
    INNER JOIN public.tipo_documentos_tecnicos TDT ON RDT.tipo_documento_id = TDT.id
    INNER JOIN public.categoria_documentos_tecnicos CDT ON TDT.categoria_id = CDT.id
    INNER JOIN public.proyectos P ON RDT.proyecto_id = P.id
    INNER JOIN public.requerimiento_equipo_rol RER
        ON RDT.id = RER.requerimiento_id AND RER.activo = TRUE
    INNER JOIN public.roles_ciclodocumento RR ON RER.rol_id = RR.id
    WHERE RER.usuario_id = %s
    ORDER BY P.nombre, RDT.fecha_registro DESC;
    """

    with connection.cursor() as cursor:
        cursor.execute(sql, [user_id])
        columns = [col[0] for col in cursor.description]
        resultados = [dict(zip(columns, row)) for row in cursor.fetchall()]

    # Filtro de visibilidad según rol y estado
    def visible_para_rol(rol, estado):
        reglas = {
            "Redactor": ["Pendiente de Inicio", "En Elaboración", "Re Estructuración"],
            "Revisor": ["En Revisión"],
            "Aprobador": [
                "En Aprobación",
                "Aprobado. Listo para Publicación",
                "Publicado",
            ],
        }
        return estado in reglas.get(rol, [])

    resultados = [
        doc
        for doc in resultados
        if visible_para_rol(doc.get("rol_asignado"), doc.get("estado_actual"))
    ]

    # Agrupar documentos por proyecto
    documentos_por_proyecto = {}
    for doc in resultados:
        proyecto = doc.get("nombre_proyecto", "Sin Proyecto")
        documentos_por_proyecto.setdefault(proyecto, []).append(doc)

    # KPIs
    total_docs = len(resultados)

    # Distribución por estado
    por_estado = {}
    for doc in resultados:
        estado = doc.get("estado_actual", "Desconocido")
        por_estado[estado] = por_estado.get(estado, 0) + 1
    chart_estado = json.dumps(
        to_json_safe({"labels": list(por_estado.keys()), "values": list(por_estado.values())})
    )

    # Distribución por rol
    por_rol = {}
    for doc in resultados:
        rol = doc.get("rol_asignado", "Sin Rol")
        por_rol[rol] = por_rol.get(rol, 0) + 1
    chart_rol = json.dumps(
        to_json_safe({"labels": list(por_rol.keys()), "values": list(por_rol.values())})
    )

    # Actividad últimos 7 días
    dias = [
        (datetime.now() - timedelta(days=i)).strftime("%d-%b")
        for i in reversed(range(7))
    ]
    actividad = []
    for i in reversed(range(7)):
        dia = (datetime.now() - timedelta(days=i)).date()
        count = 0
        for doc in resultados:
            fecha_doc = doc.get("fecha_registro")
            if isinstance(fecha_doc, datetime) and fecha_doc.date() == dia:
                count += 1
        actividad.append(count)
    chart_actividad = json.dumps(
        to_json_safe({"labels": dias, "values": actividad})
    )

    # Tiempo promedio por etapa
    with connection.cursor() as cursor:
        cursor.execute("""
            WITH duraciones AS (
                SELECT
                    l.requerimiento_id,
                    e.nombre AS estado,
                    (EXTRACT(EPOCH FROM (l.fecha_cambio - LAG(l.fecha_cambio) OVER (
                        PARTITION BY l.requerimiento_id ORDER BY l.fecha_cambio
                    ))) / 3600.0) AS horas_transicion
                FROM public.log_estado_requerimiento_documento l
                INNER JOIN public.estado_documento e ON l.estado_destino_id = e.id
            )
            SELECT
                estado,
                ROUND(AVG(horas_transicion)::numeric, 2) AS horas_promedio
            FROM duraciones
            WHERE horas_transicion IS NOT NULL
            GROUP BY estado
            ORDER BY estado;
        """)
        tiempos_data = cursor.fetchall()

    etapas, tiempos = [], []
    for estado, horas in tiempos_data:
        valor = float(horas) if isinstance(horas, Decimal) else (horas or 0)
        etapas.append(estado)
        tiempos.append(round(valor, 2))
    chart_tiempos = json.dumps(
        to_json_safe({"labels": etapas, "values": tiempos})
    )

    # Radar desempeño
    if tiempos and max(tiempos) > 0:
        max_tiempo = max(tiempos)
        radar_values = [
            round((1 - (t / max_tiempo)) * 100, 1) for t in tiempos
        ]
    else:
        radar_values = [0 for _ in tiempos]
    chart_radar = json.dumps(
        to_json_safe({"labels": etapas, "values": radar_values})
    )

    # Cumplimiento real
    publicados = sum(
        1
        for doc in resultados
        if doc.get("estado_actual")
        in ["Publicado", "Aprobado. Listo para Publicación"]
    )
    cumplimiento = (
        round((publicados / total_docs * 100), 1) if total_docs > 0 else 0
    )

    colores_estado = {
        "Pendiente de Inicio": "secondary",
        "En Elaboración": "info",
        "En Revisión": "warning",
        "En Aprobación": "primary",
        "Aprobado. Listo para Publicación": "success",
        "Publicado": "success",
        "Re Estructuración": "danger",
    }

    sin_datos_radar = not any(t > 0 for t in tiempos)

    context = {
        "documentos_por_proyecto": documentos_por_proyecto,
        "total_docs": total_docs,
        "cumplimiento": cumplimiento,
        "chart_estado": chart_estado,
        "chart_rol": chart_rol,
        "chart_actividad": chart_actividad,
        "chart_tiempos": chart_tiempos,
        "chart_radar": chart_radar,
        "colores_estado": colores_estado,
        "sin_datos_radar": sin_datos_radar,
    }

    return render(request, "lista_documentos_asignados.html", context)



# ============================================================
# 2) D

@login_required
def detalle_documento(request, requerimiento_id):
    """
    Vista principal del RQ:
      - Muestra datos generales, historial de estados, historial de versiones.
      - Calcula eventos permitidos con la StateMachine.
      - En POST:
          * Valida permisos via StateMachine.
          * Valida estructura del archivo (estructura 2025) para eventos con archivo.
          * Ejecuta transición real (getattr(sm, evento)()).
          * Registra log_estado_requerimiento_documento.
          * Maneja versionamiento con VersionManager (excepto iniciar_elaboracion).
          * Sube archivo a GCS cuando corresponde.
    """
    mensaje = ""

    # ===============================
    # DATOS PRINCIPALES DEL RQ
    # ===============================
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT 
                R.id,
                R.fecha_registro,
                R.observaciones,
                TDT.nombre AS tipo_documento,
                CDT.nombre AS categoria_documento,
                COALESCE(RER.rol_id, 0),
                COALESCE(RR.nombre, 'Sin Rol'),
                COALESCE(E.nombre, 'Pendiente de Inicio'),
                P.nombre
            FROM requerimiento_documento_tecnico R
            JOIN tipo_documentos_tecnicos TDT ON R.tipo_documento_id = TDT.id
            JOIN categoria_documentos_tecnicos CDT ON TDT.categoria_id = CDT.id
            JOIN proyectos P ON R.proyecto_id = P.id
            LEFT JOIN requerimiento_equipo_rol RER 
                  ON RER.requerimiento_id = R.id 
                 AND RER.usuario_id = %s 
                 AND RER.activo = TRUE
            LEFT JOIN roles_ciclodocumento RR ON RR.id = RER.rol_id
            LEFT JOIN (
                SELECT requerimiento_id, estado_destino_id
                FROM log_estado_requerimiento_documento
                WHERE requerimiento_id = %s
                ORDER BY fecha_cambio DESC
                LIMIT 1
            ) ULT ON TRUE
            LEFT JOIN estado_documento E ON E.id = ULT.estado_destino_id
            WHERE R.id = %s;
        """, [request.user.id, requerimiento_id, requerimiento_id])

        row = cursor.fetchone()

    if not row:
        messages.error(request, "Documento no encontrado.")
        return redirect("documentos:lista_documentos_asignados")

    documento = dict(zip([
        "requerimiento_id", "fecha_registro", "observaciones", "tipo_documento",
        "categoria_documento", "rol_id", "rol_asignado", "estado_actual",
        "nombre_proyecto"
    ], row))

    # ===============================
    # PLANTILLA USADA EN EL RQ
    # ===============================
    try:
        plantilla_usada = obtener_plantilla_usada(requerimiento_id)
    except Exception:
        plantilla_usada = None

    # ===============================
    # HISTORIAL ESTADOS
    # ===============================
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT 
                l.fecha_cambio,
                e.nombre AS estado_destino,
                u.nombre AS usuario_nombre,
                l.observaciones AS comentario
            FROM log_estado_requerimiento_documento l
            JOIN estado_documento e ON e.id = l.estado_destino_id
            JOIN usuarios u ON u.id = l.usuario_id
            WHERE l.requerimiento_id = %s
            ORDER BY l.fecha_cambio ASC
        """, [requerimiento_id])
        historial_estados = [
            dict(zip(
                ["fecha_cambio", "estado_destino", "usuario_nombre", "comentario"],
                r
            ))
            for r in cursor.fetchall()
        ]

    # ===============================
    # HISTORIAL VERSIONES
    # ===============================
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT version, e.nombre, fecha,
                   u.nombre, comentario, signed_url
            FROM version_documento_tecnico v
            JOIN estado_documento e ON v.estado_id = e.id
            LEFT JOIN usuarios u ON v.usuario_id = u.id
            WHERE v.requerimiento_documento_id = %s
            ORDER BY fecha ASC
        """, [requerimiento_id])
        historial_versiones = [
            dict(zip(
                ["version", "estado_nombre", "fecha", "usuario_nombre", "comentario", "signed_url"],
                r
            ))
            for r in cursor.fetchall()
        ]

    # ===============================
    # STATE MACHINE
    # ===============================
    sm = DocumentoTecnicoStateMachine(
        rol_id=documento["rol_id"],
        estado_inicial=documento["estado_actual"]
    )

    eventos_map = {
        "iniciar_elaboracion": "Iniciar Elaboración",
        "enviar_revision": "Enviar a Revisión",
        "reenviar_revision": "Re-enviar a Revisión",
        "revision_aceptada": "Aceptar Revisión",
        "rechazar_revision": "Rechazar Revisión",
        "aprobar_documento": "Aprobar Documento",
        "rechazar_aprobacion": "Rechazar Aprobación",
        "publicar_documento": "Publicar Documento",
    }

    eventos_tuplas = [
        (e, t) for e, t in eventos_map.items()
        if sm.puede_transicionar(e)
    ]

    eventos_con_comentario = ["rechazar_revision", "rechazar_aprobacion"]
    eventos_con_archivo = [
        "enviar_revision",
        "reenviar_revision",
        "rechazar_revision",
        "rechazar_aprobacion"
    ]

    # ===============================
    # GET → sólo mostrar detalle
    # ===============================
    if request.method == "GET":
        return render(request, "detalle_documento.html", {
            "documento": documento,
            "historial_estados": historial_estados,
            "plantilla_usada": plantilla_usada,
            "historial_versiones": historial_versiones,
            "eventos_con_archivo": eventos_con_archivo,
            "eventos_tuplas": eventos_tuplas,
            "eventos_con_comentario": eventos_con_comentario,
        })

    # ===============================
    # POST → EJECUTAR EVENTO
    # ===============================
    evento = request.POST.get("evento")
    comentario = request.POST.get("comentario", "").strip()
    archivo = request.FILES.get("archivo")

    # Validaciones base
    if evento in eventos_con_comentario and not comentario:
        messages.error(request, f"Debes ingresar un comentario para '{eventos_map.get(evento, evento)}'.")
        return redirect("documentos:detalle_documento", requerimiento_id)

    if evento in eventos_con_archivo and not archivo:
        messages.error(request, f"Debes adjuntar un archivo para '{eventos_map.get(evento, evento)}'.")
        return redirect("documentos:detalle_documento", requerimiento_id)

    if not sm.puede_transicionar(evento):
        messages.error(
            request,
            f"No tienes permiso para ejecutar '{eventos_map.get(evento, evento)}' "
            f"desde el estado '{documento['estado_actual']}'."
        )
        return redirect("documentos:detalle_documento", requerimiento_id)

    # ===============================
    # VALIDACIÓN ESTRUCTURAL 2025 (solo eventos con archivo)
    # ===============================
    if evento in eventos_con_archivo and archivo:
        try:
            # 1) Guardar archivo temporal
            tmp_path = Path(tempfile.gettempdir()) / f"{uuid.uuid4()}.docx"
            with open(tmp_path, "wb") as f:
                for chunk in archivo.chunks():
                    f.write(chunk)

            # 2) Estructura del archivo subido
            estructura_archivo = generar_estructura(str(tmp_path))
        except Exception as e:
            messages.error(request, f"Error al leer estructura del archivo: {e}")
            try:
                tmp_path.unlink(missing_ok=True)
            except Exception:
                pass
            return redirect("documentos:detalle_documento", requerimiento_id)
        finally:
            try:
                tmp_path.unlink(missing_ok=True)
            except Exception:
                pass

        # 3) Estructura de la plantilla usada (versión REAL del RQ)
        try:
            estructura_plantilla, version_id = obtener_estructura_plantilla_usada(requerimiento_id)
        except Exception as e:
            messages.error(request, f"⚠ No se encontró la estructura de la plantilla usada: {e}")
            return redirect("documentos:detalle_documento", requerimiento_id)

        # 4) Validación estructural estricta
        dif = validar_contra_plantilla(estructura_archivo, estructura_plantilla)

        if dif.get("status") == "ERROR":
            messages.error(
                request,
                (
                    "⚠ El documento NO corresponde a la estructura registrada para este RQ. "
                    "Revisa el panel de comparación para ver detalles "
                    "(tablas Word, Excels embebidos, imágenes, filas y columnas)."
                )
            )
            return redirect("documentos:detalle_documento", requerimiento_id)

    # ===============================
    # EJECUTAR TRANSICIÓN REAL + BD
    # ===============================
    try:
        with transaction.atomic():
            with connection.cursor() as cursor:

                # 1) Ejecutar transición en la StateMachine
                getattr(sm, evento)()
                nuevo_estado = sm.current_state.name

                # 2) Registrar log de estado
                estado_id = obtener_estado_id(nuevo_estado)
                if not estado_id:
                    raise Exception(f"No se encontró estado_documento para '{nuevo_estado}'.")

                cursor.execute("""
                    INSERT INTO log_estado_requerimiento_documento
                        (requerimiento_id, usuario_id, estado_destino_id, observaciones, fecha_cambio)
                    VALUES (%s, %s, %s, %s, NOW())
                """, [
                    requerimiento_id,
                    request.user.id,
                    estado_id,
                    comentario or f"Evento: {evento}",
                ])

                # 3) Caso especial: iniciar_elaboracion → inicializar_version_inicial
                if evento == "iniciar_elaboracion":
                    cursor.execute("""
                        SELECT 
                            P.nombre AS proyecto,
                            CL.nombre AS cliente,
                            CDT.nombre AS categoria,
                            TDT.nombre AS tipo,
                            R.codigo_documento
                        FROM requerimiento_documento_tecnico R
                        JOIN proyectos P ON R.proyecto_id = P.id
                        JOIN contratos C ON P.contrato_id = C.id
                        JOIN clientes CL ON C.cliente_id = CL.id
                        JOIN tipo_documentos_tecnicos TDT ON R.tipo_documento_id = TDT.id
                        JOIN categoria_documentos_tecnicos CDT ON TDT.categoria_id = CDT.id
                        WHERE R.id = %s
                    """, [requerimiento_id])
                    proyecto, cliente, categoria, tipo_doc, codigo_doc = cursor.fetchone()

                    cliente   = clean(cliente)
                    proyecto  = clean(proyecto)
                    categoria = clean(categoria)
                    tipo_doc  = clean(tipo_doc)
                    codigo_doc = clean(codigo_doc or f"RQ-{requerimiento_id}")

                    ruta_plantilla = (
                        f"DocumentosProyectos/{cliente}/{proyecto}/"
                        f"Documentos_Tecnicos/{categoria}/{tipo_doc}/"
                        f"RQ-{requerimiento_id}/Plantilla/"
                    )

                    from plantillas_documentos_tecnicos.utils_documentos import inicializar_version_inicial

                    storage_client = storage.Client()
                    bucket = storage_client.bucket(settings.GCP_BUCKET_NAME)

                    inicializar_version_inicial(
                        cursor=cursor,
                        bucket=bucket,
                        requerimiento_id=requerimiento_id,
                        ruta_plantilla=ruta_plantilla,
                        codigo_documento=codigo_doc,
                    )

                # 4) Versionamiento (para el resto de eventos que generan versión)
                if sm.evento_genera_version(evento) and evento != "iniciar_elaboracion":
                    vm = VersionManager(requerimiento_id=requerimiento_id, cursor=cursor)

                    nueva_version = vm.registrar_version(
                        evento,
                        nuevo_estado,
                        request.user.id,
                        comentario or f"Evento: {evento}",
                    )

                    # 5) Si el evento lleva archivo, subirlo a GCS y asociar signed_url
                    if evento in eventos_con_archivo and archivo:
                        cursor.execute("""
                            SELECT 
                                P.nombre AS proyecto,
                                CL.nombre AS cliente,
                                CDT.nombre AS categoria,
                                TDT.nombre AS tipo
                            FROM requerimiento_documento_tecnico R
                            JOIN proyectos P ON R.proyecto_id = P.id
                            JOIN contratos C ON P.contrato_id = C.id
                            JOIN clientes CL ON C.cliente_id = CL.id
                            JOIN tipo_documentos_tecnicos TDT ON R.tipo_documento_id = TDT.id
                            JOIN categoria_documentos_tecnicos CDT ON TDT.categoria_id = CDT.id
                            WHERE R.id = %s
                        """, [requerimiento_id])
                        proyecto, cliente, categoria, tipo = cursor.fetchone()

                        cliente = clean(cliente)
                        proyecto = clean(proyecto)
                        categoria = clean(categoria)
                        tipo = clean(tipo)
                        nombre_archivo = clean(archivo.name)

                        ruta_final = (
                            f"DocumentosProyectos/{cliente}/{proyecto}/"
                            f"Documentos_Tecnicos/{categoria}/{tipo}/"
                            f"RQ-{requerimiento_id}/{nueva_version}/{nombre_archivo}"
                        )

                        storage_client = storage.Client()
                        bucket = storage_client.bucket(settings.GCP_BUCKET_NAME)

                        mime_type, _ = mimetypes.guess_type(nombre_archivo)
                        blob = bucket.blob(ruta_final)

                        try:
                            archivo.file.seek(0)
                        except Exception:
                            pass

                        blob.upload_from_file(
                            archivo.file,
                            content_type=mime_type or "application/octet-stream",
                        )

                        signed_url = blob.generate_signed_url(
                            version="v4",
                            expiration=timedelta(days=7)
                        )

                        cursor.execute("""
                            UPDATE version_documento_tecnico
                            SET signed_url = %s
                            WHERE requerimiento_documento_id = %s
                              AND version = %s
                        """, [signed_url, requerimiento_id, nueva_version])

        mensaje = f"Evento '{eventos_map.get(evento, evento)}' ejecutado correctamente. Nuevo estado: {sm.current_state.name}"
        messages.success(request, mensaje)

    except Exception as e:
        import traceback
        traceback.print_exc()
        messages.error(request, f"Error al ejecutar '{eventos_map.get(evento, evento)}': {e}")

    return redirect("documentos:detalle_documento", requerimiento_id)











# ============================================================
# 3) DESCARGAR PLANTILLA RQ
# ============================================================

@login_required
def descargar_plantilla_rq(request, requerimiento_id):
    """
    Genera una signed URL a la última plantilla copiada en el RQ.
    """
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT ruta_gcs
            FROM documentos_generados
            WHERE ruta_gcs LIKE %s
            ORDER BY fecha_generacion DESC, id DESC
            LIMIT 1
        """, [f"%RQ-{requerimiento_id}/%"])

        row = cursor.fetchone()

    if not row:
        messages.error(request, "No se encontró la plantilla en el RQ.")
        return redirect("documentos:detalle_documento", requerimiento_id=requerimiento_id)

    ruta = row[0]

    client = storage.Client()
    bucket = client.bucket(settings.GCP_BUCKET_NAME)
    blob = bucket.blob(ruta)

    signed_url = blob.generate_signed_url(
        version="v4",
        expiration=timedelta(minutes=15),
    )

    return redirect(signed_url)


# ============================================================
# 4) SUBIR ARCHIVO A UNA VERSIÓN
# ============================================================

@login_required
def subir_archivo_documento(request, requerimiento_id):

    if request.method != "POST":
        return redirect("documentos:detalle_documento", requerimiento_id=requerimiento_id)

    archivo = request.FILES.get("archivo")
    if not archivo:
        messages.error(request, "❌ Debes seleccionar un archivo.")
        return redirect("documentos:detalle_documento", requerimiento_id=requerimiento_id)

    try:
        with connection.cursor() as cursor:

            cursor.execute("""
                SELECT 
                    RDT.id,
                    P.nombre AS proyecto,
                    CL.nombre AS cliente,
                    CDT.nombre AS categoria,
                    TDT.nombre AS tipo
                FROM requerimiento_documento_tecnico RDT
                JOIN proyectos P ON RDT.proyecto_id = P.id
                JOIN contratos C ON P.contrato_id = C.id
                JOIN clientes CL ON C.cliente_id = CL.id
                JOIN tipo_documentos_tecnicos TDT ON RDT.tipo_documento_id = TDT.id
                JOIN categoria_documentos_tecnicos CDT ON TDT.categoria_id = CDT.id
                WHERE RDT.id = %s
            """, [requerimiento_id])

            row = cursor.fetchone()
            if not row:
                messages.error(request, "❌ Documento no encontrado.")
                return redirect("documentos:detalle_documento", requerimiento_id)

            doc_id, proyecto, cliente, categoria, tipo = row

            # obtener última version
            cursor.execute("""
                SELECT version
                FROM version_documento_tecnico
                WHERE requerimiento_documento_id = %s
                ORDER BY fecha DESC
                LIMIT 1
            """, [doc_id])

            row = cursor.fetchone()
            version_actual = row[0] if row else "v0.0.1"

        cliente = clean(cliente)
        proyecto = clean(proyecto)
        categoria = clean(categoria)
        tipo = clean(tipo)
        nombre_archivo = clean(archivo.name)

        ruta_final = (
            f"DocumentosProyectos/{cliente}/{proyecto}/"
            f"Documentos_Tecnicos/{categoria}/{tipo}/"
            f"{version_actual}/{nombre_archivo}"
        )

        client = storage.Client()
        bucket = client.bucket(settings.GCP_BUCKET_NAME)
        blob = bucket.blob(ruta_final)

        archivo.file.seek(0)
        blob.upload_from_file(archivo.file)

        signed_url = blob.generate_signed_url(
            version="v4", expiration=timedelta(days=7)
        )

        with connection.cursor() as cursor:
            cursor.executemany("""
                UPDATE version_documento_tecnico
                SET signed_url = %s
                WHERE requerimiento_documento_id = %s
                  AND version = %s
            """, [[signed_url, requerimiento_id, version_actual]])

        messages.success(request, f"Archivo subido correctamente a la versión {version_actual}")

    except Exception as e:
        import traceback
        traceback.print_exc()
        messages.error(request, f"⚠ Error al subir archivo: {e}")

    return redirect("documentos:detalle_documento", requerimiento_id=requerimiento_id)


# ============================================================
# 5) PREVALIDAR CONTROLES (FLUJO VISUAL)
# ============================================================

@login_required
def prevalidar_controles(request, requerimiento_id):
    """
    Compara un archivo subido vs la plantilla oficial del tipo_documento.
    No ejecuta el evento: solo muestra diferencias.
    """

    archivo = request.FILES.get("archivo")
    evento = request.POST.get("evento")

    if not archivo:
        messages.error(request, "❌ Debes seleccionar un archivo.")
        return redirect("documentos:detalle_documento", requerimiento_id=requerimiento_id)

    # guardar temporalmente
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        for chunk in archivo.chunks():
            tmp.write(chunk)
        tmp_path = tmp.name

    controles_subido = set(extraer_controles_contenido_desde_file(tmp_path) or [])
    os.unlink(tmp_path)

    # obtener plantilla RQ o versión_actual
    ruta_plantilla = None

    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT V.gcs_path
            FROM documentos_generados DG
            JOIN plantilla_tipo_doc_versiones V
                ON V.id = DG.plantilla_documento_tecnico_version_id
            WHERE DG.ruta_gcs LIKE %s
            LIMIT 1
        """, [f"%RQ-{requerimiento_id}/%"])

        row = cursor.fetchone()
        if row:
            ruta_plantilla = row[0]
        else:
            cursor.execute("""
                SELECT V.gcs_path
                FROM requerimiento_documento_tecnico R
                JOIN tipo_documentos_tecnicos TDT ON R.tipo_documento_id = TDT.id
                JOIN plantilla_tipo_doc P ON P.tipo_documento_id = TDT.id
                JOIN plantilla_tipo_doc_versiones V ON V.id = P.version_actual_id
                WHERE R.id = %s
            """, [requerimiento_id])

            row2 = cursor.fetchone()
            ruta_plantilla = row2[0] if row2 else None

    controles_esperados = set()
    if ruta_plantilla:
        controles_esperados = set(extraer_controles_contenido_desde_gcs(ruta_plantilla) or [])

    faltantes = sorted(controles_esperados - controles_subido)
    sobrantes = sorted(controles_subido - controles_esperados)
    coincidencias = sorted(controles_esperados & controles_subido)

    return render(request, "comparacion_controles.html", {
        "requerimiento_id": requerimiento_id,
        "evento": evento,
        "archivo": archivo,
        "controles_subido": sorted(controles_subido),
        "controles_cuerpo": sorted(controles_esperados),
        "coincidencias": coincidencias,
        "faltantes": faltantes,
        "sobrantes": sobrantes,
    })


# ============================================================
# 6) VALIDAR CONTROLES AJAX
# ============================================================

@login_required
def validar_controles_doc_ajax(request, requerimiento_id):
    """
    Valida el archivo subido contra la estructura EXACTA
    de la versión de plantilla usada en este requerimiento.
    """

    if request.method != "POST":
        return JsonResponse({"error": "Método no permitido"}, status=405)

    archivo = request.FILES.get("archivo")
    if not archivo:
        return JsonResponse({"error": "No se recibió archivo"}, status=400)

    # ------------------------------------------------------------
    # 1) Obtener estructura de la PLANTILLA USADA (desde BD)
    # ------------------------------------------------------------
    try:
        estructura_plantilla, version_id = obtener_estructura_plantilla_usada(requerimiento_id)
    except Exception as e:
        return JsonResponse({"error": str(e)})

    # ------------------------------------------------------------
    # 2) Generar estructura del archivo SUBIDO
    # ------------------------------------------------------------
    try:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        for chunk in archivo.chunks():
            tmp.write(chunk)
        tmp.close()

        estructura_archivo = generar_estructura(tmp.name)
    except Exception as e:
        return JsonResponse({"error": f"Error leyendo archivo: {e}"})

    # ------------------------------------------------------------
    # 3) Validación estructural
    # ------------------------------------------------------------
    dif = validar_contra_plantilla(estructura_archivo, estructura_plantilla)

    return JsonResponse(dif)

