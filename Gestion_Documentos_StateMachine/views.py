# C:\Users\jonat\Documents\gestion_docs\Gestion_Documentos_StateMachine\views.py

from .state_machine import DocumentoTecnicoStateMachine
from django.http import HttpResponse, JsonResponse
from datetime import datetime, timedelta
from decimal import Decimal
from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required
from django.db import connection, transaction
from django.contrib import messages
from django.views.decorators.http import require_POST
import json
import mimetypes
import re

"""
NOTA IMPORTANTE:
- Ya NO se usan las tablas plantillas_utilidad_old ni plantillas_documentos_tecnicos_old.
- Toda la lógica de este módulo asume las tablas actuales:
    - plantilla_tipo_doc
    - plantilla_portada
"""


def clean(x):
    if not x:
        return ""
    x = str(x)
    x = re.sub(r"[\/\\]+", " ", x)
    x = re.sub(r"\s+", "_", x)
    x = re.sub(r"[:*?\"<>|]+", "_", x)
    x = re.sub(r"_+", "_", x)
    return x.strip("_")


def to_json_safe(data):
    """Convierte Decimals, fechas y None en tipos seguros para json.dumps"""
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


def validar_controles_archivo(requerimiento_id, archivo):
    """
    Compara los controles de contenido del archivo subido (DOCX)
    contra la combinación de:
      - Portada utilidad Word correspondiente (tipo_id = 1, formato_id del tipo)
      - Plantilla de cuerpo del tipo_documento correspondiente

    Si hay diferencias, devuelve (False, mensaje_error).
    Si todo ok o no hay plantillas de referencia, devuelve (True, None).
    """

    # Import local para evitar problemas de import circular
    from plantillas_documentos_tecnicos.views import (
        extraer_controles_contenido_desde_gcs,
        extraer_controles_contenido_desde_file,
    )
    import tempfile
    import os

    # 1) Guardar archivo subido en un .docx temporal
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        for chunk in archivo.chunks():
            tmp.write(chunk)
        tmp_path = tmp.name

    # MUY IMPORTANTE: resetear puntero para que luego pueda subirse normalmente
    try:
        archivo.seek(0)
    except Exception:
        pass

    try:
        # 2) Extraer controles del archivo que sube el usuario
        controles_subido = set(extraer_controles_contenido_desde_file(tmp_path) or [])

    finally:
        # Eliminar archivo temporal
        try:
            os.remove(tmp_path)
        except Exception:
            pass

    # Si el archivo no tiene controles, igual seguimos y comparamos con referencia
    # 3) Obtener tipo_documento_id y formato_id del requerimiento
    with connection.cursor() as cursor:
        cursor.execute(
            """
            SELECT 
                TDT.id AS tipo_id,
                F.id   AS formato_id
            FROM requerimiento_documento_tecnico R
            JOIN tipo_documentos_tecnicos TDT ON R.tipo_documento_id = TDT.id
            JOIN formato_archivo F           ON TDT.formato_id       = F.id
            WHERE R.id = %s
        """,
            [requerimiento_id],
        )
        row = cursor.fetchone()

    if not row:
        # No podemos validar sin datos, dejamos pasar
        return True, None

    tipo_id, formato_id = row

    # 4) Obtener GCS path de la plantilla de cuerpo (última versión)
    cuerpo_path = None
    portada_path = None

    with connection.cursor() as cursor:
        # Cuerpo: última plantilla del tipo de documento
        cursor.execute(
            """
            SELECT gcs_path
            FROM plantilla_tipo_doc
            WHERE tipo_documento_id = %s
            ORDER BY version DESC, id DESC
            LIMIT 1
        """,
            [tipo_id],
        )
        row = cursor.fetchone()
        if row and row[0]:
            cuerpo_path = row[0]

        # Portada utilidad correspondiente (tipo_id = 1 = Portada Word)
        cursor.execute(
            """
            SELECT gcs_path
            FROM plantilla_portada
            WHERE tipo_id = 1 AND formato_id = %s
            ORDER BY version DESC, id DESC
            LIMIT 1
        """,
            [formato_id],
        )
        row = cursor.fetchone()
        if row and row[0]:
            portada_path = row[0]

    # 5) Construir set de controles de referencia (portada + cuerpo)
    controles_ref = set()

    if cuerpo_path:
        try:
            controles_ref |= set(
                extraer_controles_contenido_desde_gcs(cuerpo_path) or []
            )
        except Exception:
            pass

    if portada_path:
        try:
            controles_ref |= set(
                extraer_controles_contenido_desde_gcs(portada_path) or []
            )
        except Exception:
            pass

    # Si no hay controles de referencia definidos, no validamos nada
    if not controles_ref:
        return True, None

    # 6) Comparar
    faltantes = sorted(controles_ref - controles_subido)
    sobrantes = sorted(controles_subido - controles_ref)

    if not faltantes and not sobrantes:
        # Todo ok
        return True, None

    partes_msg = ["❌ El archivo no coincide con las plantillas configuradas."]

    if faltantes:
        partes_msg.append("Faltan controles: " + ", ".join(faltantes))
    if sobrantes:
        partes_msg.append("Controles adicionales: " + ", ".join(sobrantes))

    return False, " ".join(partes_msg)


@login_required
def lista_documentos_asignados(request):
    """
    Dashboard real: muestra estadísticas y gráficos reales 
    de los documentos asignados al usuario logueado.
    """
    user_id = request.user.id

    # === CONSULTA PRINCIPAL: documentos asignados ===
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

    # ===  FILTRO DE VISIBILIDAD SEGÚN ROL Y ESTADO ===
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

    # === Agrupar documentos por proyecto ===
    documentos_por_proyecto = {}
    for doc in resultados:
        proyecto = doc.get("nombre_proyecto", "Sin Proyecto")
        documentos_por_proyecto.setdefault(proyecto, []).append(doc)

    # === KPIs ===
    total_docs = len(resultados)

    #  Distribución por estado
    por_estado = {}
    for doc in resultados:
        estado = doc.get("estado_actual", "Desconocido")
        por_estado[estado] = por_estado.get(estado, 0) + 1
    chart_estado = json.dumps(
        to_json_safe(
            {"labels": list(por_estado.keys()), "values": list(por_estado.values())}
        )
    )

    #  Distribución por rol
    por_rol = {}
    for doc in resultados:
        rol = doc.get("rol_asignado", "Sin Rol")
        por_rol[rol] = por_rol.get(rol, 0) + 1
    chart_rol = json.dumps(
        to_json_safe({"labels": list(por_rol.keys()), "values": list(por_rol.values())})
    )

    #  Actividad últimos 7 días
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

    #  Tiempo promedio por etapa (logs reales)
    with connection.cursor() as cursor:
        cursor.execute(
            """
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
        """
        )
        tiempos_data = cursor.fetchall()

    etapas, tiempos = [], []
    for estado, horas in tiempos_data:
        valor = float(horas) if isinstance(horas, Decimal) else (horas or 0)
        etapas.append(estado)
        tiempos.append(round(valor, 2))
    chart_tiempos = json.dumps(
        to_json_safe({"labels": etapas, "values": tiempos})
    )

    # ️ Desempeño por etapa
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

    # ✅ Cumplimiento real
    publicados = sum(
        1
        for doc in resultados
        if doc.get("estado_actual")
        in ["Publicado", "Aprobado. Listo para Publicación"]
    )
    cumplimiento = (
        round((publicados / total_docs * 100), 1) if total_docs > 0 else 0
    )

    #  Colores de estado
    colores_estado = {
        "Pendiente de Inicio": "secondary",
        "En Elaboración": "info",
        "En Revisión": "warning",
        "En Aprobación": "primary",
        "Aprobado. Listo para Publicación": "success",
        "Publicado": "success",
        "Re Estructuración": "danger",
    }

    # ⚠️ Bandera si no hay datos
    sin_datos_radar = not any(t > 0 for t in tiempos)

    # === Contexto final ===
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


def generar_signed_url(documento_id, version):
    # Aquí va la lógica de tu storage para obtener URL firmada
    # Por ahora podemos simular
    return f"https://storage.simulado.com/doc_{documento_id}_{version}.pdf"


class VersionManager:
    """
    Control de versionamiento sin crear archivos dummy en GCS.
    Las carpetas se crean automáticamente cuando se sube un archivo real.
    """

    def __init__(self, requerimiento_id, cursor):
        self.requerimiento_id = requerimiento_id
        self.cursor = cursor
        self.version_actual = self.obtener_ultima_version()

    # ============================================================
    # Obtener la última versión registrada
    # ============================================================
    def obtener_ultima_version(self):
        self.cursor.execute(
            """
            SELECT version
            FROM version_documento_tecnico
            WHERE requerimiento_documento_id = %s
            ORDER BY fecha DESC
            LIMIT 1
        """,
            [self.requerimiento_id],
        )
        row = self.cursor.fetchone()
        return row[0] if row else "v0.0.0"

    # ============================================================
    # Contar versiones con cierto sufijo (REV, REJREV, etc)
    # ============================================================
    def _count_suffix(self, token):
        self.cursor.execute(
            """
            SELECT COUNT(*) 
            FROM version_documento_tecnico
            WHERE requerimiento_documento_id = %s 
            AND version LIKE %s
        """,
            [self.requerimiento_id, f"%-{token}%"],
        )
        return self.cursor.fetchone()[0]

    # ============================================================
    # Generación de versiones según evento
    # ============================================================
    def nueva_version(self, evento):
        base = self.version_actual.split("-")[0]  # Ej: v1.0.0
        c1, c2, c3 = map(int, base.replace("v", "").split("."))

        if evento == "iniciar_elaboracion":
            return "v0.0.0-ELAB"

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

        if evento == "publicar_documento":
            return f"{base}-PUB"

        return self.version_actual

    # ============================================================
    # Registrar versión sin crear carpetas físicas
    # ============================================================
    def registrar_version(self, evento, estado_nombre, usuario_id, comentario):

        nueva_version = self.nueva_version(evento)

        # Registrar versión en la BD
        self.cursor.execute(
            """
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
        """,
            [
                self.requerimiento_id,
                nueva_version,
                estado_nombre,
                comentario,
                usuario_id,
            ],
        )

        self.version_actual = nueva_version
        return nueva_version

@login_required
def detalle_documento(request, requerimiento_id):
    from plantillas_documentos_tecnicos.views import (
        extraer_controles_contenido_desde_gcs,
        extraer_controles_contenido_desde_file,
    )
    import tempfile
    import os
    import traceback

    mensaje = ""

    # ============================================================
    #   DATOS PRINCIPALES DEL REQUERIMIENTO
    # ============================================================
    with connection.cursor() as cursor:
        cursor.execute(
            """
            SELECT
                RDT.id AS requerimiento_id,
                RDT.fecha_registro,
                RDT.observaciones,
                TDT.nombre AS tipo_documento,
                CDT.nombre AS categoria_documento,
                COALESCE(RR.id, 0) AS rol_id,
                COALESCE(RR.nombre, 'Sin Rol Asignado') AS rol_asignado,
                COALESCE(EA.nombre, 'Pendiente de Inicio') AS estado_actual,
                P.nombre AS nombre_proyecto
            FROM public.requerimiento_documento_tecnico RDT
            INNER JOIN public.tipo_documentos_tecnicos TDT 
                ON RDT.tipo_documento_id = TDT.id
            INNER JOIN public.categoria_documentos_tecnicos CDT 
                ON TDT.categoria_id = CDT.id
            LEFT JOIN public.requerimiento_equipo_rol RER
                ON RDT.id = RER.requerimiento_id 
               AND RER.usuario_id = %s 
               AND RER.activo = TRUE
            LEFT JOIN public.roles_ciclodocumento RR ON RER.rol_id = RR.id
            INNER JOIN public.proyectos P ON RDT.proyecto_id = P.id
            LEFT JOIN public.estado_documento EA ON EA.id = (
                SELECT estado_destino_id
                FROM public.log_estado_requerimiento_documento
                WHERE requerimiento_id = RDT.id
                ORDER BY fecha_cambio DESC
                LIMIT 1
            )
            WHERE RDT.id = %s
            LIMIT 1
            """,
            [request.user.id, requerimiento_id],
        )
        row = cursor.fetchone()
        documento = dict(zip([col[0] for col in cursor.description], row)) if row else None

    if not documento:
        messages.error(request, "❌ Documento no encontrado.")
        return redirect("documentos:lista_documentos_asignados")

    # ============================================================
    # HISTORIAL DE ESTADOS
    # ============================================================
    with connection.cursor() as cursor:
        cursor.execute(
            """
            SELECT 
                LER.fecha_cambio, 
                E.nombre AS estado_destino, 
                U.nombre AS usuario_nombre, 
                LER.observaciones AS comentario
            FROM public.log_estado_requerimiento_documento LER
            LEFT JOIN public.estado_documento E 
                ON LER.estado_destino_id = E.id
            LEFT JOIN public.usuarios U 
                ON LER.usuario_id = U.id
            WHERE LER.requerimiento_id = %s
            ORDER BY LER.fecha_cambio ASC
            """,
            [requerimiento_id],
        )
        historial_estados = [
            dict(zip([col[0] for col in cursor.description], row))
            for row in cursor.fetchall()
        ]

    # ============================================================
    # HISTORIAL VERSIONES
    # ============================================================
    with connection.cursor() as cursor:
        cursor.execute(
            """
            SELECT 
                VDT.version,
                E.nombre AS estado_nombre,
                VDT.fecha,
                U.nombre AS usuario_nombre,
                VDT.comentario,
                VDT.signed_url
            FROM public.version_documento_tecnico VDT
            LEFT JOIN public.estado_documento E 
                ON VDT.estado_id = E.id
            LEFT JOIN public.usuarios U 
                ON VDT.usuario_id = U.id
            WHERE VDT.requerimiento_documento_id = %s
            ORDER BY VDT.fecha ASC
            """,
            [requerimiento_id],
        )
        historial_versiones = [
            dict(zip([col[0] for col in cursor.description], row))
            for row in cursor.fetchall()
        ]

    # ============================================================
    # OBTENER PLANTILLA PORTADA / CUERPO (VERSIÓN REAL)
    # ============================================================
    plantilla_portada = None
    plantilla_cuerpo = None

    with connection.cursor() as cursor:
        cursor.execute(
            """
            SELECT 
                DG.ruta_gcs,

                -- columnas REALES
                DG.plantilla_documento_tecnico_version_id,
                DG.plantilla_utilidad_version_id,

                -- CUERPO
                PTDV.version AS version_cuerpo,
                PTDV.gcs_path AS plantilla_cuerpo_gcs,

                -- PORTADA
                PPV.version AS version_portada,
                PPV.gcs_path AS plantilla_portada_gcs

            FROM documentos_generados DG

            LEFT JOIN plantilla_tipo_doc_versiones PTDV
                ON DG.plantilla_documento_tecnico_version_id = PTDV.id

            LEFT JOIN plantilla_portada_versiones PPV
                ON DG.plantilla_utilidad_version_id = PPV.id

            WHERE DG.ruta_gcs LIKE %s
            """,
            [f"%RQ-{requerimiento_id}/%"],
        )

        for (
            ruta,
            cuerpo_ver_id,
            portada_ver_id,
            ver_cuerpo,
            cuerpo_gcs,
            ver_portada,
            portada_gcs,
        ) in cursor.fetchall():

            if cuerpo_ver_id:
                plantilla_cuerpo = {
                    "ruta": cuerpo_gcs,
                    "version": ver_cuerpo,
                }

            if portada_ver_id:
                plantilla_portada = {
                    "ruta": portada_gcs,
                    "version": ver_portada,
                }

    # ============================================================
    # ARCHIVO ENVIADO A REVISIÓN (última versión REV)
    # ============================================================
    archivo_revision = None
    with connection.cursor() as cursor:
        cursor.execute(
            """
            SELECT version, signed_url, fecha
            FROM version_documento_tecnico
            WHERE requerimiento_documento_id = %s
              AND version LIKE '%%REV%%'
            ORDER BY fecha DESC
            LIMIT 1
            """,
            [requerimiento_id],
        )
        row = cursor.fetchone()
        if row:
            archivo_revision = {"version": row[0], "url": row[1], "fecha": row[2]}

    # ============================================================
    # MÁQUINA DE ESTADOS
    # ============================================================
    estado_inicial = documento["estado_actual"] or "Pendiente de Inicio"
    rol_id = documento.get("rol_id") or 0
    machine = DocumentoTecnicoStateMachine(rol_id=rol_id, estado_inicial=estado_inicial)
    historial_simulador = request.session.get("historial_simulador", [])

    eventos_con_comentario = ["rechazar_revision", "rechazar_aprobacion"]
    eventos_con_archivo = [
        "enviar_revision",
        "reenviar_revision",
        "rechazar_revision",
        "rechazar_aprobacion",
    ]

    # ============================================================
    # POST — PROCESAR EVENTOS (VALIDACIÓN + TRANSICIÓN REAL)
    # ============================================================
    if request.method == "POST":

        evento = request.POST.get("evento")
        comentario = request.POST.get("comentario", "").strip()
        archivo = request.FILES.get("archivo")

        # -----------------------------
        # VALIDACIONES BÁSICAS
        # -----------------------------
        if evento in eventos_con_comentario and not comentario:
            mensaje = f"❌ Debes ingresar un comentario para '{evento}'."

        elif evento in eventos_con_archivo and not archivo:
            mensaje = f"❌ Debes adjuntar archivo para '{evento}'."

        elif not machine.puede_transicionar(evento):
            mensaje = f"❌ No tienes permiso para ejecutar '{evento}' desde el estado '{estado_inicial}'."

        else:
            errores_controles = False

            # =====================================================
            # VALIDACIÓN ESTRICTA DE CONTROLES (PORTADA + CUERPO)
            # =====================================================
            if evento in eventos_con_archivo and archivo:

                try:
                    # Guardar archivo temporal
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                        for chunk in archivo.chunks():
                            tmp.write(chunk)
                        tmp_path = tmp.name

                    # Controles del archivo subido
                    controles_archivo = set(
                        extraer_controles_contenido_desde_file(tmp_path) or []
                    )

                    # Controles esperados (portada + cuerpo)
                    controles_esperados = set()

                    if plantilla_portada and plantilla_portada.get("ruta"):
                        try:
                            controles_esperados.update(
                                extraer_controles_contenido_desde_gcs(
                                    plantilla_portada["ruta"]
                                )
                                or []
                            )
                        except Exception:
                            pass

                    if plantilla_cuerpo and plantilla_cuerpo.get("ruta"):
                        try:
                            controles_esperados.update(
                                extraer_controles_contenido_desde_gcs(
                                    plantilla_cuerpo["ruta"]
                                )
                                or []
                            )
                        except Exception:
                            pass

                    if controles_esperados:
                        faltantes = sorted(controles_esperados - controles_archivo)
                        sobrantes = sorted(controles_archivo - controles_esperados)

                        if faltantes or sobrantes:
                            errores_controles = True
                            msg_html = "❌ El archivo NO cumple:<br>"
                            if faltantes:
                                msg_html += "• Faltantes: " + ", ".join(faltantes) + "<br>"
                            if sobrantes:
                                msg_html += "• No esperados: " + ", ".join(sobrantes)
                            mensaje = msg_html
                            messages.error(request, mensaje)

                    # limpiar temporal
                    try:
                        os.unlink(tmp_path)
                    except Exception:
                        pass

                except Exception as e:
                    errores_controles = True
                    mensaje = f"⚠ Error validando controles: {e}"
                    messages.error(request, mensaje)

            # =====================================================
            # SI NO HAY ERRORES DE CONTROLES → EJECUTAR EVENTO
            # =====================================================
            if not errores_controles:
                try:
                    with transaction.atomic():
                        with connection.cursor() as cursor:

                            # -------------------
                            # 1) Ejecutar transición
                            # -------------------
                            getattr(machine, evento)()
                            nuevo_estado = machine.current_state.name

                            resultado = None

                            # -------------------
                            # 2) iniciar_elaboracion → crea estructura y plantillas
                            # -------------------
                            if evento == "iniciar_elaboracion":
                                from google.cloud import storage

                                storage_client = storage.Client()
                                bucket = storage_client.bucket("sgdmtso_jova")

                                resultado = inicializar_plantillas_requerimiento(
                                    cursor=cursor,
                                    bucket=bucket,
                                    requerimiento_id=requerimiento_id,
                                )

                            # -------------------
                            # 3) Registrar log de estado
                            # -------------------
                            cursor.execute(
                                """
                                INSERT INTO public.log_estado_requerimiento_documento
                                    (requerimiento_id, estado_destino_id, usuario_id, fecha_cambio, observaciones)
                                SELECT %s, id, %s, NOW(), %s
                                FROM public.estado_documento
                                WHERE nombre = %s
                                """,
                                [
                                    requerimiento_id,
                                    request.user.id,
                                    comentario or f"Evento: {evento}",
                                    nuevo_estado,
                                ],
                            )

                            # -------------------
                            # 4) Versionamiento
                            # -------------------
                            if machine.evento_genera_version(evento):

                                vm = VersionManager(
                                    requerimiento_id=requerimiento_id,
                                    cursor=cursor,
                                )

                                nueva_version = vm.registrar_version(
                                    evento,
                                    nuevo_estado,
                                    request.user.id,
                                    comentario or f"Evento: {evento}",
                                )

                                # 5) Subir archivo asociado (si corresponde)
                                if evento in eventos_con_archivo and archivo:
                                    cursor.execute(
                                        """
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
                                        """,
                                        [requerimiento_id],
                                    )
                                    row = cursor.fetchone()
                                    (
                                        proyecto,
                                        cliente,
                                        categoria,
                                        tipo,
                                    ) = row

                                    cliente = clean(cliente)
                                    proyecto = clean(proyecto)
                                    categoria = clean(categoria)
                                    tipo = clean(tipo)
                                    nombre_archivo = clean(archivo.name)

                                    ruta_final = (
                                        f"DocumentosProyectos/{cliente}/{proyecto}/"
                                        f"Documentos_Tecnicos/{categoria}/{tipo}/RQ-{requerimiento_id}/"
                                        f"{nueva_version}/{nombre_archivo}"
                                    )

                                    from google.cloud import storage
                                    from google.cloud import storage as gcs_storage

                                    storage_client = gcs_storage.Client()
                                    bucket = storage_client.bucket("sgdmtso_jova")

                                    mime_type, _ = mimetypes.guess_type(nombre_archivo)
                                    blob = bucket.blob(ruta_final)

                                    # IMPORTANTÍSIMO: volver al inicio antes de subir
                                    try:
                                        archivo.file.seek(0)
                                    except Exception:
                                        pass

                                    blob.upload_from_file(
                                        archivo.file,
                                        content_type=mime_type
                                        or "application/octet-stream",
                                    )

                                    signed_url = blob.generate_signed_url(
                                        version="v4",
                                        expiration=timedelta(days=7),
                                    )

                                    cursor.execute(
                                        """
                                        UPDATE version_documento_tecnico
                                        SET signed_url = %s
                                        WHERE requerimiento_documento_id = %s
                                          AND version = %s
                                        """,
                                        [signed_url, requerimiento_id, nueva_version],
                                    )

                    # -------------------
                    # Fuera de la transacción
                    # -------------------
                    if resultado:
                        for w in resultado.get("warnings", []):
                            messages.warning(request, w)

                    historial_simulador.append(
                        {
                            "evento": evento,
                            "nuevo_estado": nuevo_estado,
                            "comentario": comentario,
                            "timestamp": datetime.now().strftime(
                                "%Y-%m-%d %H:%M:%S"
                            ),
                        }
                    )
                    request.session["historial_simulador"] = historial_simulador

                    # Actualizar estado que se muestra en pantalla
                    documento["estado_actual"] = nuevo_estado

                    mensaje = f"✅ Evento '{evento}' ejecutado. Nuevo estado: {nuevo_estado}"
                    messages.success(request, mensaje)

                except Exception as e:
                    traceback.print_exc()
                    mensaje = f"❌ Error al ejecutar '{evento}': {str(e)}"
                    messages.error(request, mensaje)

    # ============================================================
    # EVENTOS DISPONIBLES
    # ============================================================
    todos_eventos = [
        "iniciar_elaboracion",
        "enviar_revision",
        "revision_aceptada",
        "aprobar_documento",
        "publicar_documento",
        "rechazar_revision",
        "rechazar_aprobacion",
        "reenviar_revision",
    ]
    eventos_disponibles = [
        ev for ev in todos_eventos if machine.puede_transicionar(ev)
    ]
    eventos_tuplas = list(
        zip(
            eventos_disponibles,
            [ev.replace("_", " ").capitalize() for ev in eventos_disponibles],
        )
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
    estado_css = colores_estado.get(documento["estado_actual"], "secondary")

    if not eventos_disponibles and not mensaje:
        mensaje = "⚙️ No tienes acciones disponibles para este estado."

    # ============================================================
    # CONTEXTO FINAL
    # ============================================================
    context = {
        "documento": documento,
        "estado_actual": documento["estado_actual"],
        "mensaje": mensaje,
        "estado_css": estado_css,
        "eventos_tuplas": eventos_tuplas,
        "historial_estados": historial_estados,
        "historial_versiones": historial_versiones,
        "historial_simulador": historial_simulador,
        "eventos_con_comentario": eventos_con_comentario,
        "plantilla_portada": plantilla_portada,
        "plantilla_cuerpo": plantilla_cuerpo,
        "archivo_revision": archivo_revision,
    }

    return render(request, "detalle_documento.html", context)


@login_required
def subir_archivo_documento(request, requerimiento_id):
    """
    Sube archivo a:
    DocumentosProyectos/Cliente/Proyecto/Documentos_Tecnicos/Categoria/Tipo/VERSION/Archivo.pdf
    """

    if request.method != "POST":
        return redirect(
            "documentos:detalle_documento", requerimiento_id=requerimiento_id
        )

    archivo = request.FILES.get("archivo")
    if not archivo:
        messages.error(request, "❌ Debes seleccionar un archivo.")
        return redirect(
            "documentos:detalle_documento", requerimiento_id=requerimiento_id
        )

    try:
        with connection.cursor() as cursor:

            # === Obtener ruta base del documento ===
            cursor.execute(
                """
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
            """,
                [requerimiento_id],
            )

            row = cursor.fetchone()
            if not row:
                messages.error(request, "❌ Documento no encontrado.")
                return redirect(
                    "documentos:detalle_documento",
                    requerimiento_id=requerimiento_id,
                )

            doc_id, proyecto, cliente, categoria, tipo = row

            # === Obtener versión actual ===
            vm = VersionManager(requerimiento_id=doc_id, cursor=cursor)
            version_actual = vm.version_actual or "Plantilla"

        cliente = clean(cliente)
        proyecto = clean(proyecto)
        categoria = clean(categoria)
        tipo = clean(tipo)
        nombre_archivo = clean(archivo.name)

        # === Ruta final ===
        ruta_final = (
            f"DocumentosProyectos/{cliente}/{proyecto}/"
            f"Documentos_Tecnicos/{categoria}/{tipo}/"
            f"{version_actual}/{nombre_archivo}"
        )

        # === Subir archivo a Google Cloud Storage ===
        from google.cloud import storage

        storage_client = storage.Client()
        bucket = storage_client.bucket("sgdmtso_jova")

        mime_type, _ = mimetypes.guess_type(archivo.name)
        blob = bucket.blob(ruta_final)

        # Asegurar puntero al inicio
        archivo.file.seek(0)

        blob.upload_from_file(
            archivo.file,
            content_type=mime_type or "application/octet-stream",
        )

        # === Generar URL firmada (v4 obligatorio) ===
        signed_url = blob.generate_signed_url(
            version="v4", expiration=timedelta(days=7)
        )

        # === Guardar URL firmada en la última versión ===
        with connection.cursor() as cursor:
            cursor.execute(
                """
                UPDATE version_documento_tecnico
                SET signed_url = %s
                WHERE requerimiento_documento_id = %s
                AND version = %s
            """,
                [signed_url, requerimiento_id, version_actual],
            )

        messages.success(
            request, f" Archivo subido correctamente a la versión {version_actual}"
        )

    except Exception as e:
        import traceback

        traceback.print_exc()
        messages.error(request, f"⚠ Error al subir archivo: {e}")

    return redirect("documentos:detalle_documento", requerimiento_id=requerimiento_id)






def inicializar_plantillas_requerimiento(cursor, bucket, requerimiento_id):
    """
    Crea estructura RQ-ID/Plantilla/, copia y procesa portada y cuerpo (si existen),
    registra SOLO 1 registro fusionado en documentos_generados,
    y devuelve un resumen.
    """

    # ==========================================================
    # Obtener datos del requerimiento
    # ==========================================================
    cursor.execute(
        """
        SELECT 
            P.nombre AS proyecto,
            CL.nombre AS cliente,
            CDT.nombre AS categoria,
            TDT.nombre AS tipo,
            TDT.id    AS tipo_id,
            F.id      AS formato_id,
            R.proyecto_id,
            R.tipo_documento_id
        FROM requerimiento_documento_tecnico R
        JOIN proyectos P ON R.proyecto_id = P.id
        JOIN contratos C ON P.contrato_id = C.id
        JOIN clientes CL ON C.cliente_id = CL.id
        JOIN tipo_documentos_tecnicos TDT ON R.tipo_documento_id = TDT.id
        JOIN categoria_documentos_tecnicos CDT ON TDT.categoria_id = CDT.id
        JOIN formato_archivo F ON TDT.formato_id = F.id
        WHERE R.id = %s
        """,
        [requerimiento_id],
    )

    (
        proyecto,
        cliente,
        categoria,
        tipo,
        tipo_id,
        formato_id,
        proyecto_id,
        tipo_doc_id,
    ) = cursor.fetchone()

    # Normalizar nombres
    cliente = clean(cliente)
    proyecto = clean(proyecto)
    categoria = clean(categoria)
    tipo = clean(tipo)

    # Ruta base
    base = (
        f"DocumentosProyectos/{cliente}/{proyecto}/"
        f"Documentos_Tecnicos/{categoria}/{tipo}/RQ-{requerimiento_id}/Plantilla/"
    )

    # Crear carpetas
    for sub in ["", "Portada/", "Portada/Word/", "Cuerpo/"]:
        bucket.blob(base + sub).upload_from_string("")

    resumen = {
        "cuerpo": None,
        "portada": None,
        "doc_fusionado_id": None,
        "warnings": [],
    }

    # ==========================================================
    # OBTENER ÚLTIMA VERSIÓN REAL DEL CUERPO
    # ==========================================================
    cursor.execute(
        """
        SELECT 
            v.id AS version_id,
            v.gcs_path
        FROM plantilla_tipo_doc p
        JOIN plantilla_tipo_doc_versiones v
            ON v.id = p.version_actual_id
        WHERE p.tipo_documento_id = %s
        ORDER BY v.creado_en DESC, v.id DESC
        LIMIT 1
        """,
        [tipo_id],
    )

    row = cursor.fetchone()
    if row:
        version_cuerpo_id, cuerpo_path = row
    else:
        version_cuerpo_id, cuerpo_path = None, None

    # Copia cuerpo al RQ
    if cuerpo_path:
        cuerpo_blob = bucket.blob(cuerpo_path)
        cuerpo_bytes = cuerpo_blob.download_as_bytes()

        cuerpo_nombre = cuerpo_path.split("/")[-1]
        new_cuerpo_path = base + "Cuerpo/" + cuerpo_nombre

        bucket.blob(new_cuerpo_path).upload_from_string(
            cuerpo_bytes,
            content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        resumen["cuerpo"] = new_cuerpo_path
    else:
        resumen["warnings"].append("⚠ No existe plantilla de cuerpo para este tipo.")

    # ==========================================================
    # OBTENER ÚLTIMA VERSIÓN REAL DE LA PORTADA (WORD)
    # ==========================================================
    cursor.execute(
        """
        SELECT 
            v.id AS version_id,
            v.gcs_path
        FROM plantilla_portada p
        JOIN plantilla_portada_versiones v
            ON v.id = p.version_actual_id
        WHERE p.utilidad_id = 1 AND p.formato_id = %s
        ORDER BY v.creado_en DESC, v.id DESC
        LIMIT 1
        """,
        [formato_id],
    )

    row = cursor.fetchone()
    if row:
        version_portada_id, portada_path = row
    else:
        version_portada_id, portada_path = None, None

    # Procesar portada
    if portada_path:

        portada_blob = bucket.blob(portada_path)
        portada_bytes = portada_blob.download_as_bytes()

        from io import BytesIO
        from .docx_filler import process_template_docx

        # ================== DATOS DEL REQUERIMIENTO ==================
        cursor.execute(
            """
            SELECT 
                P.nombre AS proyecto,
                CL.nombre AS cliente,
                F.nombre AS faena,
                C.numero_contrato,
                P.numero_servicio,
                U.nombre AS administrador_servicio,
                TDT.nombre AS tipo_documento,
                RDT.codigo_documento
            FROM requerimiento_documento_tecnico RDT
            JOIN proyectos P ON RDT.proyecto_id = P.id
            JOIN contratos C ON P.contrato_id = C.id
            JOIN clientes CL ON C.cliente_id = CL.id
            JOIN faenas F ON P.faena_id = F.id
            JOIN tipo_documentos_tecnicos TDT ON RDT.tipo_documento_id = TDT.id
            LEFT JOIN usuarios U ON U.id = P.administrador_id
            WHERE RDT.id = %s
            """,
            [requerimiento_id],
        )

        (
            nombre_proyecto,
            nombre_cliente,
            nombre_faena,
            numero_contrato,
            numero_servicio,
            administrador_servicio,
            tipo_documento,
            codigo_documento,
        ) = cursor.fetchone()

        # Roles
        cursor.execute(
            """
            SELECT rol_id, U.nombre
            FROM requerimiento_equipo_rol RER
            JOIN usuarios U ON U.id = RER.usuario_id
            WHERE requerimiento_id = %s AND activo = TRUE
            """,
            [requerimiento_id],
        )
        roles_data = cursor.fetchall()

        equipo = {
            "redactores_equipo": ", ".join([r[1] for r in roles_data if r[0] == 1]),
            "revisores_equipo": ", ".join([r[1] for r in roles_data if r[0] == 2]),
            "aprobadores_equipo": ", ".join([r[1] for r in roles_data if r[0] == 3]),
        }

        # Historial
        cursor.execute(
            """
            SELECT version, fecha, comentario,
            (SELECT nombre FROM estado_documento WHERE id = estado_id)
            FROM version_documento_tecnico
            WHERE requerimiento_documento_id = %s
            ORDER BY fecha ASC
            """,
            [requerimiento_id],
        )

        historial = []
        for v, fecha, comentario, estado in cursor.fetchall():
            historial.append(
                {
                    "h.version": v,
                    "h.estado": estado,
                    "h.fecha": fecha.strftime("%d-%m-%Y %H:%M"),
                    "h.comentario": comentario or "",
                }
            )

        # ==================================================================
        # Rellenar portada con datos reales del RQ
        # ==================================================================
        simple_data = {
            "tipo_documento": tipo_documento,
            "codigo_documento": codigo_documento,
            "nombre_proyecto": nombre_proyecto,
            "nombre_cliente": nombre_cliente,
            "nombre_faena": nombre_faena,
            "numero_contrato": numero_contrato,
            "numero_servicio": numero_servicio,
            "administrador_servicio": administrador_servicio,
        }
        simple_data.update(equipo)

        portada_final_bytes = process_template_docx(
            template_bytes=BytesIO(portada_bytes),
            simple_data=simple_data,
            historial_versiones=historial,
        )

        # Subir portada generada
        new_portada_path = base + "Portada/Word/" + portada_path.split("/")[-1]

        bucket.blob(new_portada_path).upload_from_string(
            portada_final_bytes,
            content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        resumen["portada"] = new_portada_path

    else:
        resumen["warnings"].append("⚠ No existe portada Word para este formato.")

    # ==========================================================
    # INSERTAR REGISTRO (VERSIÓN REAL)
    # ==========================================================
    cursor.execute(
        """
        INSERT INTO documentos_generados
            (proyecto_id,
             tipo_documento_id,
             ruta_gcs,
             fecha_generacion,
             plantilla_documento_tecnico_version_id,
             formato_id,
             plantilla_utilidad_version_id
            )
        VALUES (%s, %s, %s, NOW(), %s, %s, %s)
        RETURNING id
        """,
        [
            proyecto_id,
            tipo_doc_id,
            resumen["cuerpo"],       # ruta final del cuerpo
            version_cuerpo_id,       # ✔ versión REAL del CUERPO
            formato_id,
            version_portada_id,      # ✔ versión REAL de la PORTADA
        ],
    )

    resumen["doc_fusionado_id"] = cursor.fetchone()[0]

    return resumen





@login_required
def prevalidar_controles(request, requerimiento_id):
    """
    Paso 1 del flujo A:
    El usuario sube un archivo y pide 'Ver comparación'.
    NO ejecuta el evento aún.
    Solo muestra controles esperados vs controles reales.
    """

    from plantillas_documentos_tecnicos.views import (
        extraer_controles_contenido_desde_gcs,
        extraer_controles_contenido_desde_file,
    )
    import tempfile, os

    # ==============================
    # DATOS NECESARIOS
    # ==============================
    evento = request.POST.get("evento")
    archivo = request.FILES.get("archivo")

    if not archivo:
        messages.error(request, "❌ Debes seleccionar un archivo.")
        return redirect(
            "documentos:detalle_documento", requerimiento_id=requerimiento_id
        )

    # Guardar archivo temporal
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        for chunk in archivo.chunks():
            tmp.write(chunk)
        tmp_path = tmp.name

    # Controles que trae el archivo del usuario
    controles_subido = set(extraer_controles_contenido_desde_file(tmp_path))

    os.unlink(tmp_path)

    # ==============================
    # OBTENER PLANTILLAS ORIGINAL DEL RQ
    # ==============================
    plantilla_portada = None
    plantilla_cuerpo = None

    with connection.cursor() as cursor:
        cursor.execute(
            """
            SELECT 
                DG.plantilla_documento_tecnico_id,
                DG.plantilla_utilidad_id,
                PDT.gcs_path,
                PUT.gcs_path
            FROM documentos_generados DG
            LEFT JOIN plantilla_tipo_doc PDT 
                ON DG.plantilla_documento_tecnico_id = PDT.id
            LEFT JOIN plantilla_portada PUT
                ON DG.plantilla_utilidad_id = PUT.id
            WHERE DG.ruta_gcs LIKE %s
        """,
            [f"%RQ-{requerimiento_id}/%"],
        )

        for cuerpo_id, portada_id, cuerpo_path, portada_path in cursor.fetchall():
            if cuerpo_id:
                plantilla_cuerpo = cuerpo_path
            if portada_id:
                plantilla_portada = portada_path

    # ==============================
    # EXTRAER CONTROLES ESPERADOS
    # ==============================
    controles_portada = set()
    controles_cuerpo = set()

    if plantilla_portada:
        controles_portada = set(
            extraer_controles_contenido_desde_gcs(plantilla_portada)
        )

    if plantilla_cuerpo:
        controles_cuerpo = set(
            extraer_controles_contenido_desde_gcs(plantilla_cuerpo)
        )

    controles_esperados = controles_portada | controles_cuerpo

    # ==============================
    # COMPARACIÓN COMPLETA
    # ==============================
    faltantes = sorted(controles_esperados - controles_subido)
    sobrantes = sorted(controles_subido - controles_esperados)
    coincidencias = sorted(controles_esperados & controles_subido)

    # ==============================
    # Render comparativo
    # ==============================
    context = {
        "requerimiento_id": requerimiento_id,
        "evento": evento,
        "archivo": archivo,
        "controles_subido": sorted(controles_subido),
        "controles_portada": sorted(controles_portada),
        "controles_cuerpo": sorted(controles_cuerpo),
        "coincidencias": coincidencias,
        "faltantes": faltantes,
        "sobrantes": sobrantes,
    }

    return render(request, "comparacion_controles.html", context)

@login_required
@require_POST
def validar_controles_doc_ajax(request, requerimiento_id):
    from plantillas_documentos_tecnicos.views import (
        extraer_controles_contenido_desde_gcs,
        extraer_controles_contenido_desde_file,
    )
    import tempfile, os

    archivo = request.FILES.get("archivo")
    if not archivo:
        return JsonResponse({"error": "No se recibió archivo."})

    # ============================================================
    # 1) Extraer controles del archivo subido
    # ============================================================
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        for chunk in archivo.chunks():
            tmp.write(chunk)
        tmp_path = tmp.name

    controles_archivo = set(extraer_controles_contenido_desde_file(tmp_path))

    try:
        os.unlink(tmp_path)
    except:
        pass

    # ============================================================
    # 2) Obtener IDs de VERSION usados en el RQ (CORREGIDO)
    # ============================================================
    with connection.cursor() as cursor:
        cursor.execute(
            """
            SELECT 
                DG.plantilla_documento_tecnico_version_id,
                DG.plantilla_utilidad_version_id
            FROM documentos_generados DG
            WHERE DG.ruta_gcs LIKE %s
            LIMIT 1
            """,
            [f"%RQ-{requerimiento_id}/%"],
        )

        row = cursor.fetchone()

        if not row:
            return JsonResponse({"error": "No se pudo obtener metadatos del requerimiento."})

        cuerpo_version_id, portada_version_id = row

    # ============================================================
    # 3) Obtener rutas GCS de las versiones reales (CORREGIDO)
    # ============================================================
    ruta_cuerpo = None
    if cuerpo_version_id:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT gcs_path
                FROM plantilla_tipo_doc_versiones
                WHERE id = %s
                """,
                [cuerpo_version_id],
            )
            row = cursor.fetchone()
            ruta_cuerpo = row[0] if row else None

    ruta_portada = None
    if portada_version_id:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT gcs_path
                FROM plantilla_portada_versiones
                WHERE id = %s
                """,
                [portada_version_id],
            )
            row = cursor.fetchone()
            ruta_portada = row[0] if row else None

    # ============================================================
    # 4) Extraer controles esperados (portada + cuerpo)
    # ============================================================
    controles_esperados = set()

    controles_cuerpo = []
    controles_portada = []

    try:
        if ruta_cuerpo:
            controles_cuerpo = extraer_controles_contenido_desde_gcs(ruta_cuerpo)
            controles_esperados.update(controles_cuerpo)

        if ruta_portada:
            controles_portada = extraer_controles_contenido_desde_gcs(ruta_portada)
            controles_esperados.update(controles_portada)

        # ================================
        # 5) Comparación por origen
        # ================================
        set_cuerpo = set(controles_cuerpo)
        set_portada = set(controles_portada)

        coincidencias_cuerpo = sorted(controles_archivo & set_cuerpo)
        faltantes_cuerpo = sorted(set_cuerpo - controles_archivo)
        sobrantes_cuerpo = sorted(controles_archivo - set_cuerpo)

        coincidencias_portada = sorted(controles_archivo & set_portada)
        faltantes_portada = sorted(set_portada - controles_archivo)
        sobrantes_portada = sorted(controles_archivo - set_portada)

        # ================================
        # 6) Combinados
        # ================================
        coincidencias = sorted(set(coincidencias_cuerpo) | set(coincidencias_portada))
        faltantes = sorted(set(faltantes_cuerpo) | set(faltantes_portada))
        sobrantes = sorted(set(sobrantes_cuerpo) | set(sobrantes_portada))

    except Exception as e:
        return JsonResponse({"error": f"Error al procesar controles: {str(e)}"})

    # ============================================================
    # 7) Respuesta final → Compatible con todo tu JS actual
    # ============================================================
    return JsonResponse({
        # Combinados
        "coincidencias": coincidencias,
        "faltantes": faltantes,
        "sobrantes": sobrantes,

        # Separados por origen
        "coincidencias_portada": coincidencias_portada,
        "coincidencias_cuerpo": coincidencias_cuerpo,
        "faltantes_portada": faltantes_portada,
        "faltantes_cuerpo": faltantes_cuerpo,
        "sobrantes_portada": sobrantes_portada,
        "sobrantes_cuerpo": sobrantes_cuerpo,

        # Debug
        "controles_archivo": list(controles_archivo),
        "controles_esperados": list(controles_esperados),
        "controles_portada": controles_portada,
        "controles_cuerpo": controles_cuerpo,
        "ruta_cuerpo": ruta_cuerpo,
        "ruta_portada": ruta_portada,
    })
