from .state_machine import DocumentoTecnicoStateMachine
from django.http import HttpResponse
from datetime import datetime, timedelta
import json
import random
from django.shortcuts import render
from django.contrib.auth.decorators import login_required
from django.db import connection
from django.contrib import messages
from django.shortcuts import render, redirect




# ---------------------------
# Vista: cambiar estado de un requerimiento específico
# ---------------------------
def cambiar_estado_requerimiento(request, requerimiento_id, evento):
    rol_id = int(request.GET.get("rol", 1))
    estado_actual = request.GET.get("estado", "Borrador")
    machine = DocumentoTecnicoStateMachine(rol_id=rol_id, estado_inicial=estado_actual)

    try:
        machine.trigger(evento)
    except PermissionError as e:
        return HttpResponse(f"❌ {str(e)}")
    except Exception as e:
        return HttpResponse(f"⚠️ Error: {str(e)}")

    return HttpResponse(
        f"Simulación para requerimiento {requerimiento_id}: "
        f"Evento '{evento}' ejecutado. Nuevo estado: {machine.current_state.name}"
    )
   


@login_required
def simulador_estado(request):
    estado_inicial = request.GET.get("estado", "Borrador")
    evento = request.GET.get("evento")
    rol_id = int(request.GET.get("rol", 1))

    machine = DocumentoTecnicoStateMachine(rol_id=rol_id, estado_inicial=estado_inicial)
    mensaje = ""
    
    if "historial" not in request.session:
        request.session["historial"] = []
        # Si se solicita reiniciar historial

    if evento:
        try:
            # Llamamos al evento dinámicamente
            event_method = getattr(machine, evento)
            event_method()
            mensaje = f"✅ Transición '{evento}' ejecutada. Nuevo estado: {machine.current_state.name}"

            # Guardamos en historial
            historial.append({
                "evento": evento,
                "nuevo_estado": machine.current_state.name
            })
            request.session["historial"] = historial

        except PermissionError as e:
            mensaje = f"❌ {str(e)}"
        except Exception as e:
            mensaje = f"⚠️ Error al ejecutar '{evento}': {str(e)}"

    # Todos los eventos posibles
    todos_eventos = [
        "crear_documento",
        "enviar_revision",
        "revision_aceptada",
        "aprobar_documento",
        "publicar_documento",
        "rechazar_revision",
        "rechazar_aprobacion",
        "reenviar_revision",
    ]
    
    mensajes_estado = {}
    eventos_disponibles = []

    for ev in todos_eventos:
        if machine.puede_transicionar(ev):
            eventos_disponibles.append(ev)
        else:
            mensajes_estado[ev] = "No tienes permiso para ejecutar este evento."


    # Eventos disponibles según rol y estado
    eventos_disponibles = [ev for ev in todos_eventos if machine.puede_transicionar(ev)]

    return render(request, "simulador_estado.html", {
        "estado_actual": machine.current_state.name,
        "mensaje": mensaje,
        "eventos": eventos_disponibles,
        "rol_id": rol_id,
        "todos_eventos": todos_eventos,
        "mensajes_estado": mensajes_estado,
        "historial": historial,
    })




#################
@login_required
def lista_documentos_asignados(request):
    """
    Lista los documentos asignados al usuario logueado mostrando solo
    aquellos donde su rol tiene transiciones posibles según el estado actual,
    y agrupa los documentos por proyecto para el template.
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
            COALESCE(E.nombre, 'Borrador') AS estado_actual
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
    INNER JOIN EstadoActual EA
        ON EA.requerimiento_id = RDT.id
    INNER JOIN public.tipo_documentos_tecnicos TDT
        ON RDT.tipo_documento_id = TDT.id
    INNER JOIN public.categoria_documentos_tecnicos CDT
        ON TDT.categoria_id = CDT.id
    INNER JOIN public.proyectos P
        ON RDT.proyecto_id = P.id
    INNER JOIN public.requerimiento_equipo_rol RER
        ON RDT.id = RER.requerimiento_id
    INNER JOIN public.roles_ciclodocumento RR
        ON RER.rol_id = RR.id
    WHERE RER.usuario_id = %s
      AND (
            EXISTS (
                SELECT 1
                FROM public.transiciones_permitidas TP
                WHERE TP.estado_origen_id = EA.estado_id
                  AND TP.rol_id = RR.id
            )
            OR (EA.estado_actual = 'Borrador' AND RR.nombre ILIKE 'Redactor')
      )
    ORDER BY P.nombre, RDT.fecha_registro DESC;
    """

    with connection.cursor() as cursor:
        cursor.execute(sql, [user_id])
        columns = [col[0] for col in cursor.description]
        resultados = [dict(zip(columns, row)) for row in cursor.fetchall()]

    # --- Agrupar documentos por proyecto ---
    documentos_por_proyecto = {}
    for doc in resultados:
        proj_name = doc["nombre_proyecto"]
        documentos_por_proyecto.setdefault(proj_name, []).append(doc)

    # --- KPIs y gráficos ---
    total_docs = len(resultados)

    # Distribución por estado
    por_estado = {}
    for doc in resultados:
        por_estado[doc["estado_actual"]] = por_estado.get(doc["estado_actual"], 0) + 1
    chart_estado = json.dumps({"labels": list(por_estado.keys()), "values": list(por_estado.values())})

    # Distribución por rol
    por_rol = {}
    for doc in resultados:
        por_rol[doc["rol_asignado"]] = por_rol.get(doc["rol_asignado"], 0) + 1
    chart_rol = json.dumps({"labels": list(por_rol.keys()), "values": list(por_rol.values())})

    # Actividad últimos 7 días
    dias = [(datetime.now() - timedelta(days=i)).strftime("%d-%b") for i in reversed(range(7))]
    actividad = [
        sum(1 for doc in resultados if doc["fecha_registro"].date() == (datetime.now() - timedelta(days=i)).date())
        for i in reversed(range(7))
    ]
    chart_actividad = json.dumps({"labels": dias, "values": actividad})

    # Simulación para tiempos y radar
    etapas = ["Redacción", "Revisión", "Aprobación", "Publicación"]
    chart_tiempos = json.dumps({"labels": etapas, "values": [round(random.uniform(4, 18), 1) for _ in etapas]})
    chart_radar = json.dumps({"labels": etapas, "values": [random.randint(50, 100) for _ in etapas]})

    # Cumplimiento estimado (simulado)
    cumplimiento = random.randint(60, 95)

    # --- Mapeo de estado a clase de color Bootstrap ---
    colores_estado = {
        "Borrador": "secondary",
        "En Elaboración": "info",
        "En Revisión": "warning",
        "En Aprobación": "primary",
        "Aprobado": "success",
        "Publicado": "success",
        "Re Estructuración": "danger",
    }



    context = {
        "documentos_por_proyecto": documentos_por_proyecto,
        "total_docs": total_docs,
        "cumplimiento": cumplimiento,
        "chart_estado": chart_estado,
        "colores_estado": colores_estado,
        "chart_rol": chart_rol,
        "chart_actividad": chart_actividad,
        "chart_tiempos": chart_tiempos,
        "chart_radar": chart_radar,
    }

    return render(request, "lista_documentos_asignados.html", context)

def generar_signed_url(documento_id, version):
    # Aquí va la lógica de tu storage para obtener URL firmada
    # Por ahora podemos simular
    return f"https://storage.simulado.com/doc_{documento_id}_{version}.pdf"


class VersionManager:
    def __init__(self, requerimiento_id, cursor):
        self.requerimiento_id = requerimiento_id
        self.cursor = cursor
        self.version_actual = self.obtener_ultima_version()

    def obtener_ultima_version(self):
        """Obtiene la última versión registrada para este documento"""
        self.cursor.execute("""
            SELECT version
            FROM version_documento_tecnico
            WHERE requerimiento_documento_id = %s
            ORDER BY fecha DESC
            LIMIT 1
        """, [self.requerimiento_id])
        row = self.cursor.fetchone()
        return row[0] if row else "v0.0.0"

    def nueva_version(self, evento):
        """Calcula la nueva versión según la lógica de FSM y eventos"""
        base, *sufijo = self.version_actual.split("-")
        c1, c2, c3 = map(int, base.strip("v").split("."))

        if evento == "revision_aceptada":
            c1 += 1
            c2 = 0
            c3 = 0
            sufijo = []
        elif evento == "rechazar_revision":
            sufijo = ["RR"]
        elif evento == "reenviar_revision":
            c2 += 1
        elif evento == "aprobar_documento":
            c3 += 1
            sufijo = []
        elif evento == "rechazar_aprobacion":
            sufijo = ["RA"]

        nueva = f"v{c1}.{c2}.{c3}"
        if sufijo:
            nueva += f"-{'-'.join(sufijo)}"
        self.version_actual = nueva
        return nueva

    def registrar_version(self, evento, estado_nombre, usuario_id, comentario):
        nueva_version = self.nueva_version(evento)
        url = generar_signed_url(self.requerimiento_id, nueva_version)
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
                %s
            )
        """, [self.requerimiento_id, nueva_version, estado_nombre, comentario, usuario_id, url])
        return nueva_version



@login_required
def detalle_documento(request, requerimiento_id):
    mensaje = ""
    
    # --- Obtener datos del documento ---
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT
                RDT.id AS requerimiento_id,
                RDT.fecha_registro,
                RDT.observaciones,
                TDT.nombre AS tipo_documento,
                CDT.nombre AS categoria_documento,
                RR.id AS rol_id,
                RR.nombre AS rol_asignado,
                COALESCE(EA.nombre, 'Borrador') AS estado_actual,
                P.nombre AS nombre_proyecto
            FROM public.requerimiento_documento_tecnico RDT
            INNER JOIN public.tipo_documentos_tecnicos TDT ON RDT.tipo_documento_id = TDT.id
            INNER JOIN public.categoria_documentos_tecnicos CDT ON TDT.categoria_id = CDT.id
            LEFT JOIN public.requerimiento_equipo_rol RER 
                ON RDT.id = RER.requerimiento_id AND RER.usuario_id = %s
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
        """, [request.user.id, requerimiento_id])
        row = cursor.fetchone()
        documento = dict(zip([col[0] for col in cursor.description], row)) if row else None

    if not documento:
        messages.error(request, "❌ Documento no encontrado.")
        return redirect("lista_documentos_asignados")

    # --- Historial de estados ---
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT 
                LER.fecha_cambio, 
                E.nombre AS estado_destino, 
                U.nombre AS usuario_nombre, 
                LER.observaciones AS comentario
            FROM public.log_estado_requerimiento_documento LER
            LEFT JOIN public.estado_documento E ON LER.estado_destino_id = E.id
            LEFT JOIN public.usuarios U ON LER.usuario_id = U.id
            WHERE LER.requerimiento_id = %s
            ORDER BY LER.fecha_cambio ASC
        """, [requerimiento_id])
        historial_estados = [
            dict(zip([col[0] for col in cursor.description], row))
            for row in cursor.fetchall()
        ]

    # --- Historial de versiones ---
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT 
                VDT.version,
                E.nombre AS estado_nombre,
                VDT.fecha,
                U.nombre AS usuario_nombre,
                VDT.comentario,
                VDT.signed_url
            FROM public.version_documento_tecnico VDT
            LEFT JOIN public.estado_documento E ON VDT.estado_id = E.id
            LEFT JOIN public.usuarios U ON VDT.usuario_id = U.id
            WHERE VDT.requerimiento_documento_id = %s
            ORDER BY VDT.fecha ASC
        """, [requerimiento_id])
        historial_versiones = [
            dict(zip([col[0] for col in cursor.description], row))
            for row in cursor.fetchall()
        ]

    # --- Máquina de estados ---
    estado_inicial = documento["estado_actual"] or "Borrador"
    rol_id = documento.get("rol_id")
    machine = DocumentoTecnicoStateMachine(rol_id=rol_id, estado_inicial=estado_inicial)
    historial_simulador = request.session.get("historial_simulador", [])

    # --- Manejo de POST ---
    if request.method == "POST":
        evento = request.POST.get("evento")
        comentario = request.POST.get("comentario", "").strip()

        if evento:
            if evento in ["rechazar_revision", "rechazar_aprobacion"] and not comentario:
                mensaje = f"❌ Debes ingresar un comentario para ejecutar '{evento}'."
            elif not machine.puede_transicionar(evento):
                mensaje = f"❌ No tienes permiso para ejecutar '{evento}' desde el estado '{estado_inicial}'."
            else:
                try:
                    getattr(machine, evento)()
                    nuevo_estado = machine.current_state.name
                    documento["estado_actual"] = nuevo_estado

                    with connection.cursor() as cursor:
                        # Registrar log de estado
                        cursor.execute("""
                            INSERT INTO public.log_estado_requerimiento_documento
                                (requerimiento_id, estado_destino_id, usuario_id, fecha_cambio, observaciones)
                            SELECT %s, id, %s, NOW(), %s
                            FROM public.estado_documento
                            WHERE nombre = %s
                        """, [requerimiento_id, request.user.id, comentario or f"Evento: {evento}", nuevo_estado])

                        # Registrar versión si aplica
                        if machine.evento_genera_version(evento):
                            vm = VersionManager(requerimiento_id=requerimiento_id, cursor=cursor)
                            vm.registrar_version(evento, nuevo_estado, request.user.id, comentario or f"Evento: {evento}")

                    historial_simulador.append({
                        "evento": evento,
                        "nuevo_estado": nuevo_estado,
                        "comentario": comentario,
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    })
                    request.session["historial_simulador"] = historial_simulador
                    mensaje = f"✅ Evento '{evento}' ejecutado. Nuevo estado: {nuevo_estado}"

                except Exception as e:
                    mensaje = f"❌ Error al ejecutar '{evento}': {str(e)}"

    # --- Reset historial ---
    if request.GET.get("reset_historial"):
        request.session["historial_simulador"] = []
        historial_simulador = []
        mensaje = "🗑 Historial de simulación reiniciado."

    # --- Eventos disponibles ---
    todos_eventos = [
        "crear_documento", "enviar_revision", "revision_aceptada",
        "aprobar_documento", "publicar_documento",
        "rechazar_revision", "rechazar_aprobacion", "reenviar_revision"
    ]
    eventos = [ev for ev in todos_eventos if machine.puede_transicionar(ev)]
    eventos_formateados = [ev.replace("_", " ").capitalize() for ev in eventos]
    eventos_tuplas = list(zip(eventos, eventos_formateados))
    eventos_con_comentario = ["rechazar_revision", "rechazar_aprobacion"]

    # --- Mapeo de estado a color Bootstrap ---
    colores_estado = {
        "Borrador": "secondary",
        "En Elaboración": "info",
        "En Revisión": "warning",
        "En Aprobación": "primary",
        "Aprobado": "success",
        "Publicado": "success",
        "Re Estructuración": "danger",
    }
    estado_css = colores_estado.get(documento["estado_actual"], "secondary")

    context = {
        "documento": documento,
        "estado_actual": documento["estado_actual"],
        "mensaje": mensaje,
        "estado_css": estado_css,
        "eventos_tuplas": eventos_tuplas,
        "historial_estados": historial_estados,
        "historial_versiones": historial_versiones,  # <-- Aquí se pasa al template
        "historial_simulador": historial_simulador,
        "eventos_con_comentario": eventos_con_comentario,
    }

    return render(request, "detalle_documento.html", context)