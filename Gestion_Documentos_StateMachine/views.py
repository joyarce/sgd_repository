from statemachine import StateMachine, State
from django.http import HttpResponse
from django.shortcuts import render
from django.contrib.auth.decorators import login_required
from django.db import connection



# ---------------------------
# Máquina de estados
# ---------------------------
class DocumentoTecnicoStateMachine(StateMachine):
    """Máquina de estados para simular el ciclo de vida del documento técnico."""

    # Estados
    borrador = State("Borrador", initial=True)
    en_elaboracion = State("En Elaboración")
    en_revision = State("En Revisión")
    revisado = State("Revisado")
    aprobado = State("Aprobado. Listo para Publicación")
    publicado = State("Publicado")

    # Transiciones
    comenzar_elaboracion = borrador.to(en_elaboracion)
    enviar_revision = en_elaboracion.to(en_revision)
    aprobar_revision = en_revision.to(revisado)
    aprobar_documento = revisado.to(aprobado)
    publicar_documento = aprobado.to(publicado)

    # Rechazos
    rechazar_revision = en_revision.to(en_elaboracion)
    rechazar_aprobacion = aprobado.to(en_elaboracion)

    def __init__(self, rol_id=1, estado_inicial=None):
        super().__init__()
        self.rol_id = rol_id
        if estado_inicial:
            for s in self.states:
                if s.name == estado_inicial:
                    self.current_state = s
                    break

    def puede_transicionar(self, evento):
        permisos = {
            "comenzar_elaboracion": [1],  # Redactor
            "enviar_revision": [1],
            "rechazar_revision": [2],     # Revisor
            "aprobar_revision": [2],
            "aprobar_documento": [3],     # Aprobador
            "rechazar_aprobacion": [3],
            "publicar_documento": [3],    # Solo Aprobador
        }

        # Evitar publicar si no está aprobado
        if evento == "publicar_documento" and self.current_state != self.aprobado:
            return False

        return self.rol_id in permisos.get(evento, [])


# ---------------------------
# Vista: cambiar estado de un requerimiento específico
# ---------------------------
def cambiar_estado_requerimiento(request, requerimiento_id, evento):
    rol_id = int(request.GET.get("rol", 1))
    estado_actual = request.GET.get("estado", "Borrador")
    machine = DocumentoTecnicoStateMachine(rol_id=rol_id, estado_inicial=estado_actual)

    try:
        machine.ejecutar_evento(evento)
    except PermissionError as e:
        return HttpResponse(f"❌ {str(e)}")
    except Exception as e:
        return HttpResponse(f"⚠️ Error: {str(e)}")

    return HttpResponse(
        f"Simulación para requerimiento {requerimiento_id}: "
        f"Evento '{evento}' ejecutado. Nuevo estado: {machine.current_state.name}"
    )


# ---------------------------
# Vista: simulador web
# ---------------------------
@login_required
def simulador_estado(request):
    estado_actual = request.GET.get("estado", "Borrador")
    evento = request.GET.get("evento")
    rol_id = int(request.GET.get("rol", 1))

    machine = DocumentoTecnicoStateMachine(rol_id=rol_id, estado_inicial=estado_actual)
    mensaje = ""

    if evento:
        try:
            machine.ejecutar_evento(evento)
            mensaje = f"✅ Transición '{evento}' ejecutada. Nuevo estado: {machine.current_state.name}"
        except Exception as e:
            mensaje = f"⚠️ Error: {str(e)}"

    # Todos los eventos
    todos_eventos = [
        "comenzar_elaboracion",
        "enviar_revision",
        "aprobar_revision",
        "aprobar_documento",
        "publicar_documento",
        "rechazar_revision",
        "rechazar_aprobacion",
    ]

    # Mensajes para mostrar por qué no están disponibles
    mensajes_estado = {}
    for ev in todos_eventos:
        if not machine.puede_transicionar(ev):
            if ev == "publicar_documento":
                mensajes_estado[ev] = "No se puede publicar sin aprobar el documento primero."
            else:
                mensajes_estado[ev] = "No tienes permiso para ejecutar este evento."

    # Solo eventos permitidos
    eventos_disponibles = [ev for ev in todos_eventos if machine.puede_transicionar(ev)]

    return render(request, "simulador_estado.html", {
        "estado_actual": machine.current_state.name,
        "mensaje": mensaje,
        "eventos": eventos_disponibles,
        "rol_id": rol_id,
        "todos_eventos": todos_eventos,
        "mensajes_estado": mensajes_estado,
    })













#################

@login_required
def lista_documentos_asignados(request):
    """
    Lista todos los requerimientos de documentos técnicos asignados
    al usuario logueado, mostrando su tipo, estado y fecha de registro.
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
    )
    SELECT
        RDT.id AS requerimiento_id,
        RDT.fecha_registro,
        RDT.observaciones,
        TDT.nombre AS tipo_documento,
        E.nombre AS estado_actual,
        RDT.proyecto_id,
        P.nombre AS nombre_proyecto
    FROM public.requerimiento_documento_tecnico RDT
    INNER JOIN public.tipo_documentos_tecnicos TDT
        ON RDT.tipo_documento_id = TDT.id
    LEFT JOIN UltimoEstado UE
        ON RDT.id = UE.requerimiento_id AND UE.rn = 1
    LEFT JOIN public.estado_documento E
        ON UE.estado_destino_id = E.id
    INNER JOIN public.proyectos P
        ON RDT.proyecto_id = P.id
    -- Filtrar por usuario asignado si equipo_trabajo_id corresponde a usuario
    LEFT JOIN public.requerimiento_equipo_rol RER
        ON RDT.equipo_trabajo_id = RER.id
    WHERE RER.usuario_id = %s
    ORDER BY RDT.fecha_registro DESC;
    """

    with connection.cursor() as cursor:
        cursor.execute(sql, [user_id])
        columns = [col[0] for col in cursor.description]
        resultados = [dict(zip(columns, row)) for row in cursor.fetchall()]

    context = {
        'documentos': resultados
    }

    return render(request, "lista_documentos_asignados.html", context)

