# state_machine.py C:\Users\jonat\Documents\gestion_docs\Gestion_Documentos_StateMachine\state_machine.py
from statemachine import StateMachine, State

class DocumentoTecnicoStateMachine(StateMachine):
    # === ESTADOS ===
    pendiente_inicio = State("Pendiente de Inicio", initial=True)
    en_elaboracion = State("En Elaboración")
    en_revision = State("En Revisión")
    en_aprobacion = State("En Aprobación")
    re_estructuracion = State("Re Estructuración")
    aprobado = State("Aprobado. Listo para Publicación")
    publicado = State("Publicado", final=True)

    # === TRANSICIONES ===
    iniciar_elaboracion = pendiente_inicio.to(en_elaboracion)
    enviar_revision = en_elaboracion.to(en_revision)
    revision_aceptada = en_revision.to(en_aprobacion)
    rechazar_revision = en_revision.to(re_estructuracion)
    aprobar_documento = en_aprobacion.to(aprobado)
    rechazar_aprobacion = en_aprobacion.to(re_estructuracion)
    reenviar_revision = re_estructuracion.to(en_revision)
    publicar_documento = aprobado.to(publicado)

    # === CONSTRUCTOR ===
    def __init__(self, rol_id, estado_inicial=None):
        super().__init__()
        self.rol_id = rol_id
        if estado_inicial:
            for s in self.states:
                if s.name == estado_inicial:
                    self.current_state = s
                    break

    # === PERMISOS ===
    def puede_transicionar(self, evento: str) -> bool:
        """
        Devuelve True si el rol puede ejecutar la transición desde el estado actual.
        """
        # Permisos por rol (1=Redactor, 2=Revisor, 3=Aprobador)
        permisos = {
            "iniciar_elaboracion": [1],   # El redactor puede comenzar
            "enviar_revision": [1],
            "reenviar_revision": [1],
            "revision_aceptada": [2],
            "rechazar_revision": [2],
            "aprobar_documento": [3],
            "rechazar_aprobacion": [3],
            "publicar_documento": [3],
        }

        if evento not in permisos:
            return False

        # Restricciones específicas por rol y estado
        estado = self.current_state.name

        if self.rol_id == 1:  # Redactor
            if estado == "Pendiente de Inicio":
                return evento == "iniciar_elaboracion"
            if estado == "En Elaboración":
                return evento == "enviar_revision"
            if estado == "Re Estructuración":
                return evento == "reenviar_revision"

        if self.rol_id not in permisos[evento]:
            return False

        # Evitar publicar si no está aprobado
        if evento == "publicar_documento" and self.current_state != self.aprobado:
            return False

        return True

    # === GENERACIÓN DE VERSIONES ===
    def evento_genera_version(self, evento: str) -> bool:
        """
        Indica si la transición genera nueva versión del documento.
        """
        return evento in [
            "iniciar_elaboracion",   # Genera versión inicial (plantilla abierta)
            "enviar_revision",
            "revision_aceptada",
            "rechazar_revision",
            "aprobar_documento",
            "rechazar_aprobacion",
            "reenviar_revision",
        ]

    # === CONTROL DE SUBIDA DE ARCHIVOS ===
    def puede_subir_archivo(self, evento_actual=None):
        """
        Determina si se puede subir archivo según el estado, evento en curso y rol.
        """
        estado = self.current_state.name

        # --- 1️⃣ Redactor ---
        if self.rol_id == 1:
            # a) Al iniciar elaboración (crear plantilla inicial)
            if estado == "Pendiente de Inicio" and evento_actual == "iniciar_elaboracion":
                return True
            # b) Durante la elaboración del documento
            if estado == "En Elaboración":
                return True
            # c) En reestructuración (tras rechazo)
            if estado == "Re Estructuración":
                return True
            # d) Al reenviar documento a revisión
            if evento_actual in ["enviar_revision", "reenviar_revision"]:
                return True

        # --- 2️⃣ Revisor ---
        if self.rol_id == 2:
            # Puede adjuntar documento de rechazo (por ejemplo con comentarios)
            if estado == "En Revisión" and evento_actual == "rechazar_revision":
                return True

        # --- 3️⃣ Aprobador ---
        if self.rol_id == 3:
            # Puede adjuntar archivo con observaciones en rechazo
            if estado == "En Aprobación" and evento_actual == "rechazar_aprobacion":
                return True

        # --- Por defecto, no puede subir ---
        return False

