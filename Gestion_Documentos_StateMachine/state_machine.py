# Gestion_Documentos_StateMachine/state_machine.py
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

        # Forzar estado inicial manualmente (según BD)
        if estado_inicial:
            for s in self.states:
                if s.name == estado_inicial:
                    self.current_state = s
                    break

    # === PERMISOS ===
    def puede_transicionar(self, evento: str) -> bool:
        """
        Determina si el usuario (según rol) puede ejecutar
        la transición desde el estado actual.
        """

        estado = self.current_state.name

        # Regla especial: en APROBADO solo PUBLICAR
        if estado == "Aprobado. Listo para Publicación":
            return evento == "publicar_documento" and self.rol_id == 3

        # Permisos por rol
        permisos = {
            "iniciar_elaboracion": [1],
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

        if self.rol_id not in permisos[evento]:
            return False

        # Lógica por estado
        if self.rol_id == 1:
            if estado == "Pendiente de Inicio":
                return evento == "iniciar_elaboracion"
            if estado == "En Elaboración":
                return evento == "enviar_revision"
            if estado == "Re Estructuración":
                return evento == "reenviar_revision"

        if self.rol_id == 2:
            if estado == "En Revisión":
                return evento in ["revision_aceptada", "rechazar_revision"]

        if self.rol_id == 3:
            if estado == "En Aprobación":
                return evento in ["aprobar_documento", "rechazar_aprobacion"]

        # Evitar publicar fuera de APROBADO
        if evento == "publicar_documento" and estado != "Aprobado. Listo para Publicación":
            return False

        return True

    # === CONTROL: genera versión? ===
    def evento_genera_version(self, evento: str) -> bool:
        return evento in [
            "iniciar_elaboracion",
            "enviar_revision",
            "revision_aceptada",
            "rechazar_revision",
            "aprobar_documento",
            "rechazar_aprobacion",
            "reenviar_revision",
        ]

    # === CONTROL: permitir subir archivos ===
    def puede_subir_archivo(self, evento_actual=None):
        estado = self.current_state.name

        # Redactor
        if self.rol_id == 1:
            if estado == "Pendiente de Inicio" and evento_actual == "iniciar_elaboracion":
                return True
            if estado in ["En Elaboración", "Re Estructuración"]:
                return True
            if evento_actual in ["enviar_revision", "reenviar_revision"]:
                return True

        # Revisor
        if self.rol_id == 2:
            if estado == "En Revisión" and evento_actual == "rechazar_revision":
                return True

        # Aprobador
        if self.rol_id == 3:
            if estado == "En Aprobación" and evento_actual == "rechazar_aprobacion":
                return True

        return False
