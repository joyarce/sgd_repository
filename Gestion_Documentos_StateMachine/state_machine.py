
from statemachine import StateMachine, State

class DocumentoTecnicoStateMachine(StateMachine):
    # Estados
    borrador = State("Borrador", initial=True)
    en_elaboracion = State("En Elaboración")
    en_revision = State("En Revisión")
    en_aprobacion = State("En Aprobación")
    re_estructuracion = State("Re Estructuración")
    aprobado = State("Aprobado. Listo para Publicación")
    publicado = State("Publicado", final=True)

    # Transiciones
    crear_documento = borrador.to(en_elaboracion)
    enviar_revision = en_elaboracion.to(en_revision)
    revision_aceptada = en_revision.to(en_aprobacion)
    rechazar_revision = en_revision.to(re_estructuracion)
    aprobar_documento = en_aprobacion.to(aprobado)
    rechazar_aprobacion = en_aprobacion.to(re_estructuracion)
    reenviar_revision = re_estructuracion.to(en_revision)
    publicar_documento = aprobado.to(publicado)

    def __init__(self, rol_id, estado_inicial=None):
        super().__init__()
        self.rol_id = rol_id
        if estado_inicial:
            # Inicializar en un estado específico
            for s in self.states:
                if s.name == estado_inicial:
                    self.current_state = s
                    break

    def puede_transicionar(self, evento: str) -> bool:
        """
        Devuelve True si el rol puede ejecutar la transición
        desde el estado actual.
        """
        # Permisos por rol (1=Redactor, 2=Revisor, 3=Aprobador)
        permisos = {
            "crear_documento": [1],
            "enviar_revision": [1],
            "reenviar_revision": [1],
            "revision_aceptada": [2],
            "rechazar_revision": [2],
            "aprobar_documento": [3],
            "rechazar_aprobacion": [3],
            "publicar_documento": [3],
        }

        # Verifica que el evento exista
        if evento not in permisos:
            return False

        # Restricciones especiales por estado para Redactor
        if self.rol_id == 1:
            estado = self.current_state.name
            if estado == "Borrador":
                return evento == "crear_documento"
            if estado == "En Elaboración":
                return evento == "enviar_revision"
            if estado == "Re Estructuración":
                return evento == "reenviar_revision"

        # Restricciones generales por rol
        if self.rol_id not in permisos[evento]:
            return False

        # Evitar publicar si no está aprobado
        if evento == "publicar_documento" and self.current_state != self.aprobado:
            return False

        return True

    def evento_genera_version(self, evento: str) -> bool:
        """
        Indica si la transición genera nueva versión del documento.
        Ahora incluye enviar_revision para crear versión inicial con URL.
        """
        return evento in [
            "enviar_revision",
            "revision_aceptada",
            "rechazar_revision",
            "aprobar_documento",
            "rechazar_aprobacion",
            "reenviar_revision"
        ]
    
