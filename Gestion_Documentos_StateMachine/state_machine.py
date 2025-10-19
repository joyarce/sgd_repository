from statemachine import StateMachine, State

class DocumentoTecnicoStateMachine(StateMachine):
    """Máquina de estados ajustada según PlantUML."""

    # Estados
    borrador = State("Borrador", initial=True)
    en_elaboracion = State("En Elaboración")
    en_revision = State("En Revisión")
    en_aprobacion = State("En Aprobación")
    re_estructuracion = State("Re Estructuración")
    aprobado = State("Aprobado. Listo para Publicación")
    publicado = State("Publicado", final=True)

    # Transiciones (nombres de eventos coinciden con PlantUML)
    crear_documento = borrador.to(en_elaboracion)
    enviar_revision = en_elaboracion.to(en_revision)
    revision_aceptada = en_revision.to(en_aprobacion)
    rechazar_revision = en_revision.to(re_estructuracion)
    aprobar_documento = en_aprobacion.to(aprobado)
    rechazar_aprobacion = en_aprobacion.to(re_estructuracion)
    reenviar_revision = re_estructuracion.to(en_revision)
    publicar_documento = aprobado.to(publicado)

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
            # Redactor
            "crear_documento": [1],
            "enviar_revision": [1],
            "reenviar_revision": [1],
            # Revisor
            "revision_aceptada": [2],
            "rechazar_revision": [2],
            # Aprobador
            "aprobar_documento": [3],
            "rechazar_aprobacion": [3],
            "publicar_documento": [3],
        }

        # Evitar publicar si no está aprobado
        if evento == "publicar_documento" and self.current_state != self.aprobado:
            return False

        return self.rol_id in permisos.get(evento, [])