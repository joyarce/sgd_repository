# flows.py
from viewflow import flow
from viewflow.fsm import FSMMixin
from .models import DocumentoTecnico

class DocumentoFlow(flow.Flow, FSMMixin):
    process_class = DocumentoTecnico

    start = flow.Start(this.comenzar_elaboracion).Next(this.en_elaboracion)

    en_elaboracion = flow.View(this.enviar_revision).Next(this.en_revision)

    en_revision = flow.View(this.aprobar_revision).Next(this.revisado)\
        .Next(this.reestructuracion, condition=lambda p: p.estado == "rechazado")

    reestructuracion = flow.View(this.reenviar_revision).Next(this.en_revision)

    revisado = flow.View(this.enviar_aprobacion).Next(this.en_aprobacion)

    en_aprobacion = flow.View(this.aprobar_documento).Next(this.aprobado)\
        .Next(this.reestructuracion, condition=lambda p: p.estado == "rechazado")

    aprobado = flow.View(this.publicar_documento).Next(this.publicado)

    publicado = flow.End()
