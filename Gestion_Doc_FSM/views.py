from django.shortcuts import get_object_or_404
from django.http import HttpResponse
from django_fsm import can_proceed

def avanzar_estado(request, doc_id, accion):
    doc = get_object_or_404(DocumentoTecnico, id=doc_id)
    
    # Mapeo de la acción enviada por el usuario a la función de transición
    # Esto simula un formulario o botón que el usuario presiona.
    transiciones_posibles = {
        'enviar_a_revision': doc.enviar_a_revision,
        'iniciar_revision': doc.iniciar_revision,
        'revision_conforme': doc.revision_conforme,
        'rechazar_revision': doc.rechazar_revision,
        'aprobar_documento': doc.aprobar_documento,
        'rechazar_aprobacion': doc.rechazar_aprobacion,
        'volver_a_borrador': doc.volver_a_borrador,
        'archivar': doc.archivar,
    }
    
    if accion not in transiciones_posibles:
        return HttpResponse("Acción no válida.", status=400)
    
    transicion_func = transiciones_posibles[accion]
    
    # 1. Verificar si la transición es posible desde el estado actual
    if not can_proceed(transicion_func):
        return HttpResponse(
            f"La acción '{accion}' no es válida desde el estado actual: {doc.get_estado_display()}", 
            status=409 # Conflicto
        )
        
    try:
        # 2. Ejecutar la transición (llama a la función con el decorador @transition)
        transicion_func()
        doc.save()
        return HttpResponse(
            f"Documento '{doc.titulo}' pasó a estado: {doc.get_estado_display()}",
            status=200
        )
    except Exception as e:
        # Manejo de errores durante la ejecución de la transición (ej. permisos)
        return HttpResponse(f"Error al ejecutar la transición: {e}", status=500)