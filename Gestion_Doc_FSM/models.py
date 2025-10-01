from django.db import models
from django_fsm import FSMField, transition

class Equipo(models.Model):
    # ... (El modelo Equipo se mantiene igual)
    nombre = models.CharField(max_length=100)
    
    def __str__(self):
        return self.nombre

class DocumentoTecnico(models.Model):
    # 📝 Estados alineados con Confección, Revisión y Aprobación
    ESTADOS = [
        ('borrador', 'Borrador (Confección)'),
        ('pendiente_revision', 'Pendiente de Revisión'),
        ('en_revision', 'En Revisión'),
        ('requiere_cambios', 'Requiere Cambios (Rechazo)'),
        ('listo_para_aprobacion', 'Listo para Aprobación'),
        ('aprobado', 'Aprobado (Final)'),
        ('obsoleto', 'Obsoleto/Archivado'), # Estado final
    ]

    titulo = models.CharField(max_length=200)
    descripcion = models.TextField()
    estado = FSMField(default='borrador', choices=ESTADOS, protected=True)

    equipos_redactores = models.ManyToManyField(Equipo, related_name='docs_redactados')
    equipos_revisores = models.ManyToManyField(Equipo, related_name='docs_revisados')
    equipos_aprobadores = models.ManyToManyField(Equipo, related_name='docs_aprobados')

    # --- Transiciones de Confección ---
    
    @transition(field=estado, source='borrador', target='pendiente_revision', 
                permission='myapp.can_submit_for_review') # Asumiendo un permiso
    def enviar_a_revision(self):
        """El Confeccionador termina y envía el documento a Revisión."""
        print(f"Documento '{self.titulo}' enviado a Revisión.")

    # --- Transiciones de Revisión ---
    
    @transition(field=estado, source='pendiente_revision', target='en_revision', 
                permission='myapp.can_start_review')
    def iniciar_revision(self):
        """El Revisor toma el documento para revisarlo."""
        print(f"Documento '{self.titulo}' pasa a En Revisión.")
        
    @transition(field=estado, source='en_revision', target='listo_para_aprobacion',
                permission='myapp.can_finish_review')
    def revision_conforme(self):
        """El Revisor aprueba el contenido técnico y lo pasa a Aprobación."""
        print(f"Revisión de '{self.titulo}' finalizada. Listo para Aprobación.")

    @transition(field=estado, source='en_revision', target='requiere_cambios',
                permission='myapp.can_reject_review')
    def rechazar_revision(self):
        """El Revisor devuelve el documento al Confeccionador."""
        print(f"Documento '{self.titulo}' rechazado, requiere cambios.")

    # --- Transición de Corrección a Confección ---
    
    @transition(field=estado, source='requiere_cambios', target='borrador',
                permission='myapp.can_reedit')
    def volver_a_borrador(self):
        """El Confeccionador recibe el rechazo y vuelve a editar."""
        print(f"Documento '{self.titulo}' vuelve a Borrador para correcciones.")

    # --- Transiciones de Aprobación ---
    
    @transition(field=estado, source='listo_para_aprobacion', target='aprobado',
                permission='myapp.can_approve')
    def aprobar_documento(self):
        """El Aprobador da el visto bueno final."""
        print(f"Documento '{self.titulo}' ha sido APROBADO.")

    @transition(field=estado, source='listo_para_aprobacion', target='requiere_cambios',
                permission='myapp.can_reject_approval')
    def rechazar_aprobacion(self):
        """El Aprobador lo devuelve al Confeccionador (o al revisor, dependiendo del flujo exacto)."""
        print(f"Documento '{self.titulo}' rechazado por Aprobador, requiere cambios.")
        
    # --- Transición Final (Opcional) ---
    
    @transition(field=estado, source='aprobado', target='obsoleto',
                permission='myapp.can_archive')
    def archivar(self):
        """Mueve el documento de 'aprobado' a un estado histórico."""
        print(f"Documento '{self.titulo}' pasa a Obsoleto/Archivado.")