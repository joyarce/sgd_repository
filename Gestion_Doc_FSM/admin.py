from fsm_admin.mixins import FSMTransitionMixin
from django.contrib import admin
from .models import DocumentoTecnico, Equipo

@admin.register(DocumentoTecnico)
class DocumentoTecnicoAdmin(FSMTransitionMixin, admin.ModelAdmin):
    list_display = ['titulo', 'estado']
    list_filter = ['estado']

@admin.register(Equipo)
class EquipoAdmin(admin.ModelAdmin):
    list_display = ['nombre']
