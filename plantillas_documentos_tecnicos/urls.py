from django.urls import path
from . import views

app_name = "plantillas"

urlpatterns = [
    path("", views.lista_plantillas, name="lista_plantillas"),
    path("categorias/<int:categoria_id>/", views.categoria_detalle, name="detalle_categoria"),
    path("crear-categoria/", views.crear_categoria, name="crear_categoria"),
    path("crear-tipo/", views.crear_tipo_documento, name="crear_tipo_documento"),
    path("tipos/<int:tipo_id>/", views.tipo_detalle, name="detalle_tipo"),
    path('validar-etiquetas/', views.validar_etiquetas_archivo, name='validar_etiquetas_archivo'),
        # Obtener automáticamente todas las columnas / etiquetas del archivo
    path('obtener-columnas/', views.obtener_columnas_archivo, name='obtener_columnas_archivo'),
    path('obtener-tablas-y-columnas/', views.obtener_tablas_y_columnas, name='obtener_tablas_y_columnas'),


    path('test-tablas/', views.test_tablas, name='test_tablas'),

]
