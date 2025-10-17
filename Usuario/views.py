from django.shortcuts import render, redirect
from django.http import HttpResponse, JsonResponse # Keep JsonResponse
from django.contrib.auth.decorators import login_required
from django.utils import timezone
from datetime import timedelta
from google.cloud import storage
from django.conf import settings
import openpyxl
from .models import FilePreview
from django.db import connection 
from collections import defaultdict
import json 
# You need to import the decorator:
from django.views.decorators.http import require_POST 


# Duración de previsualización en minutos
PREVIEW_EXPIRATION_MINUTES = 60  

# Tipos MIME permitidos (puedes ampliar)
ALLOWED_MIME_TYPES = [
    "application/vnd.ms-excel",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/msword",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/pdf",
    "image/jpeg",
    "image/png",
]

# Cliente de Google Cloud Storage
storage_client = storage.Client.from_service_account_json(settings.GCP_SERVICE_ACCOUNT_JSON)
bucket = storage_client.bucket(settings.GCP_BUCKET_NAME)


def get_or_create_preview_url(blob):
    """Genera o reutiliza enlace de previsualización temporal"""
    try:
        preview = FilePreview.objects.get(blob_name=blob.name)
        if preview.is_expired():
            raise FilePreview.DoesNotExist
        remaining_minutes = int((preview.expires_at - timezone.now()).total_seconds() / 60)
        return preview.signed_url, remaining_minutes
    except FilePreview.DoesNotExist:
        try:
            url = blob.generate_signed_url(
                version="v4",
                expiration=timedelta(minutes=PREVIEW_EXPIRATION_MINUTES),
                method="GET"
            )
        except Exception:
            url = ""
        expires_at = timezone.now() + timedelta(minutes=PREVIEW_EXPIRATION_MINUTES)
        FilePreview.objects.update_or_create(
            blob_name=blob.name,
            defaults={"signed_url": url, "expires_at": expires_at}
        )
        return url, PREVIEW_EXPIRATION_MINUTES

@login_required
def inicio(request):
    return render(request, "usuario_inicio.html")

@login_required
def lista_proyectos(request):
    proyectos = []

    with connection.cursor() as cursor:
        cursor.execute("SELECT id, nombre, descripcion, fecha_inicio, fecha_fin FROM proyectos")
        rows = cursor.fetchall()
        for row in rows:
            proyectos.append({
                "id": row[0],
                "nombre": row[1],
                "descripcion": row[2],
                "fecha_inicio": row[3],
                "fecha_fin": row[4],
            })

    return render(request, "usuario_proyectos.html", {"proyectos": proyectos})

@login_required
def crear_proyecto(request):
    numero_orden = ""
    usuarios = []
    usuarios_administrador = []
    grupos_maestros = []
    documentos = []
    form_error = None
    data_debug = None  # Para mostrar los datos en pantalla

    # 1. Obtener TODOS los USUARIOS para roles
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT u.id, u.nombre, u.email
            FROM usuarios u
            ORDER BY u.nombre
        """)
        usuarios = [{"id": r[0], "nombre": r[1], "email": r[2]} for r in cursor.fetchall()]

    # 1b. Obtener USUARIOS para Administrador de Servicio (solo área "Administrador de Contratos")
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT u.id, u.nombre, u.email
            FROM usuarios u
            JOIN area_cargos_empresa a ON u.cargo_id = a.id
            WHERE LOWER(a.nombre) = 'administrador de contratos'
            ORDER BY u.nombre
        """)
        usuarios_administrador = [{"id": r[0], "nombre": r[1], "email": r[2]} for r in cursor.fetchall()]

    # 2. Obtener GRUPOS DE TRABAJO desde las categorías de documentos técnicos
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT c.id, c.nombre, c.descripcion
            FROM categoria_documentos_tecnicos c
            ORDER BY c.nombre
        """)
        grupos_maestros = [{"id": r[0], "nombre": r[1], "descripcion": r[2]} for r in cursor.fetchall()]

    # 2b. Obtener DOCUMENTOS asociados a cada categoría
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id, categoria_id, nombre
            FROM tipo_documentos_tecnicos
            ORDER BY nombre
        """)
        documentos = [{"id": r[0], "categoria_id": r[1], "nombre": r[2]} for r in cursor.fetchall()]

    # 4. Procesar formulario
    if request.method == "POST":
        numero_orden = request.POST.get("numero_orden")
        nombre_proyecto = request.POST.get("nombre")
        descripcion_proyecto = request.POST.get("descripcion")
        fecha_inicio = request.POST.get("fecha_inicio")
        fecha_fin = request.POST.get("fecha_fin")
        administrador_id = request.POST.get("administrador")

        # Extraer los documentos seleccionados y sus roles
        documentos_roles = {}
        for doc in documentos:
            doc_id = str(doc["id"])
            redactores = request.POST.getlist(f"redactor_id_{doc_id}[]")
            revisores = request.POST.getlist(f"revisor_id_{doc_id}[]")
            aprobadores = request.POST.getlist(f"aprobador_id_{doc_id}[]")
            if redactores or revisores or aprobadores:
                documentos_roles[doc_id] = {
                    "redactores": redactores,
                    "revisores": revisores,
                    "aprobadores": aprobadores,
                }

        # VALIDACIÓN
        if not (numero_orden and nombre_proyecto and fecha_inicio):
            form_error = "Faltan campos obligatorios del proyecto."
        elif not administrador_id:
            form_error = "Debe seleccionar un Administrador de Servicio."
        elif not documentos_roles:
            form_error = "Debe seleccionar al menos un Documento con roles."
        else:
            try:
                # Preparamos la estructura con lo que se habría insertado
                data_debug = {
                    "Proyecto": {
                        "numero_orden": numero_orden,
                        "nombre": nombre_proyecto,
                        "descripcion": descripcion_proyecto,
                        "fecha_inicio": fecha_inicio,
                        "fecha_fin": fecha_fin,
                        "administrador_id": administrador_id,
                    },
                    "Documentos_y_roles": documentos_roles,
                }

                # Imprimir en consola (log de servidor) para ver fácilmente durante el desarrollo
                print("========== DATOS QUE SE GUARDARÍAN ==========")
                import pprint
                pprint.pprint(data_debug)
                print("=============================================")

                # --- INICIO BLOQUE DESACTIVADO: INSERTS EN BD (comentado para testeo) ---
                # with connection.cursor() as cursor:
                #     # Crear proyecto
                #     cursor.execute("""
                #         INSERT INTO proyectos (nombre, descripcion, fecha_inicio, fecha_fin)
                #         VALUES (%s, %s, %s, %s)
                #         RETURNING id
                #     """, [nombre_proyecto, descripcion_proyecto, fecha_inicio, fecha_fin])
                #     proyecto_id = cursor.fetchone()[0]
                #
                #     # Asociar documentos y roles
                #     for doc_id, roles in documentos_roles.items():
                #         cursor.execute("""
                #             INSERT INTO documentos_proyecto (proyecto_id, documento_id)
                #             VALUES (%s, %s)
                #             RETURNING id
                #         """, [proyecto_id, doc_id])
                #         doc_proy_id = cursor.fetchone()[0]
                #
                #         for r_id in roles["redactores"]:
                #             cursor.execute("""
                #                 INSERT INTO usuarios_documentos (usuario_id, documento_proyecto_id, rol)
                #                 VALUES (%s, %s, 'redactor')
                #             """, [r_id, doc_proy_id])
                #         for r_id in roles["revisores"]:
                #             cursor.execute("""
                #                 INSERT INTO usuarios_documentos (usuario_id, documento_proyecto_id, rol)
                #                 VALUES (%s, %s, 'revisor')
                #             """, [r_id, doc_proy_id])
                #         for r_id in roles["aprobadores"]:
                #             cursor.execute("""
                #                 INSERT INTO usuarios_documentos (usuario_id, documento_proyecto_id, rol)
                #                 VALUES (%s, %s, 'aprobador')
                #             """, [r_id, doc_proy_id])
                #
                #     # Registrar administrador del servicio
                #     cursor.execute("""
                #         INSERT INTO usuarios_grupos (usuario_id, grupo_id, rol_id)
                #         VALUES (%s, NULL, (SELECT id FROM roles_ciclodocumento WHERE LOWER(nombre) = 'administrador' LIMIT 1))
                #     """, [administrador_id])
                #
                # # No redirigimos: esto es modo testeo, se queda en la misma vista mostrando data_debug
                # --- FIN BLOQUE DESACTIVADO ---

            except Exception as e:
                form_error = f"Error al simular guardado del proyecto: {str(e)}"

    context = {
        "numero_orden": numero_orden,
        "usuarios": usuarios,
        "usuarios_administrador": usuarios_administrador,
        "grupos_maestros": grupos_maestros,
        "documentos": documentos,
        "form_error": form_error,
        "data_debug": data_debug,  # Muestra datos simulados en la vista
    }
    return render(request, "usuario_crearproyecto.html", context)

@require_POST
def validar_orden_ajax(request):
    if request.method == "POST" and request.FILES.get("archivo"):
        archivo = request.FILES.get("archivo")

        if not archivo.name.lower().endswith(".xlsx"):
            return JsonResponse({"error": "Formato de archivo no soportado. Se espera un .xlsx"}, status=400)

        try:
            wb = openpyxl.load_workbook(archivo, data_only=True)
            numero_orden = None

            for defined_name in wb.defined_names.values():
                if defined_name.name.lower() == "numordenservicio":
                    dest = list(defined_name.destinations)[0]
                    sheet_name, cell_coord = dest
                    sheet = wb[sheet_name]
                    valor_celda = sheet[cell_coord].value

                    if valor_celda is not None:
                        numero_orden = str(valor_celda).strip()
                        break

            if numero_orden:
                return JsonResponse({"numero_orden": numero_orden})
            else:
                return JsonResponse({"error": "No se encontró el nombre definido 'NumOrdenServicio' o la celda está vacía."}, status=400)

        except Exception as e:
            return JsonResponse({"error": f"Error al procesar el archivo: {str(e)}"}, status=500)

    return JsonResponse({"error": "Petición inválida o falta el archivo."}, status=400)

@login_required
def detalle_proyecto(request, proyecto_id):
    """
    Muestra los detalles completos de un proyecto incluyendo:
    - Proyecto
    - Contrato
    - Cliente
    - Máquinas
    - Requerimientos de documentos técnicos con su último estado
    """

    # Consulta SQL completa
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
        -- Proyecto
        P.id AS proyecto_id,
        P.nombre AS nombre_proyecto,
        P.descripcion AS proyecto_descripcion,
        P.fecha_inicio,
        P.fecha_fin,
        P.numero_orden,

        -- Administrador
        U.nombre AS administrador_nombre_completo,
        U.email AS administrador_email,

        -- Faena
        F.nombre AS nombre_faena,

        -- Contrato
        C.id AS contrato_id,
        C.numero_contrato,
        C.monto_total,
        C.fecha_creacion AS contrato_fecha_creacion,
        C.representante_cliente_nombre,
        C.representante_cliente_correo,
        C.representante_cliente_telefono,

        -- Cliente
        CL.id AS cliente_id,
        CL.nombre AS cliente_nombre,
        CL.rut AS cliente_rut,
        CL.direccion AS cliente_direccion,
        CL.correo_contacto AS cliente_correo,
        CL.telefono_contacto AS cliente_telefono,

        -- Requerimientos de Documento Técnico
        RDT.id AS requerimiento_id,
        TDT.nombre AS nombre_documento_tecnico,
        E.nombre AS estado_actual_documento,
        RDT.fecha_registro AS requerimiento_fecha,

        -- Máquinas
        M.id AS maquina_id,
        M.nombre AS maquina_nombre,
        M.codigo_equipo,
        M.marca,
        M.modelo,
        M.anio_fabricacion,
        M.tipo AS tipo_maquina,
        M.descripcion AS maquina_descripcion

    FROM public.proyectos P
    INNER JOIN public.usuarios U ON P.administrador_id = U.id
    LEFT JOIN public.faenas F ON P.faena_id = F.id
    INNER JOIN public.contratos C ON P.contrato_id = C.id
    INNER JOIN public.clientes CL ON C.cliente_id = CL.id
    LEFT JOIN public.requerimiento_documento_tecnico RDT ON P.id = RDT.proyecto_id
    LEFT JOIN public.tipo_documentos_tecnicos TDT ON RDT.tipo_documento_id = TDT.id
    LEFT JOIN UltimoEstado UE ON RDT.id = UE.requerimiento_id AND UE.rn = 1
    LEFT JOIN public.estado_documento E ON UE.estado_destino_id = E.id
    LEFT JOIN public.maquinas M ON P.id = M.proyecto_id
    WHERE P.id = %s
    ORDER BY RDT.fecha_registro DESC, M.id;
    """

    # Ejecutar la consulta
    with connection.cursor() as cursor:
        cursor.execute(sql, [proyecto_id])
        columns = [col[0] for col in cursor.description]
        resultados = [
            dict(zip(columns, row))
            for row in cursor.fetchall()
        ]

    # Separar información general del proyecto/contrato/cliente (tomamos la primera fila)
    proyecto_info = resultados[0] if resultados else None

    # Requerimientos de documentos y máquinas
    requerimientos = []
    maquinas = []
    req_ids = set()
    maquina_ids = set()
    for row in resultados:
        # Evitar duplicados en requerimientos
        if row['requerimiento_id'] and row['requerimiento_id'] not in req_ids:
            requerimientos.append({
                'id': row['requerimiento_id'],
                'nombre_documento_tecnico': row['nombre_documento_tecnico'],
                'estado_actual': row['estado_actual_documento'],
                'fecha_registro': row['requerimiento_fecha'],
            })
            req_ids.add(row['requerimiento_id'])

        # Evitar duplicados en maquinas
        if row['maquina_id'] and row['maquina_id'] not in maquina_ids:
            maquinas.append({
                'id': row['maquina_id'],
                'nombre': row['maquina_nombre'],
                'codigo_equipo': row['codigo_equipo'],
                'marca': row['marca'],
                'modelo': row['modelo'],
                'anio_fabricacion': row['anio_fabricacion'],
                'tipo': row['tipo_maquina'],
                'descripcion': row['maquina_descripcion'],
            })
            maquina_ids.add(row['maquina_id'])

    context = {
        'proyecto': proyecto_info,
        'requerimientos': requerimientos,
        'maquinas': maquinas,
    }

    return render(request, "usuario_proyecto_detalle.html", context)


@login_required
def detalle_documento(request, documento_id):
    """
    Muestra el historial de estados de un requerimiento técnico y los equipos asignados.
    """
    logs = []
    equipo_redactores = set()
    equipo_revisores_aprobadores = set()

    with connection.cursor() as cursor:
        # Obtener historial de estados
        cursor.execute("""
            SELECT
                rdt.id AS requerimiento_id,
                u.id AS usuario_id,
                u.nombre AS usuario_nombre,
                rol.nombre AS rol_usuario,
                eo.nombre AS estado_origen,
                ed.nombre AS estado_destino,
                led.created_at AS fecha_accion
            FROM log_estado_requerimiento_documento led
            LEFT JOIN requerimiento_documento_tecnico rdt ON led.requerimiento_id = rdt.id
            LEFT JOIN usuarios u ON led.usuario_id = u.id
            LEFT JOIN requerimiento_equipo_rol rer 
                ON rer.requerimiento_id = rdt.id AND rer.usuario_id = u.id
            LEFT JOIN roles_ciclodocumento rol ON rer.rol_id = rol.id
            LEFT JOIN estado_documento eo ON led.estado_origen_id = eo.id
            LEFT JOIN estado_documento ed ON led.estado_destino_id = ed.id
            WHERE rdt.id = %s
            ORDER BY led.id ASC
        """, [documento_id])
        columns = [col[0] for col in cursor.description]
        logs = [dict(zip(columns, row)) for row in cursor.fetchall()]

        # Obtener todos los usuarios asignados al requerimiento y sus roles desde tabla usuarios
        cursor.execute("""
            SELECT u.nombre, rol.nombre
            FROM requerimiento_equipo_rol rer
            INNER JOIN usuarios u ON rer.usuario_id = u.id
            INNER JOIN roles_ciclodocumento rol ON rer.rol_id = rol.id
            WHERE rer.requerimiento_id = %s AND rer.activo = true
        """, [documento_id])

        for usuario_nombre, rol_nombre in cursor.fetchall():
            if rol_nombre.lower() == "redactor":
                equipo_redactores.add(usuario_nombre)
            elif rol_nombre.lower() in ["revisor", "aprobador"]:
                equipo_revisores_aprobadores.add(usuario_nombre)

    context = {
        "documento_id": documento_id,
        "logs": logs,
        "equipo_redactores": sorted(equipo_redactores),
        "equipo_revisores_aprobadores": sorted(equipo_revisores_aprobadores),
    }

    return render(request, "usuario_proyecto_detalle_documento.html", context)








@login_required
def lista_usuarios(request):
    usuarios = []

    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT 
                u.id,
                u.nombre,
                u.email AS email_corporativo,
                COALESCE(u.email_secundario, '') AS email_secundario,
                COALESCE(u.telefono_corporativo, '') AS telefono_corporativo,
                COALESCE(u.telefono_secundario, '') AS telefono_secundario,
                COALESCE(a.codigo, '') AS area_trabajo,
                COALESCE(c.nombre, '') AS cargo,
                u.fecha_registro
            FROM public.usuarios u
            LEFT JOIN public.cargos_empresa c ON u.cargo_id = c.id
            LEFT JOIN public.area_cargo_empresa a ON c.area_id = a.id
            ORDER BY u.id;
        """)
        rows = cursor.fetchall()
        for row in rows:
            usuarios.append({
                "id": row[0],
                "nombre": row[1],
                "email_corporativo": row[2],
                "email_secundario": row[3],
                "telefono_corporativo": row[4],
                "telefono_secundario": row[5],
                "area_trabajo": row[6],
                "cargo": row[7],
                "fecha_registro": row[8],  # datetime o None
            })

    return render(request, "usuario_usuarios.html", {"usuarios": usuarios})





@login_required
def list_files(request):
    folder = request.GET.get("folder", "")
    folders, files = [], []

    try:
        iterator = bucket.list_blobs(prefix=folder, delimiter="/")
        page = next(iterator.pages)

        # Carpetas
        for prefix in page.prefixes:
            folder_name = prefix[len(folder):].strip("/")
            folders.append({"id": prefix, "name": folder_name})

        # Archivos
        for blob in page:
            if not blob.name.endswith("/"):
                preview_url, preview_expiration = get_or_create_preview_url(blob)
                files.append({
                    "id": blob.name,
                    "name": blob.name.split("/")[-1],
                    "preview_url": preview_url,
                    "preview_expiration": preview_expiration,
                    "size": blob.size,
                    "created_at": blob.time_created,
                })

    except Exception as e:
        print("Error GCS:", e)

    parent_folder = "/".join(folder.strip("/").split("/")[:-1])
    if parent_folder:
        parent_folder += "/"

    context = {
        "folders": folders,
        "files": files,
        "current_folder": folder,
        "parent_folder": parent_folder
    }
    return render(request, "usuario_repositorio.html", context)


@login_required
def upload_file(request):
    if request.method == "POST" and request.FILES.get("file"):
        uploaded_file = request.FILES["file"]
        folder = request.POST.get("current_folder", "")

        if uploaded_file.content_type not in ALLOWED_MIME_TYPES:
            return HttpResponse(
                "<script>alert('Tipo de archivo no permitido.'); window.history.back();</script>"
            )

        blob_name = f"{folder}{uploaded_file.name}" if folder else uploaded_file.name
        try:
            blob = bucket.blob(blob_name)
            blob.upload_from_file(uploaded_file, content_type=uploaded_file.content_type)
        except Exception as e:
            return HttpResponse(f"Error al subir archivo: {e}")

        return redirect(f"{request.path}?folder={folder}")
    return redirect("list_files")


@login_required
def download_file(request, file_id):
    try:
        blob = bucket.blob(file_id)
        if blob.exists():
            content = blob.download_as_bytes()
            response = HttpResponse(content, content_type=blob.content_type or "application/octet-stream")
            response['Content-Disposition'] = f'attachment; filename="{blob.name.split("/")[-1]}"'
            return response
    except Exception as e:
        return HttpResponse(f"Error al descargar archivo: {e}", status=500)
    return HttpResponse("Archivo no encontrado", status=404)


@login_required
def delete_file(request, file_id):
    try:
        blob = bucket.blob(file_id)
        if blob.exists():
            blob.delete()
    except Exception as e:
        return HttpResponse(f"Error al eliminar archivo: {e}", status=500)
    return redirect("list_files")


@login_required
def new_folder(request):
    if request.method == "POST":
        folder_name = request.POST.get("folder_name")
        current_folder = request.POST.get("current_folder", "")
        if folder_name:
            if not folder_name.endswith("/"):
                folder_name += "/"
            blob_name = f"{current_folder}{folder_name}" if current_folder else folder_name
            try:
                blob = bucket.blob(blob_name)
                blob.upload_from_string("")  # GCS no tiene carpetas reales
            except Exception as e:
                return HttpResponse(f"Error al crear carpeta: {e}", status=500)
    return redirect("list_files")



