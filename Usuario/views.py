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

    # 1. Obtener TODOS los USUARIOS para roles
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT u.id, u.nombre, u.email
            FROM usuarios_microsoft u
            ORDER BY u.nombre
        """)
        usuarios = [{"id": r[0], "nombre": r[1], "email": r[2]} for r in cursor.fetchall()]

    # 1b. Obtener USUARIOS para Administrador de Servicio (solo área "Administrador de Contratos")
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT u.id, u.nombre, u.email
            FROM usuarios_microsoft u
            JOIN area_cargo_empresa a ON u.cargo_id = a.id
            WHERE LOWER(a.nombre) = 'administrador de contratos'
            ORDER BY u.nombre
        """)
        usuarios_administrador = [{"id": r[0], "nombre": r[1], "email": r[2]} for r in cursor.fetchall()]

    # 2. Obtener GRUPOS DE TRABAJO desde las categorías de documentos técnicos
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT c.id, c.nombre, c.descripcion
            FROM categoria_documentos c
            ORDER BY c.nombre
        """)
        grupos_maestros = [{"id": r[0], "nombre": r[1], "descripcion": r[2]} for r in cursor.fetchall()]

    # 2b. Obtener DOCUMENTOS asociados a cada categoría
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id, categoria_id, nombre
            FROM tipo_documentos
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
                with connection.cursor() as cursor:
                    # Crear proyecto
                    cursor.execute("""
                        INSERT INTO proyectos (nombre, descripcion, fecha_inicio, fecha_fin)
                        VALUES (%s, %s, %s, %s)
                        RETURNING id
                    """, [nombre_proyecto, descripcion_proyecto, fecha_inicio, fecha_fin])
                    proyecto_id = cursor.fetchone()[0]

                    # Asociar documentos y roles
                    for doc_id, roles in documentos_roles.items():
                        cursor.execute("""
                            INSERT INTO documentos_proyecto (proyecto_id, documento_id)
                            VALUES (%s, %s)
                            RETURNING id
                        """, [proyecto_id, doc_id])
                        doc_proy_id = cursor.fetchone()[0]

                        for r_id in roles["redactores"]:
                            cursor.execute("""
                                INSERT INTO usuarios_documentos (usuario_id, documento_proyecto_id, rol)
                                VALUES (%s, %s, 'redactor')
                            """, [r_id, doc_proy_id])
                        for r_id in roles["revisores"]:
                            cursor.execute("""
                                INSERT INTO usuarios_documentos (usuario_id, documento_proyecto_id, rol)
                                VALUES (%s, %s, 'revisor')
                            """, [r_id, doc_proy_id])
                        for r_id in roles["aprobadores"]:
                            cursor.execute("""
                                INSERT INTO usuarios_documentos (usuario_id, documento_proyecto_id, rol)
                                VALUES (%s, %s, 'aprobador')
                            """, [r_id, doc_proy_id])

                    # Registrar administrador del servicio
                    cursor.execute("""
                        INSERT INTO usuarios_grupos (usuario_id, grupo_id, rol_id)
                        VALUES (%s, NULL, (SELECT id FROM roles_ciclodocumento WHERE LOWER(nombre) = 'administrador' LIMIT 1))
                    """, [administrador_id])

                return redirect("usuario:detalle_proyecto", proyecto_id=proyecto_id)

            except Exception as e:
                form_error = f"Error al guardar el proyecto: {str(e)}"

    context = {
        "numero_orden": numero_orden,
        "usuarios": usuarios,  # Para roles en documentos
        "usuarios_administrador": usuarios_administrador,  # Solo para el select de Administrador
        "grupos_maestros": grupos_maestros,
        "documentos": documentos,
        "form_error": form_error,
    }
    return render(request, "usuario_crearproyecto.html", context)





# ----------------------------------------------------------------------
# FUNCIÓN PARA VALIDACIÓN AJAX
# ----------------------------------------------------------------------
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
    proyecto = {}
    grupos = []
    total_integrantes = 0

    with connection.cursor() as cursor:
        # Traer datos del proyecto
        cursor.execute(
            "SELECT id, nombre, descripcion, fecha_inicio, fecha_fin FROM proyectos WHERE id = %s",
            [proyecto_id]
        )
        row = cursor.fetchone()
        if row:
            proyecto = {
                "id": row[0],
                "nombre": row[1],
                "descripcion": row[2],
                "fecha_inicio": row[3],
                "fecha_fin": row[4],
            }

        # Traer grupos con categoría, usuarios y roles
        cursor.execute("""
            SELECT g.id AS grupo_id,
                   c.nombre AS categoria_nombre,
                   u.id AS usuario_id, u.nombre AS usuario_nombre, u.email AS usuario_email,
                   COALESCE(r.nombre, 'Sin rol') AS rol_nombre
            FROM grupos_proyecto g
            LEFT JOIN categoria_documentos c ON g.categoria_id = c.id
            LEFT JOIN usuarios_grupos ug ON g.id = ug.grupo_id
            LEFT JOIN usuarios_microsoft u ON ug.usuario_id = u.id
            LEFT JOIN roles_ciclodocumento r ON ug.rol_id = r.id
            WHERE g.proyecto_id = %s
            ORDER BY g.id, u.id
        """, [proyecto_id])

        rows = cursor.fetchall()

        grupos_dict = {}
        for row in rows:
            grupo_id, categoria_nombre, usuario_id, usuario_nombre, usuario_email, rol_nombre = row

            if grupo_id not in grupos_dict:
                grupos_dict[grupo_id] = {
                    "id": grupo_id,
                    "nombre": categoria_nombre or "Sin categoría",
                    "usuarios": []
                }

            if usuario_id:
                grupos_dict[grupo_id]["usuarios"].append({
                    "id": usuario_id,
                    "nombre": usuario_nombre,
                    "email": usuario_email,
                    "rol": rol_nombre
                })
                total_integrantes += 1

        grupos = list(grupos_dict.values())

    # Estadísticas
    num_grupos = len(grupos)
    promedio_miembros = total_integrantes / num_grupos if num_grupos else 0
    grupo_max = max(grupos, key=lambda g: len(g["usuarios"]))["nombre"] if grupos else ''
    grupo_min = min(grupos, key=lambda g: len(g["usuarios"]))["nombre"] if grupos else ''

    # Distribución de roles
    roles_ciclodocumento = {}
    for grupo in grupos:
        for usuario in grupo["usuarios"]:
            rol = usuario["rol"]
            roles_ciclodocumento[rol] = roles_ciclodocumento.get(rol, 0) + 1

    # Datos para gráficos
    grupos_nombres = [g['nombre'] for g in grupos]
    grupos_num_usuarios = [len(g["usuarios"]) for g in grupos]
    roles_labels = list(roles_ciclodocumento.keys())
    roles_values = list(roles_ciclodocumento.values())

    context = {
        "proyecto": proyecto,
        "grupos": grupos,
        "num_grupos": num_grupos,
        "num_usuarios": total_integrantes,
        "promedio_miembros": promedio_miembros,
        "grupo_max": grupo_max,
        "grupo_min": grupo_min,
        "grupos_nombres_json": json.dumps(grupos_nombres),
        "grupos_num_usuarios_json": json.dumps(grupos_num_usuarios),
        "roles_labels_json": json.dumps(roles_labels),
        "roles_values_json": json.dumps(roles_values),
    }

    return render(request, "usuario_proyecto_detalle.html", context)



@login_required
def lista_usuarios(request):
    usuarios = []

    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT u.id,
                   u.nombre,
                   u.email AS email_corporativo,
                   COALESCE(u.email_secundario, '') AS email_secundario,
                   COALESCE(u.telefono_corporativo, '') AS telefono_corporativo,
                   COALESCE(u.telefono_secundario, '') AS telefono_secundario,
                   COALESCE(c.nombre, '') AS cargo
            FROM usuarios_microsoft u
            LEFT JOIN cargo_empresa c ON u.cargo_id = c.id
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
                "cargo": row[6],
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



