from django.shortcuts import render, redirect
from django.http import HttpResponse
from django.contrib.auth.decorators import login_required
from django.utils import timezone
from datetime import timedelta
from google.cloud import storage
from django.conf import settings
from .models import FilePreview
from django.db import connection 
import json
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

        # Traer grupos relacionados con usuarios y roles
        cursor.execute("""
            SELECT g.id, g.nombre, g.descripcion,
                   u.id, u.nombre, u.email,
                   r.nombre as rol
            FROM grupos_trabajo g
            LEFT JOIN usuarios_grupos ug ON g.id = ug.grupo_id
            LEFT JOIN usuarios_microsoft u ON ug.usuario_id = u.id
            LEFT JOIN roles r ON ug.rol_id = r.id
            WHERE g.proyecto_id = %s
            ORDER BY g.id, u.id
        """, [proyecto_id])
        rows = cursor.fetchall()

        grupos_dict = {}
        for row in rows:
            grupo_id, grupo_nombre, grupo_desc, usuario_id, usuario_nombre, usuario_email, rol_nombre = row

            if grupo_id not in grupos_dict:
                grupos_dict[grupo_id] = {
                    "id": grupo_id,
                    "nombre": grupo_nombre,
                    "descripcion": grupo_desc,
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

    # Estadísticas adicionales
    num_grupos = len(grupos)
    promedio_miembros = total_integrantes / num_grupos if num_grupos else 0
    grupo_max = max(grupos, key=lambda g: len(g["usuarios"]))["nombre"] if grupos else ''
    grupo_min = min(grupos, key=lambda g: len(g["usuarios"]))["nombre"] if grupos else ''

    # Distribución de roles
    roles = {}
    for grupo in grupos:
        for usuario in grupo["usuarios"]:
            rol = usuario["rol"] or "Sin rol"
            roles[rol] = roles.get(rol, 0) + 1

    # Datos para gráficos
    grupos_nombres = [g["nombre"] for g in grupos]
    grupos_num_usuarios = [len(g["usuarios"]) for g in grupos]
    roles_labels = list(roles.keys())
    roles_values = list(roles.values())

    context = {
        "proyecto": proyecto,
        "grupos": grupos,
        "num_grupos": num_grupos,
        "num_usuarios": total_integrantes,
        "promedio_miembros": promedio_miembros,
        "grupo_max": grupo_max,
        "grupo_min": grupo_min,
        "roles": roles,
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
        cursor.execute("SELECT id, nombre, email FROM usuarios_microsoft")
        rows = cursor.fetchall()
        for row in rows:
            usuarios.append({
                "id": row[0],
                "nombre": row[1],
                "email": row[2],
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



