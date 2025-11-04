from django.shortcuts import render, redirect
from django.http import HttpResponse, JsonResponse # Keep JsonResponse
from django.contrib.auth.decorators import login_required
from django.utils import timezone
from collections import Counter, defaultdict
from datetime import timedelta
from google.cloud import storage
from django.conf import settings
import openpyxl
from datetime import datetime
from django.shortcuts import render
from django.db import connection
from .models import FilePreview
from django.views.decorators.csrf import csrf_exempt
from django.db import connection 
from collections import defaultdict
import json 
# You need to import the decorator:
from django.views.decorators.http import require_POST 
from django.shortcuts import get_object_or_404, render
from django.contrib import messages
import re
import unidecode
import openpyxl
import csv
import json

from openpyxl import load_workbook



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
        cursor.execute("""
            SELECT id, nombre, descripcion, 
                   fecha_recepcion_evaluacion, 
                   fecha_inicio_planificacion,
                   fecha_inicio_ejecucion,
                   fecha_cierre_proyecto
            FROM proyectos
            ORDER BY nombre
        """)
        rows = cursor.fetchall()
        for row in rows:
            proyectos.append({
                "id": row[0],
                "nombre": row[1],
                "descripcion": row[2],
                "fecha_recepcion_evaluacion": row[3],
                "fecha_inicio_planificacion": row[4],
                "fecha_inicio_ejecucion": row[5],
                "fecha_cierre_proyecto": row[6],
            })

    return render(request, "usuario_proyectos.html", {"proyectos": proyectos})



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
    - Cliente (con abreviatura)
    - Máquinas (usa abreviatura)
    - Requerimientos (incluye abreviatura del documento técnico)
    """
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
        P.numero_orden,
        P.fecha_recepcion_evaluacion,
        P.fecha_inicio_planificacion,
        P.fecha_inicio_ejecucion,
        P.fecha_cierre_proyecto,
        P.abreviatura AS proyecto_abreviatura,

        -- Administrador
        U.nombre AS administrador_nombre_completo,
        U.email AS administrador_email,

        -- Faena
        F.nombre AS nombre_faena,

        -- Contrato
        C.id AS contrato_id,
        C.numero_contrato,
        C.monto_total,
        C.fecha_firma AS contrato_fecha_firma,
        C.representante_cliente_nombre,
        C.representante_cliente_correo,
        C.representante_cliente_telefono,

        -- Cliente
        CL.id AS cliente_id,
        CL.nombre AS cliente_nombre,
        CL.abreviatura AS cliente_abreviatura,
        CL.rut AS cliente_rut,
        CL.direccion AS cliente_direccion,
        CL.correo_contacto AS cliente_correo,
        CL.telefono_contacto AS cliente_telefono,

        -- Requerimientos
        RDT.id AS requerimiento_id,
        TDT.nombre AS nombre_documento_tecnico,
        TDT.abreviatura AS abreviatura_documento_tecnico,
        E.nombre AS estado_actual_documento,
        RDT.fecha_registro AS requerimiento_fecha,

        -- Máquinas
        M.id AS maquina_id,
        M.nombre AS maquina_nombre,
        M.abreviatura,
        M.marca,
        M.modelo,
        M.anio_fabricacion,
        M.tipo AS tipo_maquina,
        M.descripcion AS maquina_descripcion

    FROM public.proyectos P
    INNER JOIN public.usuarios U ON P.administrador_id = U.id
    LEFT JOIN public.faenas F ON P.faena_id = F.id
    LEFT JOIN public.contratos C ON P.contrato_id = C.id
    LEFT JOIN public.clientes CL ON C.cliente_id = CL.id
    LEFT JOIN public.requerimiento_documento_tecnico RDT ON P.id = RDT.proyecto_id
    LEFT JOIN public.tipo_documentos_tecnicos TDT ON RDT.tipo_documento_id = TDT.id
    LEFT JOIN UltimoEstado UE ON RDT.id = UE.requerimiento_id AND UE.rn = 1
    LEFT JOIN public.estado_documento E ON UE.estado_destino_id = E.id
    LEFT JOIN public.proyecto_maquina PM ON P.id = PM.proyecto_id
    LEFT JOIN public.maquinas M ON PM.maquina_id = M.id
    WHERE P.id = %s
    ORDER BY RDT.fecha_registro DESC, M.id;
    """

    with connection.cursor() as cursor:
        cursor.execute(sql, [proyecto_id])
        columns = [col[0] for col in cursor.description]
        resultados = [dict(zip(columns, row)) for row in cursor.fetchall()]

    proyecto_info = resultados[0] if resultados else {}
    requerimientos, maquinas = [], []
    req_ids, maquina_ids = set(), set()

    for row in resultados:
        # Requerimientos únicos
        if row['requerimiento_id'] and row['requerimiento_id'] not in req_ids:
            requerimientos.append({
                'id': row['requerimiento_id'],
                'nombre_documento_tecnico': row['nombre_documento_tecnico'],
                'abreviatura': row['abreviatura_documento_tecnico'],
                'estado_actual': row['estado_actual_documento'],
                'fecha_registro': row['requerimiento_fecha'],
            })
            req_ids.add(row['requerimiento_id'])

        # Máquinas únicas
        if row['maquina_id'] and row['maquina_id'] not in maquina_ids:
            maquinas.append({
                'id': row['maquina_id'],
                'nombre': row['maquina_nombre'],
                'abreviatura': row['abreviatura'],
                'marca': row['marca'],
                'modelo': row['modelo'],
                'anio_fabricacion': row['anio_fabricacion'],
                'tipo': row['tipo_maquina'],
                'descripcion': row['maquina_descripcion'],
            })
            maquina_ids.add(row['maquina_id'])

    return render(request, "usuario_proyecto_detalle.html", {
        'proyecto': proyecto_info,
        'requerimientos': requerimientos,
        'maquinas': maquinas,
    })




@login_required
def detalle_documento(request, documento_id):
    logs = []
    equipo_redactores = []
    equipo_revisores = []
    equipo_aprobadores = []
    documento_info = None

    with connection.cursor() as cursor:
        # Información del documento
        cursor.execute("""
            SELECT RDT.id, TDT.nombre, CDT.nombre, RDT.fecha_registro
            FROM requerimiento_documento_tecnico RDT
            LEFT JOIN tipo_documentos_tecnicos TDT ON RDT.tipo_documento_id = TDT.id
            LEFT JOIN categoria_documentos_tecnicos CDT ON TDT.categoria_id = CDT.id
            WHERE RDT.id = %s
        """, [documento_id])
        row = cursor.fetchone()
        if row:
            documento_info = {
                'id': row[0],
                'tipo_documento': row[1],
                'categoria': row[2],
                'created_at': row[3] or timezone.now()
            }

        # Logs del documento
        cursor.execute("""
            SELECT rdt.id, u.nombre AS usuario_nombre, rol.nombre AS rol_usuario,
                   eo.nombre AS estado_origen, ed.nombre AS estado_destino,
                   led.created_at AS fecha_accion, led.observaciones
            FROM log_estado_requerimiento_documento led
            LEFT JOIN requerimiento_documento_tecnico rdt ON led.requerimiento_id = rdt.id
            LEFT JOIN usuarios u ON led.usuario_id = u.id
            LEFT JOIN requerimiento_equipo_rol rer ON rer.requerimiento_id = rdt.id AND rer.usuario_id = u.id
            LEFT JOIN roles_ciclodocumento rol ON rer.rol_id = rol.id
            LEFT JOIN estado_documento eo ON led.estado_origen_id = eo.id
            LEFT JOIN estado_documento ed ON led.estado_destino_id = ed.id
            WHERE rdt.id = %s
            ORDER BY led.created_at ASC
        """, [documento_id])
        columns = [col[0] for col in cursor.description]
        logs = [dict(zip(columns, row)) for row in cursor.fetchall()]

        # Agregar log inicial del sistema si es necesario
        if not logs or logs[0]['estado_origen'] is None:
            logs.insert(0, {
                'fecha_accion': documento_info.get('created_at') or timezone.now(),
                'estado_origen': 'Sistema',
                'estado_destino': 'Borrador',
                'usuario_nombre': 'Sistema',
                'rol_usuario': '',
                'observaciones': 'El sistema ha habilitado la plantilla en el repositorio.'
            })

        # Reconstruir estado_origen y usuario si es NULL
        estado_anterior = "Borrador"
        for log in logs:
            if not log['estado_origen']:
                log['estado_origen'] = estado_anterior
            if not log['usuario_nombre']:
                log['usuario_nombre'] = "Sistema"
            if not log['rol_usuario']:
                log['rol_usuario'] = ""
            estado_anterior = log['estado_destino']

        # Equipos
        cursor.execute("""
            SELECT u.nombre, rol.nombre
            FROM requerimiento_equipo_rol rer
            INNER JOIN usuarios u ON rer.usuario_id = u.id
            INNER JOIN roles_ciclodocumento rol ON rer.rol_id = rol.id
            WHERE rer.requerimiento_id = %s AND rer.activo = true
        """, [documento_id])
        for usuario_nombre, rol_nombre in cursor.fetchall():
            rol = rol_nombre.lower().strip()
            if rol == "redactor": equipo_redactores.append(usuario_nombre)
            elif rol == "revisor": equipo_revisores.append(usuario_nombre)
            elif rol == "aprobador": equipo_aprobadores.append(usuario_nombre)

    # Estadísticas adicionales
    acciones_por_usuario = Counter(log['usuario_nombre'] for log in logs)

    # Conteo de estados finales
    conteo_estados = Counter(log['estado_destino'] for log in logs)

    # Tiempo promedio por estado
    tiempos_estado = defaultdict(list)
    for i in range(1, len(logs)):
        estado_anterior = logs[i-1]['estado_destino']
        tiempo = logs[i]['fecha_accion'] - logs[i-1]['fecha_accion']
        tiempos_estado[estado_anterior].append(tiempo.total_seconds())

    tiempo_promedio_estado = {estado: sum(tiempos)/len(tiempos) for estado, tiempos in tiempos_estado.items() if tiempos}

    context = {
        "documento": {
            "titulo": f"{documento_info['tipo_documento']} - {documento_info['categoria']}" if documento_info else "Documento",
            "categoria": documento_info['categoria'] if documento_info else "",
            "tipo_documento": documento_info['tipo_documento'] if documento_info else ""
        },
        "equipo_redactores": sorted(equipo_redactores),
        "equipo_revisores": sorted(equipo_revisores),
        "equipo_aprobadores": sorted(equipo_aprobadores),
        "logs": logs,
        "acciones_por_usuario": list(acciones_por_usuario.items()),
        "conteo_estados": dict(conteo_estados),
        "tiempo_promedio_estado": tiempo_promedio_estado
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


# # Evita solapamiento de Redactor y Revisor
# redactores = set(proyecto_temp["equipo_roles"]["redactores"])
# revisores = set(proyecto_temp["equipo_roles"]["revisores"])
# if redactores & revisores:
#     raise ValidationError("Un usuario no puede ser Redactor y Revisor simultáneamente.")






@login_required 
def nuevo_requerimiento(request, proyecto_id):
    from django.utils import timezone
    from django.db import transaction
    from django.db import connection

    # --- Obtener info del proyecto ---
    with connection.cursor() as cursor:
        cursor.execute("SELECT id, nombre, descripcion FROM proyectos WHERE id=%s;", [proyecto_id])
        row = cursor.fetchone()
    if not row:
        return render(request, "error.html", {"mensaje": "Proyecto no encontrado."})

    proyecto = {"id": row[0], "nombre": row[1], "descripcion": row[2]}

    # --- Obtener tipos de documento técnico ---
    with connection.cursor() as cursor:
        cursor.execute("SELECT id, nombre FROM tipo_documentos_tecnicos ORDER BY nombre;")
        tipos_documento = cursor.fetchall()

    # --- Obtener usuarios ---
    with connection.cursor() as cursor:
        cursor.execute("SELECT id, nombre, email FROM usuarios ORDER BY nombre;")
        usuarios = cursor.fetchall()

    if request.method == "POST":
        # --- Paso 1: detalles ---
        tipo_doc_id = request.POST.get("tipo_documento")
        observaciones = request.POST.get("observaciones", "")

        # --- Paso 2: planificación ---
        fecha_primera_revision = request.POST.get("fecha_primera_revision")
        alertar_dias_revision = request.POST.get("alertar_dias_revision")
        fecha_entrega = request.POST.get("fecha_entrega")
        alertar_dias_entrega = request.POST.get("alertar_dias_entrega")

        # --- Paso 3: roles ---
        redactores = request.POST.getlist("redactores")
        revisores = request.POST.getlist("revisores")
        aprobadores = request.POST.getlist("aprobadores")

        try:
            with transaction.atomic():
                with connection.cursor() as cursor:
                    # 1️⃣ Insertar requerimiento
                    cursor.execute("""
                        INSERT INTO requerimiento_documento_tecnico
                        (proyecto_id, tipo_documento_id, fecha_registro, observaciones)
                        VALUES (%s, %s, %s, %s)
                        RETURNING id;
                    """, [proyecto_id, tipo_doc_id, timezone.now(), observaciones])
                    req_id = cursor.fetchone()[0]

                    # 2️⃣ Insertar log inicial
                    cursor.execute("""
                        INSERT INTO log_estado_requerimiento_documento
                        (requerimiento_id, usuario_id, estado_origen_id, estado_destino_id, created_at, observaciones)
                        VALUES (%s, %s, %s, %s, %s, %s);
                    """, [req_id, request.user.id, None, 1, timezone.now(),
                          "El sistema ha habilitado la plantilla en el repositorio."])

                    # 3️⃣ Insertar roles asignados
                    def insertar_roles(lista_usuarios, rol_id):
                        for u in lista_usuarios:
                            cursor.execute("""
                                INSERT INTO requerimiento_equipo_rol
                                (requerimiento_id, usuario_id, rol_id, fecha_asignacion, activo)
                                VALUES (%s, %s, %s, %s, TRUE);
                            """, [req_id, u, rol_id, timezone.now()])

                    insertar_roles(redactores, 1)
                    insertar_roles(revisores, 2)
                    insertar_roles(aprobadores, 3)

            return redirect('usuario:detalle_proyecto', proyecto_id=proyecto_id)

        except Exception as e:
            return render(request, "error.html", {"mensaje": f"Error al crear el requerimiento: {str(e)}"})

    return render(request, "nuevo_requerimiento.html", {
        "proyecto": proyecto,
        "tipos_documento": tipos_documento,
        "usuarios": usuarios,
    })



@login_required
def editar_requerimiento(request, requerimiento_id):
    # --- Obtener datos del requerimiento ---
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT rdt.id, rdt.proyecto_id, rdt.tipo_documento_id, TDT.nombre, rdt.observaciones
            FROM requerimiento_documento_tecnico rdt
            LEFT JOIN tipo_documentos_tecnicos TDT ON rdt.tipo_documento_id = TDT.id
            WHERE rdt.id = %s
        """, [requerimiento_id])
        row = cursor.fetchone()

    if not row:
        return render(request, "error.html", {"mensaje": "Requerimiento no encontrado."})

    requerimiento = {
        "id": row[0],
        "proyecto_id": row[1],
        "tipo_documento_id": row[2],
        "nombre_tipo_documento": row[3],
        "observaciones": row[4]
    }

    # --- Obtener usuarios para asignar roles ---
    with connection.cursor() as cursor:
        cursor.execute("SELECT id, nombre FROM usuarios ORDER BY nombre;")
        usuarios = cursor.fetchall()

        cursor.execute("""
            SELECT usuario_id, rol_id
            FROM requerimiento_equipo_rol
            WHERE requerimiento_id = %s AND activo = TRUE
        """, [requerimiento_id])
        roles_asignados = cursor.fetchall()
        redactores = [u for u, r in roles_asignados if r == 1]
        revisores = [u for u, r in roles_asignados if r == 2]
        aprobadores = [u for u, r in roles_asignados if r == 3]

    if request.method == "POST":
        observaciones = request.POST.get("observaciones", "")
        redactores_post = request.POST.getlist("redactores")
        revisores_post = request.POST.getlist("revisores")
        aprobadores_post = request.POST.getlist("aprobadores")

        with connection.cursor() as cursor:
            # Actualizar observaciones
            cursor.execute("""
                UPDATE requerimiento_documento_tecnico
                SET observaciones=%s
                WHERE id=%s
            """, [observaciones, requerimiento_id])

            # Actualizar roles: desactivar todos primero
            cursor.execute("""
                UPDATE requerimiento_equipo_rol
                SET activo=FALSE
                WHERE requerimiento_id=%s
            """, [requerimiento_id])

            # Insertar roles seleccionados
            for u in redactores_post:
                cursor.execute("""
                    INSERT INTO requerimiento_equipo_rol
                    (requerimiento_id, usuario_id, rol_id, fecha_asignacion, activo)
                    VALUES (%s, %s, 1, NOW(), TRUE)
                """, [requerimiento_id, u])
            for u in revisores_post:
                cursor.execute("""
                    INSERT INTO requerimiento_equipo_rol
                    (requerimiento_id, usuario_id, rol_id, fecha_asignacion, activo)
                    VALUES (%s, %s, 2, NOW(), TRUE)
                """, [requerimiento_id, u])
            for u in aprobadores_post:
                cursor.execute("""
                    INSERT INTO requerimiento_equipo_rol
                    (requerimiento_id, usuario_id, rol_id, fecha_asignacion, activo)
                    VALUES (%s, %s, 3, NOW(), TRUE)
                """, [requerimiento_id, u])

        return redirect('usuario:detalle_proyecto', proyecto_id=requerimiento["proyecto_id"])

    return render(request, "editar_requerimiento.html", {
        "requerimiento": requerimiento,
        "usuarios": usuarios,
        "redactores": redactores,
        "revisores": revisores,
        "aprobadores": aprobadores
    })


@login_required
def eliminar_requerimiento(request, requerimiento_id):
    """Elimina un requerimiento de documento técnico y todas sus asociaciones de forma segura."""
    if request.method == "POST":
        # Obtener proyecto_id asociado al requerimiento directamente desde la DB
        with connection.cursor() as cursor:
            cursor.execute("""
                SELECT proyecto_id
                FROM requerimiento_documento_tecnico
                WHERE id = %s
            """, [requerimiento_id])
            row = cursor.fetchone()
            if row:
                proyecto_id = row[0]
            else:
                messages.error(request, "No se encontró el requerimiento.")
                return redirect('usuario:lista_proyectos')

            # Eliminar el requerimiento
            cursor.execute("""
                DELETE FROM requerimiento_documento_tecnico
                WHERE id = %s
            """, [requerimiento_id])

        messages.success(request, "El requerimiento y todos sus registros relacionados fueron eliminados correctamente.")
        return redirect('usuario:detalle_proyecto', proyecto_id=proyecto_id)

    # Si se intenta acceder por GET, redirigir a lista de proyectos
    return redirect('usuario:usuario_proyectos')


from django.http import JsonResponse
import json
import pprint

def _get_nombres_por_ids(valor, usuarios):
    """
    Convierte una lista o string de IDs en una cadena de nombres de usuario.
    Ejemplo: '1' -> 'Juan Pérez', ['1','2'] -> 'Juan Pérez, María Gómez'
    """
    if not valor:
        return "-"
    if isinstance(valor, str):
        valor = [valor]
    nombres = []
    for v in valor:
        for u in usuarios:
            if str(u["id"]) == str(v):
                nombres.append(u["nombre"])
    return ", ".join(nombres) if nombres else "-"


@login_required
def crear_proyecto(request):
    from django.db import connection

    # === Definir pasos ===
    steps = ["Datos Generales", "Contrato y Cliente", "Grupos y Documentos", "Confirmación"]

    # === Cargar datos para selects ===
    with connection.cursor() as cursor:
        cursor.execute("SELECT id, nombre, email FROM usuarios ORDER BY nombre")
        usuarios = [{"id": r[0], "nombre": r[1], "email": r[2]} for r in cursor.fetchall()]

        cursor.execute("SELECT id, nombre, email FROM usuarios WHERE cargo_id = 4 ORDER BY nombre")
        usuarios_administrador = [{"id": r[0], "nombre": r[1], "email": r[2]} for r in cursor.fetchall()]

        cursor.execute("SELECT id, nombre, descripcion FROM categoria_documentos_tecnicos ORDER BY nombre")
        grupos_maestros = [{"id": r[0], "nombre": r[1], "descripcion": r[2]} for r in cursor.fetchall()]

        cursor.execute("SELECT id, categoria_id, nombre FROM tipo_documentos_tecnicos ORDER BY nombre")
        documentos = [{"id": r[0], "categoria_id": r[1], "nombre": r[2]} for r in cursor.fetchall()]

        cursor.execute("SELECT id, nombre, abreviatura FROM maquinas ORDER BY nombre")
        maquinas = [{"id": r[0], "nombre": r[1], "abreviatura": r[2] or ""} for r in cursor.fetchall()]

        cursor.execute("""
            SELECT c.id, c.numero_contrato, c.monto_total, c.fecha_firma,
                   c.representante_cliente_nombre, c.representante_cliente_correo,
                   c.representante_cliente_telefono, c.cliente_id,
                   cl.nombre, cl.rut, cl.direccion, cl.correo_contacto, cl.telefono_contacto
            FROM contratos c
            JOIN clientes cl ON cl.id = c.cliente_id
            ORDER BY c.numero_contrato
        """)
        contratos = [{
            "id": r[0],
            "numero_contrato": r[1],
            "monto_total": r[2],
            "fecha_firma": r[3],
            "representante_cliente_nombre": r[4],
            "representante_cliente_correo": r[5],
            "representante_cliente_telefono": r[6],
            "cliente_id": r[7],
            "cliente_nombre": r[8],
            "cliente_rut": r[9],
            "cliente_direccion": r[10],
            "cliente_correo": r[11],
            "cliente_telefono": r[12],
        } for r in cursor.fetchall()]

        cursor.execute("SELECT id, nombre FROM clientes ORDER BY nombre")
        clientes = [{"id": r[0], "nombre": r[1]} for r in cursor.fetchall()]

        cursor.execute("SELECT id, cliente_id, nombre, ubicacion FROM faenas ORDER BY nombre")
        faenas = [{"id": r[0], "cliente_id": r[1], "nombre": r[2], "ubicacion": r[3]} for r in cursor.fetchall()]

    # Vincular documentos con su grupo
    for grupo in grupos_maestros:
        grupo_docs = [d for d in documentos if d["categoria_id"] == grupo["id"]]
        grupo["documentos"] = grupo_docs

    # === POST FINAL: simular inserciones ===
    if request.method == "POST":
        form_data = request.POST

        with connection.cursor() as cursor:
            print("\n" + "=" * 90)
            print("💡 [MODO DESARROLLO] Simulación de inserciones SQL")
            print("=" * 90)

            # CLIENTE
            cliente_id = form_data.get("cliente_id")
            if not cliente_id and form_data.get("cliente_nombre"):
                # cursor.execute("""
                #     INSERT INTO clientes (nombre, rut, direccion, correo_contacto, telefono_contacto)
                #     VALUES (%s, %s, %s, %s, %s) RETURNING id
                # """, [
                #     form_data.get("cliente_nombre"),
                #     form_data.get("cliente_rut"),
                #     form_data.get("cliente_direccion"),
                #     form_data.get("cliente_correo"),
                #     form_data.get("cliente_telefono")
                # ])
                print("🧾 INSERT CLIENTE →", {
                    "nombre": form_data.get("cliente_nombre"),
                    "rut": form_data.get("cliente_rut"),
                    "direccion": form_data.get("cliente_direccion"),
                    "correo_contacto": form_data.get("cliente_correo"),
                    "telefono_contacto": form_data.get("cliente_telefono")
                })
                cliente_id = "SIM_CLIENTE_ID"

            # CONTRATO
            contrato_id = form_data.get("contrato_id")
            if not contrato_id:
                # cursor.execute("""
                #     INSERT INTO contratos (
                #         numero_contrato, monto_total, fecha_firma,
                #         representante_cliente_nombre, representante_cliente_correo,
                #         representante_cliente_telefono, cliente_id, created_at
                #     )
                #     VALUES (%s, %s, %s, %s, %s, %s, %s, CURRENT_TIMESTAMP)
                #     RETURNING id
                # """, [
                #     form_data.get("numero_contrato"),
                #     form_data.get("monto_total"),
                #     form_data.get("contrato_fecha_firma"),
                #     form_data.get("representante_cliente_nombre"),
                #     form_data.get("representante_cliente_correo"),
                #     form_data.get("representante_cliente_telefono"),
                #     cliente_id
                # ])
                print("📑 INSERT CONTRATO →", {
                    "numero_contrato": form_data.get("numero_contrato"),
                    "monto_total": form_data.get("monto_total"),
                    "fecha_firma": form_data.get("contrato_fecha_firma"),
                    "representante_cliente_nombre": form_data.get("representante_cliente_nombre"),
                    "representante_cliente_correo": form_data.get("representante_cliente_correo"),
                    "representante_cliente_telefono": form_data.get("representante_cliente_telefono"),
                    "cliente_id": cliente_id
                })
                contrato_id = "SIM_CONTRATO_ID"

            # FAENA
            faena_id = form_data.get("faena_id")
            if not faena_id and form_data.get("faena_nombre"):
                # cursor.execute("""
                #     INSERT INTO faenas (cliente_id, nombre, ubicacion)
                #     VALUES (%s, %s, %s) RETURNING id
                # """, [cliente_id, form_data.get("faena_nombre"), form_data.get("faena_ubicacion")])
                print("🏗️ INSERT FAENA →", {
                    "cliente_id": cliente_id,
                    "nombre": form_data.get("faena_nombre"),
                    "ubicacion": form_data.get("faena_ubicacion")
                })
                faena_id = "SIM_FAENA_ID"

            # PROYECTO
            # cursor.execute("""
            #     INSERT INTO proyectos (
            #         nombre, descripcion, fecha_recepcion_evaluacion,
            #         fecha_inicio_planificacion, fecha_inicio_ejecucion,
            #         fecha_cierre_proyecto, abreviatura,
            #         administrador_id, contrato_id, cliente_id, faena_id
            #     )
            #     VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            #     RETURNING id
            # """, [
            #     form_data.get("nombre"),
            #     form_data.get("descripcion"),
            #     form_data.get("fecha_recepcion_evaluacion"),
            #     form_data.get("fecha_inicio_planificacion"),
            #     form_data.get("fecha_inicio_ejecucion"),
            #     form_data.get("fecha_cierre_proyecto"),
            #     form_data.get("abreviatura"),
            #     form_data.get("administrador") or form_data.get("administrador_id"),
            #     contrato_id,
            #     cliente_id,
            #     faena_id
            # ])
            print("📋 INSERT PROYECTO →", {
                "nombre": form_data.get("nombre"),
                "descripcion": form_data.get("descripcion"),
                "fecha_recepcion_evaluacion": form_data.get("fecha_recepcion_evaluacion"),
                "fecha_inicio_planificacion": form_data.get("fecha_inicio_planificacion"),
                "fecha_inicio_ejecucion": form_data.get("fecha_inicio_ejecucion"),
                "fecha_cierre_proyecto": form_data.get("fecha_cierre_proyecto"),
                "abreviatura": form_data.get("abreviatura"),
                "administrador_id": form_data.get("administrador") or form_data.get("administrador_id"),
                "contrato_id": contrato_id,
                "cliente_id": cliente_id,
                "faena_id": faena_id
            })
            proyecto_id = "SIM_PROYECTO_ID"

            # MÁQUINAS
            maquinas_ids = form_data.getlist("maquinas_ids[]")
            for maquina_id in maquinas_ids:
                # cursor.execute("""
                #     INSERT INTO proyecto_maquina (proyecto_id, maquina_id)
                #     VALUES (%s, %s) ON CONFLICT DO NOTHING
                # """, [proyecto_id, maquina_id])
                print("🔩 INSERT PROYECTO_MAQUINA →", {"proyecto_id": proyecto_id, "maquina_id": maquina_id})

            # DOCUMENTOS TÉCNICOS Y ROLES
            print("\n📂 Documentos técnicos seleccionados y asignaciones:")
            documentos_ids = form_data.getlist("documento_id[]")

            for doc_id in documentos_ids:
                redactores = form_data.getlist(f"redactor_id_{doc_id}[]")
                revisores = form_data.getlist(f"revisor_id_{doc_id}[]")
                aprobadores = form_data.getlist(f"aprobador_id_{doc_id}[]")

                # cursor.execute("""
                #     INSERT INTO requerimiento_documento_tecnico (proyecto_id, tipo_documento_id)
                #     VALUES (%s, %s) RETURNING id
                # """, [proyecto_id, doc_id])
                # requerimiento_id = cursor.fetchone()[0]

                print({
                    "documento_id": doc_id,
                    "redactores": redactores,
                    "revisores": revisores,
                    "aprobadores": aprobadores
                })

            print("=" * 90 + "\n")

        messages.info(request, "💡 [Modo desarrollo] Datos impresos en consola, sin guardar en la BDD.")
        return redirect("usuario:lista_proyectos")

    # === Render ===
    step_templates = [
        "crear_proyecto_paso1.html",
        "crear_proyecto_paso2.html",
        "crear_proyecto_paso3.html",
        "crear_proyecto_paso4.html",
    ]

    step_actual = int(request.GET.get("step", "1"))

    return render(request, "crear_proyecto.html", {
        "steps": steps,
        "step_templates": step_templates,
        "usuarios": usuarios,
        "usuarios_administrador": usuarios_administrador,
        "grupos_maestros": grupos_maestros,
        "documentos": documentos,
        "maquinas": maquinas,
        "contratos": contratos,
        "clientes": clientes,
        "faenas": faenas,
        "step_actual": step_actual,
    })



#Paso 1
@csrf_exempt
@login_required
def generar_abreviatura_proyecto(request):
    """
    Genera la abreviatura de un proyecto a partir de:
    - nombre de la máquina
    - descripcion
    - fecha_recepcion_evaluacion
    """
    if request.method != "POST":
        return JsonResponse({"error": "Método no permitido"}, status=405)

    try:
        data = json.loads(request.body)
        maquina = data.get("maquina", "").strip()
        descripcion = data.get("descripcion", "").strip().upper()
        fecha = data.get("fecha_recepcion_evaluacion", "").strip()

        if not maquina or not descripcion or not fecha:
            return JsonResponse({"abreviatura": ""})

        # Formato MMYY
        try:
            fecha_obj = datetime.datetime.strptime(fecha, "%Y-%m-%d")
            mes = f"{fecha_obj.month:02}"
            anio = str(fecha_obj.year)[-2:]
            fecha_formato = f"{mes}{anio}"
        except ValueError:
            fecha_formato = ""

        nombre_limpio = maquina.replace(" ", "").upper()
        descripcion_limpio = descripcion.replace(" ", "")

        abreviatura = f"{nombre_limpio}.{descripcion_limpio}.{fecha_formato}"

        return JsonResponse({"abreviatura": abreviatura})

    except Exception as e:
        return JsonResponse({"error": str(e)}, status=400)
#paso2 ####################  ####################    ####################    ####################  


def generar_abreviatura_cliente(nombre):
    nombre = unidecode.unidecode(nombre.upper().strip())
    nombre = re.sub(r'[^A-ZÑ\s]', ' ', nombre)
    nombre = re.sub(r'\s+', ' ', nombre).strip()
    palabras = nombre.split()

    stopwords = {
        "DE", "DEL", "LA", "LOS", "LAS", "Y", "E", "S", "SA", "SAA",
        "LTDA", "LIMITADA", "COMPANIA", "COMPAÑIA", "CORPORACION",
        "CORPORACIÓN", "EMPRESA", "GRUPO", "INDUSTRIAL", "SERVICIOS"
    }
    palabras = [p for p in palabras if p not in stopwords]

    if not palabras:
        base = "GEN"
    elif len(palabras) == 1:
        base = palabras[0][:6]
    else:
        partes = [p[:3] for p in palabras[:3]]
        base = ''.join(partes).upper()
        base = base[:6] if len(base) > 6 else base

    abrev_final = base
    contador = 1
    with connection.cursor() as cursor:
        while True:
            cursor.execute("SELECT COUNT(*) FROM clientes WHERE abreviatura = %s;", [abrev_final])
            if cursor.fetchone()[0] == 0:
                break
            abrev_final = f"{base}{contador}"
            contador += 1
    return abrev_final



@csrf_exempt
@login_required
def generar_abreviatura_cliente_ajax(request):
    """
    Genera la abreviatura desde el backend según el nombre enviado por AJAX.
    Mantiene las mismas reglas que generar_abreviatura_cliente().
    """
    try:
        nombre = request.GET.get("nombre", "").strip()
        if not nombre:
            return JsonResponse({"abreviatura": ""})
        
        abreviatura = generar_abreviatura_cliente(nombre)
        return JsonResponse({"abreviatura": abreviatura})

    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)



@login_required
@csrf_exempt
def obtener_abreviatura_cliente(request, cliente_id):
    """
    Devuelve la abreviatura registrada para un cliente existente.
    """
    try:
        with connection.cursor() as cursor:
            cursor.execute("SELECT abreviatura FROM clientes WHERE id = %s;", [cliente_id])
            result = cursor.fetchone()
            if result and result[0]:
                return JsonResponse({"abreviatura": result[0]})
            else:
                return JsonResponse({"abreviatura": None})
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)
    

 ####################  ####################    ####################    ####################     


def obtener_numordenservicio(path_archivo):
    wb = load_workbook(path_archivo, data_only=True)

    # Recorre todos los nombres definidos en el libro
    for name, defn in wb.defined_names.items():
        if name.lower() == "numordenservicio":
            try:
                for title, coord in defn.destinations:
                    sheet = wb[title]
                    cell = sheet[coord]
                    return str(cell.value or "")
            except Exception as e:
                print(f"⚠️ Error leyendo numordenservicio: {e}")
                return ""
    return ""

@login_required
def leer_excel_numero_orden(request):
    if request.method == "POST" and request.FILES.get("archivo"):
        archivo = request.FILES["archivo"]
        import tempfile, os

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:

            for chunk in archivo.chunks():
                tmp.write(chunk)
            tmp_path = tmp.name
        numero_orden = obtener_numordenservicio(tmp_path)
        os.remove(tmp_path)
        return JsonResponse({"numero_orden": numero_orden})
    return JsonResponse({"numero_orden": ""})


@login_required
def editar_usuario(request, user_id):
    # Traer usuario
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id, nombre, email, email_secundario,
                   telefono_corporativo, telefono_secundario, area_trabajo, cargo_id
            FROM usuarios
            WHERE id = %s
        """, [user_id])
        row = cursor.fetchone()

    if not row:
        messages.error(request, "Usuario no encontrado.")
        return redirect('usuario:lista_usuarios')

    usuario = {
        'id': row[0],
        'nombre': row[1],
        'email': row[2],
        'email_secundario': row[3],
        'telefono_corporativo': row[4],
        'telefono_secundario': row[5],
        'area_trabajo': row[6],
        'cargo_id': row[7],
    }

    # Traer áreas
    with connection.cursor() as cursor:
        cursor.execute("SELECT id, nombre FROM area_cargo_empresa")
        areas = cursor.fetchall()

    # Traer todos los cargos
    with connection.cursor() as cursor:
        cursor.execute("SELECT id, nombre, area_id FROM cargos_empresa")
        cargos = cursor.fetchall()

    # Convertir a diccionarios
    areas = [{'id': a[0], 'nombre': a[1]} for a in areas]
    cargos = [{'id': c[0], 'nombre': c[1], 'area_id': c[2]} for c in cargos]

    if request.method == "POST":
        nombre = request.POST.get('nombre', usuario['nombre'])
        email = request.POST.get('email', usuario['email'])
        email_secundario = request.POST.get('email_secundario', usuario['email_secundario'])
        telefono_corporativo = request.POST.get('telefono_corporativo', usuario['telefono_corporativo'])
        telefono_secundario = request.POST.get('telefono_secundario', usuario['telefono_secundario'])
        area_trabajo_id = request.POST.get('area_trabajo')
        cargo_id = request.POST.get('cargo_id')

        # Guardar area_trabajo como nombre de área
        area_nombre = next((a['nombre'] for a in areas if str(a['id'])==area_trabajo_id), '')

        with connection.cursor() as cursor:
            cursor.execute("""
                UPDATE usuarios
                SET nombre=%s, email=%s, email_secundario=%s,
                    telefono_corporativo=%s, telefono_secundario=%s,
                    area_trabajo=%s, cargo_id=%s
                WHERE id=%s
            """, [nombre, email, email_secundario,
                  telefono_corporativo, telefono_secundario,
                  area_nombre, cargo_id, user_id])

        messages.success(request, f"Usuario {nombre} actualizado correctamente.")
        return redirect('usuario:lista_usuarios')

    return render(request, 'editar_usuario.html', {
        'usuario': usuario,
        'areas': areas,
        'cargos': cargos,
    })



@login_required
def ver_estadisticas_usuario(request, user_id):
    # Obtener datos del usuario
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id, nombre, email, email_secundario,
                   telefono_corporativo, telefono_secundario,
                   area_trabajo, cargo_id, fecha_registro
            FROM usuarios
            WHERE id = %s
        """, [user_id])
        row = cursor.fetchone()

    if not row:
        messages.error(request, "Usuario no encontrado.")
        return redirect('usuario:lista_usuarios')

    usuario = {
        'id': row[0],
        'nombre': row[1],
        'email_corporativo': row[2],
        'email_secundario': row[3],
        'telefono_corporativo': row[4],
        'telefono_secundario': row[5],
        'area_trabajo': row[6],
        'cargo': row[7],  # Opcional: traducir cargo_id a nombre después
        'fecha_registro': row[8],
    }

    # Por ahora, solo renderizamos el HTML vacío de estadísticas
    return render(request, 'usuario_estadisticas.html', {
        'usuario': usuario,
    })


@login_required
def editar_proyecto(request, proyecto_id):
    # Obtener datos del proyecto incluyendo el id del administrador
    sql = """
    SELECT
        P.id AS proyecto_id,
        P.nombre AS nombre_proyecto,
        P.descripcion AS proyecto_descripcion,
        P.numero_orden,
        P.fecha_recepcion_evaluacion,
        P.fecha_cierre_proyecto,
        U.id AS administrador_id,
        U.nombre AS administrador_nombre_completo,
        U.email AS administrador_email
    FROM public.proyectos P
    LEFT JOIN public.usuarios U ON P.administrador_id = U.id
    WHERE P.id = %s
    """
    
    with connection.cursor() as cursor:
        cursor.execute(sql, [proyecto_id])
        row = cursor.fetchone()
        if not row:
            raise Http404("Proyecto no encontrado")

        columns = [col[0] for col in cursor.description]
        proyecto = dict(zip(columns, row))

        # Lista de administradores con cargo_id = 4
        cursor.execute("""
            SELECT u.id, u.nombre, u.email
            FROM public.usuarios u
            WHERE u.cargo_id = 4
            ORDER BY u.nombre
        """)
        administradores = [
            dict(zip([col[0] for col in cursor.description], r))
            for r in cursor.fetchall()
        ]

    # Formatear fechas para input type="date"
    proyecto['fecha_recepcion_evaluacion'] = (
        proyecto['fecha_recepcion_evaluacion'].strftime('%Y-%m-%d')
        if proyecto['fecha_recepcion_evaluacion'] else ''
    )
    proyecto['fecha_cierre_proyecto'] = (
        proyecto['fecha_cierre_proyecto'].strftime('%Y-%m-%d')
        if proyecto['fecha_cierre_proyecto'] else ''
    )

    context = {
        'proyecto': proyecto,
        'administradores': administradores
    }
    return render(request, "editar_proyecto.html", context)



def editar_contrato(request, contrato_id):
    contrato = get_object_or_404(Contrato, id=contrato_id)
    return render(request, 'usuario/editar_contrato.html', {'contrato': contrato})

def editar_cliente(request, cliente_id):
    cliente = get_object_or_404(Cliente, id=cliente_id)
    return render(request, 'usuario/editar_cliente.html', {'cliente': cliente})

@login_required
def editar_maquina(request, maquina_id):
    with connection.cursor() as cursor:
        # Buscar los datos de la máquina
        cursor.execute("""
            SELECT m.id, m.nombre, m.abreviatura, m.marca, m.modelo, m.anio_fabricacion,
                   m.tipo, m.descripcion, pm.proyecto_id
            FROM maquinas m
            LEFT JOIN proyecto_maquina pm ON pm.maquina_id = m.id
            WHERE m.id = %s
        """, [maquina_id])
        row = cursor.fetchone()

    if not row:
        raise Http404("Máquina no encontrada")

    maquina = {
        'id': row[0],
        'nombre': row[1],
        'abreviatura': row[2],
        'marca': row[3],
        'modelo': row[4],
        'anio_fabricacion': row[5],
        'tipo': row[6],
        'descripcion': row[7],
        'proyecto_id': row[8],  # ← tomado desde proyecto_maquina
    }

    return render(request, 'editar_maquina.html', {'maquina': maquina})
