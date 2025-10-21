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
from django.db import connection 
from collections import defaultdict
import json 
# You need to import the decorator:
from django.views.decorators.http import require_POST 
from django.shortcuts import get_object_or_404, render
from django.contrib import messages





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
    LEFT JOIN public.contratos C ON P.contrato_id = C.id
    LEFT JOIN public.clientes CL ON C.cliente_id = CL.id
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
    proyecto_info = resultados[0] if resultados else {
    'proyecto_id': proyecto_id,  # <-- aseguramos que siempre haya ID
    'nombre_proyecto': '',
    'proyecto_descripcion': '',
    'numero_orden': '',
    'administrador_nombre_completo': '',
    'administrador_email': '',
    'nombre_faena': '',
    'numero_contrato': '',
    'monto_total': '',
    'contrato_fecha_creacion': None,
    'representante_cliente_nombre': '',
    'representante_cliente_correo': '',
    'representante_cliente_telefono': '',
    'cliente_nombre': '',
    'cliente_rut': '',
    'cliente_direccion': '',
    'cliente_correo': '',
    'cliente_telefono': '',
    }

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

@login_required
def nuevo_requerimiento(request, proyecto_id):
    from django.utils import timezone

    # --- Obtener info del proyecto ---
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id, nombre, descripcion
            FROM proyectos
            WHERE id = %s;
        """, [proyecto_id])
        row = cursor.fetchone()

    if not row:
        return render(request, "error.html", {"mensaje": "Proyecto no encontrado."})

    proyecto = {
        "id": row[0],
        "nombre": row[1],
        "descripcion": row[2],
    }

    # --- Obtener tipos de documento técnico ---
    with connection.cursor() as cursor:
        cursor.execute("SELECT id, nombre FROM tipo_documentos_tecnicos ORDER BY nombre;")
        tipos_documento = cursor.fetchall()

    # --- Obtener usuarios ---
    with connection.cursor() as cursor:
        cursor.execute("SELECT id, nombre, email FROM usuarios ORDER BY nombre;")
        usuarios = cursor.fetchall()

    # --- Si el formulario fue enviado ---
    if request.method == "POST":
        tipo_doc_id = request.POST.get("tipo_documento")
        observaciones = request.POST.get("observaciones", "")
        redactores = request.POST.getlist("redactores")
        revisores = request.POST.getlist("revisores")
        aprobadores = request.POST.getlist("aprobadores")

        with connection.cursor() as cursor:
            # 1️⃣ Insertar en requerimiento_documento_tecnico
            cursor.execute("""
                INSERT INTO requerimiento_documento_tecnico
                (proyecto_id, tipo_documento_id, fecha_registro, observaciones)
                VALUES (%s, %s, NOW(), %s)
                RETURNING id;
            """, [proyecto_id, tipo_doc_id, observaciones])
            req_id = cursor.fetchone()[0]

            # 2️⃣ Registrar log inicial “Sistema → Borrador”
            # Usamos usuario_id=NULL para indicar que es el sistema
            cursor.execute("""
                INSERT INTO log_estado_requerimiento_documento
                (requerimiento_id, usuario_id, estado_origen_id, estado_destino_id, created_at, observaciones)
                VALUES (%s, NULL, NULL, (SELECT id FROM estado_documento WHERE nombre='Borrador'), NOW(), %s)
            """, [req_id, 'El sistema ha habilitado la plantilla en el repositorio.'])

            # 3️⃣ Asignar roles (redactor=1, revisor=2, aprobador=3)
            for u in redactores:
                cursor.execute("""
                    INSERT INTO requerimiento_equipo_rol
                    (requerimiento_id, usuario_id, rol_id, fecha_asignacion, activo)
                    VALUES (%s, %s, 1, NOW(), TRUE);
                """, [req_id, u])
            for u in revisores:
                cursor.execute("""
                    INSERT INTO requerimiento_equipo_rol
                    (requerimiento_id, usuario_id, rol_id, fecha_asignacion, activo)
                    VALUES (%s, %s, 2, NOW(), TRUE);
                """, [req_id, u])
            for u in aprobadores:
                cursor.execute("""
                    INSERT INTO requerimiento_equipo_rol
                    (requerimiento_id, usuario_id, rol_id, fecha_asignacion, activo)
                    VALUES (%s, %s, 3, NOW(), TRUE);
                """, [req_id, u])

        return redirect('usuario:detalle_proyecto', proyecto_id=proyecto_id)

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



@login_required
def crear_proyecto_wizard(request, paso=1):
    """
    Wizard multipágina usando SQL crudo
    Paso 1: Proyecto
    Paso 2: Contrato
    Paso 3: Cliente
    Paso 4: Máquinas
    Paso 5: Requerimientos (roles)
    Paso 6: Verificación y guardado final
    """
    usuarios, usuarios_administrador, grupos_maestros, documentos = [], [], [], []

    # Datos generales para formularios
    with connection.cursor() as cursor:
        cursor.execute("SELECT id, nombre, email FROM usuarios ORDER BY nombre")
        usuarios = [{"id": r[0], "nombre": r[1], "email": r[2]} for r in cursor.fetchall()]

        cursor.execute("""
            SELECT u.id, u.nombre, u.email
            FROM usuarios u
            WHERE u.cargo_id = 4
            ORDER BY u.nombre
        """)
        usuarios_administrador = [{"id": r[0], "nombre": r[1], "email": r[2]} for r in cursor.fetchall()]

        cursor.execute("SELECT id, nombre, descripcion FROM categoria_documentos_tecnicos ORDER BY nombre")
        grupos_maestros = [{"id": r[0], "nombre": r[1], "descripcion": r[2]} for r in cursor.fetchall()]

        cursor.execute("SELECT id, categoria_id, nombre FROM tipo_documentos_tecnicos ORDER BY nombre")
        documentos = [{"id": r[0], "categoria_id": r[1], "nombre": r[2]} for r in cursor.fetchall()]

    # Inicializar sesión temporal
    if 'proyecto_temp' not in request.session:
        request.session['proyecto_temp'] = {}
    proyecto_temp = request.session['proyecto_temp']

    form_error = None

    if request.method == "POST":
        # ---------------- PASO 1: Proyecto ----------------
        if paso == 1:
            proyecto_temp['nombre'] = request.POST.get('nombre')
            proyecto_temp['descripcion'] = request.POST.get('descripcion')
            proyecto_temp['fecha_inicio'] = request.POST.get('fecha_inicio')
            proyecto_temp['fecha_fin'] = request.POST.get('fecha_fin')
            proyecto_temp['administrador_id'] = request.POST.get('administrador')
            proyecto_temp['numero_orden'] = request.POST.get('numero_orden')

            if not proyecto_temp['nombre'] or not proyecto_temp['fecha_inicio']:
                form_error = "Faltan campos obligatorios."
            else:
                request.session['proyecto_temp'] = proyecto_temp
                return redirect('usuario:crear_proyecto_wizard', paso=2)

        # ---------------- PASO 2: Contrato ----------------
        elif paso == 2:
            proyecto_temp['numero_contrato'] = request.POST.get('numero_contrato')
            proyecto_temp['monto_total'] = request.POST.get('monto_total')
            proyecto_temp['contrato_fecha_creacion'] = request.POST.get('fecha_creacion')
            request.session['proyecto_temp'] = proyecto_temp
            return redirect('usuario:crear_proyecto_wizard', paso=3)

        # ---------------- PASO 3: Cliente ----------------
        elif paso == 3:
            proyecto_temp['cliente_nombre'] = request.POST.get('cliente_nombre')
            proyecto_temp['cliente_rut'] = request.POST.get('cliente_rut')
            proyecto_temp['cliente_direccion'] = request.POST.get('cliente_direccion')
            proyecto_temp['cliente_correo'] = request.POST.get('cliente_correo')
            proyecto_temp['cliente_telefono'] = request.POST.get('cliente_telefono')
            request.session['proyecto_temp'] = proyecto_temp
            return redirect('usuario:crear_proyecto_wizard', paso=4)

        # ---------------- PASO 4: Máquinas ----------------
        elif paso == 4:
            proyecto_temp['maquinas'] = request.POST.getlist('maquinas[]')
            request.session['proyecto_temp'] = proyecto_temp
            return redirect('usuario:crear_proyecto_wizard', paso=5)

        # ---------------- PASO 5: Requerimientos (roles) ----------------
        elif paso == 5:
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
                        "aprobadores": aprobadores
                    }
            proyecto_temp['documentos_roles'] = documentos_roles
            request.session['proyecto_temp'] = proyecto_temp
            # Redirigir al Paso 6 para verificación
            return redirect('usuario:crear_proyecto_wizard', paso=6)

        # ---------------- PASO 6: Verificación y guardado final ----------------
        elif paso == 6:
            try:
                with connection.cursor() as cursor:
                    cursor.execute("""
                        INSERT INTO proyectos (numero_orden, nombre, descripcion, fecha_inicio, fecha_fin, administrador_id)
                        VALUES (%s, %s, %s, %s, %s, %s) RETURNING id
                    """, [
                        proyecto_temp['numero_orden'],
                        proyecto_temp['nombre'],
                        proyecto_temp['descripcion'],
                        proyecto_temp['fecha_inicio'],
                        proyecto_temp['fecha_fin'],
                        proyecto_temp['administrador_id']
                    ])
                    proyecto_id = cursor.fetchone()[0]

                    # Guardar documentos y roles
                    for doc_id, roles in proyecto_temp.get('documentos_roles', {}).items():
                        cursor.execute("""
                            INSERT INTO requerimiento_documento_tecnico (proyecto_id, tipo_documento_id)
                            VALUES (%s, %s) RETURNING id
                        """, [proyecto_id, doc_id])
                        requerimiento_id = cursor.fetchone()[0]

                        for u_id in roles.get("redactores", []):
                            cursor.execute("""
                                INSERT INTO requerimiento_equipo_rol (requerimiento_id, usuario_id, rol_id)
                                VALUES (%s, %s, 1)
                            """, [requerimiento_id, u_id])
                        for u_id in roles.get("revisores", []):
                            cursor.execute("""
                                INSERT INTO requerimiento_equipo_rol (requerimiento_id, usuario_id, rol_id)
                                VALUES (%s, %s, 2)
                            """, [requerimiento_id, u_id])
                        for u_id in roles.get("aprobadores", []):
                            cursor.execute("""
                                INSERT INTO requerimiento_equipo_rol (requerimiento_id, usuario_id, rol_id)
                                VALUES (%s, %s, 3)
                            """, [requerimiento_id, u_id])

                del request.session['proyecto_temp']
                return redirect('proyectos_lista')
            except Exception as e:
                form_error = f"Error al guardar proyecto: {str(e)}"

    # Renderizar el template correspondiente según el paso
    template_name = f"crear_proyecto_paso{paso}.html"
    context = {
        'paso': paso,
        'form_error': form_error,
        'usuarios': usuarios,
        'usuarios_administrador': usuarios_administrador,
        'grupos_maestros': grupos_maestros,
        'documentos': documentos,
        'proyecto_temp': proyecto_temp
    }

    return render(request, template_name, context)


