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
    from django.db import transaction, connection

    # --- Obtener info del proyecto ---
    with connection.cursor() as cursor:
        cursor.execute("SELECT id, nombre, descripcion FROM proyectos WHERE id=%s;", [proyecto_id])
        row = cursor.fetchone()
    if not row:
        return render(request, "error.html", {"mensaje": "Proyecto no encontrado."})

    proyecto = {"id": row[0], "nombre": row[1], "descripcion": row[2]}

    # --- Tipos de documento técnico ---
    with connection.cursor() as cursor:
        cursor.execute("SELECT id, nombre FROM tipo_documentos_tecnicos ORDER BY nombre;")
        tipos_documento = cursor.fetchall()

    # --- Usuarios disponibles ---
    with connection.cursor() as cursor:
        cursor.execute("SELECT id, nombre, email FROM usuarios ORDER BY nombre;")
        usuarios = cursor.fetchall()

    # ===========================================================
    # POST
    # ===========================================================
    if request.method == "POST":
        paso = request.POST.get("paso_actual", "3")

        # -------------------------------------------------------
        # 🟠 PASO 4 → Confirmación final (simulación de inserción)
        # -------------------------------------------------------
        if paso == "4":
            tipo_doc_id = request.POST.get("tipo_documento")
            observaciones = request.POST.get("observaciones", "")
            fecha_primera_revision = request.POST.get("fecha_primera_revision")
            alertar_dias_revision = request.POST.get("alertar_dias_revision")
            fecha_entrega = request.POST.get("fecha_entrega")
            alertar_dias_entrega = request.POST.get("alertar_dias_entrega")
            redactores = request.POST.getlist("redactores")
            revisores = request.POST.getlist("revisores")
            aprobadores = request.POST.getlist("aprobadores")

            try:
                with transaction.atomic():
                    print("\n========== SIMULACIÓN FINAL DE INSERCIÓN ==========")
                    print(f"📄 Proyecto ID: {proyecto_id}")
                    print(f"📘 Tipo Documento Técnico ID: {tipo_doc_id}")
                    print(f"🕒 Fecha de Registro: {timezone.now()}")
                    print(f"🗒️ Observaciones: {observaciones}")
                    print(f"📆 Fechas: Revisión={fecha_primera_revision}, Entrega={fecha_entrega}")
                    print(f"🔔 Alertas: rev={alertar_dias_revision}, entr={alertar_dias_entrega}")

                    # ====================================================
                    # 🔸 Inserción del requerimiento (comentada)
                    # ====================================================
                    # cursor.execute("""
                    #     INSERT INTO requerimiento_documento_tecnico
                    #     (proyecto_id, tipo_documento_id, fecha_registro, observaciones)
                    #     VALUES (%s, %s, %s, %s)
                    #     RETURNING id;
                    # """, [proyecto_id, tipo_doc_id, timezone.now(), observaciones])
                    # req_id = cursor.fetchone()[0]

                    req_id = 9999  # ID simulado para depuración
                    print(f"✅ Se insertaría un nuevo requerimiento con ID simulado: {req_id}")

                    # ====================================================
                    # 🔸 Log inicial (comentado)
                    # ====================================================
                    print({
                        "requerimiento_id": req_id,
                        "usuario_id": request.user.id,
                        "estado_origen_id": None,
                        "estado_destino_id": 1,
                        "created_at": timezone.now(),
                        "observaciones": "El sistema ha habilitado la plantilla en el repositorio."
                    })

                    # cursor.execute("""
                    #     INSERT INTO log_estado_requerimiento_documento
                    #     (requerimiento_id, usuario_id, estado_origen_id, estado_destino_id, created_at, observaciones)
                    #     VALUES (%s, %s, %s, %s, %s, %s);
                    # """, [req_id, request.user.id, None, 1, timezone.now(), "El sistema ha habilitado la plantilla en el repositorio."])

                    # ====================================================
                    # 🔸 Asignación de roles (comentada)
                    # ====================================================
                    def imprimir_roles(lista, rol_id, nombre_rol):
                        for u in lista:
                            print(f"\n👤 Rol: {nombre_rol}")
                            print({
                                "requerimiento_id": req_id,
                                "usuario_id": u,
                                "rol_id": rol_id,
                                "fecha_asignacion": timezone.now(),
                                "activo": True
                            })
                            # cursor.execute("""
                            #     INSERT INTO requerimiento_equipo_rol
                            #     (requerimiento_id, usuario_id, rol_id, fecha_asignacion, activo)
                            #     VALUES (%s, %s, %s, %s, TRUE);
                            # """, [req_id, u, rol_id, timezone.now()])

                    imprimir_roles(redactores, 1, "Redactor")
                    imprimir_roles(revisores, 2, "Revisor")
                    imprimir_roles(aprobadores, 3, "Aprobador")

                    print("\n✅ SIMULACIÓN COMPLETA — no se ha modificado la base de datos.")
                    print("====================================================")

                return render(request, "simulacion_requerimiento.html", {
                    "proyecto": proyecto,
                    "tipo_doc_id": tipo_doc_id,
                    "observaciones": observaciones,
                    "redactores": redactores,
                    "revisores": revisores,
                    "aprobadores": aprobadores,
                    "fecha_primera_revision": fecha_primera_revision,
                    "fecha_entrega": fecha_entrega,
                })

            except Exception as e:
                print(f"❌ Error en simulación: {e}")
                return render(request, "error.html", {"mensaje": f"Error al simular creación: {str(e)}"})

        # -------------------------------------------------------
        # 🟠 PASO 3 → Avanza al paso 4 (confirmación)
        # -------------------------------------------------------
        tipo_doc_id = request.POST.get("tipo_documento")
        observaciones = request.POST.get("observaciones", "")
        fecha_primera_revision = request.POST.get("fecha_primera_revision")
        alertar_dias_revision = request.POST.get("alertar_dias_revision")
        fecha_entrega = request.POST.get("fecha_entrega")
        alertar_dias_entrega = request.POST.get("alertar_dias_entrega")
        redactores = request.POST.getlist("redactores")
        revisores = request.POST.getlist("revisores")
        aprobadores = request.POST.getlist("aprobadores")

        # 🔍 Debug: imprimir todo lo que llega al paso 4
        print("\n========== DATOS RECIBIDOS PARA EL PASO 4 ==========")
        print(f"📄 Proyecto ID: {proyecto_id}")
        print(f"📘 Tipo Documento Técnico ID: {tipo_doc_id}")
        print(f"🗒️ Observaciones: {observaciones}")
        print(f"📆 Fecha primera revisión: {fecha_primera_revision}")
        print(f"🔔 Alertar días antes revisión: {alertar_dias_revision}")
        print(f"📆 Fecha entrega final: {fecha_entrega}")
        print(f"🔔 Alertar días antes entrega: {alertar_dias_entrega}")
        print(f"👤 Redactores: {redactores}")
        print(f"🔍 Revisores: {revisores}")
        print(f"✅ Aprobadores: {aprobadores}")
        print("====================================================\n")

        return render(request, "nuevo_requerimiento_paso4.html", {
            "proyecto": proyecto,
            "tipo_doc_id": tipo_doc_id,
            "observaciones": observaciones,
            "fecha_primera_revision": fecha_primera_revision,
            "alertar_dias_revision": alertar_dias_revision,
            "fecha_entrega": fecha_entrega,
            "alertar_dias_entrega": alertar_dias_entrega,
            "redactores": redactores,
            "revisores": revisores,
            "aprobadores": aprobadores,
            # 🔸 datos para debug en navegador
            "debug_json": {
                "tipo_doc_id": tipo_doc_id,
                "observaciones": observaciones,
                "fecha_primera_revision": fecha_primera_revision,
                "alertar_dias_revision": alertar_dias_revision,
                "fecha_entrega": fecha_entrega,
                "alertar_dias_entrega": alertar_dias_entrega,
                "redactores": redactores,
                "revisores": revisores,
                "aprobadores": aprobadores,
            }
        })

    # ===========================================================
    # GET inicial
    # ===========================================================
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





from django.shortcuts import render, redirect
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import connection


@login_required
def crear_proyecto(request):
    from django.db import connection
    from django.utils import timezone
    from google.cloud import storage  # ✅ Import GCS

    # === Datos temporales en sesión ===
    if "proyecto_temp" not in request.session:
        request.session["proyecto_temp"] = {}
    temp = request.session["proyecto_temp"]

    paso_actual = request.POST.get("paso_actual", "1")
    accion = request.POST.get("accion")
    resumen = {}

    # === Cargar datos base para selects ===
    with connection.cursor() as cursor:
        # 🔹 Administradores (para Paso 1)
        cursor.execute("""
            SELECT id, nombre, email
            FROM usuarios
            WHERE cargo_id = 4
            ORDER BY nombre
        """)
        usuarios_administrador = [
            {"id": r[0], "nombre": r[1], "email": r[2]} for r in cursor.fetchall()
        ]

        # 🔹 Todos los usuarios (para Paso 3)
        cursor.execute("""
            SELECT id, nombre, email
            FROM usuarios
            ORDER BY nombre
        """)
        usuarios_todos = [
            {"id": r[0], "nombre": r[1], "email": r[2]} for r in cursor.fetchall()
        ]

        # 🔹 Máquinas
        cursor.execute("SELECT id, nombre, abreviatura FROM maquinas ORDER BY nombre")
        maquinas = [
            {"id": r[0], "nombre": r[1], "abreviatura": r[2] or ""} for r in cursor.fetchall()
        ]

        # 🔹 Contratos (con cliente)
        cursor.execute("""
            SELECT c.id, c.numero_contrato, c.monto_total, c.fecha_firma,
                   c.representante_cliente_nombre, c.representante_cliente_correo,
                   c.representante_cliente_telefono, cl.id, cl.nombre AS cliente_nombre
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
        } for r in cursor.fetchall()]

        # 🔹 Clientes
        cursor.execute("SELECT id, nombre FROM clientes ORDER BY nombre")
        clientes = [{"id": r[0], "nombre": r[1]} for r in cursor.fetchall()]

        # 🔹 Faenas
        cursor.execute("SELECT id, cliente_id, nombre, ubicacion FROM faenas ORDER BY nombre")
        faenas = [
            {"id": r[0], "cliente_id": r[1], "nombre": r[2], "ubicacion": r[3]}
            for r in cursor.fetchall()
        ]

        # 🔹 Grupos y Documentos Técnicos
        cursor.execute("""
            SELECT id, nombre, descripcion
            FROM categoria_documentos_tecnicos
            ORDER BY nombre
        """)
        grupos_maestros = [
            {"id": r[0], "nombre": r[1], "descripcion": r[2]} for r in cursor.fetchall()
        ]

        cursor.execute("""
            SELECT id, categoria_id, nombre
            FROM tipo_documentos_tecnicos
            ORDER BY nombre
        """)
        documentos = [
            {"id": r[0], "categoria_id": r[1], "nombre": r[2]} for r in cursor.fetchall()
        ]

    # === Relacionar documentos con su grupo ===
    for grupo in grupos_maestros:
        grupo_docs = [d for d in documentos if d["categoria_id"] == grupo["id"]]
        grupo["documentos"] = grupo_docs

    # === Control de pasos ===
    if request.method == "POST":
        if accion == "anterior":
            paso_actual = str(max(1, int(paso_actual) - 1))

        elif accion == "siguiente":
            if paso_actual == "1":
                temp["paso1"] = {
                    "nombre": request.POST.get("nombre"),
                    "descripcion": request.POST.get("descripcion"),
                    "abreviatura": request.POST.get("abreviatura"),
                    "fecha_recepcion_evaluacion": request.POST.get("fecha_recepcion_evaluacion"),
                    "fecha_inicio_planificacion": request.POST.get("fecha_inicio_planificacion"),
                    "fecha_inicio_ejecucion": request.POST.get("fecha_inicio_ejecucion"),
                    "fecha_cierre_proyecto": request.POST.get("fecha_cierre_proyecto"),
                    "administrador": request.POST.get("administrador"),
                    "numero_orden": request.POST.get("numero_orden"),
                    "maquinas_ids": request.POST.getlist("maquinas_ids[]"),
                }

            elif paso_actual == "2":
                temp["paso2"] = {
                    "contrato_id": request.POST.get("contrato_id"),
                    "numero_contrato": request.POST.get("numero_contrato"),
                    "monto_total": request.POST.get("monto_total"),
                    "contrato_fecha_firma": request.POST.get("contrato_fecha_firma"),
                    "representante_cliente_nombre": request.POST.get("representante_cliente_nombre"),
                    "representante_cliente_correo": request.POST.get("representante_cliente_correo"),
                    "representante_cliente_telefono": request.POST.get("representante_cliente_telefono"),
                    "cliente_id": request.POST.get("cliente_id"),
                    "cliente_nombre": request.POST.get("cliente_nombre"),
                    "cliente_abreviatura": request.POST.get("cliente_abreviatura"),
                    "cliente_rut": request.POST.get("cliente_rut"),
                    "cliente_direccion": request.POST.get("cliente_direccion"),
                    "cliente_correo": request.POST.get("cliente_correo"),
                    "cliente_telefono": request.POST.get("cliente_telefono"),
                    "faena_id": request.POST.get("faena_id"),
                    "faena_nombre": request.POST.get("faena_nombre"),
                    "faena_ubicacion": request.POST.get("faena_ubicacion"),
                }

            elif paso_actual == "3":
                temp["paso3"] = {
                    "responsable": request.POST.get("responsable"),
                    "observaciones": request.POST.get("observaciones"),
                    "documentos_ids": request.POST.getlist("documentos_ids[]"),
                }

                documentos_roles = {}
                for doc_id in request.POST.getlist("documentos_ids[]"):
                    documentos_roles[doc_id] = {
                        "redactores": request.POST.getlist(f"redactor_id_{doc_id}[]"),
                        "revisores": request.POST.getlist(f"revisor_id_{doc_id}[]"),
                        "aprobadores": request.POST.getlist(f"aprobador_id_{doc_id}[]"),
                        "observaciones": request.POST.get(f"observaciones_{doc_id}", ""),
                        "hitos": {
                            "fecha_inicio_elaboracion": request.POST.get(f"fecha_inicio_elaboracion_{doc_id}"),
                            "alertar_dias_inicio": request.POST.get(f"alertar_dias_inicio_{doc_id}"),
                            "fecha_primera_revision": request.POST.get(f"fecha_primera_revision_{doc_id}"),
                            "alertar_dias_revision": request.POST.get(f"alertar_dias_revision_{doc_id}"),
                            "fecha_entrega": request.POST.get(f"fecha_entrega_{doc_id}"),
                            "alertar_dias_entrega": request.POST.get(f"alertar_dias_entrega_{doc_id}"),
                        }
                    }
                temp["paso3"]["documentos_roles"] = documentos_roles

            request.session.modified = True
            paso_actual = str(min(4, int(paso_actual) + 1))

        elif accion == "confirmar":
            resumen = {**temp.get("paso1", {}), **temp.get("paso2", {}), **temp.get("paso3", {})}
            for k, v in resumen.items():
                if isinstance(v, (list, tuple)):
                    resumen[k] = ", ".join(map(str, v))

            print("\n========== SIMULACIÓN FINAL ==========")
            for k, v in resumen.items():
                print(f"{k}: {v}")
            print(f"🕒 {timezone.now()} — Proyecto simulado correctamente.")
            print("======================================\n")

            # === Asegurar que cliente_nombre exista antes de crear carpetas ===
            contrato_id = resumen.get("contrato_id")
            if contrato_id and not resumen.get("cliente_nombre"):
                with connection.cursor() as cursor:
                    cursor.execute("""
                        SELECT cl.nombre
                        FROM contratos c
                        JOIN clientes cl ON cl.id = c.cliente_id
                        WHERE c.id = %s
                    """, [contrato_id])
                    row = cursor.fetchone()
                    if row:
                        resumen["cliente_nombre"] = row[0]

            # === Crear estructura en Google Cloud Storage ===
            try:
                storage_client = storage.Client()
                bucket = storage_client.bucket("sgdmtso_jova")

                # 🧹 Limpieza y validación de nombres
                cliente = resumen.get("cliente_nombre", "").strip()
                proyecto = resumen.get("nombre", "").strip()
                if not cliente:
                    cliente = "Cliente_Desconocido"
                if not proyecto:
                    proyecto = "Proyecto_SinNombre"

                for simbolo in [" ", "/", "\\", ":", "*", "?", "\"", "<", ">", "|"]:
                    cliente = cliente.replace(simbolo, "_")
                    proyecto = proyecto.replace(simbolo, "_")

                base_path = f"DocumentosProyectos/{cliente}/{proyecto}/"
                bucket.blob(base_path).upload_from_string("")

                documentos_roles = temp.get("paso3", {}).get("documentos_roles", {})
                if documentos_roles:
                    with connection.cursor() as cursor:
                        for doc_id in documentos_roles.keys():
                            cursor.execute("""
                                SELECT c.nombre AS categoria, t.nombre AS tipo
                                FROM tipo_documentos_tecnicos t
                                JOIN categoria_documentos_tecnicos c ON c.id = t.categoria_id
                                WHERE t.id = %s
                            """, [doc_id])
                            row = cursor.fetchone()
                            if row:
                                categoria = row[0].strip()
                                tipo = row[1].strip()
                                for simbolo in [" ", "/", "\\", ":", "*", "?", "\"", "<", ">", "|"]:
                                    categoria = categoria.replace(simbolo, "_")
                                    tipo = tipo.replace(simbolo, "_")
                                folder_path = f"{base_path}{categoria}/{tipo}/"
                                bucket.blob(folder_path).upload_from_string("")

                print(f"✅ Árbol de carpetas creado correctamente en GCS: {base_path}")

            except Exception as e:
                print("⚠️ Error al crear estructura de carpetas GCS:", e)

            del request.session["proyecto_temp"]
            return redirect("usuario:lista_proyectos")

    # === Paso 4: Confirmación ===
    steps = ["Datos Generales", "Contrato y Cliente", "Responsables y Documentos", "Confirmación"]

    if paso_actual == "4":
        resumen = {**temp.get("paso1", {}), **temp.get("paso2", {}), **temp.get("paso3", {})}

        with connection.cursor() as cursor:
            # 🔹 Administrador
            admin_id = resumen.get("administrador")
            if admin_id:
                cursor.execute("SELECT nombre, email FROM usuarios WHERE id = %s", [admin_id])
                row = cursor.fetchone()
                if row:
                    resumen["administrador"] = f"{row[0]} ({row[1]})"

            # 🔹 Máquinas
            maquinas_ids = resumen.get("maquinas_ids", [])
            if isinstance(maquinas_ids, list) and maquinas_ids:
                ids_lista = list(map(int, maquinas_ids))
                cursor.execute("SELECT nombre FROM maquinas WHERE id = ANY(%s)", [ids_lista])
                resumen["maquinas_ids"] = [r[0] for r in cursor.fetchall()]

            # 🔹 Contrato existente
            contrato_id = resumen.get("contrato_id")
            if contrato_id:
                cursor.execute("""
                    SELECT c.numero_contrato, c.monto_total, c.fecha_firma,
                           c.representante_cliente_nombre, c.representante_cliente_correo,
                           c.representante_cliente_telefono,
                           cl.nombre, cl.abreviatura
                    FROM contratos c
                    JOIN clientes cl ON cl.id = c.cliente_id
                    WHERE c.id = %s
                """, [contrato_id])
                row = cursor.fetchone()
                if row:
                    resumen.update({
                        "contrato_existente": True,
                        "numero_contrato": row[0],
                        "monto_total": row[1],
                        "contrato_fecha_firma": row[2],
                        "representante_cliente_nombre": row[3],
                        "representante_cliente_correo": row[4],
                        "representante_cliente_telefono": row[5],
                        "cliente_nombre": row[6],
                        "cliente_abreviatura": row[7],
                    })

            # 🔹 Faena existente
            if resumen.get("faena_id"):
                cursor.execute("""
                    SELECT f.nombre AS faena_nombre, f.ubicacion AS faena_ubicacion, c.nombre AS cliente_nombre
                    FROM faenas f
                    LEFT JOIN clientes c ON c.id = f.cliente_id
                    WHERE f.id = %s
                """, [resumen["faena_id"]])
                row = cursor.fetchone()
                if row:
                    resumen.update({
                        "faena_existente": True,
                        "faena_nombre": row[0],
                        "faena_ubicacion": row[1],
                        "faena_cliente_nombre": row[2],
                    })

            # 🔹 Documentos Técnicos
            if "documentos_roles" in resumen:
                doc_roles = resumen["documentos_roles"]
                for doc_id, roles in doc_roles.items():
                    cursor.execute("SELECT nombre FROM tipo_documentos_tecnicos WHERE id = %s", [doc_id])
                    doc_row = cursor.fetchone()
                    nombre_doc = doc_row[0] if doc_row else f"Documento #{doc_id}"

                    def ids_a_nombres(ids):
                        if not ids:
                            return []
                        placeholders = ','.join(['%s'] * len(ids))
                        cursor.execute(f"SELECT nombre FROM usuarios WHERE id IN ({placeholders})", ids)
                        return [r[0] for r in cursor.fetchall()]

                    doc_roles[doc_id] = {
                        "documento_nombre": nombre_doc,
                        "redactores": ids_a_nombres(roles.get("redactores", [])),
                        "revisores": ids_a_nombres(roles.get("revisores", [])),
                        "aprobadores": ids_a_nombres(roles.get("aprobadores", [])),
                        "observaciones": roles.get("observaciones", ""),
                        "hitos": roles.get("hitos", {}),
                    }

                resumen["documentos_roles"] = doc_roles

    # === Render final ===
    return render(request, "crear_proyecto.html", {
        "paso_actual": paso_actual,
        "proyecto_temp": temp.get("paso1", {}),
        "steps": steps,
        "resumen": resumen,
        "usuarios_administrador": usuarios_administrador,
        "usuarios_todos": usuarios_todos,
        "maquinas": maquinas,
        "contratos": contratos,
        "clientes": clientes,
        "faenas": faenas,
        "grupos_maestros": grupos_maestros,
        "documentos": documentos,
    })
















####################################Crear Proyecto - Paso 2 ##################################

@login_required
def obtener_datos_contrato(request, contrato_id):
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT c.id, c.numero_contrato, c.monto_total, c.fecha_firma,
                   c.representante_cliente_nombre, c.representante_cliente_correo,
                   c.representante_cliente_telefono,
                   cl.id AS cliente_id, cl.nombre AS cliente_nombre,
                   cl.abreviatura, cl.rut, cl.direccion, cl.correo_contacto, cl.telefono_contacto
            FROM contratos c
            JOIN clientes cl ON cl.id = c.cliente_id
            WHERE c.id = %s
        """, [contrato_id])
        r = cursor.fetchone()
    if not r:
        return JsonResponse({}, status=404)
    return JsonResponse({
        "contrato_id": r[0],
        "numero_contrato": r[1],
        "monto_total": r[2],
        "fecha_firma": r[3],
        "representante_cliente_nombre": r[4],
        "representante_cliente_correo": r[5],
        "representante_cliente_telefono": r[6],
        "cliente_id": r[7],
        "cliente_nombre": r[8],
        "cliente_abreviatura": r[9],
        "cliente_rut": r[10],
        "cliente_direccion": r[11],
        "cliente_correo": r[12],
        "cliente_telefono": r[13],
    })

@login_required
def obtener_datos_cliente(request, cliente_id):
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id, nombre, abreviatura, rut, direccion, correo_contacto, telefono_contacto
            FROM clientes WHERE id = %s
        """, [cliente_id])
        r = cursor.fetchone()
    if not r:
        return JsonResponse({}, status=404)
    return JsonResponse({
        "cliente_id": r[0],
        "nombre": r[1],
        "abreviatura": r[2],
        "rut": r[3],
        "direccion": r[4],
        "correo": r[5],
        "telefono": r[6],
    })

@login_required
def obtener_datos_faena(request, faena_id):
    with connection.cursor() as cursor:
        cursor.execute("SELECT id, nombre, ubicacion FROM faenas WHERE id = %s", [faena_id])
        r = cursor.fetchone()
    if not r:
        return JsonResponse({}, status=404)
    return JsonResponse({
        "faena_id": r[0],
        "nombre": r[1],
        "ubicacion": r[2],
    })





######################################################################




@csrf_exempt
@login_required
def generar_abreviatura_proyecto(request):
    """
    Genera la abreviatura de un proyecto a partir de:
    - nombre de la máquina
    - descripción
    - fecha de recepción / evaluación
    """
    if request.method != "POST":
        return JsonResponse({"error": "Método no permitido"}, status=405)

    try:
        import json, datetime, re
        data = json.loads(request.body)

        maquina = data.get("maquina", "").strip()
        descripcion = data.get("descripcion", "").strip().upper()
        fecha = data.get("fecha_recepcion_evaluacion", "").strip()

        print("📩 Datos recibidos:", data)  # DEBUG

        if not maquina or not fecha:
            return JsonResponse({"abreviatura": ""})

        # Limpiar nombre de máquina (sin abreviatura entre paréntesis)
        maquina = re.sub(r"\(.*?\)", "", maquina).strip().upper().replace(" ", "")

        # Limpiar descripción (usar primeras 2 palabras)
        palabras = [p for p in descripcion.split() if len(p) > 2]
        descripcion_limpia = "".join(palabras[:2]).upper()

        # Formato fecha MMYY
        try:
            fecha_obj = datetime.datetime.strptime(fecha, "%Y-%m-%d")
            mes = f"{fecha_obj.month:02}"
            anio = str(fecha_obj.year)[-2:]
            fecha_formato = f"{mes}{anio}"
        except ValueError:
            fecha_formato = ""

        abreviatura = f"{maquina}.{descripcion_limpia}.{fecha_formato}"
        print("✅ Abreviatura generada:", abreviatura)

        return JsonResponse({"abreviatura": abreviatura})

    except Exception as e:
        print("❌ Error generar_abreviatura_proyecto:", e)
        return JsonResponse({"error": str(e)}, status=400)
#paso2 ####################  ####################    ####################    ####################  


@csrf_exempt
@login_required
def generar_abreviatura_cliente(request):
    """
    Genera la abreviatura de un cliente (vía AJAX) a partir de su nombre.
    Equivalente al modelo del proyecto: una sola función con request.
    """
    if request.method != "GET":
        return JsonResponse({"error": "Método no permitido"}, status=405)

    try:
        nombre = request.GET.get("nombre", "").strip()
        if not nombre:
            return JsonResponse({"abreviatura": ""})

        # --- Normalización y limpieza ---
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

        # --- Generación base ---
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

        # --- Validación de unicidad ---
        with connection.cursor() as cursor:
            while True:
                cursor.execute("SELECT COUNT(*) FROM clientes WHERE abreviatura = %s;", [abrev_final])
                if cursor.fetchone()[0] == 0:
                    break
                abrev_final = f"{base}{contador}"
                contador += 1

        return JsonResponse({"abreviatura": abrev_final})

    except Exception as e:
        print("❌ Error generar_abreviatura_cliente:", e)
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
