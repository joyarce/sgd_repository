from django.shortcuts import render
from django.db import connection
from unidecode import unidecode
import re
from django.contrib import messages
from openpyxl import load_workbook
from django.http import JsonResponse
import tempfile, os
from openpyxl.worksheet.table import Table
from django.views.decorators.csrf import csrf_exempt
from openpyxl.utils import range_boundaries

# 🔹 Genera abreviatura/nomenclatura según las reglas del documento
def generar_nomenclatura(nombre):
    nombre_normalizado = unidecode(nombre.strip().upper())
    palabras = [p for p in nombre_normalizado.split() if p not in ['Y', 'DE', 'DEL', 'LA', 'LOS']]

    # Crear sigla
    if len(palabras) == 1:
        sigla = palabras[0][:3]
    else:
        sigla = ''.join(p[0] for p in palabras)

    # Verificar duplicados en BD
    with connection.cursor() as cursor:
        base = sigla
        contador = 1
        cursor.execute("SELECT abreviatura FROM categoria_documentos_tecnicos WHERE abreviatura = %s", [sigla])
        while cursor.fetchone() is not None:
            sigla = f"{base}{contador}"
            cursor.execute("SELECT abreviatura FROM categoria_documentos_tecnicos WHERE abreviatura = %s", [sigla])
            contador += 1

    return sigla

# 🔹 Vista principal: lista categorías y tipos asociados (con abreviaturas)
def lista_plantillas(request):
    # 🔸 Comentado el bloque de inserción (solo lectura)
    # if request.method == "POST":
    #     nombre = request.POST.get("nombre")
    #     descripcion = request.POST.get("descripcion")
    #     if nombre:
    #         abreviatura = generar_nomenclatura(nombre)
    #         with connection.cursor() as cursor:
    #             cursor.execute("""
    #                 INSERT INTO categoria_documentos_tecnicos (nombre, descripcion, abreviatura)
    #                 VALUES (%s, %s, %s)
    #             """, [nombre, descripcion, abreviatura])
    #     return redirect('plantillas:lista_plantillas')

    categorias = []
    with connection.cursor() as cursor:
        # Obtener categorías con abreviaturas
        cursor.execute("""
            SELECT id, nombre, descripcion, abreviatura
            FROM categoria_documentos_tecnicos
            ORDER BY id ASC
        """)
        categorias_data = cursor.fetchall()

        # Por cada categoría, obtener sus tipos asociados (también con abreviatura)
        for cat_id, cat_nombre, cat_desc, cat_abrev in categorias_data:
            cursor.execute("""
                SELECT id, nombre, abreviatura
                FROM tipo_documentos_tecnicos
                WHERE categoria_id = %s
                ORDER BY id ASC
            """, [cat_id])
            tipos = cursor.fetchall()

            categorias.append({
                "id": cat_id,
                "nombre": cat_nombre,
                "abreviatura": cat_abrev,
                "descripcion": cat_desc,
                "tipos": [{"id": t[0], "nombre": t[1], "abreviatura": t[2]} for t in tipos]
            })

    return render(request, "lista_plantillas.html", {"categorias": categorias})

def categoria_detalle(request, categoria_id):
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id, nombre, descripcion, abreviatura
            FROM categoria_documentos_tecnicos
            WHERE id = %s
        """, [categoria_id])
        categoria = dictfetchone(cursor)

        cursor.execute("""
            SELECT id, nombre, descripcion, abreviatura
            FROM tipo_documentos_tecnicos
            WHERE categoria_id = %s
            ORDER BY nombre;
        """, [categoria_id])
        tipos = dictfetchall(cursor)

    return render(request, 'categoria_detalle.html', {
        'categoria': categoria,
        'tipos': tipos,
    })

def dictfetchall(cursor):
    columns = [col[0] for col in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]

def dictfetchone(cursor):
    row = cursor.fetchone()
    if row is None:
        return None
    columns = [col[0] for col in cursor.description]
    return dict(zip(columns, row))

def tipo_detalle(request, tipo_id):
    """
    Muestra la información detallada de un tipo de documento técnico,
    incluyendo etiquetas, tablas y columnas asociadas.
    """
    with connection.cursor() as cursor:
        # Datos principales del tipo de documento
        cursor.execute("""
            SELECT 
                t.id,
                t.nombre,
                t.descripcion,
                t.creado_en,
                t.abreviatura,
                c.id AS categoria_id,
                c.nombre AS categoria_nombre,
                c.descripcion AS categoria_descripcion
            FROM tipo_documentos_tecnicos t
            JOIN categoria_documentos_tecnicos c ON t.categoria_id = c.id
            WHERE t.id = %s
        """, [tipo_id])
        row = cursor.fetchone()

        if not row:
            return render(request, "404.html", status=404)

        tipo = {
            "id": row[0],
            "nombre": row[1],
            "descripcion": row[2],
            "creado_en": row[3],
            "abreviatura": row[4],
            "categoria_id": row[5],
            "categoria_nombre": row[6],
            "categoria_descripcion": row[7],
        }

        # Etiquetas (Metadatos de Campos)
        cursor.execute("""
            SELECT nombre FROM metadato_documento_campo
            WHERE documento_id = %s
            ORDER BY nombre
        """, [tipo_id])
        etiquetas = [r[0] for r in cursor.fetchall()]

        # Tablas asociadas
        cursor.execute("""
            SELECT id, nombre FROM metadato_documento_tabla
            WHERE documento_id = %s
            ORDER BY nombre
        """, [tipo_id])
        tablas = [{"id": r[0], "nombre": r[1], "columnas": []} for r in cursor.fetchall()]

        # Columnas por tabla
        for tabla in tablas:
            cursor.execute("""
                SELECT nombre FROM metadato_documento_columna
                WHERE tabla_id = %s
                ORDER BY nombre
            """, [tabla["id"]])
            tabla["columnas"] = [c[0] for c in cursor.fetchall()]

    return render(request, "tipo_detalle.html", {
        "tipo": tipo,
        "etiquetas": etiquetas,
        "tablas": tablas
    })




PALABRAS_IGNORAR = {
    "y", "de", "del", "la", "los", "las", "el", "en", "por",
    "para", "ltlda", "ltda", "sa", "empresa", "asociacion",
    "compania", "hermanos"
}

def generar_abreviatura(nombre, tipo="categoria"):
    """
    Genera abreviatura estandarizada.
    tipo="categoria" → abreviatura 3 letras si 1 palabra
    tipo="tipo_documento" → abreviatura 4 letras si 1 palabra
    Para más de una palabra → iniciales hasta 4 palabras.
    Ignora palabras comunes según PALABRAS_IGNORAR.
    """
    if not nombre:
        return ""

    limpio = re.sub(r"[^A-Za-zÁÉÍÓÚáéíóúÑñ\s]", "", nombre)
    palabras = [p for p in limpio.split() if p.lower() not in PALABRAS_IGNORAR]

    if not palabras:
        return ""

    if len(palabras) == 1:
        return palabras[0][:3].upper() if tipo == "categoria" else palabras[0][:4].upper()

    # Más de una palabra → tomar iniciales de hasta 4
    return "".join(p[0].upper() for p in palabras[:4])

def crear_categoria(request):
    abreviatura_generada = ""

    if request.method == "POST":
        nombre = request.POST.get("nombre", "").strip()
        descripcion = request.POST.get("descripcion", "").strip()

        # Generar abreviatura
        abreviatura = generar_abreviatura(nombre, tipo="categoria")
        abreviatura_generada = abreviatura

        if not abreviatura:
            messages.error(request, "El nombre ingresado no genera una abreviatura válida.")
        else:
            # Validar unicidad
            with connection.cursor() as cursor:
                cursor.execute("""
                    SELECT COUNT(*) 
                    FROM categoria_documentos_tecnicos
                    WHERE UPPER(abreviatura) = UPPER(%s)
                """, [abreviatura])
                existe = cursor.fetchone()[0]

            if existe > 0:
                messages.warning(request, f"La abreviatura '{abreviatura}' ya existe. Agrega un sufijo distintivo.")
            else:
                messages.success(request, f"✅ Abreviatura generada automáticamente: {abreviatura}")
                # Mostrar en consola lo que se guardaría
                print(f"[INFO] Valores a guardar: {{'nombre': '{nombre}', 'descripcion': '{descripcion}', 'abreviatura': '{abreviatura}'}}")

    # Este return se ejecuta tanto en GET como POST
    return render(request, "crear_categoria_documento.html", {
        "abreviatura": abreviatura_generada
    })

def crear_tipo_documento(request):
    with connection.cursor() as cursor:
        cursor.execute("""
            SELECT id, nombre, abreviatura 
            FROM categoria_documentos_tecnicos
            ORDER BY nombre
        """)
        categorias = [
            {"id": r[0], "nombre": r[1], "abreviatura": r[2]}
            for r in cursor.fetchall()
        ]

    if request.method == "POST":
        nombre = request.POST.get("nombre", "").strip()
        descripcion = request.POST.get("descripcion", "").strip()
        categoria_id = request.POST.get("categoria_id")
        abreviatura_manual = request.POST.get("abreviatura", "").strip().upper()

        if not categoria_id:
            messages.error(request, "Debes seleccionar una categoría. ⚠️")
        elif not nombre:
            messages.error(request, "El nombre del tipo de documento es obligatorio. ⚠️")
        else:
            abreviatura_generada = generar_abreviatura(nombre, tipo="tipo_documento")
            abreviatura_final = abreviatura_manual or abreviatura_generada

            # Validar duplicado dentro de la categoría
            with connection.cursor() as cursor:
                cursor.execute("""
                    SELECT COUNT(*) 
                    FROM tipo_documentos_tecnicos
                    WHERE categoria_id = %s AND UPPER(abreviatura) = UPPER(%s)
                """, [categoria_id, abreviatura_final])
                existe = cursor.fetchone()[0]

            if existe > 0:
                messages.error(
                    request,
                    f"La abreviatura '{abreviatura_final}' ya existe en esta categoría. ❌ "
                    "Agrega un sufijo distintivo (por ejemplo: INF-PROY)."
                )
            else:
                messages.success(
                    request,
                    f"✅ Abreviatura generada automáticamente: '{abreviatura_final}'"
                )

                # --- Imprimir en consola lo que se guardaría ---
                print(f"[INFO] Valores a guardar en tipo_documentos_tecnicos: {{'categoria_id': {categoria_id}, 'nombre': '{nombre}', 'descripcion': '{descripcion}', 'abreviatura': '{abreviatura_final}'}}")
                # No se ejecuta insert real

    return render(request, "crear_tipo_documento.html", {"categorias": categorias})




def obtener_etiquetas_excel(path_archivo, etiquetas):
    """
    Verifica si las etiquetas/metadatos existen en el archivo Excel.
    Devuelve dict {etiqueta: True/False}
    """
    wb = load_workbook(path_archivo, data_only=True)
    resultados = {tag: False for tag in etiquetas}

    # Recorre todas las hojas y celdas
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=True):
            for cell in row:
                if cell is None:
                    continue
                cell_str = str(cell).strip()
                for tag in etiquetas:
                    if tag.lower() in cell_str.lower():
                        resultados[tag] = True

    return resultados




def validar_etiquetas_archivo(request):
    if request.method == "POST" and request.FILES.get("archivo"):
        archivo = request.FILES["archivo"]
        etiquetas = request.POST.getlist("tags[]")  # recibir lista de etiquetas
        import tempfile, os

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            for chunk in archivo.chunks():
                tmp.write(chunk)
            tmp_path = tmp.name

        resultados = obtener_etiquetas_excel(tmp_path, etiquetas)
        os.remove(tmp_path)

        return JsonResponse({"resultados": resultados})
    return JsonResponse({"error": "Archivo no recibido"}, status=400)




def obtener_columnas_archivo(request):
    """
    Devuelve solo las etiquetas detectadas en un archivo Excel:
    - nombres definidos (defined_names)
    - columnas de tablas
    - encabezados de la primera fila como fallback
    """
    if request.method == "POST" and request.FILES.get("archivo"):
        archivo = request.FILES["archivo"]

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            for chunk in archivo.chunks():
                tmp.write(chunk)
            tmp_path = tmp.name

        try:
            wb = load_workbook(tmp_path, data_only=True)
            etiquetas = set()

            # Nombres definidos
            for name, defn in wb.defined_names.items():
                try:
                    for sheet_title, coord in defn.destinations:
                        etiquetas.add(name)
                except Exception:
                    continue

            # Columnas de tablas
            for sheet in wb.worksheets:
                for table_name, table_obj in getattr(sheet, 'tables', {}).items():
                    if isinstance(table_obj, Table):
                        for col in getattr(table_obj, 'tableColumns', []):
                            if getattr(col, 'name', None):
                                etiquetas.add(col.name.strip())

            # Encabezados primera fila como fallback
            for sheet in wb.worksheets:
                if sheet.max_row >= 1:
                    for cell in sheet[1]:
                        if cell.value:
                            etiquetas.add(str(cell.value).strip())

            os.remove(tmp_path)
            return JsonResponse({"columnas": sorted(list(etiquetas))})

        except Exception as e:
            os.remove(tmp_path)
            return JsonResponse({"error": f"Error al leer el archivo: {str(e)}"}, status=400)

    return JsonResponse({"error": "Archivo no recibido"}, status=400)


@csrf_exempt
def obtener_tablas_y_columnas(request):
    if request.method == "POST" and request.FILES.get("archivo"):
        archivo = request.FILES["archivo"]

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            for chunk in archivo.chunks():
                tmp.write(chunk)
            tmp_path = tmp.name

        try:
            wb = load_workbook(tmp_path, data_only=True)
            tablas = []

            print("Hojas detectadas:", [s.title for s in wb.worksheets])

            for sheet in wb.worksheets:
                print(f"--- Hoja: {sheet.title} ---")
                print("Tablas detectadas:", list(sheet.tables.keys()))

                for table_name in sheet.tables.keys():
                    table_obj = sheet._tables[table_name]
                    ref = getattr(table_obj, 'ref', None)

                    columnas, filas = [], []

                    if ref:
                        min_col, min_row, max_col, max_row = range_boundaries(ref)

                        # Cabeceras (primera fila del rango)
                        columnas = [
                            str(sheet.cell(row=min_row, column=c).value).strip()
                            for c in range(min_col, max_col + 1)
                            if sheet.cell(row=min_row, column=c).value
                        ]

                        # Filas de datos
                        for r in range(min_row + 1, max_row + 1):
                            fila = [
                                sheet.cell(row=r, column=c).value
                                for c in range(min_col, max_col + 1)
                            ]
                            # opcional: ignorar filas completamente vacías
                            if any(fila):
                                filas.append(fila)

                    tablas.append({
                        "tabla": table_name,
                        "columnas": columnas,
                        "registros": filas
                    })

            os.remove(tmp_path)
            print("JSON de tablas a enviar:", tablas)
            return JsonResponse({"tablas": tablas}, json_dumps_params={"ensure_ascii": False, "indent": 2})

        except Exception as e:
            os.remove(tmp_path)
            return JsonResponse({"error": f"Error al leer el archivo: {str(e)}"}, status=400)

    return JsonResponse({"error": "Archivo no recibido"}, status=400)


@csrf_exempt  # para pruebas rápidas, evitar problemas con CSRF
def test_tablas(request):
    if request.method == "POST" and request.FILES.get("archivo"):
        return obtener_tablas_y_columnas(request)
    return render(request, "test_tablas.html")
