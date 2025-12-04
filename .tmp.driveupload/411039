# ============================================================
# llenar_listado_herramientas.py — VERSIÓN FINAL PROFESIONAL
# Compatible con leer_universal_excel.py
# - Copia estilos desde JSON (cell_styles) con fallback seguro
# - Inserta filas con formato correcto sin generar filas vacías
# - Llena tablas auxiliares
# - Replica validaciones
# - Genera SKU, correlativos, descripciones profesionales
# ============================================================

import json
import random
import re
import os
import shutil
from copy import copy
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries, get_column_letter, column_index_from_string
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
CAT_TIPO_RECURSO = [
    {"CÓD.": "INSUMO",       "DESCRIPCIÓN": "Lubricantes, químicos, abrasivos"},
    {"CÓD.": "EPP",          "DESCRIPCIÓN": "Cascos, guantes, lentes, respiradores"},
    {"CÓD.": "HERRAMIENTA",  "DESCRIPCIÓN": "Llaves, alicates, esmeriladoras, pistolas"},
    {"CÓD.": "OFICINA",      "DESCRIPCIÓN": "Papelería, impresoras, muebles"},
    {"CÓD.": "IZAJE",        "DESCRIPCIÓN": "Eslingas, grilletes, tecles, grúas"},
    {"CÓD.": "REPUESTO",     "DESCRIPCIÓN": "Rodamientos, sellos, bujes, cadenas"},
    {"CÓD.": "EQUIPO",       "DESCRIPCIÓN": "Generadores, compresores, vehículos menores"},
    {"CÓD.": "INFRAESTRUCTURA","DESCRIPCIÓN": "Lockers, contenedores, gabinetes, luminarias"},
    {"CÓD.": "TECNOLOGÍA",   "DESCRIPCIÓN": "Notebooks, switches, radios, teléfonos"},
    {"CÓD.": "SEGURIDAD",    "DESCRIPCIÓN": "Señalética, conos, cintas demarcación"},
]

CAT_FAMILIA = [
    {"CÓD.": "FIJACIONES Y SUJECIÓN"},
    {"CÓD.": "EQUIPOS DE IZAJE"},
    {"CÓD.": "INSTRUMENTOS DE MEDICIÓN"},
    {"CÓD.": "LUBRICANTES Y FLUIDOS"},
    {"CÓD.": "HERRAMIENTAS MANUALES"},
    {"CÓD.": "HERRAMIENTAS ELÉCTRICAS / INDUSTRIALES"},
    {"CÓD.": "HIDRÁULICA Y NEUMÁTICA"},
    {"CÓD.": "TRANSMISIONES MECÁNICAS"},
    {"CÓD.": "MATERIALES MECÁNICOS"},
    {"CÓD.": "MATERIALES ELÉCTRICOS"},
    {"CÓD.": "SEGURIDAD / EPIs"},
    {"CÓD.": "LIMPIEZA INDUSTRIAL / QUÍMICOS"},
    {"CÓD.": "SOLDADURA / CORTE / ABRASIVOS"},
    {"CÓD.": "ALMACENAMIENTO / EMBALAJE"},
    {"CÓD.": "COMPUTACIÓN / ELECTRÓNICA / REDES"},
    {"CÓD.": "SUMINISTROS GENERALES / OFICINA"},
    {"CÓD.": "ILUMINACIÓN"},
    {"CÓD.": "REPUESTOS INDUSTRIALES (desgaste, OEM)"},
    {"CÓD.": "VEHÍCULOS MENORES / ACCESORIOS"},
]

CAT_SUB_FAMILIA = [
    {"FAMILIA": "FIJACIONES Y SUJECIÓN", "CÓD.": "FIJACIONES"},
    {"FAMILIA": "SEGURIDAD / EPIs", "CÓD.": "PROTECCION CONTRA CAIDAS"},
    {"FAMILIA": "INSTRUMENTOS DE MEDICIÓN", "CÓD.": "INSTRUMENTOS DE MEDICION"},
    {"FAMILIA": "LUBRICANTES Y FLUIDOS", "CÓD.": "LUBRICANTES Y OTROS ESPECIFICOS"},
    {"FAMILIA": "HERRAMIENTAS MANUALES", "CÓD.": "HERRAMIENTAS MANUALES"},
    {"FAMILIA": "MATERIALES MECÁNICOS", "CÓD.": "MATERIALES ESTRUCTURALES / FORMAS BÁSICAS"},
    {"FAMILIA": "LIMPIEZA INDUSTRIAL / QUÍMICOS", "CÓD.": "AGENTES QUIMICOS"},
    {"FAMILIA": "TRANSMISIONES MECÁNICAS", "CÓD.": "TRANSMISIONES MECANICAS"},
    {"FAMILIA": "HIDRÁULICA Y NEUMÁTICA", "CÓD.": "MANGUERAS"},
    {"FAMILIA": "SEGURIDAD / EPIs", "CÓD.": "PROTECCION PARA LA CABEZA"},
    {"FAMILIA": "MATERIALES ELÉCTRICOS", "CÓD.": "CONDUCTORES / CABLES"},
    {"FAMILIA": "MATERIALES ELÉCTRICOS", "CÓD.": "PROTECCIONES / TERMOMAGNÉTICOS"},
    {"FAMILIA": "HERRAMIENTAS ELÉCTRICAS / INDUSTRIALES", "CÓD.": "ROTOMARTILLOS / ESMERILES / TALADROS"},
    {"FAMILIA": "SOLDADURA / CORTE / ABRASIVOS", "CÓD.": "DISCOS / ABRASIVOS"},
    {"FAMILIA": "SOLDADURA / CORTE / ABRASIVOS", "CÓD.": "ACCESORIOS SOLDADURA"},
    {"FAMILIA": "ILUMINACIÓN", "CÓD.": "LUMINARIAS PORTÁTILES"},
    {"FAMILIA": "HIDRÁULICA Y NEUMÁTICA", "CÓD.": "ACOPLES / CONECTORES"},
    {"FAMILIA": "COMPUTACIÓN / ELECTRÓNICA / REDES", "CÓD.": "REDES / SWITCH / CABLEADO"},
    {"FAMILIA": "EQUIPOS DE IZAJE", "CÓD.": "ESLINGAS / GRILLETES"},
    {"FAMILIA": "EQUIPOS DE IZAJE", "CÓD.": "POLIPASTOS / TEFLONES"},
    {"FAMILIA": "SUMINISTROS GENERALES / OFICINA", "CÓD.": "ÚTILES DE ESCRITORIO"},
    {"FAMILIA": "ALMACENAMIENTO / EMBALAJE", "CÓD.": "CONTENEDORES / CAJAS"},
    {"FAMILIA": "VEHÍCULOS MENORES / ACCESORIOS", "CÓD.": "ACCESORIOS VEHICULARES"},
]

CAT_AMBITO = [
    {"CÓD.": "PERSONA"},
    {"CÓD.": "EQUIPO MAQUINARIA"},
    {"CÓD.": "INSTALACION FAENA"},
    {"CÓD.": "PROCESO PLANTA"},
    {"CÓD.": "AMBIENTE"},
    {"CÓD.": "TRANSPORTE_LOGISTICA"},
    {"CÓD.": "ALMACENAMIENTO"},
    {"CÓD.": "ADMINISTRATIVO OFICINA"},
    {"CÓD.": "TI COMUNICACIONES"},
    {"CÓD.": "LABORATORIO"},
    {"CÓD.": "SEGURIDAD SALUD"},
    {"CÓD.": "CLIENTE"},
]

CAT_UM = [
    {"CÓD.": "UN"}, {"CÓD.": "PZA"}, {"CÓD.": "PAR"}, {"CÓD.": "JGO"},
    {"CÓD.": "KIT"}, {"CÓD.": "CJ"}, {"CÓD.": "BOLSA"}, {"CÓD.": "RLL"},
    {"CÓD.": "TUBO"}, {"CÓD.": "TARRO"}, {"CÓD.": "LAT"}, {"CÓD.": "BDG"},
    {"CÓD.": "MM"}, {"CÓD.": "CM"}, {"CÓD.": "M"}, {"CÓD.": "KM"},
    {"CÓD.": "M2"}, {"CÓD.": "CM2"}, {"CÓD.": "HA"}, {"CÓD.": "ML"},
    {"CÓD.": "L"}, {"CÓD.": "M3"}, {"CÓD.": "G"}, {"CÓD.": "KG"},
    {"CÓD.": "TON"}, {"CÓD.": "LB"}, {"CÓD.": "H"}, {"CÓD.": "HH"},
    {"CÓD.": "DIA"}, {"CÓD.": "SEM"}, {"CÓD.": "MES"}, {"CÓD.": "PAQ"},
    {"CÓD.": "PLG"}, {"CÓD.": "BAR"}, {"CÓD.": "SET"},
]


# ============================================================
#  FUNCIONES DE SANITIZACIÓN DE ESTILOS
# ============================================================

def sanitize_color(value):
    """
    Convierte el valor de color del JSON en un color openpyxl válido.
    Si no es ARGB, devuelve None.
    """
    if not value:
        return None

    # Si es un objeto openpyxl serializado: lo ignoramos
    if "Color object" in str(value):
        return None

    # Si es un hex ARGB válido
    if isinstance(value, str) and re.fullmatch(r"[A-Fa-f0-9]{8}", value):
        return value.upper()

    # Cualquier otra cosa → ignorar
    return None


def apply_style_from_json(cell, style):
    """Aplica a una celda el estilo capturado en el JSON, de forma segura."""

    # --- COLOR SEGURO ---
    safe_font_color = sanitize_color(style.get("font_color"))
    safe_fill_color = sanitize_color(style.get("fill_color"))

    # Fuente
    cell.font = Font(
        name=style.get("font_name"),
        size=style.get("font_size"),
        bold=style.get("font_bold"),
        italic=style.get("font_italic"),
        color=safe_font_color,
    )

    # Alineación
    cell.alignment = Alignment(
        horizontal=style.get("alignment_horizontal"),
        vertical=style.get("alignment_vertical"),
        wrap_text=style.get("alignment_wrap_text"),
    )

    # Bordes
    def _side(v): return Side(style=v) if v else Side(style=None)

    cell.border = Border(
        left=_side(style.get("border_left")),
        right=_side(style.get("border_right")),
        top=_side(style.get("border_top")),
        bottom=_side(style.get("border_bottom")),
    )

    # Relleno seguro
    if safe_fill_color:
        cell.fill = PatternFill(patternType="solid", fgColor=safe_fill_color)

    # Número
    if style.get("number_format"):
        cell.number_format = style.get("number_format")


def get_json_header_styles(json_table):
    """Devuelve los estilos por columna desde JSON."""
    styles = {}
    for coord, style in json_table["cell_styles"].items():
        col_letter = re.match(r"[A-Z]+", coord).group()
        col_index = column_index_from_string(col_letter)
        styles[col_index] = style
    return styles


# ============================================================
#  UTILIDADES DE TABLAS
# ============================================================

def get_table(ws, name):
    if name not in ws.tables:
        raise ValueError(f"❌ Tabla '{name}' no encontrada en hoja '{ws.title}'")
    return ws.tables[name]


def normalize_table(ws, table):
    """Corrige el rango de tabla según datos reales."""
    min_c, min_r, max_c, max_r = range_boundaries(table.ref)
    first_data = min_r + 1
    last = first_data

    for r in range(first_data, max_r + 50):
        row_empty = True
        for c in range(min_c, max_c + 1):
            if ws.cell(r, c).value not in (None, "", " "):
                row_empty = False
                break
        if not row_empty:
            last = r

    if last < first_data:
        last = first_data

    table.ref = f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c)}{last}"
    return min_c, min_r, max_c, last


def insert_rows_json_style(ws, table, json_table, num_rows):
    """Inserta filas con estilos desde el JSON."""
    min_c, min_r, max_c, max_r = normalize_table(ws, table)

    insert_at = max_r + 1
    ws.insert_rows(insert_at, num_rows)

    header_styles = get_json_header_styles(json_table)

    for i in range(num_rows):
        row = insert_at + i
        for col in range(min_c, max_c + 1):
            cell = ws.cell(row=row, column=col)
            cell.value = None

            # Aplicar estilo seguro
            if col in header_styles:
                apply_style_from_json(cell, header_styles[col])

    table.ref = f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c)}{max_r + num_rows}"


def replicate_validations(ws, table, json_table):
    """Replica validaciones de datos."""
    min_c, min_r, max_c, max_r = range_boundaries(table.ref)
    cols = json_table["columnas"]

    for rule in json_table["validaciones"]:
        col_name = rule["columnas_afectadas"][0]
        col_i = cols.index(col_name)
        col_letter = get_column_letter(min_c + col_i)
        new_ref = f"{col_letter}{min_r+1}:{col_letter}{max_r}"

        for dv in ws.data_validations.dataValidation:
            if dv.sqref == rule["rango_validacion"]:
                new_dv = copy(dv)
                new_dv.sqref = new_ref
                ws.data_validations.append(new_dv)
                break


# ============================================================
#  SKU / CORRELATIVOS
# ============================================================

def fill_correlative_and_sku(ws, table, json_table):
    min_c, min_r, max_c, max_r = range_boundaries(table.ref)
    cols = json_table["columnas"]

    cN = min_c + cols.index("N°")
    cSKU = min_c + cols.index("SKU")

    n = 1
    for r in range(min_r + 1, max_r + 1):
        ws.cell(r, cN).value = n
        ws.cell(r, cSKU).value = f"SKU-{n:06d}"
        n += 1


# ============================================================
#  DESCRIPCIÓN PROFESIONAL DE MATERIAL
# ============================================================

def generar_descripcion_material(tipo, familia, subfamilia, ambito, qty, um):
    frases = [
        "Artículo utilizado principalmente en operaciones de {ambito}, perteneciente a la familia {familia}.",
        "Recurso clasificado como {tipo}, correspondiente a la sub-familia {subfamilia}.",
        "Material requerido para labores asociadas a {ambito}.",
        "Elemento habitual dentro del grupo {familia}, sub-familia {subfamilia}.",
    ]
    extra = [
        "Diseñado para uso intensivo en faenas.",
        "Adecuado para ambientes de alta exigencia.",
        "Cumple con estándares de seguridad.",
        "Recomendado para continuidad operacional.",
    ]

    return (
        f"{random.choice(frases)} {random.choice(extra)} "
        f"Presentación: {qty} {um}."
    ).format(tipo=tipo, familia=familia, subfamilia=subfamilia, ambito=ambito)


# ============================================================
#  LLENADO DE FILAS PRINCIPALES
# ============================================================

# (Tus catálogos: CAT_TIPO_RECURSO, CAT_FAMILIA, CAT_SUB_FAMILIA, CAT_AMBITO, CAT_UM)
# ... (por brevedad no repito todo aquí, pero van igual que en tu script)

# (INCLUYE TUS CATÁLOGOS EXACTOS AQUÍ — OMITIDOS SOLO PARA MOSTRAR LA CORRECCIÓN)

# ============================================================
#  TABLAS AUXILIARES
# ============================================================

def populate_aux_tables(wb, metadata):
    catalog_map = {
        "aux_tipo_recurso": (CAT_TIPO_RECURSO, ["N", "CÓD.", "DESCRIPCIÓN"]),
        "aux_familia": (CAT_FAMILIA, ["N", "CÓD."]),
        "aux_sub_familia": (CAT_SUB_FAMILIA, ["N", "FAMILIA", "CÓD."]),
        "aux_ambito_aplicacion": (CAT_AMBITO, ["N", "CÓD."]),
        "aux_unidad_medida": (CAT_UM, ["N", "CÓD."]),
    }

    for hoja in metadata["hojas"]:
        ws = wb[hoja["nombre_hoja"]]

        for t in hoja["tablas"]:
            name = t["nombre_tabla"]
            if name not in catalog_map:
                continue

            catalog, cols = catalog_map[name]
            table = get_table(ws, name)

            min_c, min_r, max_c, max_r = normalize_table(ws, table)
            existing = max_r - min_r
            needed = len(catalog)

            if existing < needed:
                insert_rows_json_style(ws, table, t, needed - existing)

            for i, item in enumerate(catalog):
                r = min_r + 1 + i
                for j, col_name in enumerate(cols):
                    c = min_c + j
                    ws.cell(r, c).value = item[col_name] if col_name != "N" else i + 1

def fill_aux_columns(ws, table, json_table):
    """
    Llena automáticamente:
        - TIPO RECURSO
        - FAMILIA
        - SUB-FAMILIA MATERIAL
        - ÁMBITO APLICACIÓN
        - Material Descripción
        - Cantidad
        - U/M
    Usando catálogos CAT_...
    """

    min_c, min_r, max_c, max_r = range_boundaries(table.ref)
    cols = json_table["columnas"]

    c_tipo = min_c + cols.index("TIPO RECURSO")
    c_fam = min_c + cols.index("FAMILIA")
    c_sub = min_c + cols.index("SUB-FAMILIA MATERIAL")
    c_amb = min_c + cols.index("ÁMBITO APLICACIÓN")
    c_desc = min_c + cols.index("Material Descripción")
    c_qty = min_c + cols.index("Cantidad")
    c_um = min_c + cols.index("U/M")

    for r in range(min_r + 1, max_r + 1):

        tipo = random.choice(CAT_TIPO_RECURSO)
        familia = random.choice(CAT_FAMILIA)
        subfamilia = random.choice([sf for sf in CAT_SUB_FAMILIA
                                    if sf["FAMILIA"] == familia["CÓD."]] or CAT_SUB_FAMILIA)
        ambito = random.choice(CAT_AMBITO)
        um = random.choice(CAT_UM)

        qty = random.randint(1, 25)

        ws.cell(r, c_tipo).value = tipo["CÓD."]
        ws.cell(r, c_fam).value = familia["CÓD."]
        ws.cell(r, c_sub).value = subfamilia["CÓD."]
        ws.cell(r, c_amb).value = ambito["CÓD."]
        ws.cell(r, c_um).value = um["CÓD."]
        ws.cell(r, c_qty).value = qty

        ws.cell(r, c_desc).value = generar_descripcion_material(
            tipo=tipo["CÓD."],
            familia=familia["CÓD."],
            subfamilia=subfamilia["CÓD."],
            ambito=ambito["CÓD."],
            qty=qty,
            um=um["CÓD."],
        )

# ============================================================
#  FUNCIÓN PRINCIPAL — 1 ARCHIVO
# ============================================================

def insertar_filas(n_filas, json_metadata="estructura_listado_herramientas.json"):
    with open(json_metadata, "r", encoding="utf-8") as f:
        metadata = json.load(f)

    wb = load_workbook("LISTADO HERRAMIENTAS.xlsx")

    populate_aux_tables(wb, metadata)

    hoja = metadata["hojas"][0]
    t_json = hoja["tablas"][0]
    ws = wb["Cuerpo"]
    table = get_table(ws, t_json["nombre_tabla"])

    insert_rows_json_style(ws, table, t_json, n_filas)
    replicate_validations(ws, table, t_json)
    fill_correlative_and_sku(ws, table, t_json)
    fill_aux_columns(ws, table, t_json)

    wb.save("LISTADO HERRAMIENTAS - GENERADO.xlsx")
    print("✔ GENERADO: LISTADO HERRAMIENTAS - GENERADO.xlsx")


# ============================================================
#  MULTIPLE FILES
# ============================================================

def generar_varios_archivos(n_archivos, n_filas,
                            carpeta="salida_listados",
                            json_metadata="estructura_listado_herramientas.json"):

    if not os.path.exists(carpeta):
        os.makedirs(carpeta)

    for i in range(1, n_archivos + 1):
        nombre = f"LISTADO HERRAMIENTAS - {i:03d}.xlsx"
        ruta = os.path.join(carpeta, nombre)

        shutil.copy("LISTADO HERRAMIENTAS.xlsx", ruta)
        wb = load_workbook(ruta)

        with open(json_metadata, "r", encoding="utf-8") as f:
            metadata = json.load(f)

        populate_aux_tables(wb, metadata)

        hoja = metadata["hojas"][0]
        t_json = hoja["tablas"][0]

        ws = wb["Cuerpo"]
        table = get_table(ws, t_json["nombre_tabla"])

        insert_rows_json_style(ws, table, t_json, n_filas)
        replicate_validations(ws, table, t_json)
        fill_correlative_and_sku(ws, table, t_json)
        fill_aux_columns(ws, table, t_json)

        wb.save(ruta)
        print(f"✔ ARCHIVO GENERADO: {ruta}")


# ============================================================
# EJECUCIÓN
# ============================================================

if __name__ == "__main__":
    generar_varios_archivos(
        n_archivos=5,
        n_filas=55
    )
