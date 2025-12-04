# ============================================================
# llenar_curva_poblamiento.py — VERSIÓN DEFINITIVA + JSON-DRIVEN
# - Usa descripcion_curva_poblamiento.json (leer_universal_excel.py)
# - Mantiene lógica de negocios original (contratos, bajas, cargos)
# - Toma desde JSON:
#       * columnas reales
#       * listas válidas (CONTRATO, ANEXO, JORNADA, etc.)
#       * fórmula de DETALLE
# ============================================================

import json
import random
import shutil
import os
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries, get_column_letter
from copy import copy
from openpyxl.worksheet.datavalidation import DataValidation

CARGOS_USADOS = []  # opcional para debug


# ============================================================
# CATÁLOGOS FIJOS (EVOLUCIONADOS)
# ============================================================

CARGOS_MINERIA_CRITICOS = [
    "Téc. Mecánico", "Téc. Senior Mecánico", "Supervisor Mecánico",
    "Téc. Especialista Chancadores", "Téc. Especialista HPGR",
    "Téc. Montaje/Desmontaje",
]

CARGOS_MINERIA_APOYO = [
    "Prevencionista HSE", "Planificador", "Programador MP",
    "Téc. Lubricación", "Téc. Hidráulico", "Téc. Instrumentación",
    "Bodeguero / Logística",
]

# -------------------------------------
# Variabilidad extendida de ciudades
# -------------------------------------
CIUDADES = [
    "Antofagasta", "Calama", "Iquique", "Mejillones", "Tocopilla",
    "Sierra Gorda", "Chuquicamata", "La Negra", "María Elena",
    "Taltal", "Copiapó", "Vallenar", "La Serena", "Coquimbo",
]

# -------------------------------------
# Zonas extendidas de faena real
# -------------------------------------
ZONAS_MINERAS = [
    "Planta Concentradora", "Línea de Chancado", "Correas Transportadoras",
    "SAG / Bolas", "Harneros", "Stock Pile", "Patio Chancadores",
    "Sala Motores", "Sala Lubricación", "Taller de Revestimientos",
    "Plataforma de Mantenimiento",
]

TURNOS = ["10x5", "7x7", "5x2"]
JORNADAS = ["Día", "Noche"]

# -------------------------------------
# Motivos y comentarios más variados
# -------------------------------------
MOTIVOS_AS = [
    "Accidente laboral con lesión",
    "Incidente en faena con incapacidad temporal",
    "Golpe o atrapamiento durante operación",
    "Evento HSE con detención inmediata",
]

COMENTARIOS_AS = [
    "Se activó protocolo HSE.",
    "Ingreso formal a mutual confirmado.",
    "Supervisor de turno completa investigación.",
    "Pendiente informe final de HSE corporativo.",
    "Acompañamiento del área HSE durante jornada.",
]

MOTIVOS_DS = [
    "Licencia médica por enfermedad común",
    "Reposo médico extendido",
    "Licencia por maternidad",
    "Cuadro viral severo",
]

COMENTARIOS_DS = [
    "RRHH validó licencia médica.",
    "Notificado al cliente.",
    "Reposo coordinado con salud ocupacional.",
    "Pendiente alta médica oficial.",
]

MOTIVOS_BAJA_GENERICOS = [
    "Término de faena", "Finalización de contrato",
    "Desempeño insuficiente", "Renuncia voluntaria",
    "Reemplazo operativo", "Falta grave / Seguridad",
    "Reasignación interna",
    "Reemplazo por mayor experiencia requerida",
    "Salud incompatible",
]

COMENTARIOS_GENERICOS = [
    "Coordinado con supervisor de turno.",
    "Notificado al área HSE.",
    "Requiere revisión de documentos.",
    "Gestión realizada con RRHH.",
    "Reemplazo debe ingresar en próxima jornada.",
    "Pendiente cierre administrativo.",
    "Confirmado con área de operaciones.",
    "Sin observaciones.",
]

# -------------------------------------
# Contratos según rol
# -------------------------------------
LOGICA_CONTRATO = {
    "Indefinido": {
        "prob_baja": 0.05,
        "turnos": ["7x7", "10x5"],
        "jornadas": ["Día"],
        "cargos": CARGOS_MINERIA_CRITICOS + CARGOS_MINERIA_APOYO,
    },
    "Eventual": {
        "prob_baja": 0.10,
        "turnos": ["7x7", "10x5", "5x2"],
        "jornadas": ["Día", "Noche"],
        "cargos": CARGOS_MINERIA_CRITICOS,
    },
    "P. Fijo": {
        "prob_baja": 0.03,
        "turnos": ["10x5", "7x7"],
        "jornadas": ["Día"],
        "cargos": CARGOS_MINERIA_CRITICOS + ["Prevencionista HSE"],
    },
    "Baja AS": {
        "prob_baja": 1.00,
        "turnos": TURNOS,
        "jornadas": JORNADAS,
        "cargos": CARGOS_MINERIA_CRITICOS,
    },
    "Baja DS": {
        "prob_baja": 1.00,
        "turnos": TURNOS,
        "jornadas": JORNADAS,
        "cargos": CARGOS_MINERIA_APOYO + CARGOS_MINERIA_CRITICOS,
    },
}

# ============================================================
# CATÁLOGO DE CARGOS (por especialidad)
# ============================================================

CATALOGO_CARGOS_POR_ESPECIALIDAD = {
    "Operaciones": [
        ("Planificador", "Planificador de mantenimiento de contratos, paradas y recursos."),
        ("Programador MP", "Programador de mantenimiento preventivo/correctivo en sistema MP."),
        ("Coordinador Operaciones", "Coordinación táctica, recursos y control operacional."),
        ("Supervisor General", "Supervisor general responsable de seguridad y producción."),
        ("Líder de Cuadrilla", "Responsable de cuadrilla y ejecución en terreno."),
    ],

    "Mecánica": [
        ("Téc. Mecánico", "Mecánico de mantención de equipos industriales."),
        ("Téc. Senior Mecánico", "Mecánico senior referente técnico en faena."),
        ("Supervisor Mecánico", "Supervisor de mantenimiento mecánico."),
        ("Téc. Montaje/Desmontaje", "Montaje y desmontaje de equipos de conminución."),
        ("Mecánico Ajustador", "Mecánico ajustador en precisión y armado fino."),
        ("Téc. Torqueo", "Torqueo controlado y tensión hidráulica de pernos."),
        ("Téc. Alineamiento", "Alineamiento láser de ejes y poleas."),
        ("Mecánico Planta", "Mecánico planta concentradora."),
    ],

    "Conminución": [
        ("Téc. Especialista Chancadores", "Especialista en overhaul de chancadores MP/HP/GP."),
        ("Téc. Especialista HPGR", "Especialista en HPGR y cambio de rodillos."),
        ("Téc. Mantención Chancadores", "Mantención preventiva/correctiva en chancadores."),
        ("Téc. Cambio Revestimientos", "Cambio revestimientos en molienda SAG/Bolas."),
        ("Téc. Overhaul SAG", "Overhaul de molinos SAG, tapas, piñón, chumaceras."),
        ("Téc. Overhaul Bolas", "Overhaul de molinos de bolas."),
        ("Rigger Faena", "Rigger para izaje crítico."),
        ("Gruero Operador", "Operador de grúa pluma/puente grúa."),
    ],

    "Eléctrica": [
        ("Téc. Eléctrico", "Técnico eléctrico industrial."),
        ("Electricista Mantención", "Electricista de mantención en faena."),
        ("Téc. Control Eléctrico", "Control de fuerza y sistemas eléctricos."),
    ],

    "Hidráulica": [
        ("Téc. Hidráulico", "Actuadores, cilindros, lubricación y sellos."),
        ("Lubricador Industrial", "Lubricación y filtración de sistemas de alta presión."),
        ("Téc. PowerPack", "Powerpacks hidráulicos de chancadores."),
    ],

    "Instrumentación": [
        ("Téc. Instrumentación", "Sensores, loops y control de procesos."),
        ("Téc. Monitoreo Vibracional", "Análisis de vibraciones y condición."),
    ],

    "HSE": [
        ("Prevencionista HSE", "Prevencionista en faena minera."),
        ("Asesor SST", "Asesor de seguridad y salud ocupacional."),
        ("Vigilante SST", "Vigilante de seguridad en trabajos críticos."),
    ],

    "Logística": [
        ("Bodeguero", "Control de inventario y repuestos."),
        ("Administrativo Faena", "Gestión documental en faena."),
        ("Control de Materiales", "Control de materiales, repuestos críticos y herramientas."),
    ],

    "Operadores": [
        ("Operador Planta", "Operador de planta concentradora."),
        ("Operador Equipos Auxiliares", "Operador de manipuladores y plataformas."),
        ("Operador Camión Pluma", "Operador de camión pluma."),
        ("Operador Telehandler", "Operador de telehandler."),
    ]
}

CATALOGO_CARGOS = []
for esp, items in CATALOGO_CARGOS_POR_ESPECIALIDAD.items():
    for codigo, desc in items:
        CATALOGO_CARGOS.append({"especialidad": esp, "codigo": codigo, "descripcion": desc})


# ============================================================
# GENERACIÓN DE NOMBRES, EMAILS, TELÉFONOS, DOMICILIO
# ============================================================

NOMBRES = [
    "Juan", "Pedro", "Carlos", "Francisco", "José", "Cristian", "Diego",
    "Marcelo", "Luis", "Andrés", "Gustavo", "Sebastián", "Felipe",
    "Roberto", "Victor", "Patricio", "Claudio", "Rodrigo", "Iván",
    "Alexis", "Gabriel", "Matías", "Hernán", "Leonardo"
]

APELLIDOS = [
    "Soto", "González", "Rojas", "Muñoz", "Araya", "Orellana", "Pérez",
    "Gutiérrez", "Castillo", "Vega", "Bravo", "Fuentes", "Reyes",
    "Paredes", "Campos", "Aguilera", "Cortés", "Saavedra",
    "Zúñiga", "Molina", "Riquelme", "Salazar", "Bustos"
]

CALLES = [
    "Av. Los Minerales", "Calle Chancadores", "Pasaje SAG 40x22",
    "Av. Lomas del Mineral", "Calle Fundición", "Pasaje Palas",
    "Calle Concentradora", "Av. Camino Minero", "Av. Fundición Norte"
]

DOMINIOS = [
    "metso.com", "outotec.com", "metsochile.cl",
    "contratista-minera.cl", "mantencion-industrial.cl"
]


def generar_nombre():
    nombre1 = random.choice(NOMBRES)
    nombre2 = random.choice(NOMBRES)
    apellido1 = random.choice(APELLIDOS)
    apellido2 = random.choice(APELLIDOS)

    nombres_completos = f"{nombre1} {nombre2}"
    apellidos_completos = f"{apellido1} {apellido2}"

    return nombres_completos, apellidos_completos


def generar_email(nombre, apellido, cargo):
    n1 = nombre.split()[0]
    a1 = apellido.split()[0]

    base = f"{n1.lower()}.{a1.lower()}"

    if "Supervisor" in cargo:
        base = f"{a1.lower()}.{n1.lower()}"

    dominio = random.choice(DOMINIOS)
    return f"{base}@{dominio}"


def generar_domicilio():
    calle = random.choice(CALLES)
    num = random.randint(10, 9999)
    depto = f"Depto {random.randint(101, 1904)}" if random.random() < 0.25 else ""
    return f"{calle} #{num} {depto}".strip()


def generar_celular():
    inicio = random.choice(["+5693", "+5694", "+5695", "+5696"])
    final = random.randint(1000000, 9999999)
    return f"{inicio}{final}"


def generar_rut():
    base = random.randint(1000000, 26000000)
    dv = "0123456789K"[base % 11]
    return f"{base}-{dv}"


# ============================================================
# UTILIDADES PARA TABLAS
# ============================================================

def leer_cargos_desde_excel(ws, table, json_table):
    """
    Lee los códigos de cargos realmente presentes en la tabla 'Cargos'
    (columna CÓD.) para que la NOMINA solo use valores válidos
    y consistentes con la validación INDIRECT("Cargos[CÓD.]").
    """
    min_c, min_r, max_c, max_r = range_boundaries(table.ref)
    cols = json_table["columnas"]

    for name in ["CÓD.", "COD", "CÓDIGO"]:
        if name in cols:
            idx = cols.index(name)
            break
    else:
        raise Exception("❌ No existe columna 'CÓD.' en la tabla cargos.")

    cargos = []
    for r in range(min_r + 1, max_r + 1):
        val = ws.cell(row=r, column=min_c + idx).value
        if val:
            cargos.append(str(val).strip())
    return cargos


def copy_style(src, dst):
    dst.font = copy(src.font)
    dst.border = copy(src.border)
    dst.fill = copy(src.fill)
    dst.number_format = copy(src.number_format)
    dst.protection = copy(src.protection)
    dst.alignment = copy(src.alignment)


def get_table(ws, name):
    if name in ws.tables:
        return ws.tables[name]
    raise Exception(f"❌ Tabla '{name}' no encontrada en hoja '{ws.title}'")


def insert_rows_with_format(ws, table, num_rows):
    """
    Inserta filas al final de la tabla copiando el formato
    de la última fila existente (modelo).
    """
    if num_rows <= 0:
        return

    min_c, min_r, max_c, max_r = range_boundaries(table.ref)
    insert_at = max_r + 1
    model_row = max_r

    ws.insert_rows(insert_at, num_rows)

    for i in range(num_rows):
        new_r = insert_at + i
        for col in range(min_c, max_c + 1):
            dst = ws.cell(row=new_r, column=col)
            src = ws.cell(row=model_row, column=col)
            dst.value = None
            copy_style(src, dst)

    new_max_r = max_r + num_rows
    table.ref = f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c)}{new_max_r}"


def apply_validations(ws, table, json_table):
    """
    Reaplica las validaciones de datos de la tabla usando
    la definición almacenada en el JSON (tipo, fórmula, columnas).
    """
    min_c, min_r, max_c, max_r = range_boundaries(table.ref)

    for val in json_table["validaciones"]:
        col_name = val["columnas_afectadas"][0]
        idx = json_table["columnas"].index(col_name)
        col_letter = get_column_letter(min_c + idx)
        rango = f"{col_letter}{min_r+1}:{col_letter}{max_r}"

        dv = DataValidation(
            type=val["tipo"],
            formula1=val["formula1"],
            allow_blank=val["permitir_nulos"]
        )
        dv.sqref = rango
        ws.add_data_validation(dv)


def get_allowed_values_from_json(json_table):
    """
    Extrae listas de validación explícitas desde el JSON
    (solo aquellas definidas como "valores separados por coma" entre comillas).
    """
    result = {}
    for val in json_table["validaciones"]:
        if val["tipo"] == "list":
            formula = val["formula1"]
            # Formato: "\"Indefinido,Eventual,Baja AS, Baja DS, P. Fijo\""
            if (
                isinstance(formula, str)
                and formula.startswith("\"")
                and formula.endswith("\"")
                and "," in formula
                and "INDIRECT" not in formula
            ):
                lista_raw = formula.strip('"')
                valores = [v.strip() for v in lista_raw.split(",")]
                for col in val["columnas_afectadas"]:
                    result[col] = valores
    return result


# ============================================================
# RELLENO TABLA NOMINA (JSON-DRIVEN)
# ============================================================

def fill_nomina(ws, table, json_table, n_rows, cargos_disponibles=None):
    """
    Llena la tabla NOMINA usando:
      - estructura + validaciones desde JSON (columnas/valores permitidos)
      - lógica de negocios (LOGICA_CONTRATO, bajas AS/DS, etc.)
    """
    # Insertar filas y replicar validaciones de JSON
    insert_rows_with_format(ws, table, n_rows)
    apply_validations(ws, table, json_table)

    valid = get_allowed_values_from_json(json_table)
    lista_bajas = []

    min_c, min_r, max_c, max_r = range_boundaries(table.ref)
    cols = json_table["columnas"]

    # --- Listas desde JSON (CONTRATO, ACREDITACION, etc.) ---
    contratos_json = valid.get("CONTRATO", [])
    # Nos quedamos solo con los contratos para los cuales tenemos lógica:
    contratos_vals = [c for c in contratos_json if c in LOGICA_CONTRATO] or list(LOGICA_CONTRATO.keys())

    acredit_vals = valid.get("ACREDITACION", ["Acreditado", "Revisión", "Stand-By", "Rechazado"])
    hab_vals = valid.get("HABILITACION", ["Revisión", "Stand-By"])
    tipo_vals = valid.get("TIPO", ["Directo", "Indirecto"])
    anexo_vals = valid.get("ANEXO", TURNOS + ["4x3"])
    jornada_vals = valid.get("JORNADA", JORNADAS)

    # Fórmula DETALLE desde JSON (plantilla)
    detalle_formula = None
    if json_table["filas"]:
        detalle_formula_candidate = json_table["filas"][0].get("DETALLE")
        if isinstance(detalle_formula_candidate, str) and detalle_formula_candidate.startswith("="):
            detalle_formula = detalle_formula_candidate

    if cargos_disponibles is None:
        cargos_disponibles = []

    for row in range(min_r + 1, max_r + 1):

        nombre, apellido = generar_nombre()
        rut = generar_rut()

        # Seleccionar contrato coherente y con lógica
        contrato = random.choice(contratos_vals)
        log = LOGICA_CONTRATO.get(contrato, LOGICA_CONTRATO["Indefinido"])

        # Selección de cargo coherente
        if cargos_disponibles:
            preferidos = [c for c in log["cargos"] if c in cargos_disponibles]
            pool = preferidos if preferidos else cargos_disponibles
        else:
            pool = log["cargos"]

        cargo = random.choice(pool)

        # Turnos coherentes con JSON + lógica
        turnos_log = [t for t in anexo_vals if t in log["turnos"]] or anexo_vals

        # Jornadas coherentes con JSON + lógica
        jornadas_log = [j for j in jornada_vals if j in log["jornadas"]] or jornada_vals

        anexo = random.choice(turnos_log)
        jornada = random.choice(jornadas_log)

        # Si es baja
        es_baja = contrato in ["Baja AS", "Baja DS"]

        # Motivo y comentario
        if contrato == "Baja AS":
            motivo = random.choice(MOTIVOS_AS)
            comentario = random.choice(COMENTARIOS_AS)
        elif contrato == "Baja DS":
            motivo = random.choice(MOTIVOS_DS)
            comentario = random.choice(COMENTARIOS_DS)
        else:
            motivo = random.choice(MOTIVOS_BAJA_GENERICOS) if random.random() < 0.15 else ""
            comentario = random.choice(COMENTARIOS_GENERICOS + [""]) if random.random() < 0.4 else ""

        # Tipo de trabajador
        if any(k in cargo for k in ["Operador", "Mecánico", "Téc."]):
            tipo = "Directo"
        else:
            tipo = random.choice(tipo_vals)

        # Armar diccionario según columnas de JSON
        data = {
            "RUT": rut,
            "NOMBRES": nombre,
            "APELLIDOS": apellido,
            "CARGO": cargo,
            "CONTRATO": contrato,
            "ACREDITACION": random.choice(acredit_vals),
            "HABILITACION": random.choice(hab_vals),
            "TIPO": tipo,
            "CIUDAD": random.choice(CIUDADES),
            "DOMICILIO": generar_domicilio(),
            "CELULAR": generar_celular(),
            "EMAIL": generar_email(nombre, apellido, cargo),
            "ANEXO": anexo,
            "JORNADA": jornada,
            "MOTIVO BAJA": motivo,
            "COMENTARIO": comentario,
        }

        # DETALLE: si hay fórmula en JSON, se usa SIEMPRE (marca BAJA dinámicamente)
        if "DETALLE" in cols:
            if detalle_formula is not None:
                data["DETALLE"] = detalle_formula
            else:
                # fallback ultra simple
                data["DETALLE"] = ""

        # Escribir en la hoja según orden definido en JSON
        for col, key in enumerate(cols, start=min_c):
            ws.cell(row=row, column=col).value = data.get(key)

        # Registrar bajas coherentes (para tabla BAJAS)
        if es_baja:
            lista_bajas.append(data)

    return lista_bajas


# ============================================================
# TABLA CARGOS
# ============================================================

def _buscar_indice_columna(nombre_columna, columnas, alternativas):
    if nombre_columna in columnas:
        return columnas.index(nombre_columna)
    for alt in alternativas:
        if alt in columnas:
            return columnas.index(alt)
    return None


def fill_cargos(ws, table, json_table, n_rows=None):
    """
    Llena la tabla 'Cargos' usando CATALOGO_CARGOS:
      - CÓD.  -> código corto del cargo
      - CARGO -> descripción larga
    Usando columnas y rango obtenidos desde el JSON.
    """

    catalogo = CATALOGO_CARGOS
    total_rows = len(catalogo)

    min_c, min_r, max_c, max_r = range_boundaries(table.ref)
    cols = json_table["columnas"]

    idx_cod = _buscar_indice_columna("CÓD.", cols, ["COD", "CÓDIGO"])
    idx_cargo = _buscar_indice_columna("CARGO", cols, ["DESCRIPCION", "DESCRIPCIÓN"])

    if idx_cod is None or idx_cargo is None:
        raise Exception("❌ Falta columna CÓD. o CARGO en tabla 'cargos'.")

    existing_rows = max_r - min_r

    # Limpia datos actuales
    for r in range(min_r + 1, max_r + 1):
        for c in range(min_c, max_c + 1):
            ws.cell(row=r, column=c).value = None

    # Ajusta cantidad de filas
    if total_rows > existing_rows:
        insert_rows_with_format(ws, table, total_rows - existing_rows)
        min_c, min_r, max_c, max_r = range_boundaries(table.ref)

    # Escribe catálogo
    for i, cargo in enumerate(catalogo):
        row_idx = min_r + 1 + i
        ws.cell(row=row_idx, column=min_c + idx_cod).value = cargo["codigo"]
        ws.cell(row=row_idx, column=min_c + idx_cargo).value = cargo["descripcion"]

    # (la tabla Cargos no tiene validaciones en JSON, pero por consistencia)
    if json_table.get("validaciones"):
        apply_validations(ws, table, json_table)


# ============================================================
# TABLA BAJAS (EXTENDIDA)
# ============================================================

def fill_bajas(ws, table, json_table, lista_bajas):
    """
    Llena tabla BAJAS con la lista de colaboradores marcados como baja
    en NOMINA, manteniendo coherentes:
      - RUT / NOMBRE / APELLIDOS
      - CARGO / CONTRATO / ACREDITACION / HABILITACION / TIPO
      - ANEXO / JORNADA
    Y agregando una ZONA minera coherente.
    """

    if len(lista_bajas) == 0:
        return

    min_c, min_r, max_c, max_r = range_boundaries(table.ref)
    cols = json_table["columnas"]

    if len(lista_bajas) == 1:
        baja = lista_bajas[0]
        row_idx = min_r + 1
        for col, key in enumerate(cols, start=min_c):
            if key == "ZONA":
                valor = random.choice(ZONAS_MINERAS)
            else:
                valor = baja.get(key, "")
            ws.cell(row=row_idx, column=col).value = valor
        apply_validations(ws, table, json_table)
        return

    insert_rows_with_format(ws, table, len(lista_bajas) - 1)
    apply_validations(ws, table, json_table)

    row_idx = min_r + 1
    for baja in lista_bajas:
        for col, key in enumerate(cols, start=min_c):
            if key == "ZONA":
                valor = random.choice(ZONAS_MINERAS)
            else:
                valor = baja.get(key, "")
            ws.cell(row=row_idx, column=col).value = valor
        row_idx += 1


# ============================================================
# FUNCIÓN PRINCIPAL (1 archivo)
# ============================================================

def generar_curva(n_nomina=30, n_cargos=20, salida_personalizada=None):

    json_path = "estructura_curva_poblamiento.json"
    plantilla = "CURVA DE POBLAMIENTO.xlsx"

    if salida_personalizada:
        salida = salida_personalizada
    else:
        salida = "CURVA DE POBLAMIENTO - GENERADA.xlsx"

    with open(json_path, "r", encoding="utf-8") as f:
        meta = json.load(f)

    shutil.copy(plantilla, salida)
    wb = load_workbook(salida)

    lista_bajas = []
    cargos_disponibles = []

    # 1) Llenar CARGOS primero (para que la validación INDIRECT apunte a algo real)
    for hoja in meta["hojas"]:
        if hoja["nombre_hoja"] != "Cargos":
            continue
        ws = wb[hoja["nombre_hoja"]]
        for tabla in hoja["tablas"]:
            if tabla["nombre_tabla"] != "cargos":
                continue
            table_obj = get_table(ws, tabla["nombre_tabla"])
            fill_cargos(ws, table_obj, tabla, n_cargos)
            cargos_disponibles = leer_cargos_desde_excel(ws, table_obj, tabla)

    # 2) Llenar NOMINA y BAJAS usando estructura y validaciones del JSON
    for hoja in meta["hojas"]:
        ws = wb[hoja["nombre_hoja"]]
        for tabla in hoja["tablas"]:
            name = tabla["nombre_tabla"]
            table_obj = get_table(ws, name)

            if name == "nomina":
                lista_bajas = fill_nomina(ws, table_obj, tabla, n_nomina, cargos_disponibles)
            elif name == "bajas":
                fill_bajas(ws, table_obj, tabla, lista_bajas)

    wb.save(salida)
    print(f"✔ Archivo generado correctamente: {salida}")


# ============================================================
# GENERAR VARIOS DOCUMENTOS
# ============================================================

def generar_varios_documentos(
        cantidad_docs=5,
        n_nomina=80,
        n_cargos=20,
        carpeta_salida="SALIDA"
    ):

    if not os.path.exists(carpeta_salida):
        os.makedirs(carpeta_salida)

    print(f"\n📁 Carpeta de salida: {carpeta_salida}")
    print(f"📄 Generando {cantidad_docs} documentos...\n")

    for i in range(1, cantidad_docs + 1):

        nombre_archivo = f"CURVA_POBLAMIENTO_{i:03d}.xlsx"
        ruta_salida = os.path.join(carpeta_salida, nombre_archivo)

        print(f"➡ Generando archivo {i}/{cantidad_docs}: {nombre_archivo}")

        generar_curva(
            n_nomina=n_nomina,
            n_cargos=n_cargos,
            salida_personalizada=ruta_salida
        )

    print("\n✔✔ Todos los documentos fueron generados correctamente ✔✔\n")


# ============================================================
# EJECUCIÓN DIRECTA
# ============================================================

if __name__ == "__main__":
    generar_varios_documentos(
        cantidad_docs=10,
        n_nomina=110,
        n_cargos=25,
        carpeta_salida="OUTPUT_CURVAS"
    )
