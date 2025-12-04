# ================================================================
# llenar_sdi.py — VERSIÓN PROFESIONAL FINAL (SIN HARDCODEO DE CTLS)
# ---------------------------------------------------------------
# - nombre_proyecto fijo (regla de negocio)
# - Urgencias se leen desde estructura_sdi.json (dropDownList "urgencia")
# - Disciplinas se leen desde estructura_sdi.json (checkbox disciplina_*)
# - Impactos se leen desde estructura_sdi.json (checkbox impacto_*)
# - Siempre hay al menos 1 impacto marcado (plazo/costo/seguridad)
# - Tablas Excel embebidas con referencias sintéticas variadas
# ================================================================

import json
import zipfile
from io import BytesIO
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
import random
from openpyxl import load_workbook
import os

NS_ALL = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml"
}

# ================================================================
# 1) CONFIGURACIÓN DE NEGOCIO (NO VIENE DEL JSON)
# ================================================================

NOMBRE_PROYECTO_FIJO = "Overhaul Chancador Primario MP1000"

EMAILS_ORIGEN = [
    "mantencion@metso-services.cl",
    "overhaul@metso-services.cl",
    "ingenieria.cm@metso-services.cl",
    "servicio.tecnico@metso-services.cl",
    "analisis.fallas@metso-services.cl",
]

EMAILS_DESTINO_EXTRA = [
    "operaciones@codelco.cl",
    "turno.supervision@bhp.com",
    "cliente.mantenimiento@anglo.cl",
    "ingenieria.planta@cap.cl",
]

TEMAS_EXTRA = [
    "Se requiere evaluación técnica complementaria.",
    "Se solicita levantar riesgos asociados.",
    "Adjuntar documentación soporte.",
    "Incluye revisión dimensional adicional.",
    "Involucra coordinación con operaciones.",
    "Requiere aprobación de cliente.",
]

# ================================================================
# 2) CARGA DE CATÁLOGOS DESDE estructura_sdi.json
# ================================================================

def cargar_definiciones_desde_json(json_path):
    """
    Lee el JSON de estructura (generado por leer.py) y extrae:
      - DISCIPLINAS_TOTALES: ['calidad', 'mecanico', ...] desde disciplina_*
      - IMPACTOS: ['impacto_plazo', 'impacto_costo', 'impacto_seguridad']
      - URGENCIAS: lista de valores del dropDownList 'urgencia'
    """
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    controles = data.get("controles", [])

    disciplinas = []
    impactos = []
    urgencias = []

    for c in controles:
        tag = c.get("tag", "")
        tipo = c.get("tipo", "")
        valores = c.get("valores", None)

        # Disciplinas: todos los checkbox con tag "disciplina_xxx"
        if tipo == "checkbox" and tag.startswith("disciplina_"):
            disciplinas.append(tag.replace("disciplina_", ""))

        # Impactos: todos los checkbox con tag "impacto_xxx"
        elif tipo == "checkbox" and tag.startswith("impacto_"):
            impactos.append(tag)

        # Urgencia: control desplegable con tag "urgencia"
        elif tipo == "dropDownList" and tag == "urgencia":
            if isinstance(valores, list):
                urgencias = valores

    return {
        "disciplinas": disciplinas,
        "impactos": impactos,
        "urgencias": urgencias,
    }

# ================================================================
# 3) AUXILIARES WORD
# ================================================================

def set_dropdown_value(sdt, value):
    """
    Cambia correctamente el valor seleccionado de un contenido de lista desplegable.
    Reemplaza el <w:t> interno y reordena las <w:listItem> para que Word seleccione la correcta.
    """
    # 1) Cambiar texto visible
    content = sdt.find("w:sdtContent", NS_ALL)
    if content is not None:
        for t in content.findall(".//w:t", NS_ALL):
            t.text = value

    # 2) Reordenar items internos
    props = sdt.find("w:sdtPr", NS_ALL)
    if props is None:
        return

    dropdown = props.find("w:dropDownList", NS_ALL)
    if dropdown is None:
        return

    items = dropdown.findall("w:listItem", NS_ALL)
    if not items:
        return

    for itm in items:
        if itm.get(f"{{{NS_ALL['w']}}}value") == value:
            dropdown.remove(itm)
            dropdown.insert(0, itm)
            break

def insertar_valor_control(sdt, valor):
    content = sdt.find("w:sdtContent", NS_ALL)
    if content is None:
        return
    for t in content.findall(".//w:t", NS_ALL):
        t.text = str(valor)

def set_checkbox_value(sdt, checked):
    new_val = "1" if checked else "0"
    props = sdt.find("w:sdtPr", NS_ALL)

    if props is not None:
        chk14 = props.find(".//w14:checkbox", NS_ALL)
        if chk14 is not None:
            node = chk14.find("w14:checked", NS_ALL)
            if node is None:
                node = ET.SubElement(chk14, f"{{{NS_ALL['w14']}}}checked")
            node.set(f"{{{NS_ALL['w14']}}}val", new_val)

    content = sdt.find("w:sdtContent", NS_ALL)
    if content is None:
        return

    for t in content.findall(".//w:t", NS_ALL):
        if t.text in ("☐", "☑"):
            t.text = "☑" if checked else "☐"

# ================================================================
# 4) LÓGICA DE IMPACTOS (SIEMPRE AL MENOS UNO)
# ================================================================

def elegir_impactos_validos(impactos_disponibles):
    """
    Devuelve una lista de impactos seleccionados.
    Siempre incluye AL MENOS uno de los impactos presentes en la plantilla
    (impacto_plazo, impacto_costo, impacto_seguridad según JSON).
    A veces devuelve 2 o 3 para mayor realismo.
    """
    if not impactos_disponibles:
        return []

    impacto_principal = random.choice(impactos_disponibles)

    adicionales = []
    for imp in impactos_disponibles:
        if imp != impacto_principal and random.random() < 0.35:
            adicionales.append(imp)

    return [impacto_principal] + adicionales

# ================================================================
# 5) EXCEL EMBEBIDO – REFERENCIAS VARIADAS
# ================================================================

def generar_referencias_varias():
    base = [
        ("DOC-REF-001", "Procedimiento general"),
        ("DOC-REF-002", "Especificación técnica"),
        ("DOC-REF-003", "Informe de condición"),
        ("DOC-REF-004", "Plano asociado"),
    ]

    n = random.randint(2, 5)
    refs = []
    for i in range(n):
        c, d = random.choice(base)
        refs.append(
            (
                c + f"-{random.randint(10,99)}",
                d + " – " + random.choice(TEMAS_EXTRA),
            )
        )
    return refs

def llenar_excel_embebido(content, refs):
    wb = load_workbook(BytesIO(content), data_only=False)
    ws = wb["Hoja1"]

    tabla = ws._tables.get("docs_referencia")
    if tabla is None:
        return content

    min_row = int(tabla.ref.split(":")[0][1:])
    row = min_row + 1

    for code, desc in refs:
        ws[f"A{row}"].value = code
        ws[f"B{row}"].value = desc
        row += 1

    tabla.ref = f"A1:B{row-1}"

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# ================================================================
# 6) GENERAR 1 SDI
# ================================================================

def generar_sdi_unico(plantilla, salida, catalogos):
    """
    catalogos: dict con claves:
      - 'disciplinas': ['calidad', 'mecanico', ...]
      - 'impactos': ['impacto_plazo', 'impacto_costo', 'impacto_seguridad']
      - 'urgencias': ['Mediana', 'Normal', 'Inmediata']
    """
    disciplinas_totales = catalogos.get("disciplinas", [])
    impactos_disponibles = catalogos.get("impactos", [])
    urgencias_disponibles = catalogos.get("urgencias", [])

    proyecto = NOMBRE_PROYECTO_FIJO

    fecha = (datetime.now() - timedelta(days=random.randint(0, 30))).strftime("%Y-%m-%d")

    # Urgencia desde JSON; fallback mínimo si JSON viniera vacío
    if urgencias_disponibles:
        urgencia = random.choice(urgencias_disponibles)
    else:
        urgencia = "Mediana"

    tema = (
        f"Solicitud técnica relacionada con {proyecto.lower()}. "
        f"{random.choice(TEMAS_EXTRA)}"
    )

    # Disciplinas: subset aleatorio dinámico
    if disciplinas_totales:
        disciplinas = random.sample(
            disciplinas_totales,
            random.randint(1, len(disciplinas_totales))
        )
    else:
        disciplinas = []

    # Impactos: siempre al menos uno (si existen en el JSON)
    impactos_seleccionados = elegir_impactos_validos(impactos_disponibles)

    email_origen = random.choice(EMAILS_ORIGEN)
    email_destino = random.choice(EMAILS_DESTINO_EXTRA)

    referencias = generar_referencias_varias()

    print(f" → Generando SDI: {salida}")

    with zipfile.ZipFile(plantilla, "r") as zin:
        with zipfile.ZipFile(salida, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename == "word/document.xml":
                    root = ET.fromstring(data)

                    for sdt in root.findall(".//w:sdt", NS_ALL):
                        props = sdt.find("w:sdtPr", NS_ALL)
                        if props is None:
                            continue

                        tag_node = props.find("w:tag", NS_ALL)
                        if tag_node is None:
                            continue

                        tag = tag_node.get(f"{{{NS_ALL['w']}}}val")

                        if tag == "nombre_proyecto":
                            insertar_valor_control(sdt, proyecto)

                        elif tag == "fecha":
                            insertar_valor_control(sdt, fecha)

                        elif tag == "tema_solicitud":
                            insertar_valor_control(sdt, tema)

                        elif tag == "urgencia":
                            set_dropdown_value(sdt, urgencia)

                        elif tag.startswith("disciplina_"):
                            disc = tag.replace("disciplina_", "")
                            set_checkbox_value(sdt, disc in disciplinas)

                        elif tag.startswith("impacto_"):
                            # Marcamos TRUE sólo si este impacto fue seleccionado
                            set_checkbox_value(sdt, tag in impactos_seleccionados)

                        elif tag == "email_origen":
                            insertar_valor_control(sdt, email_origen)

                        elif tag == "email_destino":
                            insertar_valor_control(sdt, email_destino)

                    zout.writestr(
                        "word/document.xml",
                        ET.tostring(root, encoding="utf-8")
                    )
                    continue

                if item.filename.startswith("word/embeddings/") and item.filename.endswith(".xlsx"):
                    zout.writestr(item, llenar_excel_embebido(data, referencias))
                    continue

                zout.writestr(item, data)

# ================================================================
# 7) MULTI DOCUMENTOS
# ================================================================

def generar_n_documentos(n, plantilla, json_estructura, carpeta_salida):
    os.makedirs(carpeta_salida, exist_ok=True)

    # Cargar catálogos dinámicamente UNA sola vez
    catalogos = cargar_definiciones_desde_json(json_estructura)

    for i in range(1, n + 1):
        salida = os.path.join(carpeta_salida, f"SDI_{i:03d}.docx")
        generar_sdi_unico(plantilla, salida, catalogos)

    print(f"\n=== {n} SDI generados en '{carpeta_salida}' ===\n")

# ================================================================
# 8) MAIN
# ================================================================

if __name__ == "__main__":
    generar_n_documentos(
        n=20,
        plantilla="SDI.docx",
        json_estructura="estructura_sdi.json",
        carpeta_salida="SDIs_Generados"
    )
