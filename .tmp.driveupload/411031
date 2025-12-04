# =====================================================================
# generar_cartas_contractuales_desde_json.py — VERSIÓN UNIVERSAL FINAL
# Consume el JSON generado por leer_universal.py y rellena TODAS las SDT:
#   - document.xml
#   - headers
#   - footers
# Mantiene: alias, tipo, valores (dropdown), fechas, texto, etc.
# Incluye carpeta de salida para N documentos.
# =====================================================================

import os
import json
import random
from datetime import datetime, timedelta
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO

NS_W = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

# ---------------------------------------------------------------
# Catálogos profesionales
# ---------------------------------------------------------------
PROYECTOS = [
    "Overhaul Chancador Primario MP1000",
    "Reparación Molino SAG 40x22",
    "Cambio Coraza Molino Bolas 26x38",
    "Mantención Harnero Banana DF501",
    "Reemplazo Sistema Hidráulico Chancador Secundario HP300",
    "Upgrade de Sistema de Lubricación MP1250",
    "Servicio Major Overhaul Chancador GP500"
]

CLIENTES = [
    "BHP Escondida",
    "CODELCO Chuquicamata",
    "CODELCO Andina",
    "Collahuasi",
    "Anglo American Los Bronces",
    "Minera Centinela",
    "Minera Lomas Bayas"
]

FAENAS = [
    "Faena Escondida",
    "Faena Spence",
    "Faena Andina",
    "Faena Los Bronces",
    "Faena Radomiro Tomic",
    "Faena Collahuasi"
]

ADMINISTRADORES = [
    "Carlos Muñoz", "Francisca Rivas", "Pedro González",
    "Ana Villalobos", "Luis Arrieta", "Javiera Troncoso"
]

CODIGOS_DOCUMENTO = [
    "CTC-OVH-001",
    "CTC-MNT-045",
    "CTC-SRV-129",
    "CTC-MP1-334",
    "CTC-HP3-550"
]


# =====================================================================
# GENERADOR GENERAL DE VALORES (usa alias + tipo del JSON UNIVERSAL)
# =====================================================================
def generar_valor(alias, tipo, valores=None):

    alias_low = alias.lower() if alias else ""

    # --------------------- FECHAS ---------------------
    if tipo == "date":
        dias = random.randint(0, 14)
        f = datetime.now() - timedelta(days=dias)
        return f.strftime("%Y-%m-%d")

    # --------------------- DROPDOWNS ---------------------
    if tipo in ("dropDownList", "dropdown", "comboBox") and valores:
        return random.choice(valores)

    # --------------------- REGLAS POR ALIAS ---------------------
    if "tipo_documento" in alias_low:
        return "Carta Contractual"

    if "codigo_documento" in alias_low:
        return random.choice(CODIGOS_DOCUMENTO)

    if "nombre_proyecto" in alias_low:
        return random.choice(PROYECTOS)

    if "nombre_cliente" in alias_low:
        return random.choice(CLIENTES)

    if "nombre_faena" in alias_low:
        return random.choice(FAENAS)

    if "administrador_servicio" in alias_low:
        return random.choice(ADMINISTRADORES)

    if "numero_contrato" in alias_low:
        return f"CT-{random.randint(10000,99999)}-{random.randint(10,99)}"

    if "numero_servicio" in alias_low:
        return f"SV-{random.randint(100000,999999)}"

    if "cuerpo_carta" in alias_low:
        return (
            "Por medio de la presente informamos las condiciones asociadas al servicio de "
            "mantención y overhaul del presente proyecto. Las actividades comprenden inspección, "
            "reemplazo de componentes críticos, pruebas funcionales y entrega operativa conforme "
            "a los estándares de Metso Outotec."
        )

    # Fallback
    return f"Valor generado automáticamente para {alias}"


# =====================================================================
# RELLENAR CUALQUIER XML (document.xml, header/footer)
# =====================================================================
def fill_controls_in_xml(xml_bytes, mapping):
    root = ET.fromstring(xml_bytes)

    for sdt in root.findall(".//w:sdt", NS_W):

        tag_el = sdt.find("w:sdtPr/w:tag", NS_W)
        if tag_el is None:
            continue

        key = tag_el.get("{%s}val" % NS_W["w"])
        if key not in mapping:
            continue

        nuevo = str(mapping[key])

        content = sdt.find("w:sdtContent", NS_W)
        if content is None:
            continue

        # limpiar texto previo
        for t in content.findall(".//w:t", NS_W):
            t.text = ""

        run = content.find(".//w:r", NS_W)
        if run is None:
            run = ET.SubElement(content, "{%s}r" % NS_W["w"])

        t_final = run.find("w:t", NS_W)
        if t_final is None:
            t_final = ET.SubElement(run, "{%s}t" % NS_W["w"])

        t_final.text = nuevo

    return ET.tostring(root, encoding="utf-8")


# =====================================================================
# GENERAR DOCUMENTO COMPLETO DESDE PLANTILLA + JSON UNIVERSAL
# =====================================================================
def generar_documento_desde_json(path_plantilla, estructura_json, mapping, path_salida):

    with zipfile.ZipFile(path_plantilla, "r") as zin:
        buffer = BytesIO()

        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zout:

            for item in zin.infolist():
                contenido = zin.read(item.filename)

                # DOCUMENTO PRINCIPAL
                if item.filename == "word/document.xml":
                    contenido = fill_controls_in_xml(contenido, mapping)

                # HEADERS / FOOTERS
                elif item.filename.startswith("word/header") and item.filename.endswith(".xml"):
                    contenido = fill_controls_in_xml(contenido, mapping)

                elif item.filename.startswith("word/footer") and item.filename.endswith(".xml"):
                    contenido = fill_controls_in_xml(contenido, mapping)

                zout.writestr(item, contenido)

        with open(path_salida, "wb") as f:
            f.write(buffer.getvalue())


# =====================================================================
# MAIN UNIVERSAL — CARTAS CONTRACTUALES
# =====================================================================
if __name__ == "__main__":

    PLANTILLA = "CARTA CONTRACTUAL.docx"
    JSON_ESTRUCTURA = "estructura_carta.json"  # generado por leer_universal.py
    N = 5  # cantidad documentos

    # -----------------------------  
    # Carpeta de salida
    # -----------------------------
    OUTPUT_DIR = "CARTAS_GENERADAS"
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        print(f"[OK] Carpeta creada → {OUTPUT_DIR}")

    # -----------------------------
    # Leer JSON universal
    # -----------------------------
    estructura = json.load(open(JSON_ESTRUCTURA, "r", encoding="utf-8"))

    # fusiona controles: document + headers + footers
    controles_all = []
    controles_all.extend(estructura["controles"])
    controles_all.extend(estructura["controles_header"])
    controles_all.extend(estructura["controles_footer"])

    # -----------------------------
    # Generar N documentos
    # -----------------------------
    for i in range(1, N + 1):

        mapping = {}

        for cc in controles_all:
            alias = cc["alias"]
            tipo = cc["tipo"]
            valores = cc["valores"]

            if not alias:
                continue

            mapping[alias] = generar_valor(alias, tipo, valores)

        salida = os.path.join(OUTPUT_DIR, f"CARTA_CONTRACTUAL_{i}.docx")
        generar_documento_desde_json(PLANTILLA, estructura, mapping, salida)

        print(f"[OK] Generado → {salida}")

    print("\nPROCESO COMPLETADO — TODAS LAS CARTAS CONTRACTUALES LISTAS.")
