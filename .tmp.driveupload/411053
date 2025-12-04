# ============================================================
# C:\Users\jonat\Desktop\memoria\Diseño\Formatos Documentos - Final\1 LISTOS\FINAL\Pre\test\Informe Tecnico\generar_json_reportes_diarios.py
# Versión FINAL corregida con manejo correcto de porcentajes
# MODIFICADO: Solo exportar imágenes dentro del control de contenido "imagen_evidencia"
# CORREGIDO: Evitar duplicados en imágenes y ordenar por número de imagen para asociar descripciones correctamente
# ============================================================

import os
import zipfile
import json
import re
from io import BytesIO
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
import datetime

# Namespace Word
NS_W = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

import zipfile
import os
import xml.etree.ElementTree as ET

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

def extraer_imagenes_y_descripciones(docx_path, output_dir, id_reporte):
    os.makedirs(output_dir, exist_ok=True)

    imagenes_png = []   # solo fotos verdaderas dentro de "imagen_evidencia"
    descripciones = []  # descripciones SDT
    fecha_turno = None
    processed_rids = set()  # Para evitar duplicados

    with zipfile.ZipFile(docx_path, "r") as z:

        # Leer relaciones para mapear rId a archivos de media
        rels_path = "word/_rels/document.xml.rels"
        if rels_path in z.namelist():
            rels_xml = z.read(rels_path)
            rels_root = ET.fromstring(rels_xml)
            rels = {}
            for rel in rels_root.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                rId = rel.get("Id")
                target = rel.get("Target")
                if target.startswith("media/") and target.lower().endswith(".png"):
                    rels[rId] = target

        # Leer document.xml
        xml_doc = z.read("word/document.xml")
        root = ET.fromstring(xml_doc)

        # Extraer descripciones SDT
        for sdt in root.findall(".//w:sdt", NS):
            tag = sdt.find(".//w:tag", NS)
            if tag is None:
                continue
            val = tag.get(f"{{{NS['w']}}}val")

            # Descripción de evidencia
            if val == "descripcion_imagen_evidencia":
                textos = sdt.findall(".//w:t", NS)
                texto = "".join([t.text or "" for t in textos]).strip()
                descripciones.append(texto)

            # Fecha del turno
            if val == "fecha_turno":
                textos = sdt.findall(".//w:t", NS)
                fecha_turno = "".join([t.text or "" for t in textos]).strip()

        # Extraer imágenes SOLO de SDT con tag "imagen_evidencia"
        for sdt in root.findall(".//w:sdt", NS):
            tag = sdt.find(".//w:tag", NS)
            if tag is None or tag.get(f"{{{NS['w']}}}val") != "imagen_evidencia":
                continue

            # Dentro del SDT, buscar referencias a imágenes (rId en w:drawing)
            for drawing in sdt.findall(".//w:drawing", NS):
                for blip in drawing.findall(".//a:blip", {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}):
                    rId = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                    if rId in rels and rId not in processed_rids:
                        processed_rids.add(rId)
                        mf = f"word/{rels[rId]}"
                        out_name = f"reporte_{id_reporte}_{os.path.basename(mf)}"
                        out_path = os.path.join(output_dir, out_name)

                        with open(out_path, "wb") as imgf:
                            imgf.write(z.read(mf))

                        imagenes_png.append({
                            "img": out_path,
                            "descripcion": None,
                            "fecha": None
                        })

    # Ordenar imágenes por el número en el nombre (e.g., image9, image10, etc.)
    def sort_key(img_dict):
        match = re.search(r'image(\d+)', img_dict["img"])
        return int(match.group(1)) if match else 0

    imagenes_png.sort(key=sort_key)

    # Asociar descripciones 1:1 en orden
    for i, foto in enumerate(imagenes_png):
        if i < len(descripciones):
            foto["descripcion"] = descripciones[i]
        foto["fecha"] = fecha_turno

    return imagenes_png


# =====================================================================
# CONVERSIÓN UNIVERSAL A FORMATO JSON SEGURO (solo para casos generales)
# =====================================================================
def make_json_safe(value):
    if isinstance(value, datetime.time):
        return value.strftime("%H:%M:%S")

    if isinstance(value, datetime.datetime):
        return value.time().strftime("%H:%M:%S")

    if isinstance(value, datetime.timedelta):
        total_seconds = int(value.total_seconds())
        h = total_seconds // 3600
        m = (total_seconds % 3600) // 60
        s = total_seconds % 60
        return f"{h:02d}:{m:02d}:{s:02d}"

    if isinstance(value, (int, float)):
        total_seconds = int(round(value * 24 * 3600))
        h = total_seconds // 3600
        m = (total_seconds % 3600) // 60
        s = total_seconds % 60
        return f"{h:02d}:{m:02d}:{s:02d}"

    if isinstance(value, str):
        if "." in value and ":" in value:
            try:
                t = datetime.datetime.strptime(value, "%H:%M:%S.%f")
                return t.strftime("%H:%M:%S")
            except:
                pass

        try:
            t = datetime.datetime.strptime(value, "%H:%M:%S")
            return t.strftime("%H:%M:%S")
        except:
            pass

        return value

    if isinstance(value, bytes):
        return value.decode("utf-8", errors="ignore")

    return str(value)


# ---------------------------------------------------------
# EXTRAER CONTROLES DE CONTENIDO (SDT)
# ---------------------------------------------------------
def extract_content_controls(document_xml_bytes):
    root = ET.fromstring(document_xml_bytes)
    datos = {}

    for sdt in root.findall(".//w:sdt", NS_W):
        tag_elem = sdt.find("./w:sdtPr/w:tag", NS_W)
        if tag_elem is None:
            continue

        tag_val = tag_elem.get(f"{{{NS_W['w']}}}val")
        if not tag_val:
            continue

        textos = sdt.findall(".//w:sdtContent//w:t", NS_W)
        full_text = "".join([t.text or "" for t in textos]).strip()

        datos[tag_val] = make_json_safe(full_text)

    return datos


# ---------------------------------------------------------
# EXTRAER NOMBRES DEFINIDOS DE EXCEL (Named Ranges)
# ---------------------------------------------------------
def extract_excel_names(wb):
    result = {}

    for name_str, defn in wb.defined_names.items():
        try:
            destinations = list(defn.destinations)
        except:
            continue

        for sheet_name, cell_ref in destinations:
            try:
                ws = wb[sheet_name]
                cell = ws[cell_ref]
                result[name_str] = make_json_safe(cell.value)
            except:
                pass

    return result


# ---------------------------------------------------------
# EXTRAER TABLAS DESDE UN EXCEL EMBEBIDO
# ---------------------------------------------------------
def extract_excel_tables(xlsx_bytes):
    try:
        wb = load_workbook(BytesIO(xlsx_bytes), data_only=True)
    except:
        return {}, {}

    tables_out = {}

    def parse_time(val):
        if val is None:
            return None

        if isinstance(val, (int, float)):
            total_seconds = int(round(val * 24 * 3600))
            h = total_seconds // 3600
            m = (total_seconds % 3600) // 60
            s = total_seconds % 60
            return datetime.time(h, m, s)

        if isinstance(val, str):
            try:
                return datetime.datetime.strptime(val.strip(), "%H:%M:%S").time()
            except:
                return None

        return None

    for ws in wb.worksheets:
        for tbl in ws._tables.values():
            tbl_name = tbl.name
            min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)

            headers = []
            for c in range(min_col, max_col + 1):
                val = ws.cell(row=min_row, column=c).value
                headers.append("" if val is None else str(val).strip())

            rows = []
            for r in range(min_row + 1, max_row + 1):
                row_data = {}

                for idx, c in enumerate(range(min_col, max_col + 1)):
                    header = headers[idx]
                    val = ws.cell(row=r, column=c).value
                    safe_header = header.strip().lower()

                    # 1) N° REPORTE DIARIO
                    if safe_header == "n° reporte diario":
                        try:
                            row_data[header] = int(val)
                        except:
                            try:
                                row_data[header] = int(float(val))
                            except:
                                row_data[header] = None
                        continue

                    # 2) Porcentajes (0–1)
                    es_porcentaje = (
                        "% avance diario planificado" in safe_header
                        or "% avance diario real" in safe_header
                        or "avance planificado (acumulado)" in safe_header
                        or "avance real (acumulado)" in safe_header
                    )

                    if es_porcentaje:
                        try:
                            if isinstance(val, (int, float)):
                                row_data[header] = float(val)
                            else:
                                s = str(val).replace("%", "").replace(",", ".")
                                num = float(s)
                                row_data[header] = num / 100 if num > 1 else num
                        except:
                            row_data[header] = 0.0
                        continue

                    # 3) Otros datos
                    row_data[header] = make_json_safe(val)

                # 4) Calcular H. TOTAL
                if "H. INICIO" in row_data and "H. FIN" in row_data:
                    t_ini = parse_time(row_data["H. INICIO"])
                    t_fin = parse_time(row_data["H. FIN"])

                    if t_ini and t_fin:
                        dt_ini = datetime.timedelta(
                            hours=t_ini.hour, minutes=t_ini.minute, seconds=t_ini.second
                        )
                        dt_fin = datetime.timedelta(
                            hours=t_fin.hour, minutes=t_fin.minute, seconds=t_fin.second
                        )

                        delta = dt_fin - dt_ini

                        if delta.total_seconds() >= 0:
                            row_data["H. TOTAL"] = make_json_safe(delta)
                        else:
                            row_data["H. TOTAL"] = None
                    else:
                        row_data["H. TOTAL"] = None

                rows.append(row_data)

            tables_out[tbl_name] = {"headers": headers, "rows": rows}

    names_out = extract_excel_names(wb)

    return tables_out, names_out


# ---------------------------------------------------------
# PROCESAR UN ARCHIVO .DOCX COMPLETO
# ---------------------------------------------------------
def process_docx(path):
    print(f"Procesando {os.path.basename(path)} ...")

    data_doc = {}

    with zipfile.ZipFile(path, "r") as z:

        # ---------------------------------------------------------
        # 1) EXTRAER CONTROLES DE CONTENIDO
        # ---------------------------------------------------------
        if "word/document.xml" in z.namelist():
            xml_doc = z.read("word/document.xml")
            data_doc["content_controls"] = extract_content_controls(xml_doc)
        else:
            data_doc["content_controls"] = {}

        # ---------------------------------------------------------
        # 2) EXTRAER TABLAS EXCEL EMBEBIDAS
        # ---------------------------------------------------------
        tablas = {}
        nombres_excel = {}

        for item in z.namelist():
            if item.startswith("word/embeddings/") and item.endswith(".xlsx"):
                xlsx_bytes = z.read(item)
                tbls, names = extract_excel_tables(xlsx_bytes)
                tablas.update(tbls)
                nombres_excel.update(names)

        data_doc["excel_tables"] = tablas
        data_doc["excel_names"] = nombres_excel

        # ---------------------------------------------------------
        # 3) CONSISTENCIA HH_TURNO
        # ---------------------------------------------------------
        hh1 = nombres_excel.get("hh_turno_directo")
        hh2 = nombres_excel.get("hh_turno_indirecto")

        if hh1 and hh2 and hh1 != hh2:
            print(f"⚠ Advertencia: hh_turno_directo != hh_turno_indirecto en {path}")

        data_doc["hh_turno"] = hh1 or hh2 or None

        # ---------------------------------------------------------
        # 4) EXTRAER IMÁGENES Y DESCRIPCIONES (MODIFICADO: Solo de "imagen_evidencia", ordenadas y sin duplicados)
        # ---------------------------------------------------------
        try:
            imagenes = extraer_imagenes_y_descripciones(
                docx_path=path,
                output_dir="imagenes_reportes",
                id_reporte=os.path.basename(path)
            )
        except Exception as e:
            print(f"⚠ Error extrayendo imágenes de {path}: {e}")
            imagenes = []

        # Añadir la fecha del turno a cada imagen
        fecha_turno = data_doc["content_controls"].get("fecha_turno", None)
        for img in imagenes:
            if fecha_turno:
                img["fecha"] = fecha_turno

        data_doc["imagenes"] = imagenes

    return data_doc



# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
def main(folder, output_file):
    documentos = []

    files = sorted(os.listdir(folder))

    for f in files:
        if f.lower().startswith("reporte_") and f.lower().endswith(".docx"):
            full_path = os.path.join(folder, f)
            info = process_docx(full_path)
            info["filename"] = f
            documentos.append(info)

    with open(output_file, "w", encoding="utf-8") as j:
        json.dump(documentos, j, indent=4, ensure_ascii=False)

    print("\n✔ JSON generado correctamente:")
    print(output_file)


# ---------------------------------------------------------
# CLI
# ---------------------------------------------------------
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Extraer datos de reportes diarios .docx")
    parser.add_argument("--dir", required=True)
    parser.add_argument("--out", default="reporte_consolidado.json")
    args = parser.parse_args()

    main(args.dir, args.out)
