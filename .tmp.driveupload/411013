#leer_universal.py

import zipfile
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
from io import BytesIO
import re
import json
import os

# Namespaces Word
NS_ALL = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
}

NS_W = {"w": NS_ALL["w"]}

# Namespace Excel
NS_XL = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


# ==========================================================
# PARSEO DE styles.xml (Excel) PARA RESOLVER ESTILOS COMPLETOS
# ==========================================================
class StylesResolver:
    """
    Resuelve estilos de celda Excel a partir de styles.xml.
    Se reutiliza tanto para SDI como para Reporte Diario, etc.
    """

    def __init__(self, styles_xml_text):
        self.fonts = []
        self.fills = []
        self.borders = []
        self.num_fmts = {}
        self.cell_xfs = []

        if not styles_xml_text:
            return

        root = ET.fromstring(styles_xml_text)

        # numFmts (formatos numéricos personalizados)
        numfmts_elem = root.find("s:numFmts", NS_XL)
        if numfmts_elem is not None:
            for nf in numfmts_elem.findall("s:numFmt", NS_XL):
                num_fmt_id = nf.get("numFmtId")
                format_code = nf.get("formatCode")
                if num_fmt_id:
                    self.num_fmts[int(num_fmt_id)] = format_code

        # fonts
        fonts_elem = root.find("s:fonts", NS_XL)
        if fonts_elem is not None:
            for f in fonts_elem.findall("s:font", NS_XL):
                name_el = f.find("s:name", NS_XL)
                sz_el = f.find("s:sz", NS_XL)
                color_el = f.find("s:color", NS_XL)
                font = {
                    "name": name_el.get("val") if name_el is not None else None,
                    "size": float(sz_el.get("val")) if sz_el is not None and sz_el.get("val") else None,
                    "bold": f.find("s:b", NS_XL) is not None,
                    "italic": f.find("s:i", NS_XL) is not None,
                    "underline": f.find("s:u", NS_XL) is not None,
                    "color": dict(color_el.attrib) if color_el is not None else None,
                }
                self.fonts.append(font)

        # fills
        fills_elem = root.find("s:fills", NS_XL)
        if fills_elem is not None:
            for fl in fills_elem.findall("s:fill", NS_XL):
                pattern_el = fl.find("s:patternFill", NS_XL)
                fill = {"patternType": None, "fgColor": None, "bgColor": None}
                if pattern_el is not None:
                    fill["patternType"] = pattern_el.get("patternType")
                    fg_el = pattern_el.find("s:fgColor", NS_XL)
                    bg_el = pattern_el.find("s:bgColor", NS_XL)
                    fill["fgColor"] = dict(fg_el.attrib) if fg_el is not None else None
                    fill["bgColor"] = dict(bg_el.attrib) if bg_el is not None else None
                self.fills.append(fill)

        # borders
        borders_elem = root.find("s:borders", NS_XL)
        if borders_elem is not None:
            for bd in borders_elem.findall("s:border", NS_XL):
                border = {}
                for side in ["left", "right", "top", "bottom"]:
                    side_el = bd.find(f"s:{side}", NS_XL)
                    border[side] = side_el.get("style") if side_el is not None else None
                self.borders.append(border)

        # cellXfs (estilos aplicados a celdas)
        cellxfs_elem = root.find("s:cellXfs", NS_XL)
        if cellxfs_elem is not None:
            for xf in cellxfs_elem.findall("s:xf", NS_XL):
                numFmtId = xf.get("numFmtId")
                fontId = xf.get("fontId")
                fillId = xf.get("fillId")
                borderId = xf.get("borderId")

                align_el = xf.find("s:alignment", NS_XL)
                alignment = dict(align_el.attrib) if align_el is not None else None

                xf_dict = {
                    "numFmtId": int(numFmtId) if numFmtId else None,
                    "fontId": int(fontId) if fontId else None,
                    "fillId": int(fillId) if fillId else None,
                    "borderId": int(borderId) if borderId else None,
                    "alignment": alignment,
                }
                self.cell_xfs.append(xf_dict)

    def get_style_by_index(self, idx):
        """
        Devuelve un dict con font/fill/border/numFmt/alignment
        para el índice de estilo 's' de la celda.
        """
        if idx is None:
            return {}

        try:
            idx = int(idx)
        except Exception:
            return {}

        if not (0 <= idx < len(self.cell_xfs)):
            return {}

        xf = self.cell_xfs[idx]
        style = {"xf_index": idx}

        numFmtId = xf.get("numFmtId")
        if numFmtId is not None:
            style["numFmtId"] = numFmtId
            style["numFmtCode"] = self.num_fmts.get(numFmtId)

        fontId = xf.get("fontId")
        if fontId is not None and 0 <= fontId < len(self.fonts):
            style["font"] = self.fonts[fontId]

        fillId = xf.get("fillId")
        if fillId is not None and 0 <= fillId < len(self.fills):
            style["fill"] = self.fills[fillId]

        borderId = xf.get("borderId")
        if borderId is not None and 0 <= borderId < len(self.borders):
            style["border"] = self.borders[borderId]

        if xf.get("alignment"):
            style["alignment"] = xf["alignment"]

        return style


# ==========================================================
# UTILIDADES BASE WORD
# ==========================================================
def read_xml_from_docx(zip_obj, path):
    if path in zip_obj.namelist():
        return zip_obj.read(path).decode("utf-8", "ignore")
    return None


def extract_list_entries(xml_text):
    """
    Extrae posibles valores de listas (dropDown / comboBox) desde
    styles.xml, settings, glossary, etc.
    """
    if xml_text is None:
        return []

    patterns = [
        r'<w:listEntry[^>]*w:value="(.*?)"',
        r'<w14:listEntry[^>]*w14:val="(.*?)"',
        r'<w15:listEntry[^>]*w15:val="(.*?)"',
        r'<w:listItem[^>]*w:value="(.*?)"',
        r'<w:listItem[^>]*w:displayText=".*?" w:value="(.*?)"',
    ]
    vals = []
    for p in patterns:
        vals.extend(re.findall(p, xml_text))
    return list(set(vals))


# ==========================================================
# LECTURA DE VALOR / FÓRMULA CELDA
# ==========================================================
def get_cell_formula_or_value(cell, sheet_xml=None):
    """
    Devuelve el valor de la celda preservando fórmulas cuando existan.
    Usa info de openpyxl y, si es necesario, busca <f> en el XML de la hoja.
    """
    # ArrayFormula
    if isinstance(cell.value, ArrayFormula):
        f = cell.value.text.strip()
        return "=" + f if not f.startswith("=") else f

    # Fórmula simple
    if isinstance(cell.value, str) and cell.value.startswith("="):
        return cell.value

    # Búsqueda directa en XML de la hoja
    if sheet_xml:
        patt = rf'<c[^>]*r="{cell.coordinate}"[^>]*>.*?<f[^>]*>(.*?)</f>'
        m = re.search(patt, sheet_xml, flags=re.DOTALL)
        if m:
            f = m.group(1).strip()
            return "=" + f if not f.startswith("=") else f

    return cell.value


# ==========================================================
# FORMATO COMPLETO DE CELDA (openpyxl + sheet.xml + styles.xml)
# ==========================================================
def safe_color_to_str(color_obj):
    try:
        if color_obj is None:
            return None
        if getattr(color_obj, "type", None) == "rgb":
            return color_obj.rgb
        return str(color_obj)
    except Exception:
        return None


def get_cell_xml_type(sheet_xml, cell_ref):
    """
    Retorna info básica de la celda según el XML <c> de la hoja:
    tipo (t), estilo (s), si tiene fórmula, y otros atributos crudos.
    """
    if not sheet_xml:
        return {"t": None, "s": None, "is_formula": False, "style_attrs": {}}

    patt = rf'<c[^>]*r="{cell_ref}"([^>]*)>(.*?)</c>'
    m = re.search(patt, sheet_xml, flags=re.DOTALL)
    if not m:
        return {"t": None, "s": None, "is_formula": False, "style_attrs": {}}

    attrs = m.group(1)
    inner = m.group(2)
    tipo = None
    estilo = None
    is_formula = "<f" in inner

    mt = re.search(r't="(.*?)"', attrs)
    if mt:
        tipo = mt.group(1)

    ms = re.search(r's="(.*?)"', attrs)
    if ms:
        estilo = ms.group(1)

    style_attrs = {k: v for k, v in re.findall(r'(\w+)="([^"]*)"', attrs)}

    return {
        "t": tipo,
        "s": estilo,
        "is_formula": is_formula,
        "style_attrs": style_attrs,
    }


def get_complete_cell_format(cell, sheet_xml=None, styles_resolver=None):
    """
    Devuelve un diccionario con TODO el estilo disponible de una celda:
    - openpyxl: font, fill, alignment, borders, number_format, data_type
    - XML <c>: t, s, attrs, is_formula
    - styles.xml: font/fill/border/numFmt resueltos por índice s
    """
    try:
        font_color = safe_color_to_str(cell.font.color) if cell.font else None
    except Exception:
        font_color = None

    try:
        fill_color = safe_color_to_str(cell.fill.fgColor) if cell.fill else None
    except Exception:
        fill_color = None

    fmt = {
        "coordinate": cell.coordinate,
        "font_name": cell.font.name if cell.font else None,
        "font_size": float(cell.font.size) if cell.font and cell.font.size else None,
        "font_bold": bool(cell.font.bold) if cell.font else None,
        "font_italic": bool(cell.font.italic) if cell.font else None,
        "font_color": font_color,
        "alignment_horizontal": cell.alignment.horizontal if cell.alignment else None,
        "alignment_vertical": cell.alignment.vertical if cell.alignment else None,
        "border_left": cell.border.left.style if cell.border else None,
        "border_right": cell.border.right.style if cell.border else None,
        "border_top": cell.border.top.style if cell.border else None,
        "border_bottom": cell.border.bottom.style if cell.border else None,
        "fill_color": fill_color,
        "number_format": cell.number_format,
        "data_type": cell.data_type,
        "is_formula": isinstance(cell.value, str) and cell.value.startswith("="),
    }

    xml_info = None
    if sheet_xml:
        xml_info = get_cell_xml_type(sheet_xml, cell.coordinate)
        fmt.update(
            {
                "xml_t": xml_info.get("t"),
                "xml_s": xml_info.get("s"),
                "xml_is_formula": xml_info.get("is_formula"),
                "xml_raw_attrs": xml_info.get("style_attrs", {}),
            }
        )

    if styles_resolver and xml_info:
        s_idx = xml_info.get("s")
        if s_idx is not None:
            resolved = styles_resolver.get_style_by_index(s_idx)
            fmt["resolved_style"] = resolved

    return fmt


# ==========================================================
# CONTROLES DE CONTENIDO (Word, versión completa)
# ==========================================================
def _extract_sdt_list_from_root(root, valores_globales):
    """
    Extrae todos los w:sdt de un root XML y los devuelve
    en un formato estándar (lista de dicts).
    """
    controles = []

    for sdt in root.findall(".//w:sdt", NS_ALL):
        props = sdt.find("w:sdtPr", NS_ALL)
        if props is None:
            continue

        alias_node = props.find("w:alias", NS_ALL)
        tag_node = props.find("w:tag", NS_ALL)
        id_node = props.find("w:id", NS_ALL)

        alias = alias_node.get(f"{{{NS_ALL['w']}}}val") if alias_node is not None else None
        tag = tag_node.get(f"{{{NS_ALL['w']}}}val") if tag_node is not None else None
        cid = id_node.get(f"{{{NS_ALL['w']}}}val") if id_node is not None else None

        tipo = None
        valores_cc = []
        checked = None

        # DETECCIÓN DE TIPOS
        checkbox_moderno = props.find("w14:checkbox", NS_ALL)
        checkbox_antiguo = props.find("w:checkbox", NS_ALL)
        date_node = props.find("w:date", NS_ALL)
        drop = props.find("w:dropDownList", NS_ALL)
        combo = props.find("w:comboBox", NS_ALL)
        picture = props.find("w:picture", NS_ALL)
        richtext = props.find("w:richText", NS_ALL)
        repeating_section = props.find("w:repeatingSection", NS_ALL)
        repeating_item = props.find("w:repeatingSectionItem", NS_ALL)

        # LISTAS (dropDown / combo)
        if drop is not None or combo is not None:
            tipo = "dropDownList" if drop is not None else "comboBox"

            for entry in props.findall(".//w:listEntry", NS_ALL):
                val = entry.get(f"{{{NS_ALL['w']}}}value")
                if val:
                    valores_cc.append(val)

            for entry in props.findall(".//w:listItem", NS_ALL):
                val = entry.get(f"{{{NS_ALL['w']}}}value")
                if val:
                    valores_cc.append(val)

            if not valores_cc:
                valores_cc = valores_globales

            valores_cc = list(set(valores_cc))

        # CHECKBOX moderno
        elif checkbox_moderno is not None:
            tipo = "checkbox"
            checked_node = checkbox_moderno.find("w14:checked", NS_ALL)
            if checked_node is not None:
                val = checked_node.get(f"{{{NS_ALL['w14']}}}val")
                checked = val == "1"
            valores_cc = [True, False]

        # CHECKBOX antiguo
        elif checkbox_antiguo is not None:
            tipo = "checkbox"
            val = checkbox_antiguo.get(f"{{{NS_ALL['w']}}}checked")
            if val is not None:
                checked = val == "1"
            valores_cc = [True, False]

        # FECHA
        elif date_node is not None:
            tipo = "date"

        # PICTURE
        elif picture is not None:
            tipo = "picture"

        # Repeating section
        elif repeating_section is not None:
            tipo = "repeatingSection"

        elif repeating_item is not None:
            tipo = "repeatingSectionItem"

        # Rich text
        elif richtext is not None:
            tipo = "richText"

        # TEXTO
        else:
            tipo = "text"

        # Inferencia lógica
        if alias and isinstance(alias, str) and "fecha" in alias.lower():
            inferred_type = "date"
        elif tipo == "dropDownList":
            inferred_type = "dropdown"
        elif tipo == "checkbox":
            inferred_type = "bool"
        else:
            inferred_type = "text"

        # Extraer texto actual dentro del SDT (balanceado: solo texto plano)
        texto_actual = []
        for t in sdt.findall(".//w:sdtContent//w:t", NS_ALL):
            if t.text:
                texto_actual.append(t.text)
        texto_actual = "".join(texto_actual).strip() if texto_actual else None

        controles.append(
            {
                "alias": alias,
                "tag": tag,
                "id": cid,
                "tipo": tipo,
                "valores": valores_cc or None,
                "checked": checked if tipo == "checkbox" else None,
                "inferred_type": inferred_type,
                "texto_actual": texto_actual,
            }
        )

    return controles


def extract_content_controls_document(docx_path):
    """
    Extrae TODOS los controles de contenido solo de document.xml
    (compatibilidad directa con los 'controles' que usabas antes).
    """
    with zipfile.ZipFile(docx_path) as z:
        xml_files = {
            "document": read_xml_from_docx(z, "word/document.xml"),
            "styles": read_xml_from_docx(z, "word/styles.xml"),
            "settings": read_xml_from_docx(z, "word/settings.xml"),
            "glossary": read_xml_from_docx(z, "word/glossary/document.xml"),
            "numbering": read_xml_from_docx(z, "word/numbering.xml"),
        }

        valores_globales = []
        for _, xml in xml_files.items():
            vals = extract_list_entries(xml)
            valores_globales.extend(vals)
        valores_globales = list(set(valores_globales))

        root_doc = ET.fromstring(xml_files["document"])
        return _extract_sdt_list_from_root(root_doc, valores_globales)


def extract_content_controls_headers_footers(docx_path):
    """
    Extrae controles de contenido desde todos los headers y footers.
    Devuelve:
        (controles_header, controles_footer)
    """
    controles_header = []
    controles_footer = []

    with zipfile.ZipFile(docx_path) as z:
        # Reutilizamos valores_globales por simplicidad
        xml_global_candidates = [
            read_xml_from_docx(z, "word/styles.xml"),
            read_xml_from_docx(z, "word/settings.xml"),
            read_xml_from_docx(z, "word/numbering.xml"),
            read_xml_from_docx(z, "word/glossary/document.xml"),
        ]
        valores_globales = []
        for xml in xml_global_candidates:
            vals = extract_list_entries(xml)
            valores_globales.extend(vals)
        valores_globales = list(set(valores_globales))

        for name in z.namelist():
            if name.startswith("word/header") and name.endswith(".xml"):
                xml = read_xml_from_docx(z, name)
                if not xml:
                    continue
                try:
                    root = ET.fromstring(xml)
                except Exception:
                    continue
                sdt_list = _extract_sdt_list_from_root(root, valores_globales)
                for c in sdt_list:
                    c["origen"] = os.path.basename(name)
                controles_header.extend(sdt_list)

            if name.startswith("word/footer") and name.endswith(".xml"):
                xml = read_xml_from_docx(z, name)
                if not xml:
                    continue
                try:
                    root = ET.fromstring(xml)
                except Exception:
                    continue
                sdt_list = _extract_sdt_list_from_root(root, valores_globales)
                for c in sdt_list:
                    c["origen"] = os.path.basename(name)
                controles_footer.extend(sdt_list)

    return controles_header, controles_footer


# ==========================================================
# TABLAS NATIVAS WORD (nivel balanceado)
# ==========================================================
def extract_word_tables(docx_path):
    """
    Extrae tablas nativas de Word desde document.xml (nivel balanceado):
    - índice de tabla
    - filas y celdas con texto simple
    - sin estilos detallados (para no inflar el JSON)
    """
    tablas = []
    with zipfile.ZipFile(docx_path) as z:
        xml_doc = read_xml_from_docx(z, "word/document.xml")
        if not xml_doc:
            return tablas

        root = ET.fromstring(xml_doc)
        for i, tbl in enumerate(root.findall(".//w:tbl", NS_ALL), start=1):
            filas = []
            for r in tbl.findall("w:tr", NS_ALL):
                celdas = []
                for c in r.findall("w:tc", NS_ALL):
                    textos = []
                    for t in c.findall(".//w:t", NS_ALL):
                        if t.text:
                            textos.append(t.text)
                    celdas.append("".join(textos).strip())
                filas.append(celdas)
            tablas.append(
                {
                    "index": i,
                    "filas": filas,
                    "origen": "document",
                }
            )
    return tablas


# ==========================================================
# EXTRACCIÓN DE EXCEL EMBEBIDOS (igualado a SDI / Reporte)
# ==========================================================
def extract_embedded_excel(docx_path):
    """
    Extrae todos los Excel embebidos (word/embeddings/*.xlsx)
    y devuelve para cada uno:
      - excel: ruta interna dentro del docx
      - nombres_definidos: lista de {nombre, ref, valor}
      - tablas: lista de tablas con:
            worksheet, tabla, rango, columnas, registros, cell_styles
    Totalmente compatible con los scripts existentes.
    """
    results = []

    with zipfile.ZipFile(docx_path) as z:
        excel_files = [
            f
            for f in z.namelist()
            if f.startswith("word/embeddings/") and f.endswith(".xlsx")
        ]

        for ef in excel_files:
            content = z.read(ef)

            try:
                wb = load_workbook(BytesIO(content), data_only=False)
            except Exception:
                # Si por alguna razón el embedded no es Excel válido, se salta
                continue

            # Leemos XMLs internos (hojas + styles.xml)
            sheet_xmls = {}
            styles_xml_text = None
            with zipfile.ZipFile(BytesIO(content)) as xz:
                for item in xz.namelist():
                    if item.startswith("xl/worksheets/sheet") and item.endswith(".xml"):
                        key = item.split("/")[-1].replace(".xml", "")  # sheet1, sheet2...
                        sheet_xmls[key] = xz.read(item).decode("utf-8", "ignore")
                    if item == "xl/styles.xml":
                        styles_xml_text = xz.read(item).decode("utf-8", "ignore")

            styles_resolver = StylesResolver(styles_xml_text)

            excel_data = {"excel": ef, "tablas": [], "nombres_definidos": []}

            # Nombres definidos
            for name, defn in wb.defined_names.items():
                for ws_name, cell_ref in list(defn.destinations):
                    ws = wb[ws_name.replace("'", "")]
                    xml = sheet_xmls.get(f"sheet{ws._id}", "")

                    try:
                        celdas = ws[cell_ref]
                        if hasattr(celdas, "value"):
                            val = get_cell_formula_or_value(celdas, xml)
                        else:
                            val = [
                                [get_cell_formula_or_value(c, xml) for c in fila]
                                for fila in celdas
                            ]
                    except Exception:
                        val = None

                    excel_data["nombres_definidos"].append(
                        {"nombre": name, "ref": f"{ws_name}!{cell_ref}", "valor": val}
                    )

            # Tablas (ListObjects)
            for ws in wb.worksheets:
                xml = sheet_xmls.get(f"sheet{ws._id}", "")
                for tbl in ws._tables.values():
                    ref = tbl.ref
                    rows = []
                    cell_styles = {}

                    excel_range = ws[ref]

                    # Primera fila (encabezados) → estilos completos celda por celda
                    if excel_range:
                        first_row = excel_range[0]
                        for c in first_row:
                            cell_styles[c.coordinate] = get_complete_cell_format(
                                c, sheet_xml=xml, styles_resolver=styles_resolver
                            )

                    # Valores de toda la tabla
                    for row in excel_range:
                        rows.append([get_cell_formula_or_value(c, xml) for c in row])

                    excel_data["tablas"].append(
                        {
                            "worksheet": ws.title,
                            "tabla": tbl.name,
                            "rango": ref,
                            "columnas": rows[0] if rows else [],
                            "registros": rows[1:] if len(rows) > 1 else [],
                            "cell_styles": cell_styles,
                        }
                    )

            results.append(excel_data)

    return results


# ==========================================================
# IMÁGENES (nivel balanceado)
# ==========================================================
def extract_images_info(docx_path):
    """
    Extrae información básica de las imágenes (word/media/*)
    Nivel balanceado: solo nombre interno y tipo (según [Content_Types].xml).
    """
    imagenes = []
    with zipfile.ZipFile(docx_path) as z:
        # Mapeo de tipo por PartName en [Content_Types].xml
        content_types_xml = read_xml_from_docx(z, "[Content_Types].xml")
        type_map = {}
        if content_types_xml:
            root_ct = ET.fromstring(content_types_xml)
            for override in root_ct.findall(
                ".//{http://schemas.openxmlformats.org/package/2006/content-types}Override"
            ):
                part_name = override.get("PartName")
                ctype = override.get("ContentType")
                if part_name and ctype and "/media/" in part_name:
                    # PartName viene con slash inicial, ej: /word/media/image1.png
                    type_map[part_name] = ctype

        for name in z.namelist():
            if name.startswith("word/media/"):
                part_name = "/" + name  # para matchear con PartName
                imagenes.append(
                    {
                        "nombre": os.path.basename(name),
                        "ruta": name,
                        "content_type": type_map.get(part_name),
                    }
                )

    return imagenes


# ==========================================================
# CAMPOS WORD (MERGEFIELD, etc.) – nivel balanceado
# ==========================================================
def extract_word_fields(docx_path):
    """
    Extrae campos de Word (nivel balanceado):
    - fldSimple/@w:instr
    - bloques w:instrText
    Solo desde document.xml (es lo más habitual).
    """
    campos = []
    with zipfile.ZipFile(docx_path) as z:
        xml_doc = read_xml_from_docx(z, "word/document.xml")
        if not xml_doc:
            return campos

        root = ET.fromstring(xml_doc)

        # Campos simples <w:fldSimple w:instr="MERGEFIELD ...">
        for fld in root.findall(".//w:fldSimple", NS_ALL):
            instr = fld.get(f"{{{NS_ALL['w']}}}instr")
            textos = []
            for t in fld.findall(".//w:t", NS_ALL):
                if t.text:
                    textos.append(t.text)
            campos.append(
                {
                    "tipo": "fldSimple",
                    "instr": instr,
                    "texto": "".join(textos).strip() if textos else None,
                }
            )

        # Instrucciones libres <w:instrText> (parte de campos más complejos)
        for instr in root.findall(".//w:instrText", NS_ALL):
            txt = instr.text or ""
            txt = txt.strip()
            if txt:
                campos.append({"tipo": "instrText", "instr": txt})

    return campos


# ==========================================================
# PÁRRAFOS (nivel balanceado)
# ==========================================================
def extract_paragraphs(docx_path, max_chars_por_parrafo=500):
    """
    Extrae párrafos de document.xml en nivel balanceado:
    - índice de párrafo
    - estilo (pStyle)
    - texto truncado a max_chars_por_parrafo
    """
    parrafos = []
    with zipfile.ZipFile(docx_path) as z:
        xml_doc = read_xml_from_docx(z, "word/document.xml")
        if not xml_doc:
            return parrafos

        root = ET.fromstring(xml_doc)

        for i, p in enumerate(root.findall(".//w:p", NS_ALL), start=1):
            ppr = p.find("w:pPr", NS_ALL)
            pstyle = None
            if ppr is not None:
                st = ppr.find("w:pStyle", NS_ALL)
                if st is not None:
                    pstyle = st.get(f"{{{NS_ALL['w']}}}val")

            textos = []
            for t in p.findall(".//w:t", NS_ALL):
                if t.text:
                    textos.append(t.text)
            txt = "".join(textos).strip()
            if len(txt) > max_chars_por_parrafo:
                txt = txt[: max_chars_por_parrafo] + "..."

            if txt:
                parrafos.append({"index": i, "estilo": pstyle, "texto": txt})

    return parrafos


# ==========================================================
# GENERAR JSON COMPLETO (UNIVERSAL)
# ==========================================================

def json_safe(obj):
    """
    Convierte objetos no serializables (datetime, date, bytes) a formatos seguros.
    """
    import datetime

    if isinstance(obj, (datetime.datetime, datetime.date)):
        return obj.isoformat()

    if isinstance(obj, bytes):
        return obj.decode("utf-8", errors="ignore")

    # fallback por defecto
    return str(obj)


def generar_json_estructura(docx_path, json_output=None, detalle_extendido=False):
    """
    Genera un JSON de estructura balanceado, compatible con SDI / Reporte Diario.
    - 'controles'           → controles de document.xml (igual que antes)
    - 'controles_header'    → controles en headers
    - 'controles_footer'    → controles en footers
    - 'tablas_word'         → tablas nativas de Word (document.xml)
    - 'excels'              → Excel embebidos (estructura igual que SDI)
    - 'imagenes'            → info básica de imágenes
    - 'campos_word'         → campos (MERGEFIELD, etc.)
    - 'parrafos'            → párrafos con texto y estilo
    detalle_extendido se deja reservado para futuras ampliaciones,
    pero en la versión balanceada actual no cambia el contenido.
    """
    controles_doc = extract_content_controls_document(docx_path)
    controles_header, controles_footer = extract_content_controls_headers_footers(docx_path)
    tablas_word = extract_word_tables(docx_path)
    excels = extract_embedded_excel(docx_path)
    imagenes = extract_images_info(docx_path)
    campos = extract_word_fields(docx_path)
    parrafos = extract_paragraphs(docx_path)

    estructura = {
        "plantilla": docx_path,
        "controles": controles_doc,
        "controles_header": controles_header,
        "controles_footer": controles_footer,
        "tablas_word": tablas_word,
        "excels": excels,
        "imagenes": imagenes,
        "campos_word": campos,
        "parrafos": parrafos,
    }

    if json_output is None:
        base = os.path.splitext(os.path.basename(docx_path))[0]
        json_output = f"estructura_{base.lower()}.json"

    with open(json_output, "w", encoding="utf-8") as f:
        json.dump(estructura, f, indent=4, ensure_ascii=False, default=json_safe)

    print(f"[OK] JSON generado correctamente → {json_output}")
    return estructura


# ==========================================================
# MAIN de prueba rápida
# ==========================================================
if __name__ == "__main__":
    # Ejemplo de uso directo:
    #   python leer_universal.py SDI.docx
    #   python leer_universal.py REPORTE.docx salida.json
    import sys

    if len(sys.argv) < 2:
        print("Uso: python leer_universal.py <plantilla.docx> [salida.json]")
        sys.exit(1)

    docx_in = sys.argv[1]
    json_out = sys.argv[2] if len(sys.argv) >= 3 else None

    generar_json_estructura(docx_in, json_out)
