from zipfile import ZipFile
from lxml import etree
from datetime import datetime

path_in = r"C:\Users\jonat\Documents\gestion_docs\plantillas_documentos_tecnicos\templates\documentos\plantilla_portada.docx"
path_out = r"C:\Users\jonat\Documents\gestion_docs\plantillas_documentos_tecnicos\templates\documentos\plantilla_portada.docx"

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
parser = etree.XMLParser(remove_blank_text=False)

# ------------------------------
# 🎯 DATOS SIMPLES
# ------------------------------
simple_data = {
    "tipo_documento": "Informe Técnico",
    "codigo_documento": "IT-4455",
    "nombre_proyecto": "Proyecto Aguas Claras",
    "nombre_cliente": "Metso Chile SPA",
    "nombre_faena": "Faena Los Bronces",
    "numero_contrato": "CTR-2025-88",
    "numero_servicio": "NS-12345",
    "administrador_servicio": "Carlos Medina",
}

# ------------------------------
# 🎯 GRUPOS (Redactores, Revisores, Aprobadores)
# ------------------------------
equipos = {
    "redactores_equipo": "J. Araya",
    "revisores_equipo": "M. Soto",
    "aprobadores_equipo": "A. Vidal",
}

simple_data.update(equipos)

# ------------------------------
# 🎯 HISTORIAL DE VERSIONES
# ------------------------------
historial_versiones = [
    {
        "h.version": "V00",
        "h.estado": "Revisión",
        "h.fecha": "15-11-2025 10:20",
        "h.comentario": "Revisión inicial",
    },
    {
        "h.version": "V01",
        "h.estado": "Aprobado",
        "h.fecha": "22-11-2025 15:40",
        "h.comentario": "Cambios aceptados",
    }
]

# ------------------------------
# 🎯 CAMPOS CALCULADOS
# ------------------------------
def compute_extra_fields():
    # Extraer la parte numérica
    rev_nums = [int(v["h.version"].replace("V", "")) for v in historial_versiones]

    revision_actual = f"{max(rev_nums):02d}"

    fecha_ultima = max(
        datetime.strptime(v["h.fecha"], "%d-%m-%Y %H:%M")
        for v in historial_versiones
    ).strftime("%d/%m/%Y")

    dias_duracion = (datetime.now() - 
                     datetime.strptime(historial_versiones[0]["h.fecha"], "%d-%m-%Y %H:%M")).days

    return {
        "revision_actual": revision_actual,
        "fecha_ultima_revision": fecha_ultima,
        "dias_duracion_revision": str(dias_duracion)
    }

simple_data.update(compute_extra_fields())

# ------------------------------
# 🟣 FUNCIONES DE REEMPLAZO
# ------------------------------

def set_text_clean(cc, value):
    # Eliminar todos los w:t
    for t in cc.xpath('.//w:t', namespaces=ns):
        t.getparent().remove(t)

    # Crear nuevo w:t
    r_tags = cc.xpath('.//w:r', namespaces=ns)
    r = r_tags[0] if r_tags else etree.SubElement(cc, f"{{{ns['w']}}}r")
    t = etree.SubElement(r, f"{{{ns['w']}}}t")
    t.text = value


# ------------------------------
# 🟣 REEMPLAZAR CONTROL DE CONTENIDO SIMPLE
# ------------------------------
def replace_simple_fields(root):
    for cc in root.xpath('.//w:sdt', namespaces=ns):
        alias = cc.xpath('.//w:alias', namespaces=ns)
        if not alias:
            continue

        name = alias[0].get(f"{{{ns['w']}}}val")

        if name in simple_data:
            set_text_clean(cc, simple_data[name])


# ------------------------------
# 🟧 HISTORIAL DE VERSIONES (Tabla dinámica)
# ------------------------------
def fill_historial(root):
    template_cc = root.xpath('.//w:sdt[.//w:alias[@w:val="h.version"]]', namespaces=ns)
    if not template_cc:
        return

    template_row = template_cc[0].xpath('./ancestor::w:tr', namespaces=ns)[0]
    table = template_row.getparent()

    insert_index = table.index(template_row) + 1

    for i, version in enumerate(historial_versiones):
        if i == 0:
            row = template_row
        else:
            row = etree.fromstring(etree.tostring(template_row))
            table.insert(insert_index, row)
            insert_index += 1

        for cc in row.xpath('.//w:sdt', namespaces=ns):
            alias = cc.xpath('.//w:alias', namespaces=ns)
            if alias:
                name = alias[0].get(f"{{{ns['w']}}}val")
                if name in version:
                    set_text_clean(cc, version[name])


# ------------------------------
# 🟦 PROCESAR XML
# ------------------------------
def process_xml(xml_bytes):
    root = etree.fromstring(xml_bytes, parser)

    replace_simple_fields(root)
    fill_historial(root)

    return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes")


# ------------------------------
# 🧱 CONSTRUIR DOCX FINAL
# ------------------------------
with ZipFile(path_in) as original:
    filelist = original.infolist()

    new_xml = {}

    for item in filelist:
        xml_bytes = original.read(item.filename)

        if item.filename.startswith("word/document") or \
           item.filename.startswith("word/header") or \
           item.filename.startswith("word/footer"):

            new_xml[item.filename] = process_xml(xml_bytes)
        else:
            new_xml[item.filename] = xml_bytes

with ZipFile(path_out, "w") as newdoc:
    for name, data_bytes in new_xml.items():
        newdoc.writestr(name, data_bytes)

print("✅ Documento generado con la NUEVA plantilla:", path_out)
