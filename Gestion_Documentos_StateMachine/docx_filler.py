#C:\Users\jonat\Documents\gestion_docs\Gestion_Documentos_StateMachine\docx_filler.py
from zipfile import ZipFile
from lxml import etree
from datetime import datetime

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
parser = etree.XMLParser(remove_blank_text=False)


def set_text_clean(cc, value):
    for t in cc.xpath('.//w:t', namespaces=ns):
        t.getparent().remove(t)

    r_tags = cc.xpath('.//w:r', namespaces=ns)
    r = r_tags[0] if r_tags else etree.SubElement(cc, f"{{{ns['w']}}}r")
    t = etree.SubElement(r, f"{{{ns['w']}}}t")
    t.text = value


def replace_simple_fields(root, simple_data):
    for cc in root.xpath('.//w:sdt', namespaces=ns):
        alias = cc.xpath('.//w:alias', namespaces=ns)
        if not alias:
            continue

        name = alias[0].get(f"{{{ns['w']}}}val")

        if name in simple_data:
            set_text_clean(cc, simple_data[name])


def fill_historial(root, historial):
    template_cc = root.xpath('.//w:sdt[.//w:alias[@w:val="h.version"]]', namespaces=ns)
    if not template_cc:
        return

    template_row = template_cc[0].xpath('./ancestor::w:tr', namespaces=ns)[0]
    table = template_row.getparent()

    insert_index = table.index(template_row) + 1

    for i, version in enumerate(historial):
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


def process_template_docx(template_bytes, simple_data, historial_versiones):
    new_xml = {}

    with ZipFile(template_bytes) as original:
        filelist = original.infolist()

        for item in filelist:
            xml_bytes = original.read(item.filename)

            if item.filename.startswith("word/document") or \
               item.filename.startswith("word/header") or \
               item.filename.startswith("word/footer"):
                root = etree.fromstring(xml_bytes, parser)

                replace_simple_fields(root, simple_data)
                fill_historial(root, historial_versiones)

                new_xml[item.filename] = etree.tostring(
                    root,
                    xml_declaration=True,
                    encoding="utf-8",
                    standalone="yes"
                )
            else:
                new_xml[item.filename] = xml_bytes

    # Construir .docx final como bytes
    from io import BytesIO
    output = BytesIO()
    with ZipFile(output, "w") as newdoc:
        for name, data_bytes in new_xml.items():
            newdoc.writestr(name, data_bytes)

    return output.getvalue()
