"""
Genera el Word (.docx) de un GBS: titulo + una tabla de 3 columnas
Goals | Barriers | Solutions, una fila por punto identificado. Lo usan (via
Bash) las secciones GENERAR de los skills "gbs-deal-maker" y "gbs-
organization-maker" -- el contenido de cada fila lo redacta Claude a partir
de los comentarios/GBS reales de vTiger (nunca inventado); este script solo
se encarga del formato del documento, no decide el contenido.

Arma el .docx a mano (OOXML minimo, sin pasar por la plantilla por defecto
de python-docx) en vez de con Document() de python-docx: la plantilla por
defecto de python-docx arrastra ~800 KB de estilos/efectos que no usamos
(word/styles.xml y word/stylesWithEffects.xml), y eso infla un documento de
4-5 filas a ~37 KB -- suficiente para que su version en base64 (~50.000
caracteres) no entre completa en una sola lectura al momento de subirlo a
SharePoint (confirmado el 28-ago-2026: se corta a los ~22.500 caracteres).

CORRECCION (28-ago-2026, tras el primer intento): la primera version de
este script omitia w:tblGrid (obligatorio por el esquema ECMA-376 para toda
w:tbl) y w:sectPr -- python-docx pudo volver a leer esos archivos igual
porque su lector es permisivo, pero Word los mostraba dañados/ilegibles de
verdad. Esta version agrega tblGrid con el ancho de columna, un sectPr
minimo de pagina carta, y docProps/core.xml + docProps/app.xml (lo que
cualquier .docx real de Word trae) mantenidos deliberadamente, mientras
sigue sin los archivos pesados que no hacen falta (stylesWithEffects.xml,
theme, fontTable, numbering, customXml, thumbnail). El resultado pesa
~3-4 KB, todavia chico de sobra para subirlo sin cortes, pero ahora
estructuralmente valido.

Uso: python generar_gbs_docx.py --input <ruta.json> --output <ruta.docx>

Esquema del JSON de entrada:
{
  "titulo": "GBS - <nombre>",
  "organizacion": "<organizacion>",
  "referencia": "<nombre del deal, o null si es a nivel de organizacion>",
  "filas": [{"goal": "...", "barrier": "...", "solution": "...",
             "departamento": "<opcional, solo GBS Organization Maker>"}, ...]
}

"departamento" es opcional: cuando una fila lo trae y es distinto del de la
fila anterior, se inserta antes una fila de encabezado (celda fusionada de
las 3 columnas) con ese nombre, para separar visualmente el resumen global
por departamento -- la tabla en si sigue teniendo 3 columnas de datos, el
departamento no agrega una columna nueva. gbs-deal-maker no usa esta llave y
sale exactamente igual que antes (una fila por pain point, sin encabezados).
"""
import argparse
import json
import zipfile
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape as _xml_escape

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Ancho de columna en twips (1/20 de punto; 1 pulgada = 1440 twips). 3 x
# 3020 = ~6.3", calza con margenes de 1.1" de cada lado en carta (8.5").
COL_WIDTH = 3020

CONTENT_TYPES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/docProps/core.xml" '
    'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    '<Override PartName="/docProps/app.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    "</Types>"
)

RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/>'
    '<Relationship Id="rId2" '
    'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
    'Target="docProps/core.xml"/>'
    '<Relationship Id="rId3" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
    'Target="docProps/app.xml"/>'
    "</Relationships>"
)

DOCUMENT_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
    'Target="styles.xml"/>'
    "</Relationships>"
)

STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:styles xmlns:w="{W_NS}">'
    '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
    '<w:qFormat/></w:style>'
    "</w:styles>"
)


def _core_xml():
    now = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties '
        'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:creator>GB Advisors Panel</dc:creator>"
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
        "</cp:coreProperties>"
    )


APP_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    "<Application>GB Advisors GBS Maker</Application>"
    "</Properties>"
)

TABLE_BORDERS = (
    '<w:tblBorders>'
    '<w:top w:val="single" w:sz="4" w:color="auto"/>'
    '<w:left w:val="single" w:sz="4" w:color="auto"/>'
    '<w:bottom w:val="single" w:sz="4" w:color="auto"/>'
    '<w:right w:val="single" w:sz="4" w:color="auto"/>'
    '<w:insideH w:val="single" w:sz="4" w:color="auto"/>'
    '<w:insideV w:val="single" w:sz="4" w:color="auto"/>'
    '</w:tblBorders>'
)

TBL_GRID = (
    "<w:tblGrid>"
    f'<w:gridCol w:w="{COL_WIDTH}"/>'
    f'<w:gridCol w:w="{COL_WIDTH}"/>'
    f'<w:gridCol w:w="{COL_WIDTH}"/>'
    "</w:tblGrid>"
)

SECT_PR = (
    "<w:sectPr>"
    '<w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" '
    'w:header="720" w:footer="720" w:gutter="0"/>'
    "</w:sectPr>"
)


def esc(v):
    return _xml_escape(str(v or ""))


def run(text, bold=False, italic=False, size=None):
    props = ""
    if bold:
        props += "<w:b/>"
    if italic:
        props += "<w:i/>"
    if size:
        props += f'<w:sz w:val="{size}"/>'
    rpr = f"<w:rPr>{props}</w:rPr>" if props else ""
    return f"<w:r>{rpr}<w:t xml:space=\"preserve\">{esc(text)}</w:t></w:r>"


def paragraph(text, bold=False, italic=False, size=None):
    return f"<w:p>{run(text, bold, italic, size)}</w:p>"


def cell(text, bold=False):
    return (
        f'<w:tc><w:tcPr><w:tcW w:w="{COL_WIDTH}" w:type="dxa"/></w:tcPr>'
        f"{paragraph(text, bold=bold)}</w:tc>"
    )


def merged_cell_row(text):
    """Fila de encabezado de departamento: una sola celda visual que ocupa
    las 3 columnas (gridSpan), en vez de agregar una columna nueva a la
    tabla."""
    ancho_total = COL_WIDTH * 3
    tc = (
        f'<w:tc><w:tcPr><w:tcW w:w="{ancho_total}" w:type="dxa"/>'
        '<w:gridSpan w:val="3"/></w:tcPr>'
        f"{paragraph(text, bold=True)}</w:tc>"
    )
    return f"<w:tr>{tc}</w:tr>"


def data_row(goal, barrier, solution):
    return f"<w:tr>{cell(goal)}{cell(barrier)}{cell(solution)}</w:tr>"


def construir_document_xml(datos):
    titulo = datos.get("titulo") or "GBS"
    partes_meta = []
    if datos.get("organizacion"):
        partes_meta.append(f"Organization: {datos['organizacion']}")
    if datos.get("referencia"):
        partes_meta.append(f"Deal: {datos['referencia']}")
    partes_meta.append(f"Generated: {datetime.now():%Y-%m-%d}")
    meta_texto = " | ".join(partes_meta)

    filas = datos.get("filas") or []

    filas_xml = []
    header_row = (
        "<w:tr>"
        + cell("Goals", bold=True)
        + cell("Barriers", bold=True)
        + cell("Solutions", bold=True)
        + "</w:tr>"
    )
    filas_xml.append(header_row)

    if not filas:
        filas_xml.append(data_row("Not documented yet", "Not documented yet", "Not documented yet"))
    else:
        departamento_previo = None
        for f in filas:
            departamento = f.get("departamento")
            if departamento and departamento != departamento_previo:
                filas_xml.append(merged_cell_row(departamento))
                departamento_previo = departamento
            filas_xml.append(data_row(f.get("goal"), f.get("barrier"), f.get("solution")))

    tabla_xml = (
        "<w:tbl>"
        f"<w:tblPr>{TABLE_BORDERS}"
        f'<w:tblW w:w="{COL_WIDTH * 3}" w:type="dxa"/>'
        "</w:tblPr>"
        + TBL_GRID
        + "".join(filas_xml)
        + "</w:tbl>"
    )

    body = (
        paragraph(titulo, bold=True, size="32")
        + paragraph(meta_texto, italic=True, size="18")
        + tabla_xml
        + "<w:p/>"
        + SECT_PR
    )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}"><w:body>{body}</w:body></w:document>'
    )


def construir_docx(datos, salida):
    salida = Path(salida)
    salida.parent.mkdir(parents=True, exist_ok=True)
    document_xml = construir_document_xml(datos)
    with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        z.writestr("_rels/.rels", RELS_XML)
        z.writestr("docProps/core.xml", _core_xml())
        z.writestr("docProps/app.xml", APP_XML)
        z.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS_XML)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/document.xml", document_xml)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", required=True)
    args = ap.parse_args()

    with open(args.input, "r", encoding="utf-8") as f:
        datos = json.load(f)

    construir_docx(datos, args.output)
    print(f"Guardado: {args.output}")


if __name__ == "__main__":
    main()
