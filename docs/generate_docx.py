"""
Genera entrega2_semana5.docx a partir de entrega2_semana5.md
Ejecutar: python docs/generate_docx.py
"""
import re
import pathlib
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE     = pathlib.Path(__file__).parent
MD_FILE  = BASE / "entrega2_semana5.md"
OUT_FILE = BASE / "entrega2_semana5.docx"

# ── Colores ────────────────────────────────────────────────────────────────────
AZUL_OSCURO  = RGBColor(0x1D, 0x4E, 0xD8)   # #1D4ED8
AZUL_CLARO   = RGBColor(0x93, 0xC5, 0xFD)   # encabezados tabla
GRIS_FONDO   = RGBColor(0xF1, 0xF5, 0xF9)   # fondo código inline
NEGRO        = RGBColor(0x0F, 0x17, 0x2A)
BLANCO       = RGBColor(0xFF, 0xFF, 0xFF)
VERDE        = RGBColor(0x16, 0xA3, 0x4A)
ROJO         = RGBColor(0xDC, 0x26, 0x26)
GRIS_TEXTO   = RGBColor(0x64, 0x74, 0x8B)
AMARILLO_BG  = RGBColor(0xFF, 0xFB, 0xEB)
NARANJA      = RGBColor(0xD9, 0x77, 0x06)
CODIGO_BG    = RGBColor(0x0F, 0x17, 0x2A)   # fondo bloque código
CODIGO_FG    = RGBColor(0xE2, 0xE8, 0xF0)   # texto bloque código


def set_cell_bg(cell, color: RGBColor):
    """Aplica color de fondo a una celda de tabla."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    hex_color = f"{color[0]:02X}{color[1]:02X}{color[2]:02X}"
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def set_para_border_left(para, color: RGBColor, size_pt=12):
    """Agrega borde izquierdo a un párrafo (blockquote)."""
    pPr  = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"),   "single")
    left.set(qn("w:sz"),    str(size_pt * 8))
    left.set(qn("w:space"), "10")
    left.set(qn("w:color"), f"{color[0]:02X}{color[1]:02X}{color[2]:02X}")
    pBdr.append(left)
    pPr.append(pBdr)


def add_inline_text(run_parent, text: str, base_bold=False, base_size=10.5):
    """
    Parsea markdown inline (**bold**, `code`) y agrega runs al párrafo.
    run_parent es un objeto paragraph de python-docx.
    """
    # Patrón: **bold** o `code`
    pattern = re.compile(r'\*\*(.+?)\*\*|`([^`]+)`')
    pos = 0
    for m in pattern.finditer(text):
        # texto antes del match
        if m.start() > pos:
            r = run_parent.add_run(text[pos:m.start()])
            r.bold = base_bold
            r.font.size = Pt(base_size)
        if m.group(1) is not None:          # **bold**
            r = run_parent.add_run(m.group(1))
            r.bold = True
            r.font.size = Pt(base_size)
            r.font.color.rgb = NEGRO
        else:                               # `code`
            r = run_parent.add_run(m.group(2))
            r.font.name  = "Courier New"
            r.font.size  = Pt(9)
            r.font.color.rgb = RGBColor(0xBE, 0x18, 0x5D)
        pos = m.end()
    # resto
    if pos < len(text):
        r = run_parent.add_run(text[pos:])
        r.bold = base_bold
        r.font.size = Pt(base_size)


def configure_styles(doc: Document):
    """Configura estilos base del documento."""
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10.5)
    style.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)

    # Márgenes de página
    for section in doc.sections:
        section.top_margin    = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin   = Cm(3.0)
        section.right_margin  = Cm(2.5)


def add_heading(doc, text, level):
    """Agrega un heading con estilo personalizado."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16 if level == 1 else 12)
    p.paragraph_format.space_after  = Pt(6)
    p.paragraph_format.keep_with_next = True

    run = p.add_run(text)
    run.bold = True

    if level == 1:
        run.font.size  = Pt(18)
        run.font.color.rgb = NEGRO
        # Borde inferior
        pPr  = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bot  = OxmlElement("w:bottom")
        bot.set(qn("w:val"),   "single")
        bot.set(qn("w:sz"),    "12")
        bot.set(qn("w:space"), "4")
        bot.set(qn("w:color"), "1D4ED8")
        pBdr.append(bot)
        pPr.append(pBdr)
    elif level == 2:
        run.font.size  = Pt(14)
        run.font.color.rgb = AZUL_OSCURO
        set_para_border_left(p, AZUL_OSCURO, size_pt=6)
        p.paragraph_format.left_indent = Cm(0.4)
    elif level == 3:
        run.font.size  = Pt(12)
        run.font.color.rgb = RGBColor(0x1E, 0x40, 0xAF)
    elif level == 4:
        run.font.size  = Pt(11)
        run.font.color.rgb = RGBColor(0x33, 0x41, 0x55)


def add_code_block(doc, lines):
    """Agrega un bloque de código con fondo oscuro."""
    code_text = "\n".join(lines)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(6)
    p.paragraph_format.left_indent  = Cm(0.5)
    p.paragraph_format.right_indent = Cm(0.5)

    # Fondo oscuro via shading en el párrafo
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  "0F172A")
    pPr.append(shd)

    r = p.add_run(code_text)
    r.font.name  = "Courier New"
    r.font.size  = Pt(8)
    r.font.color.rgb = CODIGO_FG


def add_table(doc, header_row, data_rows):
    """Agrega tabla con encabezado azul y filas alternadas."""
    cols = len(header_row)
    table = doc.add_table(rows=1, cols=cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Encabezado
    hdr_cells = table.rows[0].cells
    for i, text in enumerate(header_row):
        cell = hdr_cells[i]
        set_cell_bg(cell, AZUL_OSCURO)
        p   = cell.paragraphs[0]
        run = p.add_run(text.strip())
        run.bold = True
        run.font.color.rgb = BLANCO
        run.font.size = Pt(9.5)

    # Filas de datos
    for row_idx, row_data in enumerate(data_rows):
        row_cells = table.add_row().cells
        bg = GRIS_FONDO if row_idx % 2 == 0 else BLANCO
        for i, text in enumerate(row_data):
            cell = row_cells[i]
            set_cell_bg(cell, bg)
            p = cell.paragraphs[0]
            # Detectar SUCCESS/FAILURE para colorear
            clean = text.strip()
            if "SUCCESS" in clean or "exitoso" in clean.lower():
                r = p.add_run(clean)
                r.font.color.rgb = VERDE
                r.font.size = Pt(9.5)
            elif "FAILURE" in clean or "fallido" in clean.lower():
                r = p.add_run(clean)
                r.font.color.rgb = ROJO
                r.font.size = Pt(9.5)
            else:
                add_inline_text(p, clean, base_size=9.5)

    doc.add_paragraph()  # espacio después de tabla


def parse_table_row(line):
    """Parsea una fila de tabla markdown."""
    return [c.strip() for c in line.strip().strip("|").split("|")]


def is_separator_row(row):
    return all(re.match(r'^[-:]+$', c.strip()) for c in row if c.strip())


def add_blockquote(doc, text):
    """Agrega un blockquote con borde izquierdo amarillo."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent  = Cm(0.6)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)

    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  "FFFBEB")
    pPr.append(shd)

    set_para_border_left(p, NARANJA, size_pt=6)
    add_inline_text(p, text.lstrip("> ").strip(), base_size=9.5)
    for r in p.runs:
        r.font.color.rgb = RGBColor(0x78, 0x35, 0x0F)


# ── Parser principal ──────────────────────────────────────────────────────────
def parse_md_to_docx(md_path: pathlib.Path, out_path: pathlib.Path):
    doc = Document()
    configure_styles(doc)

    lines = md_path.read_text(encoding="utf-8").splitlines()

    i = 0
    table_buffer   = []
    code_buffer    = []
    in_code_block  = False
    list_stack     = []   # (indent, ordered)

    def flush_list():
        list_stack.clear()

    def flush_table():
        if len(table_buffer) >= 2:
            header = parse_table_row(table_buffer[0])
            rows   = []
            for ln in table_buffer[2:]:
                r = parse_table_row(ln)
                if r:
                    rows.append(r)
            add_table(doc, header, rows)
        table_buffer.clear()

    while i < len(lines):
        line = lines[i]

        # ── Código (``` ... ```) ──────────────────────────────────────────────
        if line.strip().startswith("```"):
            if in_code_block:
                add_code_block(doc, code_buffer)
                code_buffer.clear()
                in_code_block = False
            else:
                flush_list()
                in_code_block = True
            i += 1
            continue

        if in_code_block:
            code_buffer.append(line)
            i += 1
            continue

        # ── Tabla ─────────────────────────────────────────────────────────────
        if line.strip().startswith("|"):
            flush_list()
            table_buffer.append(line)
            i += 1
            continue
        elif table_buffer:
            flush_table()

        # ── Separador --- ─────────────────────────────────────────────────────
        if re.match(r'^-{3,}$', line.strip()):
            flush_list()
            doc.add_paragraph().paragraph_format.space_after = Pt(2)
            i += 1
            continue

        # ── Headings ──────────────────────────────────────────────────────────
        m = re.match(r'^(#{1,4})\s+(.*)', line)
        if m:
            flush_list()
            level = len(m.group(1))
            text  = m.group(2).strip()
            add_heading(doc, text, level)
            i += 1
            continue

        # ── Blockquote ────────────────────────────────────────────────────────
        if line.startswith(">"):
            flush_list()
            add_blockquote(doc, line)
            i += 1
            continue

        # ── Lista con viñetas (- o *) ─────────────────────────────────────────
        m = re.match(r'^(\s*)([-*])\s+(.*)', line)
        if m:
            indent  = len(m.group(1))
            content = m.group(3)
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.left_indent   = Cm(0.5 + indent * 0.4)
            p.paragraph_format.space_before  = Pt(2)
            p.paragraph_format.space_after   = Pt(2)
            add_inline_text(p, content, base_size=10.5)
            i += 1
            continue

        # ── Lista numerada (1. ...) ───────────────────────────────────────────
        m = re.match(r'^(\s*)\d+\.\s+(.*)', line)
        if m:
            content = m.group(2)
            p = doc.add_paragraph(style="List Number")
            p.paragraph_format.left_indent  = Cm(0.5)
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after  = Pt(2)
            add_inline_text(p, content, base_size=10.5)
            i += 1
            continue

        # ── Línea vacía ───────────────────────────────────────────────────────
        if line.strip() == "":
            flush_list()
            i += 1
            continue

        # ── Párrafo normal ────────────────────────────────────────────────────
        flush_list()
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = Pt(6)
        add_inline_text(p, line.strip(), base_size=10.5)
        i += 1

    # Flush final
    if table_buffer:
        flush_table()
    if code_buffer:
        add_code_block(doc, code_buffer)

    doc.save(out_path)
    print(f"Documento generado: {out_path}")
    print(f"Tamaño: {out_path.stat().st_size / 1024:.1f} KB")


if __name__ == "__main__":
    parse_md_to_docx(MD_FILE, OUT_FILE)
