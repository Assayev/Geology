"""Build documents on the CGC (Caspian Geology Center) letterhead.

Template = the client's own corrected Word file (assets/CGC_letterhead_template.docx).
Letterhead (logo, 3-language address table, dividers, 5% watermark) lives in its header,
page size / margins in its sectPr, body font/spacing in its Normal style
(Times New Roman 12, single line, 0 pt space between paragraphs). We only replace the body.

Layout conventions (taken from the client's corrected sample):
  * no "space after" - vertical gaps are made with empty paragraphs (blank(doc, n))
  * body text: justified, first-line indent 709 twips (1.25 cm), one blank line between paragraphs
  * recipient block: left indent 5812 twips (10.25 cm), centered inside that block
  * salutation and signature line: centered
  * tables: style "Table Grid", indent 704 twips, total width 8647 twips, directly under text

    from cgc_doc import *
    doc = new_doc()
    letter(doc, "Исх. № 12 от 29.09.2026 г.", ["Директору ТОО «X»", "Иванову И.И."],
           "Уважаемый Иван Иванович!", ["Текст 1", "Текст 2"],
           table_rows=[["№","Наименование"],["1","…"]], signer=("Директор", "Ф.И.О."))
    save(doc, "out.docx", pdf=True)
"""
import os, shutil, subprocess, zipfile, io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Twips

TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "CGC_letterhead_template.docx")
ALIGN = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
         "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}
RECIPIENT_INDENT = 5812   # twips
FIRST_LINE = 709          # twips
TABLE_INDENT, TABLE_WIDTH = 704, 8647


def new_doc():
    """Empty document on the letterhead: template header/section/styles, body cleared."""
    doc = Document(TEMPLATE)
    body = doc.element.body
    for el in list(body):
        if el.tag != qn("w:sectPr"):
            body.remove(el)
    return doc


def para(doc, text="", bold=False, italic=False, align="left", left=None, first=None):
    """One paragraph. left/first are indents in twips. Contextual spacing on, like the sample."""
    p = doc.add_paragraph()
    p.alignment = ALIGN[align]
    if left is not None: p.paragraph_format.left_indent = Twips(left)
    if first is not None: p.paragraph_format.first_line_indent = Twips(first)
    ppr = p._p.get_or_add_pPr()
    cs = OxmlElement("w:contextualSpacing"); ppr.append(cs)
    if text:
        r = p.add_run(text)
        if bold: r.bold = True
        if italic: r.italic = True
    return p


def blank(doc, n=1, **kw):
    for _ in range(n): para(doc, **kw)


def body_text(doc, text, **kw):
    return para(doc, text, align="justify", first=FIRST_LINE, **kw)


def table(doc, rows, widths=(567, 5102, 1134, 1844), header=True):
    """Grid table, indent 704, total width 8647 twips (widths must sum to 8647)."""
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    tblPr = t._tbl.tblPr
    for tag in ("w:tblW", "w:tblInd"):
        for e in tblPr.findall(qn(tag)): tblPr.remove(e)
    w = OxmlElement("w:tblW"); w.set(qn("w:w"), str(sum(widths))); w.set(qn("w:type"), "dxa"); tblPr.append(w)
    ind = OxmlElement("w:tblInd"); ind.set(qn("w:w"), str(TABLE_INDENT)); ind.set(qn("w:type"), "dxa"); tblPr.append(ind)
    for gc, wd in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths): gc.set(qn("w:w"), str(wd))
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = t.cell(i, j); c.width = Twips(widths[j])
            p = c.paragraphs[0]
            p._p.get_or_add_pPr().append(OxmlElement("w:contextualSpacing"))
            r = p.add_run(str(val))
            if header and i == 0: r.bold = True
    return t


def letter(doc, ref_line, recipient_lines, salutation, paragraphs,
           table_rows=None, table_widths=(567, 5102, 1134, 1844),
           subject=None, signer=("Директор", "Ф.И.О.")):
    """Standard CGC letter, spacing exactly as in the client's corrected sample."""
    para(doc, ref_line)                                   # Исх. № … от …
    blank(doc, 1); blank(doc, 2, align="center", left=RECIPIENT_INDENT)
    for ln in recipient_lines:                            # recipient block
        para(doc, ln, align="center", left=RECIPIENT_INDENT)
    blank(doc, 2, align="center", bold=True)
    if subject:                                           # optional bold centered subject
        para(doc, subject, align="center", bold=True); blank(doc, 1)
    para(doc, salutation, align="center")
    for i, t in enumerate(paragraphs):
        body_text(doc, t)
        if i < len(paragraphs) - 1: blank(doc, 1, align="justify", first=FIRST_LINE)
    if table_rows: table(doc, table_rows, table_widths)
    blank(doc, 1); blank(doc, 1, align="center")
    if signer:
        para(doc, f"{signer[0]}{' ' * 66}____________ / {signer[1]} /", align="center")


# ---------------------------------------------------------------- saving / PDF
def _restore_header(path):
    """python-docx re-serialises header parts; put back the template's original bytes."""
    keep = ("word/header1.xml", "word/_rels/header1.xml.rels")
    src = zipfile.ZipFile(TEMPLATE)
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            zout.writestr(it, src.read(it.filename) if it.filename in keep else zin.read(it.filename))
    shutil.move(tmp, path)


def _render_copy(path, tmp):
    """LibreOffice ignores <a:alphaModFix> (5% watermark) and draws 'outset' borders too
    thick. Word does both right, so the .docx stays untouched; only the copy fed to
    LibreOffice gets the watermark pre-faded and the dividers set to 0.75 pt."""
    from PIL import Image
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == "word/media/image1.png":
                im = Image.open(io.BytesIO(data)).convert("RGBA")
                im.putalpha(im.getchannel("A").point(lambda a: round(a * 0.05)))
                b = io.BytesIO(); im.save(b, "PNG"); data = b.getvalue()
            elif it.filename == "word/header1.xml":
                data = data.replace(b'<w:insideV w:val="outset" w:sz="6"', b'<w:insideV w:val="single" w:sz="6"')
                data = data.replace(b'<a:alphaModFix amt="5000"/>', b"")
            zout.writestr(it, data)


def save(doc, path, pdf=False):
    """Save docx (letterhead byte-identical to the template). pdf=True also renders a PDF
    next to it via LibreOffice (needs libreoffice-writer + fonts-liberation)."""
    doc.save(path)
    _restore_header(path)
    if pdf:
        import tempfile
        outdir = os.path.dirname(os.path.abspath(path))
        with tempfile.TemporaryDirectory() as td:
            tmp = os.path.join(td, os.path.basename(path))
            _render_copy(path, tmp)
            subprocess.run(["soffice", "--headless", "-env:UserInstallation=file:///tmp/lo_prof",
                            "--convert-to", "pdf", "--outdir", outdir, tmp],
                           check=True, capture_output=True)
    return path
