"""Ethics Certificate + VQD на бланке CGC (только английский)."""
import docx, copy, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
HERE=os.path.dirname(os.path.abspath(__file__))
SRC='/root/.claude/uploads/af002989-109f-5670-bf4b-111a3c9940db/'
ETH=SRC+'d293d741-Business_Ethics_Compliance_Certificate_Contractor_or_Suppliers_EN-RU.docx'
VQD=SRC+'0a6563b3-VQD_-_Vendor_Qualification_Undertaking_EN-RU_combined1.docx'
NAME='G. Zh. Shakhtayev'; TITLE='CEO'; COMPANY='Caspian Geology Center LLP'
DATE='29 September 2026'; PLACE='Astana, Kazakhstan'

def set_text(p,t):
    rs=p.runs; rs[0].text=t
    for r in rs[1:]: r._r.getparent().remove(r._r)
def rep(p,old,new):
    full=''.join(r.text for r in p.runs)
    if old in full: set_text(p,full.replace(old,new)); return True
def nob(tbl):
    b=OxmlElement('w:tblBorders')
    for e in ('top','left','bottom','right','insideH','insideV'):
        x=OxmlElement('w:'+e); x.set(qn('w:val'),'nil'); b.append(x)
    tbl._tbl.tblPr.append(b)
def letterhead(doc):
    from cgc_letterhead import apply_letterhead
    apply_letterhead(doc)

# ---------- Ethics certificate (EN only)
d=docx.Document(ETH); letterhead(d)
body=d.element.body
ru=[p for p in d.paragraphs if p.text.startswith('Русская версия')][0]
el=ru._p
while el is not None:                      # убрать русскую версию до sectPr
    nx=el.getnext()
    if not el.tag.endswith('sectPr'): body.remove(el)
    el=nx
for p in d.paragraphs:
    rep(p,'[name], [office or title] of [Contractor]',f'{NAME}, {TITLE} of {COMPANY}')
    rep(p,'[Contractor] (the “Contractor”)',f'{COMPANY} (the “Contractor”)')
    if p.text.startswith('Children, spouses'):
        rep(p,' On behalf of [Contractor]','')
        for br in p._p.findall('.//'+qn('w:br')): br.getparent().remove(br)
datep=[p for p in d.paragraphs if p.text.startswith('Date:')][0]
set_text(datep,f'Date: {DATE}')
rows=[(f'On behalf of {COMPANY}',''),('Signature:','…………………………………………'),('Name:',NAME),('Title:',TITLE),('Stamp:','')]
tbl=d.add_table(rows=len(rows),cols=2)
for r,(a,b) in zip(tbl.rows,rows):
    r.cells[0].text=a; r.cells[1].text=b
    for c in r.cells:
        for pp in c.paragraphs:
            pp.paragraph_format.space_after=Pt(6); pp.paragraph_format.keep_with_next=True
            for rr in pp.runs: rr.font.size=Pt(11)
    r.cells[0].width=Cm(6); r.cells[1].width=Cm(10)
for rr in tbl.rows[0].cells[0].paragraphs[0].runs: rr.bold=True
datep._p.addprevious(tbl._tbl)
datep.paragraph_format.keep_with_next=False
prev=tbl._tbl.getprevious()                 # убрать пустые абзацы перед блоком подписи
while prev is not None and prev.tag.endswith('}p') and not ''.join(prev.itertext()).strip():
    x=prev.getprevious(); prev.getparent().remove(prev); prev=x
# ужать пустые абзацы, чтобы текст и подпись уместились на одной странице
for p in d.paragraphs:
    if not p.text.strip() and not p._p.findall('.//'+qn('w:drawing')):
        p.paragraph_format.space_after=Pt(0); p.paragraph_format.space_before=Pt(0)
        for r in p.runs: r.font.size=Pt(4)
        pPr=p._p.get_or_add_pPr(); rp=OxmlElement('w:rPr'); sz=OxmlElement('w:sz'); sz.set(qn('w:val'),'8'); rp.append(sz); pPr.append(rp)
# пустые абзацы после Date
nx=datep._p.getnext()
while nx is not None and nx.tag.endswith('}p') and not ''.join(nx.itertext()).strip():
    n2=nx.getnext(); nx.getparent().remove(nx); nx=n2
d.save(os.path.join(HERE,'Business_Ethics_Compliance_Certificate_CGC_EN.docx'))

# ---------- VQD
d=docx.Document(VQD); letterhead(d)
t=d.tables[0]
def cell(r,c,text):
    ps=t.rows[r].cells[c].paragraphs
    (set_text(ps[0],text) if ps[0].runs else ps[0].add_run(text))
cell(3,1,NAME)
t.rows[4].cells[0].paragraphs[0].add_run('Title:'); t.rows[4].cells[1].paragraphs[0].add_run(f'{TITLE}, {COMPANY}')
cell(6,1,f'{DATE}, {PLACE}')
cell(9,1,'M.P. / [stamp]')
d.save(os.path.join(HERE,'Vendor_Qualification_Undertaking_CGC_EN.docx'))
