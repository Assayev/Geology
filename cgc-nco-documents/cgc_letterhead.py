"""Перенос настоящего фирменного бланка CGC (шапка: логотип, реквизиты KZ/RU/EN, водяной знак) в любой docx."""
import copy, os, io
from docx.oxml.ns import qn
HERE=os.path.dirname(os.path.abspath(__file__))
TEMPLATE=os.path.join(HERE,'CGC_blank_template.docx')
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def _q(t): return '{%s}%s'%(W,t)

def _import_styles(doc, tpl, header_el):
    """Копирует стили шапки (с префиксом CGC_) в целевой документ, чтобы шапка не зависела от его стилей."""
    ts={st.get(_q('styleId')):st for st in tpl.styles.element.findall(_q('style'))}
    used=set()
    for tag in ('pStyle','rStyle','tblStyle'):
        for e in header_el.iter(_q(tag)): used.add(e.get(_q('val')))
    dd=tpl.styles.element.find(_q('docDefaults'))
    chain=[]
    def add(i):
        if i in ts and i not in chain:
            b=ts[i].find(_q('basedOn'))
            if b is not None: add(b.get(_q('val')))
            chain.append(i)
    for i in used: add(i)
    if 'Normal' in ts and 'Normal' not in chain: chain.insert(0,'Normal')
    tgt=doc.styles.element
    have={st.get(_q('styleId')) for st in tgt.findall(_q('style'))}
    for i in chain:
        new=copy.deepcopy(ts[i]); nid='CGC_'+i
        if nid in have: continue
        new.set(_q('styleId'),nid)
        n=new.find(_q('name'))
        if n is not None: n.set(_q('val'),'CGC '+n.get(_q('val')))
        for tag in ('basedOn','next','link'):
            e=new.find(_q(tag))
            if e is not None:
                if e.get(_q('val')) in ts: e.set(_q('val'),'CGC_'+e.get(_q('val')))
                else: new.remove(e)
        new.attrib.pop(_q('default'),None)
        if new.find(_q('basedOn')) is None and dd is not None and new.get(_q('type'))!='table' :
            # корневой стиль: подмешиваем docDefaults шаблона
            for kind in ('rPr','pPr'):
                src=dd.find(_q(kind+'Default'))
                src=src.find(_q(kind)) if src is not None else None
                if src is None: continue
                cur=new.find(_q(kind))
                if cur is None: cur=copy.deepcopy(src); new.append(cur)
                else:
                    for ch in src:
                        if cur.find(ch.tag) is None: cur.append(copy.deepcopy(ch))
        tgt.append(new)
    for tag in ('pStyle','rStyle','tblStyle'):
        for e in header_el.iter(_q(tag)):
            if e.get(_q('val')) in ts: e.set(_q('val'),'CGC_'+e.get(_q('val')))

def _bake_alpha(new, hp):
    """alphaModFix (прозрачность водяного знака) не везде поддерживается: запекаем её в PNG."""
    from PIL import Image
    A='http://schemas.openxmlformats.org/drawingml/2006/main'
    for blip in new.iter('{%s}blip'%A):
        fix=blip.find('{%s}alphaModFix'%A)
        if fix is None: continue
        k=int(fix.get('amt','100000'))/100000.0
        rid=blip.get('{%s}embed'%R)
        im=Image.open(io.BytesIO(hp.related_parts[rid].blob)).convert('RGBA')
        a=im.getchannel('A').point(lambda v:int(v*k))
        im.putalpha(a)
        buf=io.BytesIO(); im.save(buf,'PNG'); buf.seek(0)
        new_rid,_=hp.get_or_add_image(buf)
        blip.set('{%s}embed'%R,new_rid); blip.remove(fix)

def apply_letterhead(doc, a4=True, margins=(2.0,3.0,1.5,2.0)):
    import docx
    tpl=docx.Document(TEMPLATE)
    th=tpl.sections[0].header
    for s in doc.sections:
        h=s.header; h.is_linked_to_previous=False
        hp=h.part
        new=copy.deepcopy(th._element)
        _import_styles(doc,tpl,new)
        # переносим картинки и ссылки в связи целевой шапки
        for el in new.iter():
            for attr in ('embed','id'):
                key='{%s}%s'%(R,attr)
                rid=el.get(key)
                if not rid: continue
                rel=th.part.rels[rid]
                if rel.is_external:
                    el.set(key,hp.relate_to(rel.target_ref,rel.reltype,is_external=True))
                else:
                    new_rid,_=hp.get_or_add_image(io.BytesIO(rel.target_part.blob)) if 'image' in rel.reltype else (None,None)
                    el.set(key,new_rid)
        _bake_alpha(new, hp)
        old=h._element
        for ch in list(old): old.remove(ch)
        for ch in list(new): old.append(ch)
        if a4:
            from docx.shared import Cm
            s.page_width=Cm(21.0); s.page_height=Cm(29.7)
            s.top_margin=Cm(margins[0]); s.left_margin=Cm(margins[1]); s.right_margin=Cm(margins[2]); s.bottom_margin=Cm(margins[3])
            s.header_distance=Cm(1.25)
