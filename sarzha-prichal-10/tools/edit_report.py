"""Переоформление отчёта «Причал №10» (GeoProGlobal -> ТОО «Caspian Geology Center»).

Правки (только то, что утвердил заказчик правок):
  с.1        — титул заменён утверждённым образцом (logo Sarzha + CGC, Директор Шахтаев Г. Ж.)
  с.2        — в оглавлении «Государственная лицензия 3 л.» -> «8 л.» (листов лицензий CGC),
               «Техническое задание 7 л.» -> «8 л.»
  с.3        — «GeoProGlobal» -> «Caspian Geology Center» (2 места), лицензия GeoProGlobal
               № 20006797 -> лицензия CGC № 23006536; абзацы перебиты по ширине
  с.28-34    — скан ТЗ (с инициалами и печатью субподрядчика) заменён ТЗ из Word без печатей (8 л.)
  с.36-38    — лицензия GeoProGlobal заменена лицензиями CGC (3 + 5 листов)
  с.3        — ТЗ «выданным ТОО «Sarzha Cargo Terminal»»; «директором –Тусупбаевым А.Е.» ->
               «директором –Шахтаевым Г.Ж.»
  Чертежи 1,2— в штампе: исполнитель -> ТОО «Caspian Geology Center»,
               Заказчик -> ТОО «Sarzha Cargo Terminal»,
               Директор Тусупбаев А. -> Шахтаев Г.Ж., подпись Тусупбаева убрана (клетка пустая)
  метаданные — заголовок «Геопроглобал» заменён, закладка «...GPG» переименована
"""
import sys
import pymupdf

SRC, TITLE, LIC1, LIC2, TZ, OUT = sys.argv[1:7]
FONT = "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf"  # метрика = Times New Roman
F = pymupdf.Font(fontfile=FONT)
CGC = "Caspian Geology Center"

doc = pymupdf.open(SRC)
old_toc = doc.get_toc(simple=False)


def spans(page):
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                yield l, s


def redact(page, rects):
    for r in rects:
        page.add_redact_annot(r, fill=False)
    page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                          graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                          text=pymupdf.PDF_REDACT_TEXT_REMOVE)


def put(page, x, y, text, size):
    page.insert_text((x, y), text, fontname="LibSerif", fontfile=FONT, fontsize=size)


def mid(bbox, k=0.3):
    """Средняя по высоте полоса bbox — чтобы не задеть соседние строки."""
    x0, y0, x1, y1 = bbox
    h = y1 - y0
    return pymupdf.Rect(x0 - 0.5, y0 + h * k, x1 + 0.5, y1 - h * k)


def retypeset(page, first_line_bbox_y, n_lines, new_text, size):
    """Перебить абзац (выравнивание по ширине) на месте исходных строк."""
    lines = [l for l, s in spans(page)]
    uniq = []
    for l in lines:
        if l not in uniq:
            uniq.append(l)
    # строки абзаца: n_lines подряд начиная с той, чья верхняя граница = first_line_bbox_y
    start = next(i for i, l in enumerate(uniq) if abs(l["bbox"][1] - first_line_bbox_y) < 0.6)
    para = uniq[start:start + n_lines]
    base = [l["spans"][0]["origin"][1] for l in para]
    x_first = para[0]["bbox"][0]
    x_left = min(l["bbox"][0] for l in para[1:])
    x_right = max(l["bbox"][2] for l in para[:-1])
    redact(page, [mid(l["bbox"], 0.25) for l in para])

    words = new_text.split()
    sp = F.text_length(" ", fontsize=size)
    out, cur, i = [], [], 0
    for w in words:
        x0 = x_first if not out else x_left
        test = cur + [w]
        width = sum(F.text_length(t, fontsize=size) for t in test) + sp * (len(test) - 1)
        if cur and x0 + width > x_right + 0.5:
            out.append(cur)
            cur = [w]
        else:
            cur = test
    out.append(cur)
    assert len(out) <= n_lines, f"абзац вырос: {len(out)} > {n_lines} строк"
    for li, ws in enumerate(out):
        x0 = x_first if li == 0 else x_left
        wl = [F.text_length(t, fontsize=size) for t in ws]
        last = li == len(out) - 1
        gap = sp if (last or len(ws) == 1) else (x_right - x0 - sum(wl)) / (len(ws) - 1)
        x = x0
        for t, w in zip(ws, wl):
            put(page, x, base[li], t, size)
            x += w + gap
    return len(out)


# ---- с.3: название организации (2 места) --------------------------------------
p3 = doc[2]
n1 = retypeset(p3, 47.1, 5,
    "Наименование юридического лица: TOO «" + CGC + "». Государственная лицензия "
    "№ 23006536 от 13 марта 2023 года на занятие изыскательной деятельностью, действует "
    "на территории Республики Казахстан, выдана Государственное учреждение «Управление "
    "контроля и качества городской среды города Астаны». Акимат города Астаны.", 14.04)
n2 = retypeset(p3, 176.6, 4,
    "В соответствии с техническим заданием, выданным ТОО «Sarzha Cargo Terminal», "
    "ТОО «" + CGC + "» были выполнены инженерные изыскания по объекту: "
    "«Многофункциональный морской терминал \"Саржа\". Универсальный терминал. "
    "Причал №10 с прилегающей площадкой».", 14.04)
n3 = retypeset(p3, 656.4, 2,
    "Текущий контроль методики, качества производства работ и соблюдения правил техники "
    "безопасности осуществлялся директором –Шахтаевым Г.Ж.", 14.04)
print("с.3 абзацы:", n1, n2, n3, "строк")

# ---- с.2: число листов лицензий в оглавлении ----------------------------------
p2 = doc[1]
tgt = [s for l, s in spans(p2) if s["text"].strip() == "3 л." and abs(s["origin"][1] - 378.4) < 1]
assert len(tgt) == 1
s = tgt[0]
redact(p2, [mid(s["bbox"])])
put(p2, s["origin"][0], s["origin"][1], "8 л.", s["size"])
tgt = [s for l, s in spans(p2) if s["text"].strip() == "7 л." and abs(s["origin"][1] - 359.9) < 1]
assert len(tgt) == 1
s = tgt[0]
redact(p2, [mid(s["bbox"])])
put(p2, s["origin"][0], s["origin"][1], "8 л.", s["size"])

# ---- штампы чертежей: с.107 (Чертёж 1), с.109-111 (Чертёж 2) ------------------
def fix_stamp(page):
    ss = [s for l, s in spans(page)]
    # исполнитель
    g = [s for s in ss if "GeoProGlobal" in s["text"]]
    # заказчик: « «Caspian Geology Center» + «»»
    c = [s for s in ss if s["text"].strip().startswith("«Caspian Geology Center")]
    assert len(g) == 1 and len(c) == 1, (len(g), len(c))
    g, c = g[0], c[0]
    close = [s for s in ss if s["text"] == "»" and abs(s["bbox"][0] - c["bbox"][2]) < 1.5]
    assert len(close) == 1
    close = close[0]
    redact(page, [mid(g["bbox"]), mid(c["bbox"]), mid(close["bbox"])])
    # Заказчик: ТОО «Sarzha Cargo Terminal»
    t1 = " «Sarzha Cargo Terminal"
    x = c["origin"][0]
    put(page, x, c["origin"][1], t1, c["size"])
    x += F.text_length(t1, fontsize=c["size"])
    put(page, x, close["origin"][1], "»", close["size"])
    # исполнитель по центру прежней надписи, с подгонкой размера под ячейку
    new = "ТОО «" + CGC + "»"
    cx = (g["bbox"][0] + g["bbox"][2]) / 2
    size = g["size"]
    return new, cx, g, size


# границы ячейки «исполнитель» в штампе (по векторным линиям чертежа)
for pno, cell in ((106, (1495.2, 1629.7)), (108, (1029.4, 1162.8)), (109, (1029.4, 1162.8)), (110, (1029.4, 1162.8))):
    page = doc[pno]
    new, cx, g, size = fix_stamp(page)
    avail = (cell[1] - cell[0]) - 8
    while F.text_length(new, fontsize=size) > avail:
        size -= 0.1
    w = F.text_length(new, fontsize=size)
    put(page, cx - w / 2, g["origin"][1], new, size)
    print(f"с.{pno+1}: исполнитель {size:.2f} pt (было {g['size']:.2f}), ширина {w:.1f}/{avail:.1f}")

# ---- директор в штампах: Тусупбаев А. -> Шахтаев Г.Ж., его подпись убрать -----
from collections import Counter

DIRECTOR = "Шахтаев Г.Ж."
SANS = "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
FS = pymupdf.Font(fontfile=SANS)


def _col(pth):
    return tuple(round(c, 2) for c in (pth.get("color") or ()))


def _sig(pth):
    r = pth["rect"]
    return (_col(pth), round(r.x0, 1), round(r.y0, 1), round(r.x1, 1), round(r.y1, 1))


def remove_paths(page, pred):
    """Удалить векторные пути, для которых pred(path) истинно (несколько проходов
    с растущим запасом — у острых углов линии выступают за bbox)."""
    total = [0, 0]
    for m in (1.2, 3.0, 6.0, None):          # None — вырожденные штрихи: удаление касанием
        n, miss, still = _remove_paths_once(page, pred, m)
        total[0] += n; total[1] += miss
        if not still:
            break
    return total[0], total[1], still


def _remove_paths_once(page, pred, margin):
    """Один проход. Пути, задетые случайно (не подходящие под pred), перерисовываются как были."""
    before = page.get_drawings()
    targets = [q for q in before if pred(q)]
    keep = [q for q in before if not pred(q)]
    if not targets:
        return 0, 0, 0
    for t in targets:
        if margin is None:                   # точка в середине первого отрезка пути
            it = t["items"][0]
            a, b = it[1], it[-1]
            c = pymupdf.Point((a.x + b.x) / 2, (a.y + b.y) / 2)
            page.add_redact_annot(pymupdf.Rect(c.x - 0.1, c.y - 0.1, c.x + 0.1, c.y + 0.1), fill=False)
        else:
            page.add_redact_annot(t["rect"] + (-margin, -margin, margin, margin), fill=False)
    page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                          graphics=(pymupdf.PDF_REDACT_LINE_ART_REMOVE_IF_TOUCHED if margin is None
                                    else pymupdf.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED),
                          text=pymupdf.PDF_REDACT_TEXT_NONE)
    left = Counter(_sig(q) for q in page.get_drawings())
    missing = []
    for q in keep:
        k = _sig(q)
        if left[k]:
            left[k] -= 1
        else:
            missing.append(q)
    for q in missing:                                   # вернуть задетое
        sh = page.new_shape()
        for it in q["items"]:
            if it[0] == "l":
                sh.draw_line(it[1], it[2])
            elif it[0] == "c":
                sh.draw_bezier(it[1], it[2], it[3], it[4])
            elif it[0] == "re":
                sh.draw_rect(it[1])
            elif it[0] == "qu":
                sh.draw_quad(it[1])
        sh.finish(color=q.get("color"), fill=q.get("fill"), width=q.get("width") or 1,
                  closePath=q.get("closePath", False))
        sh.commit()
    still = [q for q in page.get_drawings() if pred(q)]
    return len(targets), len(missing), len(still)


# Чертёж 1 (с.107): фамилия — текст Times с наклоном (Tm 1 0 0.2126 1, Tz 91.02 %, 166.09*0.06 pt)
page = doc[106]
sp = [s for l, s in spans(page) if s["text"].strip() == "Тусупбаев А."]
assert len(sp) == 1
sp = sp[0]
row1 = pymupdf.Rect(1245, 1064, 1285, 1084)               # клетка «Подп.» строки «Директор»
redact(page, [mid(sp["bbox"])])
res = remove_paths(page, lambda q: _col(q) == (0.0, 0.0, 1.0) and row1.contains(q["rect"]))
o = pymupdf.Point(sp["origin"])
page.insert_text(o, DIRECTOR, fontname="LibSerif", fontfile=FONT, fontsize=166.0932 * 0.06,
                 morph=(o, pymupdf.Matrix(0.910233, 0, 0.2126, 1, 0, 0)))
print("с.107 директор: подпись удалена/перерисовано/осталось", res)

# Чертёж 2 (с.109-111): фамилия нарисована линиями шрифта AutoCAD
NAME2 = pymupdf.Rect(723, 1524, 780, 1534.8)              # клетка фамилии строки «Директор»
SIGN2 = pymupdf.Rect(784, 1519, 824, 1537.5)              # клетка «Подп.» строки «Директор»
for pno in (108, 109, 110):
    page = doc[pno]
    g = remove_paths(page, lambda q: _col(q) == (0.0, 0.0, 0.0) and NAME2.contains(q["rect"]))
    sg = remove_paths(page, lambda q: _col(q) == (0.0, 0.0, 1.0) and SIGN2.contains(q["rect"]))
    size = 6.7 / 0.716                                     # высота прописных как у соседних фамилий
    sx = min(1.0, (782.8 - 2 - 726.1) / FS.text_length(DIRECTOR, fontsize=size))
    o = pymupdf.Point(726.1, 1531.8)
    page.insert_text(o, DIRECTOR, fontname="LibSans", fontfile=SANS, fontsize=size,
                     morph=(o, pymupdf.Matrix(sx, 0, 0, 1, 0, 0)))
    print(f"с.{pno+1} директор: буквы {g}, подпись {sg}, сжатие {sx:.2f}")

# ---- лицензии: с.36-38 -> лицензии CGC ----------------------------------------
lic1, lic2 = pymupdf.open(LIC1), pymupdf.open(LIC2)
nlic = lic1.page_count + lic2.page_count
doc.delete_pages(35, 37)                       # 0-based 35..37 = с.36-38
doc.insert_pdf(lic1, start_at=35)
doc.insert_pdf(lic2, start_at=35 + lic1.page_count)
shift = nlic - 3

# ---- ТЗ: с.28-34 (скан) -> ТЗ из Word без печатей --------------------------------
tz = pymupdf.open(TZ)
doc.delete_pages(27, 33)                       # 0-based 27..33 = с.28-34
doc.insert_pdf(tz, start_at=27)
tz_shift = tz.page_count - 7

# ---- титул ---------------------------------------------------------------------
title = pymupdf.open(TITLE)
doc.delete_page(0)
doc.insert_pdf(title, from_page=0, to_page=0, start_at=0)

# ---- закладки и метаданные -----------------------------------------------------
toc = []
for lvl, text, pg, *rest in old_toc:
    if text == "Report (5)":
        continue
    if "GPG" in text:
        text = "Лицензии ТОО «Caspian Geology Center»"
    if pg > 38:
        pg += shift
    if pg > 34:
        pg += tz_shift
    toc.append([lvl, text, pg])
doc.set_toc(toc)
meta = doc.metadata
meta["title"] = "Отчет об инженерных изысканиях. ММТ «Саржа». Причал №10"
doc.set_metadata(meta)
doc.del_xml_metadata()

doc.save(OUT, garbage=3, deflate=True)
print("saved", OUT, "pages", doc.page_count)
