"""Переоформление отчёта «Причал №10» (GeoProGlobal -> ТОО «Caspian Geology Center»).

Правки (только то, что утвердил заказчик правок):
  с.1        — титул заменён утверждённым образцом (logo Sarzha + CGC, Директор Шахтаев Г. Ж.)
  с.2        — в оглавлении «Государственная лицензия 3 л.» -> «8 л.» (листов лицензий CGC)
  с.3        — «GeoProGlobal» -> «Caspian Geology Center» (2 места), лицензия GeoProGlobal
               № 20006797 -> лицензия CGC № 23006536; абзацы перебиты по ширине
  с.36-38    — лицензия GeoProGlobal заменена лицензиями CGC (3 + 5 листов)
  Чертежи 1,2— в штампе: исполнитель -> ТОО «Caspian Geology Center»,
               Заказчик -> ТОО «Sarzha Cargo Terminal»
  метаданные — заголовок «Геопроглобал» заменён, закладка «...GPG» переименована
"""
import sys
import pymupdf

SRC, TITLE, LIC1, LIC2, OUT = sys.argv[1:6]
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
    "В соответствии с техническим заданием, выданным ТОО «Caspian Geology Center», "
    "ТОО «" + CGC + "» были выполнены инженерные изыскания по объекту: "
    "«Многофункциональный морской терминал \"Саржа\". Универсальный терминал. "
    "Причал №10 с прилегающей площадкой».", 14.04)
print("с.3 абзацы:", n1, n2, "строк")

# ---- с.2: число листов лицензий в оглавлении ----------------------------------
p2 = doc[1]
tgt = [s for l, s in spans(p2) if s["text"].strip() == "3 л." and abs(s["origin"][1] - 378.4) < 1]
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

# ---- лицензии: с.36-38 -> лицензии CGC ----------------------------------------
lic1, lic2 = pymupdf.open(LIC1), pymupdf.open(LIC2)
nlic = lic1.page_count + lic2.page_count
doc.delete_pages(35, 37)                       # 0-based 35..37 = с.36-38
doc.insert_pdf(lic1, start_at=35)
doc.insert_pdf(lic2, start_at=35 + lic1.page_count)
shift = nlic - 3

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
    toc.append([lvl, text, pg])
doc.set_toc(toc)
meta = doc.metadata
meta["title"] = "Отчет об инженерных изысканиях. ММТ «Саржа». Причал №10"
doc.set_metadata(meta)
doc.del_xml_metadata()

doc.save(OUT, garbage=3, deflate=True)
print("saved", OUT, "pages", doc.page_count)
