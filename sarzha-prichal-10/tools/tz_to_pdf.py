"""ТЗ (Word, без печатей) -> PDF через LibreOffice.

Разрывы страниц перед «Таблица 1» и «Схема участка изысканий».
В колонтитуле ТЗ поле даты договора пустое; LibreOffice подставляет туда
«30 декабря 1899». Перед конвертацией у поля убирается признак «дата», чтобы
текст остался таким, каким его показывает Word: «№______________ от       2026 г.».

    python3 tz_to_pdf.py "ТЗ Причал 10 (без печати).docx" out_dir
"""
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import docx

src, out_dir = Path(sys.argv[1]), Path(sys.argv[2])
tmp = Path(tempfile.mkdtemp())
# LibreOffice разбивает страницы иначе, чем Word: «Таблица 1» и «Схема участка
# изысканий» начинаются с новой страницы (как в скане ТЗ и в Word).
from docx.oxml.ns import qn

d = docx.Document(src)
for par in d.paragraphs:
    t = par.text.strip()
    if t.startswith("Таблица 1") or t.startswith("Схема участка изысканий"):
        par.paragraph_format.page_break_before = True
# пустые абзацы-«распорки» перед схемой (ими в Word схема сдвигалась на новую страницу)
els = list(d.element.body.iterchildren())
k = next(i for i, e in enumerate(els)
         if e.tag == qn("w:p") and "".join(x.text or "" for x in e.iter(qn("w:t"))).startswith("Схема участка"))
i = k - 1
while i >= 0 and els[i].tag == qn("w:p") and not "".join(x.text or "" for x in els[i].iter(qn("w:t"))).strip() \
        and not list(els[i].iter(qn("w:drawing"))):
    els[i].getparent().remove(els[i])
    i -= 1
paged = tmp / "paged.docx"
d.save(paged)
fixed = tmp / "tz.docx"
with zipfile.ZipFile(paged) as zin, zipfile.ZipFile(fixed, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if re.fullmatch(r"word/header\d+\.xml", item.filename):
            s = data.decode("utf-8")
            s = re.sub(r"<w:date\b[^>]*/>|<w:date\b.*?</w:date>", "", s, flags=re.S)
            data = s.encode("utf-8")
        zout.writestr(item, data)
subprocess.run(["soffice", "-env:UserInstallation=file:///tmp/lo_profile", "--headless", "--norestore",
                "--convert-to", "pdf", "--outdir", str(tmp), str(fixed)], check=True, capture_output=True)
out_dir.mkdir(parents=True, exist_ok=True)
shutil.copy(tmp / "tz.pdf", out_dir / "tz.pdf")
print(out_dir / "tz.pdf")
