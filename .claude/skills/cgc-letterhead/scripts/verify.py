"""Check that a docx still carries the untouched CGC letterhead and that its
rendered page 1 header matches the reference PDF (text block positions, pt).
usage: verify.py doc.docx [doc.pdf]"""
import sys, zipfile, hashlib, os
A = os.path.join(os.path.dirname(__file__), "..", "assets")
def h(z, n): return hashlib.md5(z.read(n)).hexdigest()
ref = zipfile.ZipFile(os.path.join(A, "CGC_letterhead_template.docx"))
z = zipfile.ZipFile(sys.argv[1])
ok = True
for n in ("word/header1.xml", "word/media/image1.png", "word/media/image2.png", "word/_rels/header1.xml.rels"):
    same = h(ref, n) == h(z, n); ok &= same; print(("OK  " if same else "DIFF"), n)
import re
sp = lambda zz: re.search(r"<w:sectPr.*?</w:sectPr>", zz.read("word/document.xml").decode(), re.S).group(0)
same = re.sub(r'w:rsid\w*="[^"]*"', "", sp(ref)) == re.sub(r'w:rsid\w*="[^"]*"', "", sp(z)); ok &= same
print(("OK  " if same else "DIFF"), "sectPr (page size, margins, header ref)")
if len(sys.argv) > 2:
    import fitz
    def blocks(p): return sorted((round(b[0]), round(b[1])) for b in fitz.open(p)[0].get_text("blocks") if b[1] < 125)
    r = blocks(os.path.join(A, "CGC_letterhead_reference.pdf")); d = blocks(sys.argv[2])
    hdr = [x for x in d if x in r or any(abs(x[0]-y[0]) <= 1 and abs(x[1]-y[1]) <= 1 for y in r)]
    same = all(any(abs(x[0]-y[0]) <= 1 and abs(x[1]-y[1]) <= 1 for y in d) for x in r if x[1] < 100)
    ok &= same; print(("OK  " if same else "DIFF"), "rendered header block positions vs reference PDF")
print("PASS" if ok else "FAIL"); sys.exit(0 if ok else 1)
