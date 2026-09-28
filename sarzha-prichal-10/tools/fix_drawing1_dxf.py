"""Чертёж 1 (DWG -> DXF через LibreDWG dwg2dxf), правка двух MTEXT в штампе на уровне байтов.

    dwg2dxf -y -o ch1.dxf "Чертеж 1. Карта фактического материала.dwg"
    python3 fix_drawing1_dxf.py ch1.dxf "Чертеж 1. Карта фактического материала_CGC.dxf"

Исполнитель: ТОО «GeoProGlobal» -> ТОО «Caspian Geology Center» (высота 1.16667x -> 0.93x, чтобы влезло в ячейку)
Заказчик:    ТОО «Caspian Geology Center» -> ТОО «Sarzha Cargo Terminal»
"""
import sys

src, out = sys.argv[1:3]
b = open(src, "rb").read()
pairs = [
    ("{\\fTimes New Roman|b0|i0|c204|p18;\\H1.16667x;\\C256;ТОО «GeoProGlobal»}",
     "{\\fTimes New Roman|b0|i0|c204|p18;\\H0.93x;\\C256;ТОО «Caspian Geology Center»}"),
    (" «Caspian Geology Center\\A1;\\H0.85714x;»",
     " «Sarzha Cargo Terminal\\A1;\\H0.85714x;»"),
]
for old, new in pairs:
    assert b.count(old.encode()) == 1, old
    b = b.replace(old.encode(), new.encode())
open(out, "wb").write(b)
