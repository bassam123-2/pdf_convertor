# -*- coding: utf-8 -*-
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

docx = Path(r"C:\Users\jouan\Desktop\PDF_TOOL\Related_Party_editable.docx")
with ZipFile(docx) as z:
    xml = z.read("word/document.xml")

root = etree.fromstring(xml)
ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
text = xml.decode("utf-8", errors="replace")
print("tbl tags", text.count("<w:tbl>"))
print("Page background", text.count("Page background"))
print("locked=1", text.count('locked="1"'))
print("behindDoc=1", text.count('behindDoc="1"'))
tables = root.findall(".//w:tbl", ns)
print("tables found", len(tables))
for ti, tbl in enumerate(tables[:2]):
    rows = tbl.findall("./w:tr", ns)
    print(f"table {ti} rows", len(rows))
    for ri, row in enumerate(rows[:3]):
        cells = []
        for cell in row.findall("./w:tc", ns):
            bits = [t.text or "" for t in cell.findall(".//w:t", ns)]
            cells.append("".join(bits)[:50])
        print(f"  r{ri}", cells)
