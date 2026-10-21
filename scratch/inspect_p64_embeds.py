import docx
import sys
sys.stdout.reconfigure(encoding='utf-8')

doc = docx.Document('scratch/test_8pages.docx')
for i, p in enumerate(doc.paragraphs):
    d = p._p.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')
    if d:
        print(f"P{i}: text={repr(p.text[:30])}, drawings={len(d)}")
        for j, draw in enumerate(d):
            ext = draw.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent')
            cx = int(ext.attrib.get('cx', 0)) / 914400 if ext is not None else 0
            cy = int(ext.attrib.get('cy', 0)) / 914400 if ext is not None else 0
            print(f"   Draw {j}: {cx:.2f}\" x {cy:.2f}\"")
