import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import shutil
import win32com.client
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Copy fresh backup to test file
test_path = r'C:\Users\Admin\Desktop\Thạc sĩ\Chương trình\Cơ sở dữ liệu nâng cao\edu_graph_tinhoc_neo4j\scratch\test_toc.docx'
shutil.copy2('BAO_CAO_DO_AN_CUOI_KY_CSDL_NANG_CAO_FINAL_BACKUP.docx', test_path)

doc = docx.Document(test_path)
print("Initial sections count:", len(doc.sections))
print("Initial paragraphs count:", len(doc.paragraphs))
