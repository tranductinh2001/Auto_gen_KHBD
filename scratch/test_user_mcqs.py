import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client
import os, sys

sys.path.insert(0, os.getcwd())
sys.stdout.reconfigure(encoding='utf-8')

# Let's extract the actual data from scratch/user_downloaded_exam.docx
user_doc = docx.Document('scratch/user_downloaded_exam.docx')

# Extract MCQs from user_doc
user_mcqs = []
curr_q = None
for p in user_doc.paragraphs:
    txt = p.text.strip()
    if txt.startswith("Câu ") and "khoanh tròn" not in txt.lower():
        # Check if MCQ (before II. Viết chữ)
        if "điểm)" in txt or "“" in txt:
            continue
        # Extract question number and content
        parts = txt.split(". ", 1)
        if len(parts) == 2 and parts[0].replace("Câu ", "").isdigit():
            q_num = int(parts[0].replace("Câu ", ""))
            if q_num <= 12:
                curr_q = {"num": q_num, "content": parts[1], "options": {}}
                user_mcqs.append(curr_q)
    elif curr_q and ("A." in txt or "B." in txt or "C." in txt or "D." in txt):
        # Parse options
        items = txt.split("\t")
        for it in items:
            it = it.strip()
            for opt_k in ["A", "B", "C", "D"]:
                prefix = f"{opt_k}. "
                if it.startswith(prefix):
                    curr_q["options"][opt_k] = it[len(prefix):].strip()

print(f"Extracted {len(user_mcqs)} MCQs from user doc:")
for q in user_mcqs:
    print(f"  Câu {q['num']}: {q['content'][:45]}... (opts: {len(q['options'])})")
