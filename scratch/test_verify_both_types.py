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

from core.exam_builder import ExamMatrixEngine, ExamDocxExporter
from scratch.test_user_mcqs import user_mcqs
from scratch.test_optimal_finder import user_tfs, est_mcq_pt, est_tf_pt, find_optimal_page1_count
from scratch.test_uniform_font_and_fit import build_test_doc

# Test user exam with k=find_optimal_page1_count
k_u1 = find_optimal_page1_count(user_mcqs[:12], user_tfs)
k_u2 = find_optimal_page1_count(user_mcqs[12:], user_tfs)
print(f"Optimal k for user TH01: {k_u1}, TH02: {k_u2}")

engine = ExamMatrixEngine(grade_level='Lớp 9')
dossier = engine.generate_exam_dossier()
meta = dossier['meta']
essays = dossier['th01']['essays']

doc1 = build_test_doc(user_mcqs[:12], user_tfs, essays, meta, 'TH01', k_u1)
doc1.save(os.path.abspath('scratch/test_user_balanced_4p.docx'))

# Also test standard exam with find_optimal_page1_count
tfs_std = dossier['th01']['tfs']
k_s1 = find_optimal_page1_count(dossier['th01']['mcqs'], tfs_std)
k_s2 = find_optimal_page1_count(dossier['th02']['mcqs'], tfs_std)
print(f"Optimal k for std TH01: {k_s1}, TH02: {k_s2}")

doc2 = build_test_doc(dossier['th01']['mcqs'], tfs_std, essays, meta, 'TH01', k_s1)
doc2.save(os.path.abspath('scratch/test_std_balanced_4p.docx'))

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
try:
    for name, fpath in [("User Short Exam", 'scratch/test_user_balanced_4p.docx'), ("Standard Long Exam", 'scratch/test_std_balanced_4p.docx')]:
        d = word.Documents.Open(os.path.abspath(fpath), ReadOnly=True)
        pgs = d.ComputeStatistics(2)
        print(f"\n>>> {name} -> TOTAL PAGES: {pgs} <<<")
        for pg in range(1, pgs + 1):
            rng = d.GoTo(1, 1, pg)
            rng2 = d.GoTo(1, 1, pg + 1) if pg < pgs else d.Content
            rng.End = rng2.Start if pg < pgs else d.Content.End
            t = [l.strip() for l in rng.Text.split('\r') if l.strip()]
            print(f"   Page {pg}: {len(t)} lines | start: {t[0][:50] if t else ''}")
        d.Close(False)
finally:
    word.Quit()
