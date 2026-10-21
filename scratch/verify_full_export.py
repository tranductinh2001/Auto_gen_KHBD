import docx
import win32com.client
import os, sys

sys.path.insert(0, os.getcwd())
sys.stdout.reconfigure(encoding='utf-8')

from core.exam_builder import ExamMatrixEngine, ExamDocxExporter

print("1. Generating dossier...")
engine = ExamMatrixEngine(
    grade_level='Lớp 9',
    exam_title='KIỂM TRA GIỮA HỌC KỲ I',
    school_name='TRƯỜNG TH&THCS ĐĂK TRÔI',
    governing_body='UBND XÃ KON CHIÊNG',
    school_year='2026 - 2027',
    teacher_name='Lê Thị Phương Thảo',
    head_name='Trần Thị Tuyết Ngân'
)
dossier = engine.generate_exam_dossier()

out_docx = os.path.abspath('scratch/test_final_full_export.docx')
print(f"2. Exporting full exam dossier to: {out_docx}")
ExamDocxExporter.export(dossier, out_docx)

print("3. Validating with Word COM...")
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
try:
    d = word.Documents.Open(out_docx, ReadOnly=True)
    total_pgs = d.ComputeStatistics(2)
    print(f"\n=======================================================")
    print(f">>> TOTAL COMPUTED PAGES OF DOSSIER: {total_pgs} <<<")
    print(f"=======================================================\n")
    for pg in range(1, total_pgs + 1):
        rng = d.GoTo(1, 1, pg)
        rng2 = d.GoTo(1, 1, pg + 1) if pg < total_pgs else d.Content
        rng.End = rng2.Start if pg < total_pgs else d.Content.End
        t = [l.strip() for l in rng.Text.split('\r') if l.strip()]
        first_l = t[0][:75] if t else '<empty>'
        print(f"  Page {pg:2d} ({len(t):2d} lines): {first_l}")
    d.Close(False)
finally:
    word.Quit()
print("\nValidation finished successfully!")
