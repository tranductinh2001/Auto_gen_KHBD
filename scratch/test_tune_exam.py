import docx
import win32com.client
import os
import sys

sys.path.insert(0, os.getcwd())
sys.stdout.reconfigure(encoding='utf-8')

from core.exam_builder import ExamMatrixEngine, ExamDocxExporter

# Let's inspect how _render_exam_paper renders TH01
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

# Test export
out_docx = os.path.abspath('scratch/test_tuned_output.docx')

# Let's see what happens if page1_count is 8
# We temporarily patch _calculate_optimal_page1_mcq_count in ExamDocxExporter
old_fn = ExamDocxExporter._calculate_optimal_page1_mcq_count

def patched_calc(mcqs):
    # Dồn câu hỏi lên: 8 câu trên Trang 1 (Câu 1 đến Câu 8)
    return min(len(mcqs), 8)

ExamDocxExporter._calculate_optimal_page1_mcq_count = classmethod(lambda cls, mcqs: patched_calc(mcqs))

ExamDocxExporter.export(dossier, out_docx)

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
try:
    d = word.Documents.Open(out_docx, ReadOnly=True)
    total_pg = d.ComputeStatistics(2)
    print(f'Total pages with page1_count=8: {total_pg}')
    for pg in range(1, total_pg + 1):
        rng = d.GoTo(1, 1, pg)
        rng2 = d.GoTo(1, 1, pg + 1) if pg < total_pg else d.Content
        end_pos = rng2.Start if pg < total_pg else d.Content.End
        rng.End = end_pos
        t = [l.strip() for l in rng.Text.split('\r') if l.strip()]
        first_line = t[0] if t else '<empty>'
        print(f'Pg {pg:2d} ({len(t):2d} lines): {first_line[:75]}')
    d.Close(False)
finally:
    word.Quit()
