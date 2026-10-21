import docx
import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

from core.exam_builder import ExamMatrixEngine, ExamDocxExporter

# Generate dossier
engine = ExamMatrixEngine(
    grade_level='Lớp 9',
    exam_title='KIỂM TRA GIỮA HỌC KỲ I',
    school_name='TRƯỜNG TH&THCS ĐĂK TRÔI',
    governing_body='UBND XÃ KON CHIÊNG',
    school_year='2025 - 2026',
    teacher_name='Lê Thị Phương Thảo',
    head_name='Trần Thị Tuyết Ngân'
)
dossier = engine.generate_exam_dossier()

mcqs = dossier['th01']['mcqs']
print(f'Total MCQs: {len(mcqs)}')

def estimate_q_height_pt(q):
    stem_len = len(q.get("content", "")) + 12
    stem_lines = max(1, (stem_len + 74) // 75)
    opts = q.get("options", {})
    opt_keys = ["A", "B", "C", "D"]
    max_opt_len = max((len(str(opts.get(k, ""))) for k in opt_keys), default=0)
    if max_opt_len <= 14:
        opt_lines = 1
    elif max_opt_len <= 54:
        opt_lines = 2
    else:
        opt_lines = 4
    total_lines = stem_lines + opt_lines
    # each line ~13.5pt in 11.5pt font with 1.10 spacing, plus 3pt paragraph spacing
    return total_lines * 14.5 + 4.0

SAFE_BUDGET_PT = 510.0
accum_pt = 0.0
p1_count = 0
for i, q in enumerate(mcqs):
    h = estimate_q_height_pt(q)
    print(f'Q{i+1}: est height = {h:.1f} pt (accum = {accum_pt + h:.1f} pt)')
    if accum_pt + h > SAFE_BUDGET_PT:
        break
    accum_pt += h
    p1_count += 1

print(f'\nOptimal Page 1 MCQ count: {p1_count} (Total pt: {accum_pt:.1f} / {SAFE_BUDGET_PT} pt)')
print(f'Remaining for Page 2: {len(mcqs) - p1_count} MCQs')
