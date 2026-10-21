import docx
import win32com.client
import os
import sys

sys.path.insert(0, os.getcwd())
sys.stdout.reconfigure(encoding='utf-8')

from core.exam_builder import ExamMatrixEngine, ExamDocxExporter

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

# Let's inspect the mcqs in dossier
mcqs = dossier['th01']['mcqs']
print(f'Total MCQs: {len(mcqs)}')
for i, q in enumerate(mcqs):
    print(f'Q{i+1}: {repr(q["content"][:35])}')
