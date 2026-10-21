import docx
import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
try:
    path = os.path.abspath('để thi mẫu/Le Thi Phuong Thao_DKT GHK 1_25-26_ TIN HỌC 9.docx')
    d = word.Documents.Open(path, ReadOnly=True)
    total_pg = d.ComputeStatistics(2)
    print(f'Total pages in sample: {total_pg}')
    for pg in range(4, min(total_pg + 1, 16)):
        rng = d.GoTo(1, 1, pg)
        rng.End = min(rng.Start + 300, d.Content.End)
        first_line = rng.Text.split('\r')[0].strip()
        print(f'Page {pg:2d}: {first_line[:80]}')
    d.Close(False)
finally:
    word.Quit()
