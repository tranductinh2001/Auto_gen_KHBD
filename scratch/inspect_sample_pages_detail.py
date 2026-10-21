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
    
    for pg in range(4, 12):
        rng = d.GoTo(1, 1, pg)
        rng2 = d.GoTo(1, 1, pg + 1) if pg < d.ComputeStatistics(2) else d.Content
        end_pos = rng2.Start if pg < d.ComputeStatistics(2) else d.Content.End
        rng.End = end_pos
        print(f'==================== PAGE {pg} ({len(rng.Text)} chars) ====================')
        lines = [line.strip() for line in rng.Text.split('\r') if line.strip()]
        for l in lines[:15]:
            print('  ', l[:100])
    d.Close(False)
finally:
    word.Quit()
