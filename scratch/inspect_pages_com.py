import docx
import win32com.client
import os, sys
sys.stdout.reconfigure(encoding='utf-8')

target = os.path.abspath('scratch/user_downloaded_exam.docx')
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
try:
    d = word.Documents.Open(target, ReadOnly=True)
    pgs = d.ComputeStatistics(2)
    print(f'Total pages: {pgs}')
    for pg in range(1, pgs + 1):
        rng = d.GoTo(1, 1, pg)
        rng2 = d.GoTo(1, 1, pg + 1) if pg < pgs else d.Content
        rng.End = rng2.Start if pg < pgs else d.Content.End
        t = [l.strip() for l in rng.Text.split('\r') if l.strip()]
        first_l = t[0][:60] if t else '<empty>'
        last_l = t[-1][:60] if t else '<empty>'
        print(f'Page {pg:2d} ({len(t):2d} lines):')
        print(f'   First: {first_l}')
        print(f'   Last:  {last_l}')
    d.Close(False)
finally:
    word.Quit()
