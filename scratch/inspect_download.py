import os, sys, glob
import win32com.client
sys.stdout.reconfigure(encoding='utf-8')

dl_files = glob.glob('C:/Users/Admin/Downloads/Ho_So_De_Thi_Chuan*.docx')
dl_files.sort(key=os.path.getmtime, reverse=True)
latest = dl_files[0]
print(f'Latest download: {latest}')

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
try:
    d = word.Documents.Open(os.path.abspath(latest), ReadOnly=True)
    pgs = d.ComputeStatistics(2)
    print(f'Total pages in downloaded file: {pgs}')
    for pg in range(1, pgs + 1):
        rng = d.GoTo(1, 1, pg)
        rng2 = d.GoTo(1, 1, pg + 1) if pg < pgs else d.Content
        rng.End = rng2.Start if pg < pgs else d.Content.End
        t = [l.strip() for l in rng.Text.split('\r') if l.strip()]
        first_l = t[0][:70] if t else '<empty>'
        last_l = t[-1][:70] if t else '<empty>'
        print(f'Page {pg:2d} ({len(t):2d} lines): start="{first_l}" | end="{last_l}"')
    d.Close(False)
finally:
    word.Quit()
