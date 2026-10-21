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
    
    anchors = [
        ('TH01 Header', 'ĐỀ CHÍNH THỨC'),
        ('TH01 Cut 1', '✂'),
        ('TH01 Phần I', 'I. Khoanh tròn'),
        ('TH01 Câu 1', 'Câu 1.'),
        ('TH01 Câu 6', 'Câu 6.'),
        ('TH01 Câu 7', 'Câu 7.'),
        ('TH01 Câu 12', 'Câu 12.'),
        ('TH01 Phần II', 'II. Viết chữ'),
        ('TH01 Tự luận Header', 'Phần tự luận'),
        ('TH01 Câu 17', 'Câu 17.'),
        ('TH01 BÀI LÀM 1', 'BÀI LÀM'),
        ('TH01 BÀI LÀM 2', 'BÀI LÀM')
    ]
    
    rng = d.Content
    for label, text in anchors:
        rng.Find.ClearFormatting()
        rng.Find.Text = text
        if rng.Find.Execute():
            pg = rng.Information(3)
            print(f'{label:<25} -> Page: {pg}')
            rng.Collapse(0)
    d.Close(False)
finally:
    word.Quit()
