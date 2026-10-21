import docx

doc = docx.Document('scratch/test_8pages.docx')
for i, s in enumerate(doc.inline_shapes):
    print(f'Shape {i}: w={s.width.inches:.2f}\", h={s.height.inches:.2f}\"')
