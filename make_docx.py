# ---
# jupyter:
#   jupytext:
#     formats: ipynb,py:percent
#     text_representation:
#       extension: .py
#       format_name: percent
#       format_version: '1.3'
#       jupytext_version: 1.13.7
#   kernelspec:
#     display_name: Python 3 (ipykernel)
#     language: python
#     name: python3
# ---

# %%
from docx import Document
from docx.enum.table import WD_ROW_HEIGHT
from docx.enum.table import WD_ALIGN_VERTICAL



# %%
from docx.shared import Cm

# %%
document = Document(f)

f = 'template.docx'
d = Document(f)
para_last = d.tables[-2]

pic = 'test.png'
para_last_pic = para_last.cell(-1,0).add_paragraph().add_run().add_picture(pic, width=Cm(19))

row = para_last.add_row()
row.height_rule = WD_ROW_HEIGHT.AT_LEAST
row.height = Cm(1)
row.cells[0].text = 'test'

row = para_last.add_row()
row.height_rule = WD_ROW_HEIGHT.AT_LEAST
row.height = Cm(1)
row.cells[0].text = 'test'
row.cells[0].vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM

d.save('template_out.docx')
