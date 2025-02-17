from docx import Document
from docx.shared import Pt
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date
import locale
import ui


doc = Document("Template Proposta Tecnica Comercial Servicos de Suporte_novo (003) (002).docx")

title = doc.add_heading('Resposta para ', level=0)
title.add_run(ui.cliente_final.get())
title.add_run(" de Hewlett Packard Enterprise \n\n\n\n")
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title2 = doc.add_heading('Projeto: ', level=0)
title2.add_run(ui.num_contrato_final.get())
title2.alignment = WD_ALIGN_PARAGRAPH.LEFT
title2.add_run('\n\n\n\n\n\n\n\n')

p = doc.add_paragraph('São Paulo, ')
p.add_run(ui.str_dt)
p = doc.add_paragraph('Proposta Técnica Comercial ')
p.add_run(ui.ope_final.get())
p.add_run('\n\n\n')
doc.add_picture('image_page1.jpg', width=Pt(500))


# Add a heading
doc.add_heading('Section 1: Introduction', level=2)

# Add a bulleted list
list_paragraph = doc.add_paragraph()
list_paragraph.add_run(ui.vendedor_final.get()).bold = True
list_paragraph.add_run(' - This is the first bullet point.')
list_paragraph.add_run('\n')
list_paragraph.add_run('Bullet 2').bold = True
list_paragraph.add_run(' - This is the second bullet point.')

# Add a table
doc.add_heading('Section 2: Data', level=2)
table = doc.add_table(rows=3, cols=3)
table.style = 'Table Grid'
table.autofit = False
table.allow_autofit = False
for row in table.rows:
    for cell in row.cells:
        cell.width = Pt(100)
table.cell(0, 0).text = 'Name'
table.cell(0, 1).text = 'Age'
table.cell(0, 2).text = 'City'
for i, data in enumerate([('Alice', '25', 'New York'), ('Bob', '30', 'San Francisco'), ('Charlie', '22', 'Los Angeles')], start=0):
    table.cell(i, 0).text = data[0]
    table.cell(i, 1).text = data[1]
    table.cell(i, 2).text = data[2]

# Add an image
#doc.add_heading('Section 3: Image', level=2)
#doc.add_paragraph('Here is an image:')
#doc.add_picture('path_to_your_image.jpg', width=Pt(300))

# Save the document
doc.save('example_document.docx')