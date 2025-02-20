
# Import docx NOT python-docx
import docx
from docx.shared import Inches, Pt
 
# Create an instance of a word document
doc = docx.Document()
 
# Add a table
doc.add_heading('Section 2: Data', level=2)
table = doc.add_table(rows=5, cols=2)
table.style = 'Table Grid'
table.autofit = False
table.allow_autofit = False
for row in table.rows:
    for cell in row.cells:
        cell.width = Pt(200)
table.cell(0, 0).text = '•	Acesso pelo telefone a especialistas'
table.cell(0, 2).text = '•	Bate-papo online com especialistas'
table.cell(0, 4).text = '•	Respostas ao fórum dadas por especialistas'
table.cell(0, 6).text = '•	Orientação técnica geral'
table.cell(0, 8).text = '•	Assistência HPE InfoSight'
table.cell(0, 1).text = '•	Alertas preventivos HPE InfoSight'
table.cell(0, 3).text = '•	Registro de incidentes automatizado'
table.cell(0, 5).text = '•	Biblioteca de dicas técnicas'
table.cell(0, 7).text = '•	Acesso a informações e serviços de suporte eletrônico'
table.cell(0, 9).text = '•	Gerenciamento de interrupções (apenas no nível de serviço de Crítico)'

doc.save('test.docx')