from docx import *
from docx.shared import Pt, Inches, Mm
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date
import locale
import ui

# Pag 1
doc = Document("Template Proposta Tecnica Comercial Servicos de Suporte_novo (003) (002).docx")

section = doc.sections[0]

section.left_margin = Inches(0.75)
section.right_margin = Inches(0.75)
section.top_margin = Inches(1)
section.bottom_margin = Inches(1)

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
doc.add_page_break()


# Pag 2

section = doc.sections[-1]
section2 = doc.add_section()

section2.left_margin = Inches(1.85)
section2.right_margin = Inches(0.5)
section2.top_margin = Inches(0.25)
section2.bottom_margin = Inches(1)

p2 = doc.add_paragraph('\n\n\nSão Paulo, ')
p2.add_run(ui.str_dt)
p2.add_run('\n\n\n')
p2 = doc.add_paragraph(ui.cliente_final.get())
p2.add_run(", ")
p2.add_run('\n\n\n\n')
vendedor = doc.add_paragraph()
vendedor.add_run(ui.vendedor_final.get()).font.size = Pt(9)
vendedor.add_run('\n').font.size = Pt(9)
vendedor.add_run(ui.cargo_final.get()).font.size = Pt(9)
vendedor.add_run('\n').font.size = Pt(9)
vendedor.add_run(ui.telefone_final.get()).font.size = Pt(9)
vendedor.add_run('\n').font.size = Pt(9)
vendedor.add_run(ui.email_final.get()).font.size = Pt(9)
vendedor.add_run('\n\n\n').font.size = Pt(9)
vendedor.alignment = WD_ALIGN_PARAGRAPH.LEFT
vendedor.paragraph_format.left_indent = Mm(-30.4)
p21 = doc.add_paragraph('Prezados (as) Senhores(as):\n\n')
p21.alignment = WD_ALIGN_PARAGRAPH.LEFT
p22 = doc.add_paragraph('Temos o prazer de apresentar a nossa proposta técnico-comercial referente ao Projeto: ')
p22.add_run(ui.num_contrato_final.get())
p22.add_run("""\nEstamos confiantes que as informações contidas nesta proposta possam atender suas necessidades, demonstrando desta maneira a potencialidade de nossa empresa nos termos de qualidade de produtos e serviços.""")
p22.add_run('\nEsta proposta foi desenvolvida por Hewlett Packard Enterprise, que analisou todos os aspectos necessários para uma implementação bem sucedida.')
p22.add_run('\nEstamos confiantes em demonstrar os benefícios de valor agregado da proposta e construir um relacionamento de negócio sólido e benéfico para ambas as partes.')
p22.add_run('\nColocamo-nos à disposição para quaisquer esclarecimentos que se faça necessário.')
p22.add_run('\n\n\n\nAtenciosamente,')
p22.add_run('\n\n\n\n\n _____________________________\n\n\n')
p22.add_run(ui.vendedor_final.get())
p22.add_run('\n')
p22.add_run(ui.cargo_final.get())
p22.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
hpe = doc.add_paragraph('\n\n\n\n')
hpe.add_run('Hewlett Packard Enterprise').bold = True
hpe.add_run('\nAlameda Rio Negro, 750')
hpe.add_run('\nBarueri, SP, 06454-000')
hpe.add_run('\nBrazil')
hpe.add_run('\nwww.hpe.com')
hpe.alignment = WD_ALIGN_PARAGRAPH.LEFT
hpe.paragraph_format.left_indent = Mm(-30.4)
doc.add_page_break()

# Pag 3

title3 = doc.add_heading(level=0)
title3.add_run("\nAviso de Confidencialidade\n")
p3 = doc.add_paragraph()
p3.add_run("As informações contidas em todas as páginas deste documento / proposta é confidencial da Hewlett Packard Enterprise e Hewlett Packard Enterprise Company (a seguir coletivamente \"Hewlett Packard Enterprise\") e seguem para fins de avaliação. Ao receber o documento, o destinatário concorda em manter tais informações em sigilo e não reproduzir ou divulgar a qualquer pessoa fora do grupo diretamente responsável pela avaliação do conteúdo, a menos que a  Hewlett Packard Enterprise tenha autorizado. Não há obrigação de manter a confidencialidade de qualquer parte da informação que o destinatário tenha tido conhecimento sem restrições antes do recebimento deste documento, como é provado através de registos escritos, de negócios ou informações de conhecimento público sem que o destinatário tenha incorrido em faltas, ou que tenha sido recebido pelo destinatário através de uma terceira parte sem restrições.").font.size = Pt(9)
p3.add_run("\n\nEste documento contém informações sobre produtos, vendas e programas de serviço da  Hewlett Packard Enterprise que podem ser melhorados ou descontinuados a critério exclusivo da  Hewlett Packard Enterprise. A  Hewlett Packard Enterprise tem feito todos os esforços para incluir materiais aqui considerados confiáveis e relevantes para fins de avaliação de seu destinatário. Nem a Hewlett Packard Enterprise nem seus representantes dão qualquer garantia quanto à exatidão ou completude das informações. Portanto, este documento é apenas para fins informativos devendo ser considerado para os negócios da  Hewlett Packard Enterprise. Nem a  Hewlett Packard Enterprise nem seus representantes serão responsáveis sobre qualquer ato do destinatário ou de seus representantes, como resultado do uso das informações aqui fornecidas. A assinatura de um acordo definitivo ou assinatura de aceitação da proposta, por representantes autorizados das partes, será o único meio pelo qual a  Hewlett Packard Enterprise ou suas afiliadas serão vinculadas à proposta/ contrato.").font.size = Pt(9)
title31 = doc.add_heading(level = 0).add_run("Restrições de cópias entregues da Proposta\n")
p31 = doc.add_paragraph().add_run("""\nA proposta da Hewlett Packard Enterprise foi enviada em formato eletrônico no formato de arquivo PDF. Se o conteúdo dos arquivos originais forem diferentes da versão em PDF, somente o conteúdo da versão PDF será respeitado pela Hewlett Packard Enterprise.""").font.size = Pt(9)
title32 = doc.add_heading(level = 0).add_run("Esclarecimentos\n")
p32 = doc.add_paragraph().add_run("\nDúvidas ou  esclarecimentos sobre esta Política de Privacidade, entre em contato com seu representante de vendas.").font.size = Pt(9)
p33 = doc.add_paragraph().add_run("\n\n\n© Copyright 2025 Hewlett-Packard Development Company, L.P.").font.size = Pt(9)
p3.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
doc.add_page_break()

# Pag 4

title4 = doc.add_heading(level=0)
title4.add_run("Índice\n\n")
p4 = doc.add_paragraph()
p4.add_run('1. Resumo Executivo	                                                                                  5\n')
p4.add_run('2. Serviços de Suporte	                                                                      7\n')
p4.add_run('3. Especificações dos Níveis de Serviços	                                            13\n')
p4.add_run('4. Suporte a Produtos Multivendor	                                                        24\n')
p4.add_run('5. Condições Comerciais         	                                                        25\n')
p4.add_run('6. Condições Gerais                	                                                        29\n')
p4.add_run('7. Termo de Aceite da Proposta / Pedido de Compra                               31\n')
p4.add_run('8. Anexos                	                                                                               32\n')
doc.add_page_break()

# Pag 5

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