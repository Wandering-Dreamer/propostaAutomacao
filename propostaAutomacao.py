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
title5 = doc.add_heading(level=0)
title5.add_run("1. Resumo Executivo\n")
p5 = doc.add_paragraph()
p5.add_run('A Hewlett Packard Enterprise acredita que as seguintes vantagens competitivas nos diferenciam da concorrência:\n\n')
p5.add_run('Amplo portfólio de soluções completas. ').bold = True
p5.add_run('Combinamos nossos recursos de infraestrutura, software e serviços para oferecer o que acreditamos ser o maior e mais completo portfólio de soluções empresariais do setor de TI. Nossa capacidade de oferecer uma ampla variedade de produtos de alta qualidade e serviços de suporte e consultoria de alto valor em um único pacote é um dos nossos principais diferenciais.')
p5.add_run('\nRoteiro de inovação para vários anos. ').bold = True
p5.add_run('Atuamos no setor de tecnologia e inovação há mais de 75 anos. Nosso amplo portfólio de propriedade intelectual e nossos recursos de pesquisa e desenvolvimento global fazem parte um roteiro de inovação mais amplo para ajudar organizações de todos os tamanhos em sua jornada de plataformas de tecnologia tradicional rumo aos sistemas de TI do futuro — denominados Novo Estilo de TI — que acreditamos serão caracterizados pelo aumento e pela proeminência inter-relacionada de computação em nuvem, Big Data, segurança empresarial, aplicativos e mobilidade.')
p5.add_run('\nDistribuição global e ecossistema de parceiros. ').bold = True
p5.add_run('Somos especialistas no fornecimento de soluções tecnológicas inovadoras para nossos clientes em ambientes complexos com a participação de vários países, vários fornecedores e/ou vários idiomas. Temos um dos conjuntos de recursos go-to-market mais completos do setor, incluindo um amplo ecossistema de parceiros de canal que nos permite comercializar e entregar ofertas de produtos a clientes localizados em, praticamente, qualquer lugar do mundo.')
p5.add_run('\nSoluções financeiras personalizadas. ').bold = True
p5.add_run('Desenvolvemos soluções financeiras inovadoras para facilitar a entrega de produtos e serviços a nossos clientes. Oferecemos soluções de investimento flexíveis e especialidade que ajudam os clientes e outros parceiros a criar implantações de tecnologia exclusivas com base em necessidades comerciais específicas.')
p5.add_run('\nEquipe de liderança experiente com histórico comprovado de desempenho de sucesso. ').bold = True
p5.add_run('Nossa equipe de gerenciamento apresenta um histórico comprovado de desempenho e execução. Nossa equipe de gerenciamento sênior soma mais de 100 anos de experiência na área e possui amplo conhecimento e experiência no setor de TI comercial e nos mercados em que competimos. Além disso, possuímos um amplo banco de talentos em gerenciamento e tecnologia que — acreditamos — nos oferece pipeline sem precedentes para futuros líderes e inovadores.')
p5.add_run('\nUm parceiro de transformação com a visão e a abrangência para ajudar os clientes a alcançar ótimos resultados comerciais').font.size = Pt(10)
p5.paragraph_format.line_spacing = Pt(13)
doc.add_picture('image_page5.png', width=Pt(350))
p5.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
doc.add_page_break()

# Pag 6 

p6 = doc.add_paragraph()
p6.add_run('\nAo escolher a Hewlett Packard Enterprise como sua parceira comercial, o cliente ganhará um consultor experiente, comprometido e confiável. A Hewlett Packard Enterprise entende os muitos desafios associados à implantação da infraestrutura complexa de tecnologia da informação no ambiente de negócios em rápida mudança e altamente competitivo dos dias de hoje. Nossa estratégia concentra-se em ampliar nossos pontos fortes para oferecer maior valor para os clientes, seja com suporte pós-venda, um conjunto maior de ferramentas de virtualização e gerenciamento para empresas ou foco elevado em serviços. Oferecemos nosso incrível portfólio da maneira que for mais adequada para você — ajudando-o a transformar desafios em oportunidades. Nossas soluções completam essa estrutura ao combinar nossas tecnologias a seus objetivos comerciais avançados de forma holística e revolucionária. Com nossa visão, estratégia, experiência e liderança, a Hewlett Packard Enterprise se destaca claramente como parceira preferencial de negócios do cliente para o futuro. A equipe de profissionais de vendas e serviços da Hewlett Packard Enterprise tem a experiência necessária para traduzir as metas de negócios do cliente em soluções de TI que aprimoram a competitividade, geram um rápido retorno do investimento e fornecem proteção de longo prazo aos ativos por meio de um caminho de crescimento assegurado.')
p6.add_run('\nEm resumo, o que diferencia a Hewlett Packard Enterprise da concorrência é nossa proposta de valor para nossos clientes. A Hewlett Packard Enterprise ocupa uma posição única — sólida e invejável — com nossa combinação de ativos, adotando uma postura baseada em padrões abertos e cética em relação a plataformas, e com nossa capacidade de fornecer insights úteis para os clientes de maneira que nenhuma outra empresa pode fazer. Um conjunto integrado e global de produtos e serviços oferece muito mais — mais responsabilidade e agilidade comercial, flexibilidade, suporte econômico para as necessidades dinâmicas da TI, um retorno mais alto dos investimentos em TI e a melhor experiência total do cliente.')
p6.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
doc.add_page_break()

# Pag 7

title7 = doc.add_heading(level=0)
title7.add_run("\n2. Serviços de Suporte\n")
if ui.contrato_final.get() == 0:
    title71 = doc.add_heading(level=0)
    title71.add_run('HPE Pointnext Tech Care')
    p7 = doc.add_paragraph()
    p7.add_run("Aplicavel aos contratos que contenham algum dos Números de produto abaixo: \n\n\HU4A3AC; HU4A4AC; HU4A5AC; HU4A6AC; HU4A7AC; HU4A8AC; HU4A9AC; HU4B0AC; HU4B1AC; HU4B2AC; HU4B3AC; HU4B4AC; HU4B5AC; HU4B6AC; HU4B7AC;")
    p7.add_run("\n\nVISÃO GERAL DO SERVIÇO\n\n").bold = True
    p7.add_run('\nO HPE Pointnext Tech Care é a experiência de suporte operacional para os produtos de hardware e software da marca HPE (produtos HPE). Com o HPE Pointnext Tech Care, as equipes de TI podem manter-se focadas no desenvolvimento da empresa, buscando proativamente formas melhores de agir, ao invés de apenas reagir a problemas.')
    p7.add_run('\nO HPE Pointnext Tech Care vai além do suporte tradicional, permitindo o acesso direto a especialistas em produtos específicos e oferecendo orientação técnica geral para ajudar os clientes não só a reduzir riscos, mas também a buscar continuamente maneiras mais eficientes de agir. Os clientes HPE Pointnext Tech Care podem obter ajuda por meio de diversos canais, incluindo o telefone, fóruns HPE moderados com tempos de resposta definidos, registro automatizado de incidentes e um recurso de bate-papo em tempo real. O serviço permite o acesso a recursos técnicos oferecidos por profissionais experientes e com conhecimento especializado no hardware ou software dentro do contexto da carga de trabalho específica, evitando que o cliente perca tempo respondendo questões de triagem ou habilitação por vezes desnecessárias. O HPE Pointnext Tech Care vai além do suporte tradicional, oferecendo orientações técnicas gerais sobre a operação, gerenciamento e segurança do produto com suporte')
    p7.add_run('\nO HPE Support Center oferece uma experiência digital aprimorada e personalizada que ajuda os clientes a gerenciarem os seus ativos por meio do reconhecimento dos vários produtos instalados no seu ambiente e da maneira como eles interagem uns com os outros. Novas ferramentas de autoatendimento permitem que os clientes desempenhem certas atividades sem precisar abrir um incidente de suporte, e oferecem também um portal de recursos e conteúdos selecionados. O HPE Pointnext Tech Care oferece acesso a recursos HPE que ajudarão a alcançar a excelência operacional e a otimização do desempenho, desde a borda até a nuvem.')
    p7.add_run("\n\nESTRUTURA DO SERVIÇO\n\n").bold = True
    p7.add_run('O serviço HPE Pointnext Tech Care, conforme destacado a seguir, oferece um conjunto geral de recursos ao lado de recursos específicos de hardware e/ou software, com base na tecnologia suportada e no fato de o produto ser um hardware, um software ou ambos. Alguns recursos do serviço são aprimorados com o uso do HPE InfoSight1, permitindo que a Hewlett Packard Enterprise ofereça níveis cada vez mais altos de orientação técnica com o uso da telemetria fornecida. Os clientes que se registram online por meio do HPE Support Center ganham acesso a recursos digitais melhorados, permitindo uma maior comodidade de gerenciamento e envolvimento direto da HPE. Os tempos de resposta remota e no local variam com base no nível do serviço selecionado, com o maior nível de serviço oferecendo assistência adicional aos clientes em casos de interrupção.')
    p7.add_run("TABELA 1. ").bold = True
    p7.add_run("Resumo dos recursos do serviço")
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
elif ui.contrato_final.get() == 1:
    title72 = doc.add_heading(level=0)
    title72.add_run("HPE Pointnext Complete Care Starter Pack ")
    p71 = doc.add_paragraph()
    p71.add_run('Aplicável aos contratos que contenham algum dos Números de produto abaixo:\nH2T05AC; H0GT7AC; HA158AC; HA156AC; HL935AC; HA162AC; HA167AC; H7G18AC; H0JC1AC; HA360AC; H2T05AC; HU4A2AC')
    p71.add_run('\n\nDESCRIÇÃO DO SERVIÇO').bold = True
    p71.add_run('\n\nEsta ficha técnica do serviço HPE Pointnext Complete Care Add-on descreve o serviço que acrescenta novos produtos a um ambiente em que o cliente esteja adquirindo um novo HPE Pointnext Complete Care ou acrescentando há um ambiente existente.')
    p71.add_run('\nPara ser elegível à aquisição do serviço HPE Pointnext Complete Care Add-on, o cliente já deve ter um ambiente com HPE Pointnext Complete Care das duas seguintes formas:')
    p71.add_run('\n•	Ter um contrato já existente com os Serviços HPE Pointnext Complete Care com um descritivo técnico (SOW) personalizado, que inclui cobertura para o ambiente em que os novos produtos serão adicionados.')
    p71.add_run('\n•	Ter adquirido um novo HPE Pointnext Complete Care Starter Pack para o ambiente ao qual os produtos serão adicionados.')
    p71.add_run('\nAs opções acima são permitidas para realização de um pedido quando:')
    p71.add_run('\n•	Para o serviço HPE Pointnext Complete Care Add-on adquirido para produtos adicionados a um contrato com um descritivo técnico(SOW) personalizado existente: Os recursos de serviço fornecidos para produtos adicionados através de HPE Pointnext Complete Care Add-on serão os estabelecidos no seu contrato personalizado existente; ou ')
    p71.add_run('\n•	 Para os produtos adicionados HPE Pointnext Complete Care Add-on adquiridos com os Serviços HPE Pointnext Complete Care Starter Pack: Os entregáveis dos serviços fornecidos para produtos adicionados através dos HPE Pointnext Complete Care Add-on serão os definidos nesta ficha técnica, conforme especificado abaixo.')
    p71.add_run('\nQuando se adquire o serviços HPE Pointnext Complete Care Add-on, estende-se a equipe de contas designada liderada por um gerente de suporte de conta (ASM, account support manager) treinado ou Consultor de Serviço(Service Advisor), gerenciamento de incidente aprimorado (EIM, enhanced incidente management) e gerenciamento de relacionamento do serviço (SRM, service relationship management). ')
    p71.add_run('\n\nAlém disso e dependendo do tipo de produto sendo adicionado ao ambiente do HPE Pointnext Complete Care do cliente, várias opções proativas específicas da tecnologia ou outras opções de suporte podem ser incluídas em uma das duas maneiras a seguir:')
    p71.add_run('\n•	Recursos de suporte proativo específicos da tecnologia ou outras opções de suporte incluídas no serviço contratual do HPE Pointnext Complete Care descritos no seu contrato personalizado (SOW) do HPE Pointnext Complete Care.')
    p71.add_run('\n•	Recursos de suporte proativo específicos da tecnologia ou outras opções de suporte adquiridas separadamente por meio dos Créditos de Serviço HPE e descritas na ficha técnica do serviço HPE Pointnext Complete Care.')
    p71.add_run('\n\nNo serviço HPE Pointnext Complete Care Add-on inclui opções de nível de serviços reativos para    cobrir os requisitos de suporte do cliente, desde os ambientes mais básicos até os mais críticos para os negócios.')
    p71.add_run('\nAo adquirir um serviço de  HPE Pointnext Complete Care Add-on juntamente ao serviço HPE Pointnext Complete Care Starter Pack, consulte a ficha técnica do serviço HPE Pointnext Complete Care para conhecer as opções de cobertura disponíveis para esses serviços de HPE Pointnext Complete Care Add-on,   incluindo os níveis de serviços reativos aplicáveis e quaisquer limitações e exclusões, conforme estabelecidas e incorporadas a esta ficha técnica e aplicáveis à sua aquisição do serviço de HPE Pointnext Complete Care Add-on.')
    p71.add_run('\n\nBENEFÍCIOS DO SERVIÇO').bold = True
    p71.add_run('\n\nO serviço HPE Pointnext Complete Care Add-on foi criado para ajudar os clientes a atingirem consistentemente seus objetivos de negócios oferecendo:')

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