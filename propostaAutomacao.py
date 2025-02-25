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
p6.add_run('\n\nEm resumo, o que diferencia a Hewlett Packard Enterprise da concorrência é nossa proposta de valor para nossos clientes. A Hewlett Packard Enterprise ocupa uma posição única — sólida e invejável — com nossa combinação de ativos, adotando uma postura baseada em padrões abertos e cética em relação a plataformas, e com nossa capacidade de fornecer insights úteis para os clientes de maneira que nenhuma outra empresa pode fazer. Um conjunto integrado e global de produtos e serviços oferece muito mais — mais responsabilidade e agilidade comercial, flexibilidade, suporte econômico para as necessidades dinâmicas da TI, um retorno mais alto dos investimentos em TI e a melhor experiência total do cliente.')
p6.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
doc.add_page_break()

# Pag 7, Pag 8, Pag 9, Pag 10, Pag 11

title7 = doc.add_heading(level=0)
title7.add_run("\n2. Serviços de Suporte\n")
if ui.contrato_final.get() == 0:
    title71 = doc.add_heading(level=0)
    title71.add_run('HPE Pointnext Tech Care')
    p7 = doc.add_paragraph()
    p7.add_run("Aplicavel aos contratos que contenham algum dos Números de produto abaixo: \n\nHU4A3AC; HU4A4AC; HU4A5AC; HU4A6AC; HU4A7AC; HU4A8AC; HU4A9AC; HU4B0AC; HU4B1AC; HU4B2AC; HU4B3AC; HU4B4AC; HU4B5AC; HU4B6AC; HU4B7AC;")
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
    title73 = doc.add_heading(level=0)
    title73.add_run("HPE Pointnext Complete Care Starter Pack")
    p72 = doc.add_paragraph()
    p72.add_run('Aplicavel aos contratos que contenham algum dos Números de produto abaixo: \nH2T12AC; H2T12BC;')
    p72.add_run('\n\nDESCRIÇÃO DO SERVIÇO').bold = True
    p72.add_run('\nEsta ficha técnica da Hewlett Packard Enterprise (HPE) descreve o serviço HPE Pointnext Complete Care Starter Pack, que é um pacote de suporte que fornece serviços de gerenciamento do ambiente e relacionamento que fazem parte da oferta de serviços de suporte HPE Pointnext Complete Care. É um mecanismo para os clientes adquirirem a cobertura para cada produto do ambiente de TI juntamente ao HPE Pointnext Complete Care Add-ons. O serviço HPE Pointnext Complete Care Starter Pack, juntamente aos serviços de Add-on, oferecem uma alternativa  à aquisição de um contrato com descritivo técnico personalizado do HPE Pointnext Complete (SOW). Ao adquirirem os serviços Starter Pack, há três diferentes níveis de experiência para escolher, cada um oferecendo um conjunto diferente de entregáveis de serviços:')
    p72.add_run('\n1.	HPE Pointnext Complete Care Starter Pack Standard (padrão)')
    p72.add_run('\n2.	HPE Pointnext Complete Care Starter Pack Básico')
    p72.add_run('\n3.	HPE Pointnext Complete Care Starter Pack de Entry (Inicial)')
    p72.add_run('\nPara mais detalhes sobre os diferentes níveis de experiência, consulte a tabela de Entregáveis do Serviço HPE Pointnext Complete Care Starter Pack.')
    p72.add_run('\nOs clientes escolhem o serviço HPE Pointnext Complete Care Starter Pack mais adequado às necessidades deles e adicionam produtos ao ambiente de IT juntamente com HPE Pointnext Complete Care Add-on aplicável a cada produto. Isso expande os recursos do serviço Starter Pack, como detalhados nesta ficha técnica, a esses produtos durante o termo do serviço Starter Pack.')
    p72.add_run('\nO serviço HPE Pointnext Complete Care Starter Pack não inclui serviços técnicos proativos para produtos HPE; no entanto, os clientes podem adicionar esses serviços por meio da aquisição em separado de Créditos de Serviço HPE(HPE Support Credits). O serviço Starter Pack e os Créditos de  Serviço da HPE são ativados por meio da realização, pelo cliente, de um pedido aceitável de acordo com esta ficha técnica.')
    p72.add_run('\n\nBENEFÍCIOS DO SERVIÇO').bold = True
    p72.add_run('\nOs serviços HPE Pointnext Complete Care Starter Pack foram criados para ajudarem os clientes a atingirem consistentemente suas metas de nível de serviço e outros objetivos de negócios, oferecendo:')
    p72.add_run('\n•	Um método rápido e fácil para adquirirem serviços de suporte HPE para todo o ambiente de TI (e não por dispositivo)')
    p72.add_run('\n•	Identificação proativa de problemas e consultoria sobre mitigação de riscos')
    p72.add_run('\n•	Acesso aos especialistas da HPE, que podem amplificar os recursos do cliente com o objetivo geral de ajudar a reduzir riscos, aumentar a produtividade, lidar com pico de cargas de trabalho, projetos emergentes, assim liberando o tempo do cliente para focar nos objetivos e negócios estratégicos')
    p72.add_run('\n•	Opções flexíveis de suporte reativo')
    p72.add_run('\n•	Acesso prioritário aos especialistas da HPE que conhecem o ambiente do cliente e podem ajudar a lidar rapidamente com quaisquer problemas críticos')
    p72.add_run('\n•	Opções flexíveis de suporte proativo prestado por especialistas da HPE, que complementam os recursos do cliente e ajudam-o a se concentrar em outras prioridades')
    p72.add_run('\n•	Tecnologias e ferramentas remotas avançadas projetadas para reduzirem o tempo de inatividade e aumentarem a produtividade')
    p72.add_run('\n•	Uma equipe de conta designada, focada no ambiente de TI e nos objetivos de negócios do cliente e que fornece um único ponto de contato na HPE ajudando a garantir que o relacionamento do cliente com a HPE atenda às expectativas do cliente e verificando o fornecimento de todas as opções de serviços, conforme acordado')

elif ui.contrato_final.get() == 2:
    title72 = doc.add_heading(level=0)
    title72.add_run("HPE Pointnext Complete Care Add-on  ")
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
    p71.add_run('\n•	Uma solução de suporte modular com eficiência de custos e feita sob medida para os requisitos e o ambiente exatos do cliente')
    p71.add_run('\n•	Uma equipe de conta designada, focada no ambiente de TI e nos objetivos de negócios do cliente, fornecendo um único ponto de contato na HPE, ajudando a garantir que o relacionamento do cliente com a HPE atenda às expectativas dele e verificando o fornecimento de todas as opções de serviços, conforme acordado')
    p71.add_run('\n•	Identificação proativa de problemas e consultoria sobre mitigação de riscos')
    p71.add_run('\n•	Acesso aos especialistas da HPE, que podem amplificar os recursos do cliente com o objetivo geral de ajudar a reduzir riscos, aumentar a produtividade, lidar com as cargas de trabalho de pico e projetos emergentes e liberar o tempo do cliente para objetivos de negócios estratégicos')
    p71.add_run('\n•	Opções flexíveis de suporte reativo')
    p71.add_run('\n•	Acesso prioritário aos especialistas da HPE que conhecem o ambiente do cliente e podem ajudar a lidar rapidamente com quaisquer problemas críticos')
    p71.add_run('\n•	Opções flexíveis de suporte proativo prestado por especialistas da HPE, que complementam os recursos do cliente e podem liberar os clientes para se concentrarem em outras prioridades')
    p71.add_run('\n•	Tecnologias e ferramentas remotas avançadas projetadas para reduzirem o tempo de inatividade e aumentarem a produtividade')


# pag 12, 13, 14


title8 = doc.add_heading(level=0)
title8.add_run('\n\n2.4.	Vistoria de Hardware / RTS')
p8 = doc.add_paragraph()
p8.add_run('\n\nEste serviço é aplicável a qualquer equipamento HEWLETT PACKARD ENTERPRISE  ou equipamento suportado pela HEWLETT PACKARD ENTERPRISE  que nunca tiveram seus equipamentos cobertos por contrato de suporte de hardware, ou que já tiveram e cancelaram seu contrato e que pretendem novamente assiná-lo.')
p8.add_run('\n\nPara incluir equipamentos que estavam sem cobertura nos últimos 45 dias é necessário que se avalie a elegibilidade do equipamento, garantido:')
p8.add_run('\n•	Não esteja obsoleto, seja produto suportado pela HEWLETT PACKARD ENTERPRISE  ')
p8.add_run('\n•	Esteja atualizado com as últimas configurações e revisões')
p8.add_run('\n•	Esteja operando sem falhas conforme determinado pela HPE')
p8.add_run('\n\nPara incluir equipamentos que estavam sem cobertura nos últimos 45 dias é necessário que se avalie a elegibilidade do equipamento, garantido:')
p8.add_run('\n•	Pagamento da taxa RTS')
p8.add_run('\n•	Carência de 30 dias, isto é, o equipamento do cliente não estará suportado pelos serviços HEWLETT PACKARD ENTERPRISE ora contratados. Durante o período de carência, e uma vez evidenciado a necessidade, o cliente terá direito aos serviços de suporte mediante o pagamento de um “chamado avulso / per call service”, o qual será faturado em complemento ao(s) valore(s) especificado(s) nesta proposta e em estrita observância à lista de preços praticada pela HPE.')
p8.add_run('\nO faturamento e pagamento referente ao valor do RTS deverá ser efetuado na sua totalidade no primeiro mês de vigência do contrato de suporte.')
p8.add_run('\n\nCollaborative Support').bold = True
p8.add_run('\n\nAplicavel aos contratos que contenham algum dos Números de produto abaixo:\nHL935AC; HU4A1AC')
p8.add_run('\nOs Serviços de Hardware da HEWLETT PACKARD ENTERPRISE fornecem assistência remota e se durante o atendimento, a HEWLETT PACKARD ENTERPRISE determinar que um problema é causado por software de terceiros e que não se aplica nenhuma da correções conhecidas e disponíveis conforme definido no Software Básico, o time de suporte, quando autorizado pelo Cliente, poderá chamar o fornecedor de software de terceiros seguindo os acordos de suporte contratado entre o cliente e o fornecedor.')
p8.add_run('\nComo parte do processo de gerenciamento do chamado, a HEWLETT PACKARD ENTERPRISE fornecerá documentação e análise realizada para que o fornecedor siga com atendimento junto ao cliente.')
p8.add_run('\nUma vez que o fornecedor do software está envolvido, o chamado aberto junto a HEWLETT PACKARD ENTERPRISE será fechado e poderá ser reaberto, se necessário a qualquer momento fazendo referência ao número original de identificação da chamada.')
p8.add_run('\n\n2.5.	Suporte de Software').bold = True
p8.add_run('\n\nAplicavel aos contratos que contenham algum dos Números de produto abaixo:\nHA156AC; HU1R5AC; HV2X8AC; HA158AC')
p8.add_run('\nO Suporte de Software da HEWLETT PACKARD ENTERPRISE  fornece serviços abrangentes para produtos de software HEWLETT PACKARD ENTERPRISE  e de terceiros selecionados. Com este serviço, sua equipe de TI conta com acesso rápido e confiável às Centrais de Atendimento da HEWLETT PACKARD ENTERPRISE . Os analistas da Central de Atendimento trabalharão com sua equipe para fornecer orientações sobre as características e utilização, diagnósticos e resolução de problemas, identificação de defeitos e acesso a patches dos produtos de software.')
p8.add_run('\n\nO Serviço também disponibiliza atualizações de software para produtos da HEWLETT PACKARD ENTERPRISE  e de terceiros elegíveis suportados pela HEWLETT PACKARD ENTERPRISE , patches de software e manuais de referência, incluindo licença de uso e cópia de novas versões de produtos de software em todos os sistemas suportados e cobertos pela licença original do mesmo.')
p8.add_run('\nO serviço também fornece acesso eletrônico às informações de suporte, permitindo que qualquer membro de sua equipe de TI localize informações essenciais disponíveis sobre produtos e suporte. Para produtos de terceiros, este acesso está sujeito à disponibilidade  de tais informações eletrônicas por parte do fornecedor.')
p8.add_run('\n\nPrincipais Características dos Serviços de Software').bold = True
p8.add_run('\n\n•	Suporte remoto\n•	Acesso a recursos técnicos\n•	Análise e resolução de problemas\n•	Gerenciamento de escalação\n•	Isolamento de problemas\n•	Suporte de orientação à instalação')
p8.add_run('\n•	Atualizações de software da HEWLETT PACKARD ENTERPRISE  e de terceiros selecionados a um custo previsível\n•	Redução dos custos de aquisição de atualizações individuais de software, devido à economia substancial de assinaturas')
p8.add_run('\n•	Notificações automáticas sobre a disponibilidade de novas versões de software\n•	Opção de janelas de cobertura\n•	Acesso a informações e serviços eletrônicos avançados de suporte que aumentam a produtividade:')
p8.add_run('\n–	Hewlett Packard Enterprise Support Center (https://h20564.www2.hpe.com/ ): É um site de suporte inovador onde os profissionais de TI podem obter informações sobre software e documentações, abertura eletrônica e acompanhamento de chamados, chat direto com os engenheiros de suporte da HEWLETT PACKARD ENTERPRISE , dentre outros.')
p8.add_run('\n\n2.5.1.	Serviços de Suporte de Software Limitado (por Incidentes)').bold = True
p8.add_run('\n\nO Suporte Técnico de Software por Incidentes fornece serviços abrangentes de suporte remoto para produtos de software selecionados de terceiros (Microsoft, Linux Red Hat ou Suse Enterprise Edition e Novell).')
p8.add_run('\nCom este serviço, sua equipe de TI conta com acesso rápido e confiável às Centrais de Atendimento da HEWLETT PACKARD ENTERPRISE . Os analistas da Central de Atendimento trabalharão com sua equipe para fornecer orientações sobre as características e utilização, diagnósticos e resolução de problemas, identificação de defeitos e acesso a patches dos produtos de software.')
p8.add_run('\nAlém disso, esta modalidade permite que o serviço de suporte seja adequado às necessidades de cada ambiente com opções de 10, 25, 50 ou 75 incidentes por ano, que podem ser utilizados para diversos equipamentos. ')
p8.add_run('\n\n2.6.	Atualização de Software / RTS ').bold = True
p8.add_run('\\n serviço é aplicável a qualquer equipamento HEWLETT PACKARD ENTERPRISE  e foi desenvolvido especificamente para clientes HEWLETT PACKARD ENTERPRISE  que nunca tiveram seus equipamentos cobertos por contrato de suporte de software, ou que já tiveram e cancelaram seu contrato e que pretendem novamente assiná-lo.')
p8.add_run('\n\nO Serviço de Atualização de Software HEWLETT PACKARD ENTERPRISE  (RTS - Return-to-Support) é destinado a atualizar a versão de software de qualquer produto HEWLETT PACKARD ENTERPRISE  para a última versão corrente. Após o recebimento  e instalação desta nova versão, o equipamento estará qualificado para que seja feito um contrato de manutenção de software.')
p8.add_run('\n\nA HEWLETT PACKARD ENTERPRISE  irá prover ao cliente a versão atualmente suportada, de sistema operacional, subsistemas ou softwares aplicativos HEWLETT PACKARD ENTERPRISE . A versão corrente suportada é definida como versão de software que consta na lista de preços HEWLETT PACKARD ENTERPRISE  no momento da compra do serviço RTS. O cliente deve optar para quais produtos deseja receber a atualização, e caso a versão do sistema operacional instalada no seu equipamento apareça como a mais atual, ele deve optar somente pelos subsistemas ou softwares aplicativos.')
p8.add_run('\n\nComprando o RTS o cliente irá receber atualização dos softwares constantes no item Configuração e Preços.')
p8.add_run('\nOs serviços de RTS deverão estar incluídos no item Configuração e Preços através dos códigos de produtos UC255AC ou UC256AC descritos como SW Updates Return to Support.')
p8.add_run('\n\nServiços não incluídos:	\n•	Manuais de software ou atualização de manuais.\n•	Software Status Bulletins não recebidos.\n•	Auxílio na instalação das novas versões ')
p8.add_run('\nTanto o serviço de instalação das novas versões quanto os manuais devem ser adquiridos a parte, aos preços vigentes da época da compra do Serviço HEWLETT PACKARD ENTERPRISE -RTS. ')
p8.add_run('\nO faturamento e pagamento referente ao valor do RTS deverá ser efetuado na sua totalidade no primeiro mês de vigência do contrato de suporte.')
doc.add_page_break()

# Pag 15

title9 = doc.add_heading(level=0)
title9.add_run("3.	Especificações dos Níveis de Serviços")

if ui.servico_final.get() == 0:
    title91 = doc.add_heading(level=0)
    title91.add_run("HPE Pointnext Tech Care")
    p9 = doc.add_paragraph()
    p9.add_run("Aplicável aos contratos que contenham algum dos Números de produto abaixo: \nHU4A3AC; HU4A4AC; HU4A5AC; HU4A6AC; HU4A7AC; HU4A8AC; HU4A9AC; HU4B0AC; HU4B1AC; HU4B2AC; HU4B3AC; HU4B4AC; HU4B5AC; HU4B6AC; HU4B7AC; ")
    p9.add_run("\n\nRECURSOS GERAIS").bold = True
    p9.add_run("\n\nTABELA 2.").bold = True
    p9.add_run("Resumo das opções de níveis de serviço.\n")
    table2 = doc.add_table(rows=3, cols=2)
    table2.style = 'Table Grid'
    table2.autofit = False
    table2.allow_autofit = False
    for row2 in table2.rows:
        for cell2 in row2.cells:
            cell2.width = Pt(200)
    table2.cell(0, 0).text = "(Critical) - PNs abaixo: \nHU4A3AC\nHU4A4AC\nHU4A5AC"
    table2.cell(1, 0).text = "(Essential) - PNs abaixo:\nHU4A6AC\nHU4A7AC\nHU4A8AC\nHU4A9AC\nHU4B0AC\nHU4B1AC"
    table2.cell(2, 0).text = "(Basic) - PNs abaixo: \nHU4B2AC\nHU4B3AC\nHU4B4AC\nHU4B5AC\nHU4B6AC\nHU4B7AC"
    table2.cell(0, 1).text = "Resposta em 15 minutos, 24x7, para incidentes com severidade 1 (conecte-se diretamente aos especialistas nos produtos, quando houver disponibilidade) \nGerenciamento de interrupções para incidentes de severidade 1 Compromisso de reparo de hardware em 6 horas, 24x7 (se necessário)"
    table2.cell(1, 1).text = "Resposta em 15 minutos, 24x7, para incidentes com severidade 1 (conecte-se diretamente aos especialistas nos produtos, quando houver disponibilidade) \nAtendimento no local em 4 horas, 24x7."
    table2.cell(2, 1).text = "Resposta em 2 horas, 9x5 (horário comercial padrão) Atendimento no local no próximo dia útil."
    p91 = doc.add_paragraph()
    p91.add_run("\nTodos os níveis de serviço oferecem acesso 24x7 a recursos de autoatendimento e autorresolução de problemas, registro de incidentes 24x7 e, para os dispositivos compatíveis, análises HPE InfoSight e submissão automatizada de incidentes 24x7. As opções de nível de serviço HPE Pointnext Tech Care destacadas são dependentes do produto. A HPE fornecerá os recursos de suporte a hardware para os produtos de hardware cobertos e os recursos de suporte a software para os produtos de software cobertos.")
    p91.add_run("\nAlguns recursos dos serviços talvez não estejam disponíveis em alguns idiomas e localidades. Todos os períodos de cobertura estão sujeitos à disponibilidade local. A elegibilidade do produto pode variar. Entre em contato com um escritório local de vendas da HPE ou com o representante de vendas HPE para obter informações detalhadas sobre a disponibilidade do serviço e a elegibilidade do produto.")
    p91.add_run("\n\nTABELA 3.").bold = True
    p91.add_run("Recursos gerais do serviço\n")  
    table3 = doc.add_table(rows=8, cols=2)
    table3.style = 'Table Grid'
    table3.autofit = False
    table3.allow_autofit = False
    for row3 in table3.rows:
        for cell3 in row3.cells:
            cell3.width = Pt(200)
    table3.cell(0, 0).text = "Recurso"
    table3.cell(1, 0).text = "Acesso pelo telefone a especialistas"
    table3.cell(2, 0).text = "Bate-papo online com especialistas"
    table3.cell(3, 0).text = "Respostas ao fórum dadas por especialistas"
    table3.cell(4, 0).text = "Orientação técnica geral"
    table3.cell(5, 0).text = "Assistência HPE InfoSight"
    table3.cell(6, 0).text = "Alertas preventivos do HPE InfoSight"
    table3.cell(7, 0).text = "Registro automatizado de incidentes"
    table3.cell(0, 1).text = "Especificações de fornecimento"
    table3.cell(1, 1).text = "Os clientes podem entrar em contato com o suporte HPE pelo telefone 24 horas por dia, 7 dias por semana, para registrar incidentes de suporte. O tempo de resposta dependerá do nível de serviço do produto coberto\nResposta aprimorada em 15 minutos, 24x7 (níveis de serviço Critical e Essential)\nPara incidentes de severidade 1, a HPE procura conectar o cliente a um especialista no produto ou ligar novamente para o cliente em até 15 minutos. Para outros incidentes, a HPE pode conectar o cliente a um especialista no produto ou ligar novamente para o cliente em até uma hora.\nResposta padrão em 2 horas (nível de serviço Basic)\nPara ligações sobre produtos cobertos por um contrato de serviço básico, a HPE fornecerá uma resposta por telefone em até 2 horas, realizada por um especialista durante a janela de cobertura. A disponibilidade pode variar para determinados produtos. Consulte hpe.com/services/expertchat para obter detalhes ou entre em contato com o representante de vendas HPE local.\n"
    table3.cell(2, 1).text = "Os clientes podem iniciar um bate-papo online com um recurso técnico especialista para fazer perguntas, obter ajuda ou orientação técnica geral. O bate papo online com especialistas é oferecido para que os clientes obtenham respostas rápidas sobre questões técnicas relacionadas aos seus produtos HPE. \nQuestões complexas que exijam respostas detalhadas podem ser elevadas a incidentes de suporte conforme a necessidade. O bate-papo online com especialistas é limitado apenas ao idioma inglês e está disponível durante a janela de cobertura do serviço. A disponibilidade pode variar para determinados produtos. Consulte hpe.com/services/expertchat para obter detalhes ou entre em contato com o representante de vendas HPE local.\n"
    table3.cell(3, 1).text = "Os clientes podem postar perguntas e problemas, ou discutir o uso dos produtos dentro dos fóruns da comunidade HPE. Os especialistas em produtos HPE respondem em dois dias úteis a qualquer questão não resolvida que for levantada dentro do fórum oficial da comunidade HPE para produtos cobertos pelos serviços de suporte da HPE. \nNos casos de postagens sobre tópicos que deveriam ser abordados por processos de suporte padrão, a HPE solicita que um incidente de suporte formal seja criado e segue os processos padrão de gerenciamento de incidentes da HPE. A resposta do recurso técnico especialista é limitada apenas ao idioma inglês e exige que o usuário seja registrado no HPE Support Center e esteja associado aos contratos de serviço.\n"
    table3.cell(4, 1).text = "A HPE se empenha para oferecer orientação técnica geral relativa às questões e perguntas específicas sobre os tópicos destacados a seguir acerca da operação e gerenciamento dos produtos dos clientes cobertos pelo HPE Pointnext Tech Care. A orientação técnica geral está disponível pelo telefone, internet e bate-papo, está sujeita à janela de cobertura do serviço constante no contrato de serviço e será tratada como um incidente severidade 3.\nQuando for relacionada aos assuntos detalhados ou descritos a seguir, a HPE identificará documentos, vídeos e artigos da base de conhecimento para ajudar nos tópicos abordados.\nAlém de qualquer limitação ou exclusão estabelecida nesta ficha técnica, qualquer orientação técnica geral da HPE será fornecida especificamente para os tópicos detalhados aqui e apenas para os produtos HPE cobertos por esses serviços:\n	Uso ou procedimentos corretos para usar os recursos dos produtos\nAssistência na identificação de documentação relevante ou artigos da base de conhecimento\nConselhos sobre as melhores práticas da HPE para ajudar você a gerenciar e manter os seus produtos	\nNavegação básica para usar a interface de gerenciamento do produto\nConselhos sobre as opções de gerenciamento de capacidade com base nas tendências de uso dos produtos (quando disponível)\nOrientações sobre a configuração geral do produto, que pode incluir recomendações de melhores práticas com base na experiência operacional da HPE\nOrientação sobre os passos potenciais para ajudar a trazer o produto para uma configuração compatível\nOs tópicos de orientação técnica geral mencionados anteriormente podem não ser aplicáveis a todos os produtos de hardware e/ou software cobertos por este serviço.\n"
    table3.cell(5, 1).text = "Para produtos HPE compatíveis com o HPE InfoSight (lista disponível no link a seguir), a HPE oferece suporte e orientação para a preparação, configuração e uso do HPE InfoSight.\nAdicionalmente, para esses produtos conectados, a HPE expande sua orientação técnica geral para incluir análises do HPE InfoSight, bem como os alertas e as recomendações fornecidas. Para produtos HPE configurados, mediante solicitação, a HPE oferece assistência aos clientes para que entendam os problemas, alertas e informações oferecidas pelo HPE InfoSight.\nNos casos em que as análises oferecerem recomendações incluídas nos insights da carga de trabalho HPE InfoSight, a HPE pode oferecer a qualificação da análise, recomendações e os melhores próximos passos gerais em acordo com as orientações técnicas gerais.\nPara mais informações sobre o HPE InfoSight, cobertura de dispositivos e recursos, visite infosight.hpe.com.\n"
    table3.cell(6, 1).text = "Para os produtos HPE cobertos por um contrato de serviço, conectado a e conforme compatibilidade com o HPE InfoSight: Os clientes ganham acesso a rotinas automatizadas de monitoramento que podem identificar problemas potenciais usando assinaturas regras e determinações exclusivas da HPE. \nPara problemas identificados pelo HPE InfoSight, o HPE InfoSight alerta os clientes sobre os problemas e identifica oportunidades para ações corretivas e, de acordo com a sua criticidade, pode automaticamente submeter os incidentes à HPE com as informações diagnosticadas para acelerar os diagnósticos e reparos. \nOs recursos podem variar de acordo com o produto; os dispositivos devem ser compatíveis com o HPE InfoSight, e exige-se conectividade com o HPE InfoSight.\nNos casos em que os clientes configuram o HPE InfoSight para os produtos HPE compatíveis cobertos pelo HPE Pointnext Tech Care, eles ganham acesso aos recursos analíticos aprimorados do HPE InfoSight, que oferecem insights de produtos e alertas de problemas detalhados, além de oportunidades de uso e configuração.\n"
    table3.cell(7, 1).text = "Para produtos HPE compatíveis usando as ferramentas de serviço proprietárias HPE (incluindo o HPE InfoSight), e quando estiverem conectados, os dispositivos podem submeter os incidentes à HPE automaticamente, com informações diagnósticas que podem acelerar os diagnósticos e reparos\nOnde o monitoramento e submissão automática de incidentes identificar problemas críticos que exijam o envolvimento da HPE, a HPE procura responder ao contato do cliente previamente identificado dentro da janela de cobertura do serviço, de acordo com as definições do nível de serviço adquirido. \nCaso o cliente não esteja disponível no momento do contato, ou caso solicite, a HPE agendará o acompanhamento para o próximo dia útil. Todos os problemas não-críticos serão acompanhados no próximo dia útil. Os clientes podem, a qualquer momento, de acordo com o seu nível de serviço, entrar em contato com a HPE para solicitar a continuidade do diagnóstico e da resolução do problema.\nPara mais informações, visite hpe.com/services/getconnected.\n"

    title92 = doc.add_heading(level=0)
    title92.add_run("3.2.	ZONAS DE DESLOCAMENTO ")
    p92 = doc.add_paragraph()
    p92.add_run("Todos os tempos de resposta do Serviço de Troca e com presença no local para hardware se aplicam somente a áreas situadas dentro de uma distância de 160 km de um centro de suporte designado pela HPE. A viagem de descolamento dentro de um raio de 320 km de distância de um centro de suporte designado pela HPE é oferecida sem custos adicionais. Se o local estiver situado a mais de 320 km do centro de suporte designado pela HPE, haverá uma cobrança adicional. As zonas de deslocamentos e os custos, se aplicáveis, podem variar em algumas regiões. Os custos de postagem (em caso de troca de peças), se aplicáveis,  podem variar em algumas regiões. Os tempos de resposta para locais situados a mais de 160 km de um centro de suporte designado pela HPE serão modificados para incluir o tempo de deslocamento, conforme mostra a tabela a seguir. ")
    p92.add_run("\n\nTABELA 9.").bold = True
    p92.add_run("Zonas de deslocamento (exceto nível de serviço crítico)")
    table4 = doc.add_table(rows=6, cols=3)
    table4.style = 'Table Grid'
    table4.autofit = False
    table4.allow_autofit = False
    for row4 in table4.rows:
        for cell4 in row4.cells:
            cell4.width = Pt(180)
    table4.cell(0, 0).text = "Distância do ponto de suporte determinado pela HPE"
    table4.cell(0, 1).text = "Tempo de resposta para Essencial/Essencial Exchange"
    table4.cell(0, 2).text = "Tempo de resposta para Basic/Basic Exchange"
    table4.cell(1, 0).text = "0 a 80 km"
    table4.cell(1, 1).text = "4 horas"
    table4.cell(1, 2).text = "Cobertura no dia seguinte"
    table4.cell(2, 0).text = "81 a 160 km"
    table4.cell(2, 1).text = "4 horas"
    table4.cell(2, 2).text = "Cobertura no dia seguinte"
    table4.cell(3, 0).text = "161 a 320 km"
    table4.cell(3, 1).text = "8 horas"
    table4.cell(3, 2).text = "Mais 1 dia de cobertura"
    table4.cell(4, 0).text = "321 a 480 km"
    table4.cell(4, 1).text = "Estabelecido no momento do pedido e sujeito à disponibilidade"
    table4.cell(4, 2).text = "Mais 2 dias de cobertura"
    table4.cell(5, 0).text = "Mais de 480 km"
    table4.cell(5, 1).text = "Estabelecido no momento do pedido e sujeito à disponibilidade"
    table4.cell(5, 2).text = "Estabelecido no momento do pedido e sujeito à disponibilidade"
    p93 = doc.add_paragraph()
    p93.add_run("\n\nO compromisso de tempo de reparo está disponível para localidades situados até 80 km de um centro de suporte designado pela HPE. Para locais entre 81 e 160 km de tal ponto de suporte, o compromisso de tempo de reparo de hardware será ajustado de acordo com a tabela a seguir. O compromisso de tempo de reparo de hardware não está disponível para locais a mais de 160 km de um ponto de suporte designado pela HPE.")
    p93.add_run("\nTABELA 10").bold = True
    p93.add_run("Zonas de deslocamento para nível de serviço crítico")

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
#doc.add_picture('path_to_your_image.jpg', width=Pt(300))/

# Save the document

doc.save('example_document.docx')