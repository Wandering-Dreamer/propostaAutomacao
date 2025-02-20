# Import docx NOT python-docx 
from docx import *
from docx.shared import Pt, Inches, Mm
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date
import locale

# Create an instance of a word document 
doc = Document() 

title3 = doc.add_heading(level=0)
title3.add_run("Aviso de Confidencialidade")
p3 = doc.add_paragraph()
p3.add_run("As informações contidas em todas as páginas deste documento / proposta é confidencial da Hewlett Packard Enterprise e Hewlett Packard Enterprise Company (a seguir coletivamente \"Hewlett Packard Enterprise\") e seguem para fins de avaliação. Ao receber o documento, o destinatário concorda em manter tais informações em sigilo e não reproduzir ou divulgar a qualquer pessoa fora do grupo diretamente responsável pela avaliação do conteúdo, a menos que a  Hewlett Packard Enterprise tenha autorizado. Não há obrigação de manter a confidencialidade de qualquer parte da informação que o destinatário tenha tido conhecimento sem restrições antes do recebimento deste documento, como é provado através de registos escritos, de negócios ou informações de conhecimento público sem que o destinatário tenha incorrido em faltas, ou que tenha sido recebido pelo destinatário através de uma terceira parte sem restrições.").font.size = Pt(9)
p3.add_run("""Este documento contém informações sobre produtos, vendas e programas de serviço da  Hewlett Packard Enterprise que podem ser melhorados
            ou descontinuados a critério exclusivo da  Hewlett Packard Enterprise. A  Hewlett Packard Enterprise tem feito todos os esforços para incluir 
           materiais aqui considerados confiáveis e relevantes para fins de avaliação de seu destinatário. Nem a Hewlett Packard Enterprise nem seus representantes
            dão qualquer garantia quanto à exatidão ou completude das informações. Portanto, este documento é apenas para fins informativos devendo ser considerado
            para os negócios da  Hewlett Packard Enterprise. Nem a  Hewlett Packard Enterprise nem seus representantes serão responsáveis sobre qualquer ato 
           do destinatário ou de seus representantes, como resultado do uso das informações aqui fornecidas. A assinatura de um acordo definitivo ou
            assinatura de aceitação da proposta, por representantes autorizados das partes, será o único meio pelo qual a  Hewlett Packard Enterprise ou
            suas afiliadas serão vinculadas à proposta/ contrato.""").font.size = Pt(9)
title31 = doc.add_heading(level = 0).add_run("Restrições de cópias entregues da Proposta")
p3.add_run("""A proposta da Hewlett Packard Enterprise foi enviada em formato eletrônico no formato de arquivo PDF. Se o conteúdo dos arquivos originais forem 
           diferentes da versão em PDF, somente o conteúdo da versão PDF será respeitado pela Hewlett Packard Enterprise.""").font.size = Pt(9)
title32 = doc.add_heading(level = 0).add_run("Esclarecimentos")
p3.add_run("Dúvidas ou  esclarecimentos sobre esta Política de Privacidade, entre em contato com seu representante de vendas.").font.size = Pt(9)
p3.add_run("© Copyright 2025 Hewlett-Packard Development Company, L.P.").font.size = Pt(9)
p3.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

# Now save the document to a location 
doc.save('test.docx')
