from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE

out = "TCC_AWS_Startup_XYZ.docx"

doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(0.75)
sec.bottom_margin = Inches(0.75)
sec.left_margin = Inches(0.85)
sec.right_margin = Inches(0.85)

# Palette
navy = "17365D"
blue = "2F75B5"
light_blue = "D9EAF7"
pale = "F2F6FA"
gray = "666666"
white = "FFFFFF"

styles = doc.styles
styles["Normal"].font.name = "Aptos"
styles["Normal"].font.size = Pt(10.5)
styles["Normal"].paragraph_format.space_after = Pt(6)
styles["Normal"].paragraph_format.line_spacing = 1.12
for style_name, size in [("Title", 28), ("Heading 1", 17), ("Heading 2", 13), ("Heading 3", 11)]:
    st = styles[style_name]
    st.font.name = "Aptos Display" if style_name in ("Title", "Heading 1") else "Aptos"
    st.font.size = Pt(size)
    st.font.bold = True
    st.font.color.rgb = RGBColor.from_string(navy if style_name in ("Title", "Heading 1") else blue)
    st.paragraph_format.space_before = Pt(12)
    st.paragraph_format.space_after = Pt(5)

# Header/footer
header = sec.header.paragraphs[0]
header.text = "PROJETO TCC AWS  |  CASE 2 — STARTUP XYZ"
header.style = styles["Caption"]
header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
for r in header.runs:
    r.font.color.rgb = RGBColor.from_string(gray)
    r.font.size = Pt(8)

footer = sec.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
footer.add_run("Documento-base para revisão e validação da equipe  •  ")
fld = OxmlElement('w:fldSimple')
fld.set(qn('w:instr'), 'PAGE')
footer._p.append(fld)
for r in footer.runs:
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor.from_string(gray)

def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), fill)
    tcPr.append(shd)

def set_cell_text(cell, text, bold=False, color=None, size=9):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    r.bold = bold
    r.font.name = "Aptos"
    r.font.size = Pt(size)
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

def table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = "Table Grid"
    for i,h in enumerate(headers):
        set_cell_text(t.rows[0].cells[i], h, True, white, 9)
        shade(t.rows[0].cells[i], navy)
    for ri,row in enumerate(rows):
        cells = t.add_row().cells
        for i,val in enumerate(row):
            set_cell_text(cells[i], str(val), False, None, 8.8)
            if ri % 2 == 1:
                shade(cells[i], pale)
    if widths:
        for row in t.rows:
            for i,w in enumerate(widths):
                row.cells[i].width = Inches(w)
    doc.add_paragraph()
    return t

def para(text="", boldlead=None, italic=False):
    p = doc.add_paragraph()
    if boldlead and text.startswith(boldlead):
        p.add_run(boldlead).bold = True
        p.add_run(text[len(boldlead):])
    else:
        r=p.add_run(text)
        r.italic = italic
    return p

def bullet(text, level=0):
    p = doc.add_paragraph(style="List Bullet" if level==0 else "List Bullet 2")
    p.paragraph_format.space_after = Pt(3)
    p.add_run(text)
    return p

def numbered(text):
    p = doc.add_paragraph(style="List Number")
    p.paragraph_format.space_after = Pt(3)
    p.add_run(text)
    return p

# Cover
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(90)
r=p.add_run("PROJETO TCC AWS")
r.bold=True; r.font.size=Pt(16); r.font.color.rgb=RGBColor.from_string(blue)
p = doc.add_paragraph(style="Title")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run("Gestão Segura e Econômica de Documentos para IA")
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r=p.add_run("Case 2 — A Jornada da Startup XYZ")
r.bold=True; r.font.size=Pt(17); r.font.color.rgb=RGBColor.from_string(navy)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(24)
p.add_run("Proposta de arquitetura serverless na Amazon Web Services (AWS)").font.size=Pt(12)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(80)
p.add_run("Equipe: [NOME DA EQUIPE]\n").bold=True
p.add_run("Integrantes: [NOME 1] • [NOME 2] • [NOME 3] • [NOME 4]\n")
p.add_run("Formação: [NOME DO PROGRAMA/CURSO]\n")
p.add_run("Data: [MÊS/ANO]")
doc.add_page_break()

doc.add_heading("Orientação de uso deste documento", level=1)
para("Este documento foi elaborado com base no Guia do Projeto TCC AWS e no enunciado do Case 2 — Gestão de Documentos para IA (Startup XYZ). Ele organiza a parte escrita da proposta para que a equipe possa revisar, adaptar e aproveitar o conteúdo no Google Docs e na apresentação.")
para("Antes da entrega, o grupo deve preencher os campos da capa, confirmar os serviços e configurações que serão realmente implementados e incluir evidências da demonstração. As estimativas de custo e os objetivos de recuperação devem ser calculados/validados pela equipe; não foram inventados valores nem resultados de uma implantação que ainda não foi demonstrada.")
doc.add_heading("Resumo executivo", level=1)
para("A Startup XYZ é uma empresa SaaS em crescimento acelerado que recebe aproximadamente 50 mil novos documentos por mês, principalmente PDFs e imagens, para processamento por um agente de inteligência artificial. Os arquivos têm valor operacional e estratégico: precisam permanecer preservados para compor uma base histórica que poderá apoiar o treinamento de modelos futuros. Ao mesmo tempo, os clientes devem conseguir consultar seus próprios documentos durante os primeiros 12 meses, sem que tenham acesso aos arquivos de outros usuários.")
para("A proposta utiliza uma arquitetura serverless com Amazon API Gateway, AWS Lambda e Amazon S3. O API Gateway expõe os endpoints da aplicação; o Lambda valida a identidade e a autorização do usuário, executa as regras de negócio e gera URLs pré-assinadas para operações autorizadas no S3. Os objetos ficam em um bucket privado, organizados por identificador de usuário. Uma regra de S3 Lifecycle transfere os documentos com mais de 365 dias para uma classe de arquivamento de menor custo, sem configurar expiração ou exclusão automática.")
para("A arquitetura prioriza isolamento entre clientes, menor esforço operacional, elasticidade e controle de custos. A solução também prevê monitoramento, trilha de auditoria, proteção dos dados e procedimentos de restauração. A classe de arquivamento deverá ser escolhida considerando o custo e o tempo de recuperação aceito pelo negócio.")
doc.add_heading("1. Contexto e entendimento do problema", level=1)
doc.add_heading("1.1 O cliente", level=2)
para("A Startup XYZ é uma empresa de software como serviço (SaaS) que está lançando um agente de IA para extração inteligente de dados. A plataforma recebe documentos enviados pelos clientes e utiliza esses arquivos como insumo para o processamento da inteligência artificial.")
doc.add_heading("1.2 Problema de negócio", level=2)
para("A empresa estima receber 50 mil novos arquivos por mês e precisa armazená-los de forma segura, mantendo o acesso do cliente aos documentos durante um ano. Diferentemente de um repositório temporário, a base deverá preservar seu histórico, pois os arquivos poderão ter valor para o treinamento de modelos futuros. O desafio é manter esse acervo sem que o custo de armazenamento cresça de forma insustentável.")
doc.add_heading("1.3 Principais desafios", level=2)
for x in [
    "Receber documentos de forma confiável e permitir que o usuário consulte apenas os próprios arquivos.",
    "Suportar o crescimento do volume sem provisionar servidores permanentemente.",
    "Preservar os arquivos históricos, sem exclusão automática por política de ciclo de vida.",
    "Reduzir o custo de armazenamento dos documentos com mais de 365 dias.",
    "Evitar exposição pública do bucket e acessos indevidos entre clientes.",
    "Permitir recuperação de documentos arquivados, considerando que classes de arquivamento podem exigir restauração antes do download."
]:
    bullet(x)
doc.add_heading("1.4 Impactos se o problema não for resolvido", level=2)
para("Uma solução inadequada pode gerar vazamento de documentos entre clientes, perda de dados de valor histórico, indisponibilidade durante picos de uso, aumento desnecessário dos custos e dificuldade para recuperar arquivos antigos. Esses riscos afetam a confiança dos clientes, a continuidade do serviço e a capacidade da Startup XYZ de utilizar o acervo como ativo estratégico.")
doc.add_heading("2. Requisitos da solução", level=1)
doc.add_heading("2.1 Requisitos funcionais", level=2)
table(["ID", "Requisito", "Critério de atendimento proposto"], [
("RF01", "Receber documentos PDF pequenos e a identificação do usuário.", "Endpoint de API para iniciar o envio e validar a solicitação."),
("RF02", "Armazenar documentos na AWS.", "Guardar os objetos em bucket Amazon S3 privado."),
("RF03", "Consultar documentos do próprio usuário.", "Lambda valida identidade, propriedade do objeto e autorização antes de permitir a leitura."),
("RF04", "Disponibilizar acesso temporário ao arquivo autorizado.", "Gerar URL pré-assinada com validade curta, após a validação de acesso."),
("RF05", "Manter os documentos históricos.", "Não configurar regra de expiração; restringir exclusões por política de IAM e processo operacional."),
("RF06", "Arquivar documentos após 365 dias.", "Regra de S3 Lifecycle para transição a uma classe de arquivamento definida pela equipe."),
("RF07", "Registrar operações relevantes.", "Logs de aplicação e métricas no Amazon CloudWatch; auditoria de API conforme configuração adotada.")
])
doc.add_heading("2.2 Requisitos não funcionais", level=2)
table(["ID", "Requisito", "Diretriz"], [
("RNF01", "Segurança e privacidade", "Bucket privado, Block Public Access, criptografia, TLS e acesso mínimo necessário."),
("RNF02", "Isolamento entre clientes", "Validar o usuário autenticado no backend; não confiar em um ID de usuário enviado livremente pelo cliente."),
("RNF03", "Escalabilidade", "Utilizar serviços gerenciados/serverless e evitar servidores permanentemente provisionados."),
("RNF04", "Custo controlado", "Transicionar objetos antigos para armazenamento de arquivamento e monitorar consumo."),
("RNF05", "Durabilidade e preservação", "Manter objetos sem expiração automática e definir controles contra exclusão acidental ou indevida."),
("RNF06", "Observabilidade", "Acompanhar erros, latência, invocações, requisições e custos."),
("RNF07", "Recuperabilidade", "Documentar o processo de restauração de objetos arquivados e testar o procedimento.")
])
doc.add_heading("3. Visão geral da solução", level=1)
para("A proposta separa a camada de entrada e regras de negócio da camada de armazenamento. O cliente solicita uma operação por meio do API Gateway. A função Lambda valida a identidade e verifica se a ação solicitada é permitida. Para upload, o backend pode emitir uma URL pré-assinada de curta duração, e o navegador ou aplicativo envia o arquivo diretamente ao S3. Essa abordagem evita trafegar o conteúdo do arquivo dentro da função Lambda e reduz a dependência dos limites de payload da API.")
para("Na consulta, o cliente informa o identificador do documento. O Lambda determina o usuário autenticado, verifica se o objeto pertence a ele e só então autoriza o acesso, por exemplo emitindo uma URL pré-assinada temporária. O bucket permanece privado; a URL não substitui a autorização do backend, e sua validade deve ser curta e compatível com o uso previsto.")
para("Após 365 dias, uma regra de ciclo de vida transfere os objetos para uma classe de arquivamento de menor custo. A transição altera a classe de armazenamento, mas não apaga o arquivo. Quando um documento arquivado precisar ser consultado, a aplicação deverá informar ao usuário se é necessário solicitar a restauração e aguardar o prazo da classe escolhida.")
doc.add_heading("4. Arquitetura proposta", level=1)
doc.add_heading("4.1 Diagrama lógico", level=2)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run(
"CLIENTE / APLICAÇÃO SAAS\n"
"        │ HTTPS\n"
"        ▼\n"
"AMAZON API GATEWAY\n"
"        │ requisição autenticada\n"
"        ▼\n"
"AWS LAMBDA — validação, autorização e regras de negócio\n"
"        │                         │\n"
"        │ URL pré-assinada        └── logs e métricas → AMAZON CLOUDWATCH\n"
"        ▼\n"
"AMAZON S3 — BUCKET PRIVADO\n"
"  ├── prefixo por usuário (ex.: usuario-123/)\n"
"  ├── documentos recentes\n"
"  └── Lifecycle após 365 dias → classe de arquivamento\n"
)
r.font.name="Consolas"; r.font.size=Pt(9); r.font.color.rgb=RGBColor.from_string(navy)
para("Fluxo de upload recomendado: cliente → API Gateway → Lambda (validação e geração de URL) → cliente envia o arquivo diretamente ao S3 usando HTTPS. Fluxo de consulta: cliente → API Gateway → Lambda (autorização e verificação de propriedade) → URL pré-assinada temporária → S3. A regra Lifecycle do S3 executa a transição de classe sem depender de uma execução diária da aplicação.")
doc.add_heading("4.2 Serviços AWS e suas funções", level=2)
table(["Serviço", "Função na arquitetura", "Justificativa e cuidados"], [
("Amazon API Gateway", "Publica endpoints para iniciar upload, consultar documentos e solicitar restauração.", "Gerencia a entrada das requisições. Deve usar HTTPS, limites de requisição e mecanismos de autenticação/autorização adequados."),
("AWS Lambda", "Executa a lógica de negócio: valida identidade, verifica propriedade, cria chaves de objeto e gera URLs pré-assinadas.", "Execução sob demanda e sem gerenciamento de servidores. A função deve usar permissões mínimas e não aceitar como confiável um user_id arbitrário enviado pelo cliente."),
("Amazon S3", "Armazena os documentos e aplica a política de ciclo de vida.", "Armazenamento de objetos durável e escalável. Manter Block Public Access ativado e controlar permissões no bucket e nos objetos."),
("Amazon S3 Lifecycle", "Transfere documentos com mais de 365 dias para uma classe de arquivamento.", "Reduz custo de armazenamento de longo prazo. A regra deve fazer transição, não expiração/exclusão."),
("AWS IAM", "Define as permissões da função e dos operadores.", "Aplicar privilégio mínimo. A função deve acessar somente os recursos e operações necessários."),
("Amazon CloudWatch", "Centraliza logs, métricas e alarmes da aplicação.", "Ajuda a detectar falhas, latência elevada e erros de acesso. Evitar registrar conteúdo dos documentos, tokens ou dados pessoais desnecessários."),
("AWS CloudTrail (complementar)", "Registra chamadas de API da AWS para auditoria.", "Ajuda a investigar alterações e acessos administrativos. A retenção e o custo dos logs devem ser definidos.")
])
doc.add_heading("5. Organização e fluxo dos dados", level=1)
doc.add_heading("5.1 Organização dos objetos", level=2)
para("Os objetos serão organizados por identificador interno do usuário, por exemplo: usuario-123/2026/10/arquivo-uuid.pdf. O nome do arquivo não deve ser usado sozinho como identificador, pois pode se repetir. Recomenda-se gerar uma chave única no backend e guardar metadados mínimos necessários para localizar o objeto, como ID do documento, ID do proprietário, data de envio e estado de arquivamento.")
para("O prefixo facilita a organização, mas não constitui sozinho uma barreira de segurança. O backend deve validar a identidade autenticada e a relação de propriedade antes de autorizar qualquer operação. O cliente não deve conseguir trocar o identificador do usuário na requisição e obter acesso a outro prefixo.")
doc.add_heading("5.2 Fluxo de envio", level=2)
for x in [
    "O cliente autenticado solicita o início do envio por meio do API Gateway.",
    "O Lambda valida a identidade, verifica as regras de negócio e gera uma chave de objeto controlada pelo servidor.",
    "O Lambda retorna uma URL pré-assinada com validade limitada, restrita à operação e ao objeto esperado.",
    "O cliente envia o PDF ao S3 usando HTTPS.",
    "A aplicação registra o resultado e disponibiliza o documento na lista do proprietário."
]:
    numbered(x)
doc.add_heading("5.3 Fluxo de consulta", level=2)
for x in [
    "O cliente solicita a consulta de um documento pelo endpoint da API.",
    "O Lambda identifica o usuário autenticado e verifica se ele é o proprietário autorizado.",
    "Se o objeto estiver em uma classe de acesso imediato, o backend retorna uma URL pré-assinada temporária.",
    "Se o objeto estiver arquivado e precisar de restauração, a aplicação informa essa condição e inicia ou orienta o fluxo de restauração autorizado.",
    "Depois da restauração, o acesso temporário é liberado conforme a política definida."
]:
    numbered(x)
doc.add_heading("6. Justificativa técnica e alternativas", level=1)
table(["Decisão", "Motivo", "Alternativa considerada / limitação"], [
("Arquitetura serverless", "Reduz a necessidade de administrar servidores e acompanha variações de demanda.", "EC2 com aplicação própria daria maior controle, mas exigiria manutenção, escalabilidade e correções do sistema operacional."),
("S3 como repositório principal", "Adequado para armazenar arquivos e aplicar políticas de ciclo de vida.", "Um sistema de arquivos ou banco relacional não é a opção principal mais simples para guardar grande volume de PDFs e imagens como objetos."),
("URL pré-assinada", "Permite upload/download temporário sem tornar o bucket público.", "Enviar todos os bytes pela API/Lambda pode adicionar custo, latência e limites de payload; por isso, o envio direto ao S3 é recomendado."),
("Arquivamento após 365 dias", "Alinha a classe de armazenamento à redução esperada de acessos após um ano.", "Classes de arquivamento podem exigir restauração e ter cobrança/prazos próprios; a escolha depende do tempo de recuperação aceito."),
("Prefixo por usuário + autorização no backend", "Ajuda a organizar e separar documentos por cliente.", "Prefixo isolado não basta: a autorização deve ocorrer em cada operação e as permissões IAM precisam ser restritas."),
("Sem expiração automática", "Preserva o valor histórico indicado no case.", "O custo total continuará aumentando com o volume acumulado; é necessário acompanhar crescimento, custos de requisição, recuperação e logs.")
])
para("A decisão final sobre a classe de arquivamento deve comparar custo por GB, cobranças de transição, duração mínima de armazenamento, custo de recuperação e tempo para disponibilização. S3 Glacier Deep Archive pode ser uma opção para dados raramente acessados e com tolerância a espera maior; uma classe Glacier com recuperação mais rápida pode ser mais adequada se a experiência do cliente exigir menor espera. A equipe deve escolher e justificar a opção final.")
doc.add_heading("7. Mapeamento dos requisitos", level=1)
table(["Requisito", "Como será atendido", "Como demonstrar"], [
("Upload de documentos", "API Gateway + Lambda para autorização e URL pré-assinada; upload direto ao S3.", "Enviar um PDF de teste e confirmar a presença do objeto no bucket."),
("Acesso individual", "Autenticação, validação de propriedade no Lambda e URL temporária.", "Testar dois usuários e comprovar que um não acessa o documento do outro."),
("Bucket não público", "S3 Block Public Access ativado e políticas restritivas.", "Mostrar as configurações do bucket e testar acesso não autorizado."),
("Preservação histórica", "Sem regra de expiração; restrição de DeleteObject; versionamento e controles adicionais conforme política.", "Revisar regras Lifecycle e políticas; explicar controles contra exclusão."),
("Arquivamento aos 365 dias", "Regra S3 Lifecycle para transição, sem expiração.", "Mostrar a regra configurada; em demonstração, pode-se usar objeto de teste e regra temporária controlada, sem alterar a política de produção."),
("Escalabilidade", "API Gateway, Lambda e S3 gerenciados.", "Mostrar métricas de requisições e invocações, ou explicar o comportamento esperado se não houver teste de carga."),
("Monitoramento", "CloudWatch Logs, métricas e alarmes definidos pela equipe.", "Mostrar logs e um alarme de teste, se implementado.")
])
doc.add_heading("8. Escalabilidade e resiliência", level=1)
doc.add_heading("8.1 Crescimento da demanda", level=2)
para("A estimativa inicial é de 50 mil arquivos por mês, equivalente a aproximadamente 600 mil arquivos por ano se o ritmo permanecer constante. Como o enunciado não informa o tamanho médio dos arquivos, não é possível calcular com precisão o volume total em GB/TB. Esse dado deve ser medido ou estimado pela equipe antes da análise de custos.")
para("API Gateway e Lambda são serviços gerenciados que podem acompanhar variações de tráfego dentro de suas cotas e limites de serviço. O S3 foi projetado para armazenar grandes volumes de objetos. Ainda assim, a solução deve monitorar limites, erros, concorrência do Lambda, throttling e latência. Se houver picos relevantes, as cotas e a concorrência precisam ser avaliadas antes da entrada em produção.")
doc.add_heading("8.2 Resiliência e falhas", level=2)
for x in [
    "Se uma invocação do Lambda falhar, a aplicação deve retornar uma mensagem controlada, registrar o erro e permitir uma nova tentativa segura.",
    "Operações de upload devem usar chaves únicas e tratar reenvios para evitar sobrescritas acidentais.",
    "A equipe deve definir como identificar uploads incompletos e como limpar metadados de operações que não terminaram, sem apagar documentos históricos válidos.",
    "Os serviços gerenciados reduzem a necessidade de administrar instâncias, mas não eliminam falhas de configuração, indisponibilidades regionais ou limites de serviço.",
    "Para requisitos de continuidade mais rigorosos, avaliar replicação entre regiões e cópias independentes, após análise de custo, residência dos dados e objetivo de recuperação."
]:
    bullet(x)
doc.add_heading("9. Segurança e privacidade", level=1)
doc.add_heading("9.1 Identidade e autorização", level=2)
para("A autenticação pode ser integrada ao mecanismo de identidade já utilizado pela plataforma SaaS. Se o case ainda não possuir esse componente, Amazon Cognito pode ser avaliado como complemento; sua adoção não é requisito explícito do enunciado e deve ser validada pelo grupo. O API Gateway deve encaminhar ao backend a identidade verificada. O Lambda autoriza a operação com base nessa identidade e na propriedade do documento.")
doc.add_heading("9.2 Controles mínimos", level=2)
for x in [
    "Manter S3 Block Public Access ativado e evitar políticas que permitam leitura pública.",
    "Exigir HTTPS/TLS na comunicação entre cliente e endpoints.",
    "Ativar criptografia em repouso no S3; avaliar SSE-S3 como configuração simples ou SSE-KMS quando houver necessidade de controle de chaves e auditoria mais específica.",
    "Aplicar privilégio mínimo nas roles IAM e restringir a função às operações e ao bucket necessários.",
    "Gerar URLs pré-assinadas com validade curta, escopo restrito e somente após autorização do proprietário.",
    "Não registrar em logs tokens, URLs pré-assinadas completas, conteúdo de arquivos ou dados pessoais desnecessários.",
    "Restringir e auditar ações administrativas, incluindo alteração de políticas e tentativa de exclusão.",
    "Definir uma política de retenção de logs e acesso operacional compatível com a necessidade de auditoria."
]:
    bullet(x)
doc.add_heading("9.3 Preservação e proteção contra exclusão", level=2)
para("A exigência do case é que nenhum arquivo seja deletado. Portanto, a regra Lifecycle deve realizar somente a transição de classe de armazenamento. Não deve existir regra de expiração automática. Para reduzir exclusões acidentais, recomenda-se restringir a permissão s3:DeleteObject por IAM e separar as funções operacionais. O versionamento do S3 pode ajudar na recuperação de sobrescritas e exclusões, mas não impede por si só toda exclusão definitiva. Se for necessária imutabilidade formal, a equipe deve avaliar S3 Object Lock e definir o modo e o período de retenção apropriados.")
doc.add_heading("10. Custos e otimização", level=1)
para("Os principais componentes de custo previstos são armazenamento no S3, requisições de API, invocações do Lambda, transferência de dados, logs e operações de arquivamento/restauração. A participação de cada item dependerá do tamanho médio dos documentos, da quantidade de consultas e downloads, da região AWS, da classe escolhida e do tempo de retenção.")
doc.add_heading("10.1 Estratégias de otimização", level=2)
for x in [
    "Aplicar a transição para arquivamento após 365 dias, conforme o requisito.",
    "Escolher a classe de arquivamento a partir do padrão de acesso e do tempo de recuperação aceitável.",
    "Evitar transferir o conteúdo do arquivo pelo Lambda; preferir URL pré-assinada e upload/download direto ao S3.",
    "Definir retenção adequada para logs e métricas, evitando coleta excessiva.",
    "Monitorar armazenamento, requisições e restaurações com alertas de orçamento e AWS Cost Explorer.",
    "Usar AWS Pricing Calculator para estimar cenários de volume e acesso antes de fechar a proposta."
]:
    bullet(x)
doc.add_heading("10.2 Premissas para estimativa", level=2)
table(["Variável", "Valor conhecido ou a levantar", "Impacto na estimativa"], [
("Novos arquivos por mês", "50.000 (informado no case)", "Determina o crescimento do número de objetos."),
("Arquivos por ano", "Cerca de 600.000, mantendo o ritmo mensal", "Ajuda a projetar o acervo acumulado."),
("Tamanho médio do arquivo", "A levantar por amostra ou requisito do produto", "Principal variável para estimar GB/TB armazenados."),
("Consultas e downloads mensais", "A levantar", "Afeta requisições e custos de transferência/recuperação."),
("Classe de arquivamento", "A decidir após comparar opções", "Afeta custo por GB, restauração e tempo de acesso."),
("Região AWS e retenção de logs", "A definir", "Afeta preços e custos operacionais.")
])
para("Não se apresenta um valor monetário final neste documento porque o case não informa tamanho médio dos arquivos, volume de consultas, região AWS ou classe de arquivamento selecionada. A equipe deve preencher esses parâmetros no AWS Pricing Calculator e registrar a data da simulação.")
doc.add_heading("11. Backup, recuperação e continuidade", level=1)
doc.add_heading("11.1 Proteção dos dados", level=2)
para("A preservação do acervo será apoiada pela ausência de expiração automática, restrição de permissões de exclusão e, conforme a política aprovada, versionamento ou controles de imutabilidade. Esses mecanismos têm funções diferentes: versionamento ajuda a recuperar versões anteriores, enquanto Object Lock pode impor retenção contra exclusão/alteração conforme sua configuração. A estratégia precisa ser validada por testes.")
doc.add_heading("11.2 Recuperação de documentos arquivados", level=2)
para("Quando um arquivo estiver em uma classe que não permita leitura imediata, o sistema deverá iniciar ou orientar uma solicitação de restauração. O documento só será disponibilizado após a restauração terminar. O prazo depende da classe e da opção de recuperação escolhida; por isso, não deve ser prometido um tempo fixo sem a seleção da classe e a validação do procedimento.")
doc.add_heading("11.3 Objetivos de recuperação", level=2)
para("O RPO (perda de dados aceitável) e o RTO (tempo máximo aceitável para restabelecer o serviço) não foram definidos no enunciado. A equipe deve propor metas realistas, obter aprovação do grupo/cliente e testar a recuperação antes de declarar esses objetivos como atendidos.")
table(["Cenário", "Resposta planejada", "Validação necessária"], [
("Erro de aplicação", "Corrigir ou reverter a função e monitorar falhas no CloudWatch.", "Testar resposta de erro e recuperação do endpoint."),
("Exclusão ou sobrescrita indevida", "Recuperar versão, se o versionamento estiver ativado, e revisar permissões.", "Executar teste controlado em ambiente de demonstração."),
("Documento arquivado solicitado", "Iniciar restauração e liberar acesso após conclusão.", "Medir o tempo real conforme a classe selecionada."),
("Falha regional grave", "Avaliar cópia/replicação em outra região conforme RTO/RPO aprovados.", "Definir custo e realizar exercício de recuperação, se implementado.")
])
doc.add_heading("12. Monitoramento e operação", level=1)
para("O Amazon CloudWatch será usado para acompanhar logs e métricas do API Gateway e Lambda. A equipe deve considerar alarmes para aumento de erros, latência, throttling e falhas de execução. O monitoramento do S3 e dos custos deve permitir observar o crescimento do acervo, a quantidade de solicitações e o uso das classes de armazenamento.")
para("Como diferencial, recomenda-se documentar a infraestrutura e, se houver tempo, automatizar a criação dos recursos com AWS CDK ou AWS CloudFormation. A infraestrutura como código facilita repetição, revisão e recuperação do ambiente, mas só deve ser apresentada como implementada se houver evidências.")
doc.add_heading("13. Plano de demonstração", level=1)
para("O guia do TCC destaca que não basta afirmar que a solução funciona: a equipe deve mostrar evidências. Recomenda-se preparar uma demonstração curta, com dados de teste e sem documentos reais ou sensíveis.")
for x in [
    "Mostrar os endpoints da API e o fluxo de autenticação adotado.",
    "Enviar um PDF de teste e confirmar que o objeto foi criado no bucket correto.",
    "Mostrar que Block Public Access está ativado e que o acesso direto não autorizado é negado.",
    "Testar com dois usuários de teste: o usuário A não deve conseguir consultar um objeto pertencente ao usuário B.",
    "Demonstrar a geração de URL pré-assinada e explicar sua validade limitada.",
    "Mostrar a regra Lifecycle configurada para transição após 365 dias, sem regra de expiração.",
    "Mostrar logs ou métricas relevantes no CloudWatch.",
    "Se o arquivamento/restauração não tiver sido implementado por limitação de tempo, apresentar a configuração planejada e identificar claramente que se trata de uma etapa proposta, não demonstrada."
]:
    numbered(x)
doc.add_heading("14. Evolução futura", level=1)
for x in [
    "Adicionar Amazon Cognito ou integrar o provedor de identidade corporativo, se necessário.",
    "Automatizar a infraestrutura com AWS CDK ou CloudFormation e manter o código versionado.",
    "Criar dashboards e alarmes de observabilidade e custo.",
    "Avaliar eventos do S3 e processamento assíncrono para integração futura com etapas de extração por IA, caso esse escopo seja aprovado.",
    "Avaliar replicação entre regiões, Object Lock e controles adicionais de governança conforme os requisitos de continuidade e retenção.",
    "Executar testes de carga e segurança antes da disponibilização para clientes reais."
]:
    bullet(x)
doc.add_heading("15. Conclusão", level=1)
para("A arquitetura proposta atende ao objetivo central do Case 2 ao combinar uma camada de API serverless com armazenamento de objetos no Amazon S3 e uma política de arquivamento de longo prazo. O API Gateway recebe as solicitações, o Lambda aplica as regras de autorização e o S3 mantém os documentos em um bucket privado. O ciclo de vida transfere arquivos com mais de 365 dias para uma classe de menor custo, sem configurar exclusão automática.")
para("A principal decisão de segurança é não tratar o prefixo do objeto como única barreira: cada consulta deve validar a identidade e a propriedade do documento antes de emitir uma URL temporária. A principal decisão de custo é alinhar a classe de armazenamento ao padrão de acesso, considerando o custo e o prazo de restauração. Com validação de permissões, monitoramento, testes de recuperação e estimativa de custos, a proposta oferece uma base serverless escalável e alinhada ao valor histórico do acervo.")
doc.add_heading("16. Referências e premissas", level=1)
bullet("Manual do Aluno: Projeto TCC AWS — Guia Completo para Entrega e Apresentação (material fornecido pela formação).")
bullet("Cases Cloud — Case 2: Case de Arquitetura: A Jornada da Startup XYZ / Gestão de Documentos para IA (material fornecido pela formação).")
para("Documentação oficial AWS recomendada para validação pela equipe antes da entrega: Amazon S3 User Guide (Block Public Access, presigned URLs, Lifecycle, classes de armazenamento e Object Lock); AWS Lambda Developer Guide; Amazon API Gateway Developer Guide; AWS Identity and Access Management User Guide; Amazon CloudWatch User Guide; AWS Pricing Calculator.")
doc.add_heading("Apêndice A — Checklist de revisão do grupo", level=1)
for x in [
    "Preencher nome da equipe, integrantes, curso e data.",
    "Confirmar se a autenticação será feita por serviço existente ou por Amazon Cognito.",
    "Escolher a classe de arquivamento e documentar prazo de restauração e custos.",
    "Revisar as políticas IAM e comprovar o isolamento entre dois usuários.",
    "Verificar que não existe regra Lifecycle de expiração/exclusão.",
    "Definir os valores de RTO e RPO, se exigidos, e testar a recuperação.",
    "Executar AWS Pricing Calculator com premissas documentadas.",
    "Inserir diagrama visual definitivo e capturas reais da demonstração.",
    "Diferenciar claramente o que foi implementado, testado e apenas proposto."
]:
    bullet("☐ " + x)

# Set heading keep with next; widow control-ish
for p in doc.paragraphs:
    if p.style.name.startswith("Heading"):
        p.paragraph_format.keep_with_next = True

doc.save(out)
print(out)
