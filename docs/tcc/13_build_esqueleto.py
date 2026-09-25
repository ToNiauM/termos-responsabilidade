# -*- coding: utf-8 -*-
"""
Gera o esqueleto CORRIGIDO do TCC (versão 13) a partir do template oficial
07_Template TCC_PT (251, 252) (1).docx, incorporando os resultados da campanha
censitária na posição de 22 set. 2026 (relatorio_20260922_211049.xlsx).
Versão 12 = versão 11 + correções da análise das regras da banca (11_analise_regras_banca.md):
referências conferidas no Crossref, acórdão errado retirado, Resumo/Abstract <= 250 palavras,
justificativa CEP (Res. CNS 510/2016), normas nas Referências, pretérito perfeito, Conclusões sem
repetir números, notas de revisão removidas, nota da Tabela 9 (29), ordem das citações múltiplas.
Versão 13 = versão 12 + subtítulos de Resultados e Discussão espelhando os da Metodologia (Manual p. 43):
a Metodologia ganhou um subtítulo por bloco de resultados e os Resultados foram renomeados/fundidos.

Pré-requisito: .venv/bin/python 11_make_figuras.py  (gera 11_figuras/ e 11_metricas.json)
Uso:           .venv/bin/python 13_build_esqueleto.py
Saída:         13_TCC_Esqueleto_corrigido.docx

Os arquivos 10_ e 11_ NÃO são alterados; figuras e métricas continuam as de 11_.
"""
import json
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

TEMPLATE = "07_Template TCC_PT (251, 252) (1).docx"
OUT = "13_TCC_Esqueleto_corrigido.docx"
FIG = "11_figuras/"
RED = RGBColor(0xC0, 0x00, 0x00)
M = json.load(open("11_metricas.json", encoding="utf-8"))

def br(v, d=0):
    s = f"{v:,.{d}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")

doc = Document(TEMPLATE)

# ---------------------------------------------------------------- limpeza
body = doc.element.body
for el in list(body):
    if el.tag.endswith("}sectPr"):
        continue
    body.remove(el)

for sec in doc.sections:
    for hdr in (sec.header, sec.first_page_header, sec.even_page_header):
        for p in hdr.paragraphs:
            if "Nome do curso" in p.text:
                for r in p.runs:
                    r.text = (r.text
                              .replace("_________ (Nome do curso)", "Data Science e Analytics")
                              .replace("____ (ano da defesa)", "2026"))

st = doc.styles["Normal"]
st.font.name = "Arial"
st.font.size = Pt(11)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")

# ---------------------------------------------------------------- helpers
def para(text="", size=11, bold=False, italic=False, align="justify",
         indent=None, spacing=1.5, space_after=0, color=None, center=False):
    p = doc.add_paragraph()
    fmt = p.paragraph_format
    fmt.alignment = {"justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
                     "left": WD_ALIGN_PARAGRAPH.LEFT,
                     "center": WD_ALIGN_PARAGRAPH.CENTER}["center" if center else align]
    if spacing == 1.5:
        fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    else:
        fmt.line_spacing_rule = WD_LINE_SPACING.SINGLE
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(space_after)
    if indent:
        fmt.first_line_indent = Cm(indent)
    if text:
        r = p.add_run(text)
        r.font.size = Pt(size)
        r.bold = bold
        r.italic = italic
        if color:
            r.font.color.rgb = color
    return p

def body_p(text):
    return para(text, indent=1.25)

def title_sec(text):
    return para(text, bold=True, align="left")

def subtitle(text):
    return para(text, bold=True, align="left", indent=1.25)

def blank(spacing=1.5):
    return para("", spacing=spacing)

def nota(texto):
    """Nota de revisão destacada em vermelho — REMOVER antes da submissão."""
    return para("[NOTA DE REVISÃO — " + texto + " Remover esta nota antes da submissão.]",
                italic=True, color=RED, indent=1.25)

def caption(text):
    return para(text, align="justify", spacing=1.0)

def fig(path, legenda, fonte="Fonte: Resultados originais da pesquisa", nota_txt=None, largura=13.0):
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.add_run().add_picture(FIG + path, width=Cm(largura))
    caption(legenda)
    caption(fonte)
    if nota_txt:
        caption("Nota: " + nota_txt)
    blank(1.0)

def _border(el, edges):
    tcPr = el.get_or_add_tcPr() if hasattr(el, "get_or_add_tcPr") else el
    borders = OxmlElement("w:tcBorders")
    for edge, val in edges.items():
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), val)
        e.set(qn("w:sz"), "8")
        e.set(qn("w:color"), "000000")
        borders.append(e)
    tcPr.append(borders)

def table(rows, col_widths=None, header=True):
    """Tabela padrão Esalq: bordas apenas acima/abaixo do cabeçalho e no fim;
    sem negrito; 1ª coluna à esquerda, demais centralizadas (números à direita)."""
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = col_widths is None
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = t.cell(i, j)
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.space_before = Pt(2)
            if j == 0:
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            elif i == 0:
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                txt = str(val)
                num = (txt.replace(".", "").replace(",", "").replace("%", "").replace("R$", "")
                       .replace("≈", "").replace("–", "").replace(" ", "").replace("s", "").replace("h", ""))
                p.paragraph_format.alignment = (WD_ALIGN_PARAGRAPH.RIGHT
                                                if num.isdigit() else WD_ALIGN_PARAGRAPH.JUSTIFY)
            r = p.add_run(str(val))
            r.font.name = "Arial"
            r.font.size = Pt(11)
            edges = {}
            if header and i == 0:
                edges = {"top": "single", "bottom": "single"}
            if i == len(rows) - 1:
                edges["bottom"] = "single"
            edges.setdefault("left", "nil"); edges.setdefault("right", "nil")
            if not edges.get("top"): edges["top"] = "nil"
            if not edges.get("bottom"): edges["bottom"] = "nil"
            _border(cell._tc, edges)
    if col_widths:
        for j, w in enumerate(col_widths):
            for i in range(len(rows)):
                t.cell(i, j).width = Cm(w)
    return t

def page_break():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

def tit(s):
    """Capitaliza rótulos de localização exportados em caixa alta."""
    small = {"de", "da", "do", "e"}
    out = []
    for w in s.lower().split(" "):
        if w in small: out.append(w)
        elif w.startswith("s") and len(w) == 2 and w[1].isdigit(): out.append(w.upper())
        elif w in ("vp", "tv", "cfc", "cgti", "ccom", "cdprof", "cci", "digeo", "cofis", "projur",
                   "cotec", "colog", "cpd", "cge", "sepat", "coreg", "egc", "depes", "depad",
                   "delic", "setran", "decont", "degep", "seprot", "coad", "defin", "diepg",
                   "cgov", "depev"): out.append(w.upper())
        else: out.append(w.capitalize())
    return " ".join(out)

# ================================================================ PÁGINA 1
TITULO = ("Automação de inventários no setor público com Data Science e "
          "infraestrutura open source replicável")

para(TITULO, bold=True, center=True, spacing=1.0, space_after=12)
blank(1.0)
para("Antônio Rodrigues de Sousa Júnior¹*; Gabriel Gomes de Oliveira²",
     center=True, spacing=1.0, space_after=12)
blank(1.0)
para("¹* Supervisor do Setor de Patrimônio do Conselho Federal de Contabilidade (CFC). "
     "E-mail autor correspondente: toniaum@gmail.com", size=9, align="left", spacing=1.0)
para("² PhD. Professor orientador — MBA USP/Esalq. E-mail: oliveiragomesgabriel@ieee.org",
     size=9, align="left", spacing=1.0)
page_break()

# ================================================================ PÁGINA 2
para(TITULO, bold=True, center=True, spacing=1.0, space_after=12)
blank(1.0)
para("Resumo", bold=True, align="left", spacing=1.0)
blank(1.0)
para("A gestão patrimonial no setor público brasileiro é regida por exigências legais de "
     "inventário anual, cuja execução manual consome tempo elevado e está sujeita a erros. "
     "Este trabalho avaliou os impactos, em produtividade e acurácia, da adoção de uma "
     "arquitetura progressiva de automação baseada em ferramentas open source no inventário "
     "patrimonial do Conselho Federal de Contabilidade. Adotou-se a pesquisa-ação, em duas "
     "etapas de campo: um inventário-piloto no setor de "
     "tecnologia e uma campanha censitária em todo o órgão. A partir dos registros temporais do "
     "sistema web desenvolvido, mediu-se a produtividade efetiva e construiu-se um modelo de "
     "estimativa por simulação de Monte Carlo com reamostragem bootstrap. O "
     "piloto registrou 94 bens por hora e acurácia locacional de 94,4%, e a simulação projetou em "
     "37,6 horas o inventário dos 3.518 bens ativos. A campanha, em 15 dias de campo, conferiu "
     "3.457 bens e alcançou 97,4% de cobertura do acervo ativo em 28,2 horas efetivas, com 113 "
     "bens por hora, concordância locacional de 83,4% e 569 divergências identificadas em campo, "
     "das quais cerca de metade decorreu da reestruturação administrativa ocorrida durante o "
     "inventário, como revelou a reconciliação com o cadastro atualizado. O tempo "
     "medido situou-se dentro do intervalo de confiança do modelo e evidenciou rendimentos "
     "decrescentes entre setores administrativos, depósitos e varredura residual. Frente ao "
     "cenário manual, estimado entre 125 e 188 horas, a redução de tempo superou 75%. A "
     "arquitetura reduziu expressivamente o tempo de inventário, produziu diagnóstico imediato "
     "do acervo e mostrou-se replicável por órgãos públicos de porte similar.",
     align="justify", spacing=1.0)
blank(1.0)
para("Palavras-chave: controle patrimonial; produtividade; simulação de Monte Carlo; "
     "acurácia; pesquisa-ação.", align="left", spacing=1.0)
blank(1.0)
para("Public sector inventory automation with Data Science and replicable open source "
     "infrastructure", bold=True, center=True, spacing=1.0, space_after=12)
blank(1.0)
para("Abstract", bold=True, align="left", spacing=1.0)
blank(1.0)
para("Asset management in the Brazilian public sector is governed by legal requirements for "
     "annual inventories, whose manual execution is time-consuming and error-prone. This study "
     "evaluated the impacts, on productivity and accuracy, of adopting a progressive automation "
     "architecture based on open source tools in the asset inventory of the Federal Accounting "
     "Council. Action research was adopted, in two field "
     "stages: a pilot inventory in the technology sector and a census campaign across the whole "
     "organization. From the time records of the web system developed, effective productivity "
     "was measured and an estimation model was built by Monte Carlo simulation with bootstrap "
     "resampling. The pilot registered 94 assets per hour and a "
     "location accuracy of 94.4%, and the simulation projected 37.6 hours for the inventory of "
     "the 3,518 active assets. The campaign, over 15 field days, checked 3,457 assets and reached "
     "97.4% coverage of the active assets in 28.2 effective hours, with 113 assets per hour, a "
     "location agreement of 83.4% and 569 discrepancies identified in the field, about half of "
     "which stemmed from the administrative restructuring that took place during the inventory, "
     "as revealed by reconciling with the updated register. The measured time fell "
     "within the confidence interval of the model and revealed diminishing returns between "
     "administrative sectors, storerooms and the residual sweep. Compared with the manual "
     "scenario, estimated at 125 to 188 hours, the time reduction exceeded 75%. The architecture "
     "significantly reduced inventory time, produced an immediate diagnosis of the assets and "
     "proved replicable by public bodies of similar size.",
     align="justify", spacing=1.0)
blank(1.0)
para("Keywords: asset control; productivity; Monte Carlo simulation; accuracy; action "
     "research.", align="left", spacing=1.0)
blank()

# ================================================================ INTRODUÇÃO
title_sec("Introdução")
blank()
body_p("A gestão patrimonial no setor público brasileiro é regida por um conjunto normativo "
       "que impõe a realização de inventários anuais de bens a todos os órgãos da Administração "
       "Pública. A Constituição Federal, em seu artigo 70, estabelece a obrigatoriedade de "
       "prestação de contas sobre o uso de recursos públicos, incluindo o patrimônio material "
       "(Brasil, 1988a). Nesse arcabouço, a Instrução Normativa SEDAP nº 205/1988 disciplina os "
       "procedimentos para o inventário físico (Brasil, 1988b) e a Norma Brasileira de "
       "Contabilidade NBC TSP 07 define os critérios de reconhecimento e mensuração do ativo "
       "imobilizado (Conselho Federal de Contabilidade [CFC], 2017). No plano internacional, a "
       "norma ISO 55000 consolida princípios de planejamento, controle e melhoria contínua para "
       "sistemas de gestão de ativos (International Organization for Standardization [ISO], "
       "2014).")
body_p("A literatura internacional sobre gestão de ativos públicos evidencia que o desafio "
       "não é apenas operacional, mas informacional: a transição de um controle patrimonial "
       "meramente cartorial para uma gestão estratégica de ativos depende da qualidade, da "
       "disponibilidade e do uso efetivo das informações sobre os bens (Kaganova e Amoils, "
       "2020). Estudos aplicados em governos locais identificam problemas recorrentes muito "
       "próximos da realidade brasileira, como ausência de dados adequados, sobreposição de "
       "responsabilidades e limitações institucionais (Hanis et al., 2011), enquanto análises "
       "recentes demonstram que os gestores públicos carecem de informações estruturadas para "
       "decidir sobre seus ativos fixos (Roje et al., 2025). Apesar dessa base normativa e "
       "conceitual, a execução prática dos inventários em órgãos públicos ainda é marcada pela "
       "precariedade dos processos. O CFC realizava seu "
       "inventário predominantemente de forma manual, com planilhas eletrônicas e conferência "
       "física bem a bem — modelo que consome tempo excessivo, é suscetível a erros humanos e "
       "gera elevado custo operacional, realidade comum a organizações de todos os portes "
       "(Zelbst et al., 2012).")
body_p("Tecnologias de identificação e automação apresentam potencial transformador na "
       "gestão de ativos. Lim et al. (2013) identificaram que a aplicação de identificação por "
       "radiofrequência [RFID] reduz erros manuais e melhora a acurácia dos registros, e Paul "
       "et al. (2024) documentaram melhoria de acurácia de 85% para 98% e reduções expressivas "
       "de tempo e perdas. Trabalhos recentes exploram arquiteturas combinando RFID, código de "
       "barras e QR Code em sistemas de inventário web e móveis (Kar et al., 2022; Ton et al., "
       "2024), majoritariamente em contextos logísticos e industriais. No contexto brasileiro, "
       "Madeira Junior e Silveira (2024) evidenciaram que o método tradicional demanda elevada "
       "mão de obra, e Brito et al. (2019) demonstraram que o custo de implantação de "
       "tecnologias de rastreamento equivale a fração reduzida das despesas correntes.")
body_p("Em paralelo, a Administração Pública passa por uma migração de sistemas que apenas "
       "armazenam registros para sistemas que utilizam dados na decisão. A Ciência de Dados "
       "aplicada ao setor público exige que a análise responda a um problema decisório concreto "
       "(Arnaboldi e Azzone, 2020), e a implantação de capacidades analíticas depende de "
       "fatores organizacionais e de qualidade de dados (Merhi, 2021; Broomfield e Reutter, "
       "2021). Quanto à infraestrutura, a adoção de software livre na Administração Pública é "
       "sustentada por fatores técnicos, econômicos e institucionais (van Loon e Toshkov, 2015; "
       "Sánchez et al., 2020), e a conteinerização favorece a reprodutibilidade dos ambientes "
       "computacionais (Boettiger, 2015; Nüst et al., 2020). Kirešová et al. (2023) "
       "demonstraram o potencial do Grafana para painéis em tempo real, e Vilela et al. (2023) "
       "validaram arquiteturas de extração, transformação e carga [ETL] automatizadas em "
       "ambientes de produção.")
body_p("Persistem, contudo, lacunas relevantes: as vertentes de gestão patrimonial pública, "
       "automação de inventários, analytics governamental e infraestrutura aberta são "
       "investigadas de forma predominantemente independente, e são escassos os estudos que as "
       "integrem em uma única abordagem replicável e quantificada aplicada ao inventário "
       "patrimonial de órgãos públicos brasileiros. Diante desse cenário, este trabalho avaliou "
       "os impactos mensuráveis — em produtividade e acurácia — decorrentes da adoção de uma "
       "arquitetura progressiva de automação, baseada em ferramentas open source, no inventário "
       "patrimonial do CFC. Especificamente, mediu-se a produtividade real do processo "
       "automatizado em um setor-piloto, estimou-se por simulação o tempo de inventário do "
       "acervo completo e validou-se essa estimativa por meio de campanha censitária que "
       "alcançou 97,4% do acervo ativo do órgão. A hipótese de trabalho foi a de que a "
       "automação reduziria o tempo de inventário em mais de 70%, patamar reportado pela "
       "literatura, sem perda de acurácia em relação ao processo manual.")
blank()

# ================================================================ METODOLOGIA
title_sec("Metodologia")
blank()
subtitle("Caracterização da pesquisa")
blank()
body_p("A pesquisa classificou-se como pesquisa-ação, de natureza mista — qualitativa e "
       "quantitativa —, de caráter explicativo e abordagem aplicada, conforme preconizado por "
       "Thiollent (2011). Essa classificação justificou-se pelo fato de o pesquisador integrar "
       "o quadro funcional do CFC na qualidade de supervisor do Setor de Patrimônio, atuando "
       "diretamente no diagnóstico, na intervenção e na avaliação dos resultados. A "
       "identificação do órgão é permitida nos termos do Manual de Instruções e Normas "
       "USP/Esalq, uma vez que a pesquisa dispõe de Termo de Anuência institucional e utiliza "
       "dados de acesso restrito interno, sem envolver seres humanos como sujeitos de pesquisa. "
       "Os empregados que participaram das conferências físicas não foram identificados, em "
       "observância às normas do programa.")
body_p("A pesquisa não foi submetida a Comitê de Ética em Pesquisa por enquadrar-se nas "
       "hipóteses do parágrafo único do artigo 1º da Resolução nº 510/2016 do Conselho Nacional "
       "de Saúde (Brasil, 2016): utilizou exclusivamente registros patrimoniais de domínio "
       "institucional, tratados de forma agregada e sem possibilidade de identificação "
       "individual, e não envolveu intervenção, entrevista ou questionário com participantes. A "
       "justificativa foi formalizada no Formulário de Dispensa de Ética [FDE] do programa, "
       "entregue com este trabalho.")
blank()
subtitle("Local do estudo e base patrimonial")
blank()
body_p("O estudo foi conduzido no CFC, autarquia federal sediada em Brasília/DF, responsável "
       "pela regulação e fiscalização do exercício da profissão contábil no Brasil. A fonte "
       "primária de dados foi a base patrimonial consolidada no sistema web de inventário "
       "desenvolvido pelo autor e em operação no órgão (Fase 1 da arquitetura proposta), "
       "exportada em relatório na posição de 22 set. 2026, ao final da campanha censitária. "
       "Nessa data, a base reunia 7.428 registros patrimoniais, dos quais 3.519 correspondiam "
       "a bens em situação ativa — o universo efetivo do inventário físico —, distribuídos por "
       "97 localizações entre setores, salas de reunião, áreas comuns, depósitos e localizações "
       "lógicas do órgão. O inventário-piloto e a simulação, realizados em etapa anterior, "
       "utilizaram a posição de 21 ago. 2026, com 3.518 bens ativos; a diferença de um registro "
       "decorre de um bem incluído sem número de tombamento durante a campanha.")
body_p("Cada registro patrimonial contemplou os campos de número de tombamento, situação, "
       "descrição e complemento do bem, classificação contábil, localização cadastrada, data "
       "de entrada, valores de compra e atual e, a partir da conferência física, o local "
       "verificado, o estado de conservação, o empregado responsável pela conferência e o "
       "instante de sincronização. Esse conjunto de campos permitiu tratar o inventário "
       "simultaneamente como obrigação legal e como fonte de dados analíticos sobre o ciclo de "
       "vida do acervo.")
body_p("A análise quantitativa empregou Python 3, com as bibliotecas pandas e NumPy para "
       "tratamento dos dados e Matplotlib para visualização. À caracterização do acervo "
       "aplicaram-se técnicas de estatística descritiva: composição por classe contábil, "
       "distribuição etária dos bens a partir da data de entrada e classificação ABC pelo valor "
       "atual, na qual os bens foram ordenados de forma decrescente de valor e agrupados nas "
       "classes A (80% do valor acumulado), B (15% seguintes) e C (5% restantes). Essas análises "
       "subsidiaram a priorização das rotas de conferência e a discussão sobre o esforço de "
       "controle proporcional ao risco patrimonial.")
blank()
subtitle("Diagnóstico do processo manual")
blank()
body_p("O processo manual vigente até o ciclo anterior foi caracterizado por levantamento "
       "documental — relatórios da comissão de inventário, listagens e planilhas dos ciclos "
       "anteriores — e por observação participante do pesquisador, responsável pelo setor. As "
       "etapas do processo foram descritas em sequência e, para cada uma, identificou-se o ponto "
       "de falha típico, base da comparação entre os cenários manual e automatizado.")
blank()
subtitle("Arquitetura progressiva de automação")
blank()
body_p("A intervenção estruturou-se em três fases progressivas. A Fase 1, já implementada, "
       "correspondeu ao sistema web de conferência física, operado em dispositivos móveis, com "
       "registro de local conferido, estado de conservação, registro fotográfico e sincronização "
       "dos dados em lotes. A Fase 2, cujos resultados este estudo analisou, compreendeu o pipeline de "
       "dados em Python — com exportação da base, validação, cálculo de indicadores e "
       "visualização —, com armazenamento em PostgreSQL, painéis gerenciais em ferramentas "
       "open source de visualização (Grafana e Streamlit) e infraestrutura conteinerizada via "
       "Docker Compose, favorecendo a replicabilidade do ambiente (Boettiger, 2015; Nüst et al., "
       "2020). A Fase 3, proposta como oportunidade de melhoria "
       "incremental, previu a integração de leitores RFID UHF ao processo de conferência. A "
       "Figura 1 apresenta o fluxo de dados da solução.")
body_p("No sistema web da Fase 1, o operador selecionava a localização em conferência, "
       "identificava cada bem pelo número de tombamento, registrava fotografia e estado de "
       "conservação e prosseguia para o bem seguinte, sem interromper o deslocamento pelo "
       "ambiente. Os registros acumulados no dispositivo eram sincronizados em lotes com o "
       "servidor, o que permitiu o trabalho contínuo mesmo em áreas com conectividade "
       "intermitente e, como efeito colateral metodológico, produziu a marcação temporal por "
       "lote utilizada nas análises de produtividade deste estudo. O mesmo sistema controlou "
       "usuários, perfis de acesso e o histórico de movimentações entre localizações, além de "
       "alimentar o módulo de geração automática de termos de responsabilidade por centro de "
       "custos e por responsável individual.")
body_p("Na Fase 2, o pipeline extraiu a base consolidada, aplicou validações de consistência "
       "(unicidade do tombamento, domínios de situação e localização, tipos de valores e "
       "datas), transformou os registros com pandas e os carregou em PostgreSQL, de onde os "
       "painéis gerenciais consumiram os indicadores. O processamento foi idempotente — podia "
       "ser reexecutado a qualquer momento sem duplicar registros — e cada execução ficou "
       "registrada, o que garantiu trilha de auditoria compatível com as exigências de "
       "prestação de contas do setor público.")
fig("figura_1_fluxo.png",
    "Figura 1. Fluxo de dados da arquitetura de automação do inventário",
    "Fonte: Dados originais da pesquisa")
body_p("A etapa final da arquitetura foi a carga das leituras da campanha no sistema de gestão "
       "patrimonial do órgão, cuja importação foi atômica (aceitou todas as leituras ou nenhuma) "
       "e rejeitou bens inexistentes, datas inválidas e estados de conservação fora do domínio, "
       "o que serviu de validação de ponta a ponta da solução.")
body_p("Todo o código da solução foi desenvolvido com ferramentas open source e organizado "
       "para publicação em repositórios públicos, de modo a permitir a reprodução do estudo por "
       "outros órgãos. A descrição da infraestrutura replicável e o endereço dos repositórios "
       "constam do Apêndice A.")
blank()
subtitle("Inventário-piloto e projeção por simulação")
blank()
body_p("A coleta de campo ocorreu em duas etapas. A primeira consistiu em um inventário-piloto "
       "conduzido em 26 maio 2026 no setor de Coordenadoria de Gestão de Tecnologia da "
       "Informação [CGTI], no qual foram verificados 266 bens, com registro temporal em 14 "
       "lotes de sincronização cobrindo 222 bens em uma sessão contínua de trabalho.")
body_p("Para projetar o tempo de inventário do acervo completo a partir da amostra-piloto, "
       "construiu-se um modelo de estimativa por simulação de Monte Carlo com reamostragem "
       "bootstrap: em cada uma das 10.000 iterações, os lotes observados foram reamostrados com "
       "reposição, recalculou-se o tempo médio por bem ponderado e projetou-se o tempo total "
       "para os 3.518 bens ativos, obtendo-se a distribuição da estimativa e o intervalo de "
       "confiança de 95%. A semente aleatória foi fixada para garantir a reprodutibilidade.")
blank()
subtitle("Campanha censitária")
blank()
body_p("A segunda "
       "etapa consistiu na campanha censitária de inventário, conduzida entre 18 ago. e 22 set. "
       "2026, em 15 dias de campo, por quatro empregados do órgão — dois dos quais responderam "
       "por 98,7% das conferências —, com 3.457 bens conferidos em 345 lotes de sincronização. "
       "A campanha organizou-se em três etapas operacionais: (i) varredura dos pavimentos "
       "administrativos, do 5º ao 12º, entre 18 e 21 ago.; (ii) conferência dos subsolos, do "
       "térreo e do 2º, 3º e 13º pavimentos — que reúnem depósitos, almoxarifado, arquivo, "
       "oficina, plenário, auditório, o setor de tecnologia e o centro de processamento de "
       "dados —, entre 24 e 28 ago.; e (iii) varredura residual de pendências e de localizações "
       "lógicas, entre 1º e 22 set.")
body_p("Em ambas as etapas, o sistema registrou, para cada bem verificado, o local conferido, "
       "o estado de conservação, o empregado responsável e o instante da sincronização. Como os "
       "registros foram sincronizados em lotes, a unidade de observação temporal adotada foi o "
       "lote de coleta — conjunto de bens com o mesmo instante de sincronização —, e o tempo "
       "por bem foi obtido pela razão entre o intervalo decorrido entre lotes consecutivos e a "
       "quantidade de bens do lote, com exclusão dos intervalos de pausa superiores a uma hora "
       "e do primeiro lote de cada dia, que não dispõe de intervalo anterior de referência. "
       "Bens reconferidos em mais de uma data preservaram apenas a última marcação temporal, de "
       "modo que os quantitativos diários refletem a posição final da base.")
body_p("A produtividade efetiva (bens por hora) foi obtida pela razão entre a quantidade de "
       "bens conferidos e o tempo efetivo de conferência, e seu recíproco, o tempo por bem, "
       "apurados por dia, por etapa da campanha e para o conjunto. A campanha foi utilizada como "
       "validação empírica externa da projeção do piloto: o tempo efetivo medido, acrescido do "
       "tempo projetado para os bens pendentes ao ritmo da etapa residual, foi confrontado com "
       "o intervalo de confiança simulado.")
blank()
subtitle("Acurácia locacional e reconciliação cadastral")
blank()
body_p("A acurácia locacional foi definida como a proporção de bens ativos cujo local físico "
       "conferido coincidiu com o registro do sistema. As divergências de localização foram "
       "analisadas pelos pares origem-destino mais frequentes e classificadas em cinco "
       "categorias — granularidade cadastral dentro do mesmo setor, recolhimento a depósito, "
       "almoxarifado ou arquivo, origem em localização lógica, remanejamento entre setores do "
       "mesmo pavimento e remanejamento entre pavimentos. Após a campanha, as leituras foram "
       "reconciliadas com o cadastro patrimonial atualizado por meio de tabela de "
       "correspondência entre as denominações antigas e atuais das localizações, derivada "
       "automaticamente da localização atual dos próprios bens (regra da maioria por "
       "localização de origem), e a concordância foi recalculada contra o cadastro "
       "reestruturado.")
blank()
subtitle("Cobertura do inventário, conservação e bens baixados")
blank()
body_p("A cobertura do inventário foi definida como a proporção de bens ativos conferidos, "
       "calculada por quantidade e por valor e desagregada por classe contábil, pavimento, "
       "localização e classe ABC. Registraram-se, ainda, o estado de conservação atribuído a "
       "cada bem conferido e os bens em situação de baixa ou doação localizados fisicamente, "
       "indicativos de descompasso entre o desfazimento contábil e o físico.")
blank()
subtitle("Comparação com o cenário manual e implicações econômicas")
blank()
body_p("À falta de registro histórico cronometrado do processo manual, o tempo do cenário "
       "anterior (AS-IS) foi estimado a partir da literatura, que reporta reduções de tempo da "
       "ordem de 70% a 80% com a adoção de automação (Madeira Junior e Silveira, 2024; Paul et "
       "al., 2024), parâmetro adotado de forma explícita e conservadora apenas para fins "
       "comparativos. O custo de implantação foi levantado a partir das horas de "
       "desenvolvimento e da infraestrutura utilizada, sem despesas de licenciamento.")
blank()
# (blocos antigos de "Análise dos dados" removidos abaixo)
subtitle("Síntese metodológica")
blank()
body_p("A fim de sintetizar a metodologia científica da presente pesquisa, utilizou-se a "
       "ferramenta 5W2H, que responde aos principais aspectos que norteiam a investigação "
       "(Tabela 1).")
caption("Tabela 1. Síntese metodológica 5W2H da pesquisa")
table([
    ["Pergunta", "Resposta"],
    ["What? (O quê?)", "Desenvolvimento e avaliação de uma arquitetura progressiva de "
     "automação do inventário patrimonial, baseada em ferramentas open source."],
    ["Why? (Por quê?)", "O processo manual é demorado, suscetível a erros e de alto custo; "
     "faltam estudos replicáveis e quantificados de soluções open source na gestão "
     "patrimonial pública brasileira."],
    ["Who? (Quem?)", "O pesquisador, supervisor do Setor de Patrimônio do CFC, sob orientação "
     "do Prof. PhD. Gabriel Gomes de Oliveira, com apoio de empregados do setor nas "
     "conferências físicas."],
    ["Where? (Onde?)", "Conselho Federal de Contabilidade, Brasília/DF, com infraestrutura "
     "conteinerizada e publicada em repositório público."],
    ["When? (Quando?)", "De março a dezembro de 2026, com piloto em maio e campanha censitária "
     "de agosto a setembro de 2026."],
    ["How? (Como?)", "Pesquisa-ação mista: diagnóstico AS-IS, desenvolvimento da solução, "
     "inventário-piloto, simulação de Monte Carlo e campanha censitária de validação."],
    ["How much? (Quanto custa?)", "Custo de implantação próximo a zero, restrito a horas do "
     "pesquisador e à infraestrutura de servidor já existente no órgão."],
], col_widths=[4.5, 11.5])
caption("Fonte: Dados originais da pesquisa")
blank()
body_p("As etapas metodológicas, do levantamento bibliográfico à análise dos resultados, "
       "estão representadas no fluxograma da Figura 2.")
fig("fluxograma_etapas.png",
    "Figura 2. Fluxograma das etapas metodológicas da pesquisa",
    "Fonte: Dados originais da pesquisa", largura=13.5)
blank()

# ================================================================ RESULTADOS
title_sec("Resultados e Discussão")
blank()
subtitle("Base patrimonial")
blank()
body_p("A base patrimonial consolidada reuniu 7.428 registros, dos quais 3.519 (47,4%) em "
       "situação ativa, 2.051 (27,6%) baixados, 1.854 (25,0%) doados e quatro classificados "
       "como inservíveis, conforme a Tabela 2. O predomínio de registros históricos já "
       "desfeitos (52,6% do total) evidencia o caráter cumulativo da base e reforça a "
       "importância do saneamento cadastral como etapa preliminar do inventário — problema "
       "informacional recorrente na gestão de ativos públicos (Hanis et al., 2011; Roje et "
       "al., 2025).")
caption("Tabela 2. Situação dos registros da base patrimonial do órgão (posição de 22 set. 2026)")
table([
    ["Situação", "Registros", "Participação (%)"],
    ["Ativo", "3.519", "47,4"],
    ["Baixado", "2.051", "27,6"],
    ["Doado", "1.854", "25,0"],
    ["Inservível", "4", "0,1"],
    ["Total", "7.428", "100,0"],
], col_widths=[6, 5, 5])
caption("Fonte: Resultados originais da pesquisa")
blank()
body_p("O acervo ativo concentrou-se em duas classes contábeis: móveis e utensílios de "
       "escritório (1.906 bens; 54,2%) e equipamentos de processamento de dados (1.189 bens; "
       "33,8%), seguidos de máquinas e equipamentos (342 bens; 9,7%), conforme a Figura 3. O "
       "valor atual contabilizado dos bens ativos somou R$ 102,8 milhões, dos quais R$ 15,6 "
       "milhões referentes a bens móveis — objeto efetivo da conferência física — e o restante "
       "concentrado em sede, terrenos e instalações. Os bens distribuíam-se por 97 localizações "
       "físicas e lógicas, das quais as cinco maiores (dois depósitos, o setor de tecnologia, o "
       "setor de comunicação e o gabinete) concentravam 30% do acervo ativo.")
fig("composicao_classes.png",
    "Figura 3. Composição do acervo ativo por classe contábil",
    "Fonte: Resultados originais da pesquisa")
body_p("A distribuição etária dos bens móveis ativos, obtida a partir da data de entrada, "
       "revelou idade média de 14,1 anos e mediana de 13,7 anos, com 60,0% do acervo em uso há "
       "mais de dez anos (Figura 4). A concentração de 1.124 bens ingressados a partir de 2020 "
       "reflete o ciclo recente de renovação de equipamentos de informática, enquanto os 521 "
       "bens anteriores a 2000 sinalizam a necessidade de avaliação sistemática de "
       "obsolescência e desfazimento — dimensão do ciclo de vida dos ativos que a norma ISO "
       "55000:2014 coloca no centro da gestão estratégica e que a literatura aponta como fator "
       "crítico da qualidade informacional patrimonial (Kaganova e Amoils, 2020).")
fig("idade_acervo.png",
    "Figura 4. Distribuição dos bens móveis ativos por período de entrada no acervo",
    "Fonte: Resultados originais da pesquisa")
body_p("A classificação ABC pelo valor atual evidenciou forte concentração patrimonial: 352 "
       "bens (10,0% do acervo móvel) responderam por 80% do valor total, enquanto 2.092 bens "
       "(59,6%) somaram apenas 5% (Figura 5). Essa assimetria tem implicação direta para a "
       "política de inventário: o esforço de controle pode ser calibrado ao risco, com "
       "conferência prioritária e mais frequente dos bens de classe A — abordagem quantitativa "
       "de priorização alinhada ao uso de métodos multicritério na gestão de ativos "
       "governamentais e à administração pública orientada por dados (Arnaboldi e Azzone, "
       "2020).")
fig("curva_abc.png",
    "Figura 5. Curva ABC do valor atual dos bens móveis ativos",
    "Fonte: Resultados originais da pesquisa")
blank()
subtitle("Diagnóstico do processo manual")
blank()
body_p("O levantamento documental e a observação participante caracterizaram o processo "
       "manual vigente até o ciclo anterior, sintetizado na Tabela 3. A conferência era "
       "conduzida com listagens impressas a partir de planilhas eletrônicas, com anotação "
       "manuscrita em campo e posterior digitação dos resultados — duplo manuseio que "
       "constitui a principal fonte de erro e retrabalho do modelo tradicional, em linha com o "
       "diagnóstico da literatura sobre controle patrimonial manual (Madeira Junior e "
       "Silveira, 2024; Zelbst et al., 2012). A ausência de registro fotográfico e de marcação "
       "temporal impedia tanto a evidenciação da conferência quanto qualquer medição de "
       "produtividade, e as divergências de localização identificadas em campo raramente "
       "retornavam ao cadastro.")
caption("Tabela 3. Etapas do processo manual de inventário e pontos de falha associados")
table([
    ["Etapa do processo manual", "Descrição", "Ponto de falha típico"],
    ["Preparação", "Impressão de listagens por setor a partir de planilhas",
     "Listagens desatualizadas em relação ao cadastro"],
    ["Conferência física", "Localização visual do bem e anotação manuscrita",
     "Erros de transcrição e bens não localizados sem registro"],
    ["Consolidação", "Digitação das anotações nas planilhas",
     "Duplo manuseio dos dados e retrabalho"],
    ["Apuração", "Comparação manual entre listagens e cadastro",
     "Divergências latentes e sem tratamento sistemático"],
    ["Relatório", "Elaboração manual do relatório da comissão",
     "Ausência de indicadores e de evidência fotográfica"],
], col_widths=[3.8, 6.0, 5.7])
caption("Fonte: Resultados originais da pesquisa")
blank()
subtitle("Inventário-piloto e projeção por simulação")
blank()
body_p("No setor-piloto (CGTI) inventariaram-se 266 bens, com registro temporal em 14 lotes de "
       "sincronização que cobriram 222 bens em uma sessão contínua de trabalho. A totalidade "
       "dos bens conferidos no piloto foi classificada em estado de conservação “Bom”. "
       "A Tabela 4 sintetiza os indicadores obtidos.")
caption("Tabela 4. Indicadores do inventário-piloto da CGTI (26 maio 2026)")
table([
    ["Indicador", "Valor"],
    ["Bens verificados no piloto", "266"],
    ["Lotes de sincronização analisados", "14 (222 bens)"],
    ["Produtividade média", "≈ 94 bens/hora"],
    ["Tempo mediano por bem", "37 segundos"],
    ["Acurácia locacional", "94,4%"],
    ["Divergências de localização", "15 bens (5,6%)"],
    ["Universo de bens ativos do órgão", "3.518"],
], col_widths=[9, 6])
caption("Fonte: Resultados originais da pesquisa")
blank()
body_p("A produtividade variou entre os lotes conforme a natureza e a disposição física dos "
       "bens, registrando ritmo médio ponderado de aproximadamente 94 bens por hora (Figura 6). "
       "A distribuição do tempo por bem apresentou mediana de 37 segundos, com cauda à direita "
       "associada aos lotes de bens de conferência mais trabalhosa. A "
       "conferência incluiu registro fotográfico de cada bem, o que qualifica a evidência de "
       "inventário sem comprometer o ritmo de trabalho.")
fig("piloto_produtividade_lote.png",
    "Figura 6. Produtividade por lote no inventário-piloto da CGTI",
    "Fonte: Resultados originais da pesquisa")
body_p("A partir dessa medição, a simulação de Monte Carlo (10.000 iterações) projetou em 37,6 "
       "horas o tempo médio necessário para inventariar os 3.518 bens ativos do órgão, com "
       "intervalo de confiança de 95% entre 30,6 e 46,6 horas (Figura 7) — equivalente a "
       "aproximadamente seis jornadas efetivas de trabalho. O intervalo relativamente amplo "
       "refletiu a incerteza decorrente de a estimativa apoiar-se em um único setor-piloto, "
       "hipótese que a campanha censitária permitiu testar empiricamente.")
fig("piloto_montecarlo.png",
    "Figura 7. Distribuição simulada (Monte Carlo) do tempo de inventário do acervo",
    "Fonte: Resultados originais da pesquisa")
blank()
subtitle("Campanha censitária")
blank()
E = {e["etapa"]: e for e in M["etapas"]}
body_p(f"A campanha censitária conferiu {br(M['conferidos'])} bens — {br(M['ativos_conferidos'])} "
       f"ativos, 27 baixados e dois doados — em 15 dias de campo distribuídos entre 18 ago. e "
       f"22 set. 2026, elevando a cobertura do inventário a {br(M['cobertura'],1)}% do acervo ativo "
       f"({br(M['ativos_conferidos'])} de {br(M['ativos'])} bens), com {M['pendentes']} bens "
       f"pendentes. A Tabela 5 consolida os indicadores da campanha, a Tabela 6 detalha as três "
       f"etapas operacionais. O volume diário de "
       f"conferências (Figura 8) situou-se entre 325 e 450 bens ativos na primeira etapa e entre "
       f"115 e 540 na segunda, caindo para 1 a 111 bens na etapa residual, quando o trabalho "
       f"passou a consistir na busca individual de bens ausentes de sua localização cadastrada.")
caption("Tabela 5. Indicadores consolidados da campanha censitária (18 ago. a 22 set. 2026)")
table([
    ["Indicador", "Valor"],
    ["Bens conferidos", f"{br(M['conferidos'])} ({br(M['ativos_conferidos'])} ativos)"],
    ["Cobertura do acervo ativo", f"{br(M['cobertura'],1)}%"],
    ["Dias de campo", str(M["dias_campo"])],
    ["Lotes de sincronização", str(M["lotes"])],
    ["Tempo efetivo de conferência", f"≈ {br(M['tempo_efetivo_h'],1)} horas"],
    ["Produtividade média ponderada", f"≈ {br(M['produtividade'])} bens/hora"],
    ["Tempo médio por bem", f"≈ {br(M['seg_por_bem'])} segundos (mediana por lote: {br(M['seg_por_bem_mediana_lote'])} s)"],
    ["Concordância locacional (bens ativos)", f"{br(M['concordancia_pct'],1)}%"],
    ["Divergências de localização (bens ativos)", f"{br(M['divergentes'])} bens ({br(M['divergencia_pct'],1)}%)"],
    ["Bens baixados ou doados localizados fisicamente", str(M["nao_ativos_encontrados"]["n"])],
    ["Empregados envolvidos na conferência", str(M["servidores"])],
], col_widths=[9, 6])
caption("Fonte: Resultados originais da pesquisa")
blank()
caption("Tabela 6. Desempenho da campanha censitária por etapa operacional")
e1, e2, e3 = E[1], E[2], E[3]
rows = [
    ["Indicador", "Etapa 1", "Etapa 2", "Etapa 3", "Campanha"],
    ["Escopo", "Pavimentos administrativos", "Depósitos, subsolos e áreas técnicas", "Varredura residual", "Acervo completo"],
    ["Período (2026)", "18 a 21 ago.", "24 a 28 ago.", "1º a 22 set.", "18 ago. a 22 set."],
    ["Dias de campo", str(int(e1["dias"])), str(int(e2["dias"])), str(int(e3["dias"])), str(M["dias_campo"])],
    ["Bens conferidos", br(e1["bens"]), br(e2["bens"]), br(e3["bens"]), br(M["conferidos"])],
    ["Lotes de sincronização", str(int(e1["lotes"])), str(int(e2["lotes"])), str(int(e3["lotes"])), str(M["lotes"])],
    ["Tempo efetivo (h)", br(e1["tempo_h"], 1), br(e2["tempo_h"], 1), br(e3["tempo_h"], 1), br(M["tempo_efetivo_h"], 1)],
    ["Produtividade (bens/h)", br(e1["prod"]), br(e2["prod"]), br(e3["prod"]), br(M["produtividade"])],
    ["Tempo por bem (s)", br(e1["seg_bem"]), br(e2["seg_bem"]), br(e3["seg_bem"]), br(M["seg_por_bem"])],
]
table(rows, col_widths=[4.4, 2.9, 3.1, 2.8, 2.8])
caption("Fonte: Resultados originais da pesquisa")
caption("Nota: a produtividade considera os intervalos entre lotes consecutivos de um mesmo dia, "
        "excluindo pausas superiores a uma hora e os bens do primeiro lote de cada dia, sem "
        "intervalo anterior de referência")
blank()
fig("figura_2_conferencias_dia.png",
    "Figura 8. Bens ativos conferidos por dia de campo na campanha censitária",
    "Fonte: Resultados originais da pesquisa")
body_p(f"A produtividade média ponderada da campanha alcançou aproximadamente "
       f"{br(M['produtividade'])} bens por hora (Figura 9) — cerca de 20% acima dos 94 bens por "
       f"hora medidos no piloto —, porém com forte heterogeneidade entre etapas: "
       f"{br(E[1]['prod'])} bens por hora na varredura dos pavimentos administrativos, "
       f"{br(E[2]['prod'])} nos ambientes de acervo denso e {br(E[3]['prod'])} na varredura "
       f"residual. Esse padrão de rendimentos decrescentes tem explicação operacional. Na "
       f"primeira etapa, os bens encontravam-se dispostos em postos de trabalho, com etiquetas "
       f"visíveis e acesso direto; na segunda, os depósitos concentravam bens empilhados, sem "
       f"ordenação por tombamento e com etiquetas de acesso difícil, o que elevou o tempo por "
       f"bem de {br(E[1]['seg_bem'])} para {br(E[2]['seg_bem'])} segundos; na terceira, o "
       f"esforço concentrou-se em localizar individualmente bens ausentes de sua localização "
       f"cadastrada — tarefa de busca cujo custo unitário é intrinsecamente maior "
       f"({br(E[3]['seg_bem'])} segundos por bem). As etapas de varredura foram executadas "
       f"predominantemente por um empregado por dia (no máximo dois), de modo que as horas "
       f"efetivas de equipe praticamente coincidem com horas-pessoa (28,8 horas); os dois "
       f"empregados que concentraram as conferências apresentaram produtividades individuais de "
       f"134 e 94 bens por hora, diferença explicada pelos ambientes que lhes foram atribuídos "
       f"(pavimentos administrativos e depósitos, respectivamente), e não pelo desempenho "
       f"pessoal.")
fig("campanha_produtividade_dia.png",
    "Figura 9. Produtividade efetiva por dia de campo da campanha censitária",
    "Fonte: Resultados originais da pesquisa",
    nota_txt="n.d. = não disponível (dia com um único lote de sincronização)")
body_p(f"A evolução da cobertura acumulada (Figura 10) torna visível o mesmo fenômeno: a "
       f"campanha atingiu 43,8% do acervo ao fim da primeira etapa, 88,4% ao fim da segunda e "
       f"avançou apenas nove pontos percentuais nas três semanas seguintes. Foi o "
       f"comportamento típico de processos de busca — os bens remanescentes são, por definição, "
       f"aqueles que não estavam onde deveriam — e tem implicação direta para o planejamento: o "
       f"esforço de conclusão de um inventário não é proporcional à fração pendente do acervo.")
fig("cobertura_acumulada.png",
    "Figura 10. Evolução da cobertura acumulada do acervo ativo ao longo da campanha censitária",
    "Fonte: Resultados originais da pesquisa")
body_p(f"As {br(M['tempo_efetivo_h'],1)} horas efetivas consumidas para {br(M['cobertura'],1)}% do "
       f"acervo, acrescidas das cerca de {br(M['horas_para_concluir_ritmo_e3'],1)} horas necessárias "
       f"para concluir os {M['pendentes']} bens pendentes ao ritmo da etapa residual, projetam o "
       f"inventário completo em aproximadamente {br(M['horas_total_projetado'])} horas — valor "
       f"situado dentro do intervalo de confiança de 95% da simulação (30,6 a 46,6 horas) e 18% "
       f"abaixo da média projetada (37,6 horas). O resultado valida empiricamente o modelo de "
       f"Monte Carlo: a estimativa construída a partir de um único setor-piloto foi ligeiramente "
       f"conservadora, mas capturou a ordem de grandeza correta do esforço, o que a qualifica "
       f"como instrumento de planejamento de campanhas em outros órgãos. Uma "
       f"extrapolação feita ao final da primeira etapa, com base nos {br(E[1]['prod'])} bens por "
       f"hora então observados, apontaria cerca de 19 horas para o acervo completo — "
       f"subestimativa de quase 40% que ilustra o risco de projetar o esforço total a partir dos "
       f"setores de conferência mais fácil e reforça a utilidade de um modelo que incorpore a "
       f"variabilidade entre lotes.")
body_p("Esse desempenho é compatível com os ganhos reportados pela literatura para processos "
       "automatizados de rastreamento de ativos (Lim et al., 2013; Paul et al., 2024) e "
       "demonstra, na prática, a escalabilidade da solução: o mesmo sistema absorveu quatro "
       "operadores, 15 dias de campo e 345 lotes de sincronização sem alteração de arquitetura, "
       "com consolidação automática dos registros e sem qualquer etapa de digitação.")
body_p(f"O patamar de {br(E[1]['prod'])} bens por hora nos ambientes "
       f"administrativos foi alcançado sem identificação automática por radiofrequência: a "
       f"conferência apoiou-se na identificação visual do tombamento, estratégia de custo "
       f"marginal nulo análoga às soluções móveis de identificação por código descritas na "
       f"literatura (Kar et al., 2022; Ton et al., 2024). O resultado indica que a maior parcela "
       f"do ganho decorre da eliminação do duplo manuseio de dados — anotação e digitação — e não "
       f"do hardware de leitura. Nos depósitos, ao contrário, a produtividade caiu para cerca de "
       f"{br(E[2]['prod'])} bens por hora e o tempo por bem praticamente dobrou, e é exatamente "
       f"nesse ambiente — bens empilhados, etiquetas de difícil acesso — que a leitura por RFID "
       f"UHF (Fase 3) oferece maior retorno, o que reposiciona a integração de leitores como "
       f"otimização incremental voltada aos ambientes de alta densidade, e não como "
       f"pré-requisito da automação.")
blank()
subtitle("Acurácia locacional e reconciliação cadastral")
blank()
body_p(f"Entre os {br(M['ativos_conferidos'])} bens ativos conferidos na campanha, a concordância "
       f"entre o local físico e o registro do sistema foi de {br(M['concordancia_pct'],1)}% "
       f"({br(M['concordantes'])} bens), com {br(M['divergentes'])} divergências "
       f"({br(M['divergencia_pct'],1)}%) — patamar inferior aos 94,4% "
       f"observados no piloto e ligeiramente superior aos 81,2% apurados na posição parcial da "
       f"primeira etapa. A diferença em relação ao piloto é explicada pela natureza dos setores "
       f"cobertos: enquanto o piloto ocorreu no setor de tecnologia, cujo acervo é acompanhado "
       f"de perto, a campanha alcançou setores cujos registros não passavam por conferência "
       f"sistemática há mais tempo, expondo o passivo informacional acumulado do processo "
       f"manual. A taxa de divergência variou de 1,9% no 11º pavimento a 28,1% no 7º "
       f"pavimento — que concentra as unidades administrativas de maior rotatividade de "
       f"mobiliário —, e foi semelhante entre móveis e utensílios (17,6%) e equipamentos de "
       f"processamento de dados (17,4%), mas inferior em máquinas e equipamentos (10,3%), "
       f"tipicamente fixos.")
body_p(f"A análise dos pares origem-destino (Tabela 7) e a classificação das divergências por "
       f"natureza (Tabela 8) revelam que o fenômeno é heterogêneo. Quase metade dos casos "
       f"(46,9%) corresponde a remanejamentos entre pavimentos e 18,1% a remanejamentos entre "
       f"setores do mesmo pavimento — movimentações reais nunca formalizadas no sistema. Outros "
       f"26,4% correspondem ao recolhimento de bens a depósitos, almoxarifado ou arquivo sem "
       f"registro da movimentação, caso do par mais frequente (37 bens registrados no Depósito "
       f"Rampa e localizados no Depósito SEPAT) e de 29 bens novos que permaneciam no "
       f"almoxarifado. Apenas 7,0% decorrem de granularidade cadastral — bens registrados em "
       f"salas de reunião e conferidos no setor contíguo, como os 30 bens da sala de reunião da "
       f"COTEC — e, portanto, refletem escolha de modelagem, e não erro físico. Longe de indicar "
       f"falha do método, as divergências evidenciam o valor diagnóstico do inventário "
       f"automatizado, que localizou de imediato bens deslocados — informação que, no processo "
       f"manual, permanecia latente. Espera-se que, após a reconciliação dos registros e a "
       f"padronização da malha de localizações, a conformidade convirja para o patamar superior "
       f"a 98% reportado pela literatura para processos automatizados maduros (Paul et al., "
       f"2024).")
caption("Tabela 7. Pares de divergência locacional mais frequentes na campanha censitária")
rows = [["Local registrado no sistema", "Local físico conferido", "Bens"]]
for a, b, n in M["div_top10"]:
    rows.append([tit(a), tit(b), str(n)])
table(rows, col_widths=[6.5, 6.5, 2.5])
caption("Fonte: Resultados originais da pesquisa")
caption(f"Nota: os {br(M['divergentes'] - M['div_top10_sum'])} demais casos distribuem-se em "
        f"{M['div_pares_n'] - 10} pares, dos quais {M['div_pares_ate5']} com cinco ou menos ocorrências")
blank()
caption("Tabela 8. Classificação das divergências locacionais por natureza")
rows = [["Natureza da divergência", "Bens", "Participação (%)"]]
order = ["Remanejamento entre pavimentos", "Recolhimento a depósito, almoxarifado ou arquivo",
         "Remanejamento entre setores do mesmo pavimento",
         "Granularidade cadastral (sala de reunião ou corredor do mesmo setor)",
         "Origem em localização lógica (bens novos, termos, CFC)"]
for k in order:
    n = M["div_categorias"][k]
    rows.append([k, str(n), br(n / M["divergentes"] * 100, 1)])
rows.append(["Total", br(M["divergentes"]), "100,0"])
table(rows, col_widths=[10, 2.5, 3.5])
caption("Fonte: Resultados originais da pesquisa")
blank()
body_p("A experiência de campo evidenciou um fator que a literatura de gestão de ativos públicos "
       "trata como problema informacional (Kaganova e Amoils, 2020; Hanis et al., 2011), mas "
       "que raramente é quantificado: a estrutura administrativa do órgão mudou durante a "
       "própria campanha. Ao longo do inventário, setores foram renomeados, fundidos e "
       "desmembrados, e o cadastro patrimonial só refletiu a nova estrutura após o término das "
       "conferências. Das 91 localizações em que houve leitura, 38 deixaram de existir com o "
       "nome utilizado em campo — entre elas as de maior acervo, como o setor de tecnologia, o "
       "de comunicação e o principal depósito —, o que afetou 1.970 leituras, 57% do total.")
body_p("Para medir o efeito, construiu-se uma tabela de correspondência entre denominações "
       "antigas e atuais, derivada automaticamente da localização em que cada bem passou a "
       "figurar no cadastro atualizado (regra da maioria por localização de origem), e "
       "recalculou-se a concordância das mesmas 3.456 leituras válidas contra o cadastro "
       "reestruturado (Tabela 9). As divergências caíram de 569 (16,6%) para 280 (8,1%): "
       "cerca de metade dos casos apontados durante a campanha não correspondia a bens fora do "
       "lugar, mas ao cadastro que ainda não havia absorvido a reorganização — os pares mais "
       "frequentes da Tabela 7 são justamente setores fundidos, desmembrados ou renomeados, "
       "como as unidades de pessoal e contabilidade, as salas de reunião vinculadas às "
       "gerências e o gabinete e sua assessoria.")
caption("Tabela 9. Divergências locacionais das leituras da campanha antes e depois da reconciliação com o cadastro reestruturado")
table([
    ["Base de comparação", "Leituras", "Divergentes", "Taxa (%)"],
    ["Cadastro vigente durante a campanha (denominações antigas)", "3.428", "569", "16,6"],
    ["Cadastro atualizado após a reestruturação (denominações atuais)", "3.456", "280", "8,1"],
    ["Localizações lidas renomeadas ou extintas", "38 de 91", "—", "—"],
    ["Leituras em localizações renomeadas", "1.970", "—", "57,0"],
], col_widths=[8.5, 2.5, 2.5, 2.5])
caption("Fonte: Resultados originais da pesquisa")
caption("Nota: a segunda linha inclui os 29 bens baixados ou doados localizados fisicamente e exclui um registro "
        "sem número de tombamento, tratado como sobra; a base atualizada corresponde à posição de 24 set. 2026")
blank()
body_p("O achado tem duas implicações. A primeira é metodológica: a acurácia locacional medida "
       "por um inventário não é apenas função da disciplina física sobre os bens, mas também da "
       "estabilidade da malha de localizações durante a coleta, e deve ser reportada com "
       "referência explícita ao cadastro contra o qual foi calculada. A segunda é prática: o "
       "processo automatizado, ao registrar cada leitura com data, hora, responsável e local "
       "físico, permitiu reconciliar retroativamente todo o inventário com a nova estrutura sem "
       "nova ida a campo — operação inviável no processo manual, em que as listagens impressas "
       "já nasciam vinculadas à estrutura antiga. Nesse sentido, a campanha funcionou como "
       "experimento de validação do sistema: os dados coletados sobreviveram intactos a uma "
       "reorganização administrativa e foram absorvidos pelo sistema de gestão patrimonial do "
       "órgão, conforme descrito na subseção sobre a arquitetura de automação.")
blank()
subtitle("Cobertura do inventário, conservação e bens baixados")
blank()
body_p(f"A cobertura de {br(M['cobertura'],1)}% do acervo ativo distribuiu-se de forma homogênea "
       f"entre classes contábeis — de 97,1% em móveis e utensílios a 99,1% em "
       f"máquinas e equipamentos — e entre pavimentos (Tabela 10): sete pavimentos "
       f"foram integralmente conferidos e nenhum ficou abaixo de 92%. Das {M['loc_n']} "
       f"localizações com bens ativos, {M['loc_100']} alcançaram 100% de cobertura e "
       f"{M['loc_90']} superaram 90%. A sequência de varredura foi deliberada: iniciou-se pelos "
       f"setores de maior circulação de pessoas e bens, onde o risco de divergência é maior, "
       f"deixando para a segunda etapa os ambientes de acervo estático e concentrado — "
       f"depósitos, plenário, auditório e áreas técnicas do subsolo.")
caption("Tabela 10. Cobertura do inventário por pavimento ao final da campanha censitária")
rows = [["Pavimento", "Bens ativos", "Conferidos", "Cobertura (%)"]]
for r in M["cobertura_pavimento"]:
    rows.append([r["pav"], br(r["ativos"]), br(r["conf"]), br(r["cob"], 1)])
rows.append(["Total", br(M["ativos"]), br(M["ativos_conferidos"]), br(M["cobertura"], 1)])
table(rows, col_widths=[5.5, 3.5, 3.2, 3.3])
caption("Fonte: Resultados originais da pesquisa")
caption("Nota: “Localizações lógicas” agrupa registros sem vínculo com pavimento físico, como "
        "bens novos, softwares, termos individuais e bens em apuração de responsabilidade")
blank()
body_p(f"Os {M['pendentes']} bens pendentes concentravam-se no Depósito Rampa (28), no setor de "
       f"comunicação (14) e no Depósito SEPAT (5); {M['pendentes_logicos']} deles encontravam-se em "
       f"localizações lógicas que não comportam conferência física — quatro licenças de software "
       f"amortizadas aguardando baixa, quatro bens em apuração de responsabilidade e quatro bens "
       f"novos ainda não entregues —, de modo que restavam {M['pendentes'] - M['pendentes_logicos']} "
       f"bens físicos a localizar, majoritariamente móveis (56) e equipamentos de informática "
       f"(26). Em valor, os pendentes somavam R$ {br(M['pendentes_valor_k'])} mil, dos quais "
       f"R$ {br(M['pendentes_valor_softwares_k'])} mil correspondem às licenças de software, o que "
       f"reduz a pendência física a menos de R$ 110 mil.")
body_p(f"A dimensão financeira confirma a qualidade da cobertura: os bens móveis conferidos "
       f"representam {br(M['cobertura_valor_moveis_pct'],1)}% do valor do acervo móvel ativo "
       f"(R$ {br(M['valor_moveis_conferido_M'],2)} milhões de R$ {br(M['valor_moveis_M'],2)} "
       f"milhões), e a cobertura foi uniforme entre as classes da curva ABC — 97,4% dos bens de "
       f"classe A, 97,7% da classe B e 97,3% da classe C —, o que indica que a estratégia de "
       f"varredura não privilegiou nem negligenciou os itens de maior valor. Na posição parcial "
       f"da primeira etapa, a cobertura por valor era de apenas 13,8%, porque os bens de maior "
       f"valor concentravam-se nos depósitos e no centro de processamento de dados, então "
       f"pendentes; a segunda etapa reverteu esse quadro.")
blank()
body_p(f"A totalidade dos {br(M['conferidos'])} bens conferidos foi classificada em estado de "
       f"conservação “Bom”. Diferentemente do piloto, a campanha censitária não realizou "
       f"registro fotográfico dos bens — opção operacional que privilegiou o ritmo de "
       f"conferência em uma equipe reduzida —, de modo que a evidência de inventário "
       f"restringiu-se ao local, ao estado de conservação, ao empregado responsável e ao instante "
       f"da sincronização. O resultado, ainda que positivo, sugere baixa discriminação da escala "
       f"em uso e recomenda, para os próximos ciclos, a adoção de gradação mais granular "
       f"(ótimo, bom, regular, ruim e inservível) associada ao registro fotográfico já "
       f"disponível no sistema — insumo necessário às decisões de manutenção, reavaliação e "
       f"desfazimento previstas no ciclo de vida dos ativos. A limitação foi de "
       f"parametrização e de procedimento, e não do método: os campos já estavam estruturados e "
       f"a alteração não exigiria mudança de arquitetura.")
body_p(f"A campanha localizou fisicamente, ainda, {M['nao_ativos_encontrados']['n']} bens já "
       f"baixados ou doados contabilmente — 27 baixados e dois doados —, dos quais "
       f"{M['nao_ativos_encontrados']['em_deposito']} nos depósitos. O achado revela descompasso "
       f"entre o desfazimento contábil e o físico: bens retirados do ativo permanecem ocupando "
       f"espaço e sendo conferidos, o que onera o inventário e expõe o órgão a risco de "
       f"reutilização indevida. A identificação automática desses casos, impossível no processo "
       f"manual baseado em listagens de bens ativos, permite programar o desfazimento físico e "
       f"encerrar o ciclo de vida desses itens.")
body_p("A confiabilidade do dado locacional produziu também desdobramentos administrativos "
       "imediatos. A partir da base conferida, o módulo de termos do sistema passou a gerar "
       "automaticamente os termos de responsabilidade por centro de custos e os termos "
       "individuais de bens sob guarda pessoal — como notebooks e smartphones —, documentos "
       "antes elaborados manualmente a cada movimentação. O histórico de movimentações "
       "registrado pelo sistema também passou a alimentar a apuração de responsabilidade, com "
       "separação entre quem confere, quem movimenta e quem responde pelo bem.")
blank()
subtitle("Comparação com o cenário manual e implicações econômicas")
blank()
body_p(f"Na comparação com o cenário anterior, o tempo estimado do processo manual (AS-IS), "
       f"projetado a partir dos parâmetros da literatura, situou-se na faixa de 125 a 188 horas "
       f"para o mesmo universo de bens, ante as 37,6 horas projetadas pelo modelo e as "
       f"{br(M['tempo_efetivo_h'],1)} horas efetivas medidas na campanha censitária — cerca de "
       f"{br(M['horas_total_projetado'])} horas projetadas para o acervo completo (Figura 11). "
       f"Ainda que o baseline manual constitua estimativa, e não medição direta, a magnitude da "
       f"diferença sustenta a hipótese de redução de tempo superior a 75% — patamar coerente com "
       f"o reportado por Paul et al. (2024) e Madeira Junior e Silveira (2024). A Tabela 11 "
       f"sintetiza a comparação entre os cenários.")
fig("piloto_comparacao_asis.png",
    "Figura 11. Tempo de inventário automatizado (medido e projetado) versus manual AS-IS (estimado pela literatura)",
    "Fonte: Resultados originais da pesquisa")
caption("Tabela 11. Comparação entre o processo manual (AS-IS) e o automatizado (TO-BE)")
table([
    ["Dimensão", "Manual (AS-IS)", "Automatizado (TO-BE)"],
    ["Tempo total de conferência", "125 a 188 horas (estimado)", f"{br(M['tempo_efetivo_h'],1)} horas medidas; ≈ {br(M['horas_total_projetado'])} horas projetadas"],
    ["Registro da conferência", "Planilhas e anotações", "Sistema web com hora, responsável e local"],
    ["Localização de divergências", "Latente, sob demanda", "Imediata, em campo (569 casos)"],
    ["Reorganização administrativa durante a coleta", "Exige nova conferência", "Reconciliação retroativa (569 → 280)"],
    ["Bens baixados ainda presentes", "Invisíveis às listagens", "Identificados (29 casos)"],
    ["Consolidação dos dados", "Digitação manual", "Sincronização automática em lotes"],
    ["Indicadores gerenciais", "Inexistentes", "Painéis e relatórios do pipeline"],
    ["Custo de licenciamento", "—", "Zero (open source)"],
], col_widths=[5.0, 5.0, 5.5])
caption("Fonte: Resultados originais da pesquisa")
blank()
body_p("Do ponto de vista econômico, o custo de implantação da solução restringiu-se a horas "
       "de desenvolvimento do próprio pesquisador e à infraestrutura de servidor já existente "
       "no órgão, sem despesas de licenciamento — em linha com o argumento de viabilidade "
       "financeira de tecnologias de rastreamento no setor público apresentado por Brito et al. "
       "(2019) e com os fatores de adoção de software livre discutidos por Sánchez et al. "
       "(2020). Considerando-se apenas a economia de horas de trabalho da equipe de "
       "conferência, a redução estimada aproxima-se de cem horas ou mais por ciclo anual de "
       "inventário, sem contabilizar os ganhos indiretos de eliminação de retrabalho, de "
       "redução de digitação manual e de disponibilidade imediata da informação para a comissão "
       "de inventário.")
blank()
subtitle("Arquitetura progressiva de automação: pipeline e carga no sistema de gestão")
blank()
body_p("Os dados consolidados das conferências alimentaram o pipeline em Python da Fase 2, "
       "responsável pela validação dos registros, pelo cálculo automático dos indicadores "
       "apresentados nesta seção e pela geração dos painéis gerenciais. A disponibilidade de "
       "indicadores de cobertura, produtividade e divergência em tempo próximo ao real "
       "converteu o inventário de um evento anual de prestação de contas em um instrumento "
       "contínuo de gestão — movimento alinhado à literatura de administração pública orientada "
       "por dados, que condiciona o valor da análise à existência de um problema decisório "
       "concreto (Arnaboldi e Azzone, 2020) e de capacidade organizacional para explorá-la "
       "(Merhi, 2021; Broomfield e Reutter, 2021). Os painéis atenderam, ainda, à demanda por "
       "informação estruturada para decisão sobre ativos fixos identificada por Roje et al. "
       "(2025).")
body_p("A etapa final do pipeline foi a carga das leituras no sistema de gestão patrimonial do "
       "órgão — o mesmo que emite os termos de responsabilidade —, na qual o inventário passa a "
       "existir como evento formal, com comissão, salas no escopo e leituras vinculadas aos "
       "bens do cadastro vigente. A importação foi atômica e validada: cada leitura precisou "
       "referenciar um bem existente, uma sala do cadastro, um integrante da comissão, uma data "
       "válida e um estado de conservação do domínio, e qualquer rejeição interromperia a carga "
       "inteira. No ensaio de carga com as 3.457 conferências da campanha, aplicada a tabela "
       "de correspondência de localizações, todas as 3.456 leituras com número de tombamento "
       "foram aceitas, nenhuma foi rejeitada por bem inexistente e as 98 salas resultantes "
       "coincidiram exatamente com as localizações ativas do cadastro atualizado. Essa foi a "
       "validação de ponta a ponta da arquitetura: dados coletados em campo por quatro "
       "operadores, ao longo de 15 dias e atravessando uma reestruturação administrativa, "
       "chegaram íntegros ao sistema de gestão sem qualquer redigitação.")
body_p("A Tabela 12 consolida os indicadores-chave de desempenho definidos para o "
       "acompanhamento contínuo do inventário no painel gerencial, com os valores apurados ao "
       "final da campanha censitária.")
caption("Tabela 12. Indicadores-chave do painel gerencial e valores ao final da campanha")
table([
    ["Indicador", "Definição", "Valor apurado"],
    ["Cobertura do inventário", "Bens ativos conferidos / bens ativos", f"{br(M['cobertura'],1)}%"],
    ["Cobertura por valor", "Valor conferido / valor do acervo móvel", f"{br(M['cobertura_valor_moveis_pct'],1)}%"],
    ["Taxa de concordância", "Bens no local cadastrado / bens conferidos", f"{br(M['concordancia_pct'],1)}%"],
    ["Taxa de divergência", "Bens fora do local cadastrado / bens conferidos", f"{br(M['divergencia_pct'],1)}%"],
    ["Taxa de divergência reconciliada", "Idem, contra o cadastro reestruturado", "8,1%"],
    ["Produtividade", "Bens conferidos / hora efetiva", f"≈ {br(M['produtividade'])} bens/h"],
    ["Tempo médio por bem", "Hora efetiva / bens conferidos", f"≈ {br(M['seg_por_bem'])} s"],
    ["Bens pendentes", "Bens ativos não conferidos", str(M["pendentes"])],
    ["Bens baixados localizados", "Bens não ativos conferidos fisicamente", str(M["nao_ativos_encontrados"]["n"])],
], col_widths=[4.5, 7.0, 4.0])
caption("Fonte: Resultados originais da pesquisa")
blank()
subtitle("Limitações do estudo")
blank()
body_p("Cinco limitações devem ser registradas. Primeira: o baseline do processo manual foi "
       "estimado a partir da literatura, e não de cronometragem histórica, de modo que a "
       "comparação AS-IS/TO-BE tem natureza indicativa. Segunda: a medição de produtividade "
       "apoia-se nos instantes de sincronização dos lotes, e não em marcações de início e fim "
       "por bem; a regra de exclusão de pausas superiores a uma hora é uma convenção, e os "
       "dias com poucos lotes produzem estimativas instáveis, como o valor atípico de 10 set. "
       "Terceira: a campanha cobria, até o fechamento desta análise, 97,4% do acervo ativo, e o "
       "desfecho das 569 divergências ainda não havia sido classificado entre movimentação "
       "legítima, erro de cadastro ou bem não localizado. Quarta: a acurácia locacional é "
       "sensível à granularidade e à estabilidade da malha de localizações — parte das "
       "divergências reflete escolhas de modelagem ou a reestruturação administrativa ocorrida "
       "durante a coleta, e a tabela de correspondência usada na reconciliação foi derivada por "
       "regra de maioria, de modo que setores desmembrados admitem mais de uma correspondência "
       "válida. Quinta: por se tratar de pesquisa-ação "
       "em um único órgão, a generalização dos resultados requer cautela, ainda que a natureza "
       "open source e conteinerizada da solução favoreça sua replicação (Nüst et al., 2020; "
       "Sánchez et al., 2020).")
blank()

# ================================================================ CONCLUSÕES
title_sec("Conclusões")
blank()
body_p("A arquitetura progressiva de automação baseada em ferramentas open source reduziu "
       "expressivamente o tempo do inventário patrimonial do órgão estudado e confirmou a "
       "hipótese de trabalho: a redução de tempo frente ao cenário manual estimado superou o "
       "patamar reportado pela literatura, sem perda de acurácia atribuível ao método. A maior "
       "parcela do ganho decorreu da eliminação do duplo manuseio dos dados, e não do hardware "
       "de leitura, o que reposicionou a identificação por radiofrequência como otimização "
       "incremental para ambientes de alta densidade.")
body_p("O modelo de estimativa por simulação de Monte Carlo mostrou-se válido como "
       "instrumento de planejamento: construído a partir de um único setor-piloto, capturou a "
       "ordem de grandeza do esforço da campanha completa. A campanha revelou rendimentos "
       "decrescentes entre setores administrativos, depósitos e varredura residual, padrão que "
       "deve ser incorporado ao planejamento de inventários em outros órgãos, sob pena de "
       "subestimar o esforço total a partir dos setores de conferência mais fácil.")
body_p("O processo automatizado revelou valor diagnóstico imediato ao identificar bens "
       "deslocados e bens já baixados ainda fisicamente presentes, convertendo em informação "
       "acionável um passivo que o processo manual mantinha latente. A reconciliação com o "
       "cadastro reestruturado mostrou que a acurácia locacional depende também da estabilidade "
       "da malha de localizações durante a coleta e deve ser reportada com referência explícita "
       "ao cadastro contra o qual foi calculada; o registro estruturado de cada leitura permitiu "
       "absorver a reorganização administrativa sem nova ida a campo, com carga íntegra no "
       "sistema de gestão patrimonial, validação de ponta a ponta da arquitetura proposta.")
body_p("A solução demonstrou escalabilidade ao absorver mais operadores e dias de campo sem "
       "alteração de arquitetura, e sua natureza open source e conteinerizada, com custo de "
       "implantação próximo a zero, torna o modelo replicável por órgãos públicos de porte "
       "similar. Como desdobramentos futuros, recomendam-se a conclusão dos bens físicos "
       "pendentes, a classificação do desfecho das divergências, o desfazimento físico dos bens "
       "baixados localizados, a adoção de escala de conservação mais granular, a consolidação "
       "dos painéis gerenciais e a avaliação da integração de leitores RFID UHF prevista na "
       "Fase 3, prioritariamente nos depósitos.")
blank()

# ================================================================ AGRADECIMENTO
title_sec("Agradecimento")
blank()
body_p("Ao orientador, pela condução criteriosa, e à equipe do Setor de Patrimônio, pelo "
       "empenho nas conferências físicas que viabilizaram esta pesquisa.")
blank()

# ================================================================ REFERÊNCIAS
title_sec("Referências")
blank()
REFS = [
 "Arnaboldi, M.; Azzone, G. 2020. Data science in the design of public policies: dispelling "
 "the obscurity in matching policy demand and data offer. Heliyon 6(6): e04300. Disponível "
 "em: https://doi.org/10.1016/j.heliyon.2020.e04300. Acesso em: 24 ago. 2026.",

 "Boettiger, C. 2015. An introduction to Docker for reproducible research, with examples "
 "from the R environment. ACM SIGOPS Operating Systems Review 49(1): 71-79. Disponível em: "
 "https://doi.org/10.1145/2723872.2723882. Acesso em: 24 ago. 2026.",

 "Brasil. 1988a. Constituição da República Federativa do Brasil de 1988. Diário Oficial da "
 "União, Brasília, 5 out. 1988. Seção 1, p. 1.",

 "Brasil. Secretaria de Administração Pública da Presidência da República [SEDAP]. 1988b. "
 "Instrução Normativa nº 205, de 8 de abril de 1988. Racionalização do uso de material e "
 "controle patrimonial. Diário Oficial da União, Brasília, 11 abr. 1988. Seção 1.",

 "Brasil. Conselho Nacional de Saúde. 2016. Resolução nº 510, de 7 de abril de 2016. Dispõe "
 "sobre as normas aplicáveis a pesquisas em Ciências Humanas e Sociais. Diário Oficial da "
 "União, Brasília, 24 maio 2016. Seção 1, p. 44-46.",

 "Brito, C.V.S.P.; Santos, W.B.; Galhardo, C.X.; Santos, V.M.L. 2019. Etiquetas inteligentes "
 "na administração pública: análise da viabilidade no controle patrimonial da UNIVASF. "
 "ForScience: Revista Científica do IFMG 7(2): e00661. Disponível em: "
 "https://doi.org/10.29069/forscience.2019v7n2.e661. Acesso em: 25 set. 2026.",

 "Broomfield, H.; Reutter, L. 2021. Towards a data-driven public administration: an "
 "empirical analysis of nascent phase implementation. Scandinavian Journal of Public "
 "Administration 25(2): 73-97. Disponível em: https://doi.org/10.58235/sjpa.v25i2.7117. "
 "Acesso em: 24 ago. 2026.",

 "Conselho Federal de Contabilidade [CFC]. 2017. NBC TSP 07 – Ativo Imobilizado. CFC, "
 "Brasília, DF, Brasil.",

 "Hanis, M.H.; Trigunarsyah, B.; Susilawati, C. 2011. The application of public asset "
 "management in Indonesian local government: a case study in South Sulawesi province. "
 "Journal of Corporate Real Estate 13(1): 36-47. Disponível em: "
 "https://doi.org/10.1108/14630011111120332. Acesso em: 24 ago. 2026.",

 "International Organization for Standardization [ISO]. 2014. ISO 55000:2014 – Asset "
 "management – Overview, principles and terminology. ISO, Genebra, Suíça.",

 "Kaganova, O.; Amoils, J.M. 2020. Central government property asset management: a review "
 "of international changes. Journal of Corporate Real Estate 22(3): 239-260. Disponível em: "
 "https://doi.org/10.1108/JCRE-09-2019-0038. Acesso em: 24 ago. 2026.",

 "Kar, S.; Bhimrajka, S.; Kumar, A.; Mukherjee, S. 2022. Mobile based inventory management "
 "system with QR code. In: IEEE International Conference on Electronics, Computing and "
 "Communication Technologies [CONECCT], 2022, Bangalore, Índia. Anais... p. 1-6. Disponível "
 "em: https://doi.org/10.1109/CONECCT55679.2022.9865739. Acesso em: 25 set. 2026.",

 "Kirešová, S.; Guzan, M.; Fecko, B.; Somka, O.; Rusyn, V.; Yatsiuk, R. 2023. Grafana as a "
 "visualization tool for measurements. In: IEEE International Conference on Modern "
 "Electrical and Energy Systems, 2023, Košice, Eslováquia. Anais... p. 1-6. Disponível em: "
 "https://doi.org/10.1109/MEES61502.2023.10402486. Acesso em: 24 ago. 2026.",

 "Lim, M.K.; Bahr, W.; Leung, S.C.H. 2013. RFID in the warehouse: a literature analysis "
 "(1995-2010) of its applications, benefits, challenges and future trends. International "
 "Journal of Production Economics 145(1): 409-430. Disponível em: "
 "https://doi.org/10.1016/j.ijpe.2013.05.006. Acesso em: 24 ago. 2026.",

 "Madeira Junior, J.J.; Silveira, L.E.C. 2024. Tecnologia RFID na gestão de estoques: uma "
 "alternativa ao controle tradicional na UFSC. Trabalho de Conclusão de Curso. Universidade "
 "Federal de Santa Catarina, Florianópolis, SC, Brasil. Disponível em: "
 "https://repositorio.ufsc.br/handle/123456789/262052. Acesso em: 24 ago. 2026.",

 "Merhi, M.I. 2021. Evaluating the critical success factors of data intelligence "
 "implementation in the public sector using analytical hierarchy process. Technological "
 "Forecasting and Social Change 173: 121180. Disponível em: "
 "https://doi.org/10.1016/j.techfore.2021.121180. Acesso em: 24 ago. 2026.",

 "Nüst, D.; Sochat, V.; Marwick, B.; Eglen, S.J.; Head, T.; Hirst, T.; Evans, B.D. 2020. "
 "Ten simple rules for writing Dockerfiles for reproducible data science. PLOS "
 "Computational Biology 16(11): e1008316. Disponível em: "
 "https://doi.org/10.1371/journal.pcbi.1008316. Acesso em: 24 ago. 2026.",

 "Paul, P.O.; Ogugua, J.O.; Eyo-Udo, N.L. 2024. Innovations in fixed asset management: "
 "enhancing efficiency through advanced tracking and maintenance systems. International "
 "Journal of Science and Technology Research Archive 7(1): 19-26. Disponível em: "
 "https://doi.org/10.53771/ijstra.2024.7.1.0053. Acesso em: 24 ago. 2026.",

 "Roje, G.; Anessi Pessina, E.; Botica Redmayne, N. 2025. Information needs for managing "
 "fixed public sector assets: an exploratory analysis in South-Eastern Europe. Journal of "
 "Public Budgeting, Accounting & Financial Management 37(1): 109-128. Disponível em: "
 "https://doi.org/10.1108/JPBAFM-06-2024-0096. Acesso em: 24 ago. 2026.",

 "Sánchez, V.R.; Neira Ayuso, P.; Galindo, J.A.; Benavides, D. 2020. Open source adoption "
 "factors: a systematic literature review. IEEE Access 8: 94594-94609. Disponível em: "
 "https://doi.org/10.1109/ACCESS.2020.2993248. Acesso em: 24 ago. 2026.",

 "Thiollent, M. 2011. Metodologia da Pesquisa-Ação. 18ed. Cortez, São Paulo, SP, Brasil.",

 "Ton, N.T.N.; Le, M.T.; Lam, T.T.; Do, T.D. 2024. Inventory management system using RFID "
 "and barcode: a modeling and simulation approach. In: International Conference on Green "
 "Technology and Sustainable Development [GTSD], 7., 2024, Ho Chi Minh, Vietnã. Anais... "
 "p. 362-367. Disponível em: https://doi.org/10.1109/GTSD62346.2024.10675053. Acesso em: "
 "25 set. 2026.",

 "van Loon, A.; Toshkov, D. 2015. Adopting open source software in public administration: "
 "the importance of boundary spanners and political commitment. Government Information "
 "Quarterly 32(2): 207-215. Disponível em: https://doi.org/10.1016/j.giq.2015.01.004. "
 "Acesso em: 24 ago. 2026.",

 "Vilela, F.A.; Times, V.C.; Bernardi, A.C.C.; Freitas, A.P.; Ciferri, R.R. 2023. A "
 "non-intrusive and reactive architecture to support real-time ETL processes in data "
 "warehousing environments. Heliyon 9(5): e15728. Disponível em: "
 "https://doi.org/10.1016/j.heliyon.2023.e15728. Acesso em: 26 ago. 2026.",

 "Zelbst, P.J.; Green Junior, K.W.; Sower, V.E.; Reyes, P.M. 2012. Impact of RFID on "
 "manufacturing effectiveness and efficiency. International Journal of Operations & "
 "Production Management 32(3): 329-350. Disponível em: "
 "https://doi.org/10.1108/01443571211212600. Acesso em: 25 set. 2026.",
]
for r in REFS:
    para(r, align="left", spacing=1.0, space_after=10)
blank()

# ================================================================ APÊNDICE A
title_sec("Apêndice A. Infraestrutura replicável e repositórios da solução")
blank()
body_p("A solução foi construída exclusivamente com ferramentas open source: Python 3 (pandas, "
       "NumPy e Matplotlib) para o pipeline de dados e as análises; PostgreSQL para "
       "armazenamento; Grafana e Streamlit para os painéis gerenciais; e Docker Compose para a "
       "conteinerização da infraestrutura, com versões fixadas e dependências explícitas, "
       "conforme as boas práticas de reprodutibilidade computacional. A semente aleatória da "
       "simulação de Monte Carlo foi fixada, de modo que os resultados quantitativos deste "
       "trabalho podem ser regenerados integralmente a partir da base de dados e dos scripts.")
body_p("Os componentes públicos da solução estão disponíveis nos seguintes repositórios: "
       "sistema web de conferência física (Fase 1), em operação em "
       "https://sistemadeinventario.com.br; dados públicos do acervo para construção de "
       "painéis, em https://github.com/ToNiauM/acervo-cfc; e o sistema de gestão patrimonial que "
       "recebeu a carga do inventário e gera os termos de responsabilidade por centro de custos e "
       "individuais, em https://github.com/ToNiauM/sistema-de-inventario (o README do repositório "
       "registra o método de desenvolvimento e os achados desta pesquisa). Os scripts que "
       "recalculam todos os indicadores e figuras deste trabalho a partir do relatório da "
       "campanha, com os valores agregados resultantes, acompanham esse último repositório, na "
       "pasta docs/tcc, junto com o registro da origem de cada número.")
blank()

doc.save(OUT)
print("gerado:", OUT)
