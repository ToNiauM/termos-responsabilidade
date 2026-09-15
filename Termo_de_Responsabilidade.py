from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook

import config
import textos as textos_mod

# Função para formatar moeda no estilo brasileiro sem usar locale
def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")


def _runs_com_negrito(paragrafo, texto, campos):
    """Texto com marcadores substituídos; o {responsavel} sai em negrito."""
    antes, nome, depois = textos_mod.com_nome(texto, campos, "responsavel")
    if antes:
        paragrafo.add_run(antes)
    if nome is not None:
        paragrafo.add_run(nome).bold = True
        if depois:
            paragrafo.add_run(depois)


def gerar_termo_centro(ccustos, responsavel, bens, destino, textos=None):
    """Um termo para um centro de custo. responsavel: dict de db.responsaveis; bens: dicts de db.bens."""
    documento = Document(str(config.caminho_timbrado()))
    t = textos or textos_mod.PADRAO
    campos = {k: responsavel.get(k) or "" for k in ("responsavel", "matricula", "funcao")}
    campos.update(textos_mod.campos_gerais(t), ccustos=ccustos, responsavel=textos_mod.nome_proprio(campos["responsavel"]))
    cabecalho = documento.add_paragraph()
    cabecalho_run = cabecalho.add_run(t["ccusto_titulo"].format_map(campos))
    cabecalho_run.font.bold = True
    cabecalho.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cabecalho_run.font.size = Pt(16)

    soma_valores = sum(b["valor_atual"] or 0 for b in bens)
    grupo_ordenado = sorted(bens, key=lambda b: b["numero"])

    for paragraph in textos_mod.paragrafos(t["ccusto_paragrafos"]):
        paragrafo = documento.add_paragraph()
        paragrafo.paragraph_format.first_line_indent = Inches(0.59)
        paragrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _runs_com_negrito(paragrafo, paragraph, campos)

    tabela = documento.add_table(rows=1, cols=5)
    tabela.style = 'Table Grid'

    hdr_cells = tabela.rows[0].cells
    for i, heading in enumerate(['Número Bem', 'Descrição', 'Complemento', 'Localização', 'Valor Atual']):
        hdr_cells[i].text = heading
        paragraph = hdr_cells[i].paragraphs[0]
        paragraph.clear()
        run = paragraph.add_run(heading)
        run.font.bold = True
        run.font.size = Pt(10)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for bem in grupo_ordenado:
        row_cells = tabela.add_row().cells
        row_cells[0].text = str(bem['numero'])
        row_cells[1].text = bem['descricao'] or ""
        row_cells[2].text = bem['complemento'] or ""
        row_cells[3].text = bem['localizacao'] or ""

        valor_formatado = formatar_moeda(bem['valor_atual'] or 0)
        paragrafo_valor = row_cells[4].paragraphs[0]
        paragrafo_valor.clear()
        run_valor = paragrafo_valor.add_run(valor_formatado)
        run_valor.font.size = Pt(9)
        paragrafo_valor.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        for cell in row_cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9)

    ultima_linha = tabela.add_row().cells
    ultima_linha[0].merge(ultima_linha[1]).merge(ultima_linha[2]).merge(ultima_linha[3])
    ultima_linha[0].text = "TOTAL"
    ultima_linha[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    valor_total_formatado = formatar_moeda(soma_valores)
    ultima_linha[4].text = valor_total_formatado
    ultima_linha[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    for cell in ultima_linha:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(10)
                run.bold = True

    documento.add_paragraph()
    paragrafo_assinatura = documento.add_paragraph()
    for i, linha in enumerate(textos_mod.linhas(t["ccusto_assinatura"])):
        if i:
            paragrafo_assinatura.add_run("\n")
        _runs_com_negrito(paragrafo_assinatura, linha, campos)
    paragrafo_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in paragrafo_assinatura.runs:
        run.font.size = Pt(12)

    documento.save(destino)
    return destino


def gerar_planilha_centro(bens, destino):
    """Planilha com os bens do termo, para quem quiser analisar os dados."""
    wb = Workbook()
    ws = wb.active
    ws.title = "bens"
    ws.append(["numero", "descricao", "complemento", "localizacao", "valor_atual"])
    for b in sorted(bens, key=lambda b: b["numero"]):
        ws.append([b["numero"], b["descricao"], b["complemento"], b["localizacao"], b["valor_atual"]])
    wb.save(destino)
    return destino
