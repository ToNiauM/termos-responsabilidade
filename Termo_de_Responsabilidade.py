from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook

import config

# Função para formatar moeda no estilo brasileiro sem usar locale
def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

texto_padrao = """
Pelo presente termo, eu, {responsavel}, matrícula n.º {matricula}, {funcao} do(a) {ccustos} do CFC, declaro que os bens patrimoniais abaixo discriminados se encontram na localização sob a minha guarda e responsabilidade.
Assumo TOTAL responsabilidade pelos referidos bens, comprometendo-me a informar o Setor de Patrimônio quanto a qualquer alteração e/ou irregularidade, bem como zelar pela guarda e bom uso do patrimônio público.
Em caso de extravio ou dano a bem sob a minha responsabilidade, comprometo-me a ressarcir o CFC dos prejuízos causados.
Observações:
Em caso de perda ou roubo do bem, o responsável deverá registrar boletim de ocorrência policial e apresentar ao Setor de Patrimônio;
Ao final do mandato, função ou designação, o responsável deverá devolver o bem, se for o caso.
No caso de movimentação e transferência de bens entre as unidades administrativas, o Setor de Patrimônio utilizará o Termo de Transferência disponível no SEI, que será apensado a processo específico até a emissão de um novo termo atualizado.
"""


def gerar_termo_centro(ccustos, responsavel, bens, destino):
    """Um termo para um centro de custo. responsavel: dict de db.responsaveis; bens: dicts de db.bens."""
    documento = Document(str(config.caminho_timbrado()))
    cabecalho = documento.add_paragraph()
    cabecalho_run = cabecalho.add_run(f'Termo de Responsabilidade - {ccustos}')
    cabecalho_run.font.bold = True
    cabecalho.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cabecalho_run.font.size = Pt(16)

    soma_valores = sum(b["valor_atual"] or 0 for b in bens)
    grupo_ordenado = sorted(bens, key=lambda b: b["numero"])

    for paragraph in texto_padrao.strip().split('\n'):
        paragrafo = documento.add_paragraph()
        paragrafo.paragraph_format.first_line_indent = Inches(0.59)
        paragrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragrafo.add_run(paragraph.format(
            responsavel=responsavel['responsavel'],
            matricula=responsavel['matricula'],
            funcao=responsavel['funcao'],
            ccustos=ccustos
        ))

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
    paragrafo_assinatura.add_run(f"{responsavel['responsavel']}\n{responsavel['funcao']} do(a) {ccustos} do CFC")
    paragrafo_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in paragrafo_assinatura.runs:
        run.font.size = Pt(12)

    documento.save(str(destino))
    return Path(destino)


def gerar_planilha_centro(bens, destino):
    """Planilha com os bens do termo, para quem quiser analisar os dados."""
    wb = Workbook()
    ws = wb.active
    ws.title = "bens"
    ws.append(["numero", "descricao", "complemento", "localizacao", "valor_atual"])
    for b in sorted(bens, key=lambda b: b["numero"]):
        ws.append([b["numero"], b["descricao"], b["complemento"], b["localizacao"], b["valor_atual"]])
    wb.save(str(destino))
    return Path(destino)
