from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Pt

import config
import textos as textos_mod

# Formata o valor como moeda brasileira, sem usar locale
def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

def centralizar_celula(celula):
    for paragraph in celula.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tc = celula._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = OxmlElement("w:vAlign")
    vAlign.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "center")
    tcPr.append(vAlign)

def criar_termo_responsabilidade(nome, bens, destino, textos=None):
    """bens: lista de dicts com numero, descricao, complemento, valor_atual. Grava em destino."""
    t = textos or textos_mod.PADRAO
    doc = Document(str(config.caminho_timbrado()))
    campos = {"nome": nome, "orgao_sigla": t["orgao_sigla"]}

    p = doc.add_heading(level=1)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(t["individual_titulo"].format_map(campos))
    run.bold = True
    run.italic = False
    run.font.size = Pt(14)

    doc.add_paragraph("")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    antes, nome_negrito, depois = textos_mod.com_nome(t["individual_abertura"], campos)
    p.add_run(antes)
    if nome_negrito is not None:
        p.add_run(nome_negrito).bold = True
        p.add_run(depois)

    doc.add_paragraph("")

    tabela = doc.add_table(rows=1, cols=4)
    tabela.style = 'Table Grid'
    tabela.alignment = WD_ALIGN_PARAGRAPH.CENTER
    larguras_colunas = [0.5, 4.5, 10, 4]

    hdr_cells = tabela.rows[0].cells
    for i, heading in enumerate(['Patrimônio', 'Descrição', 'Complemento', 'Valor Atual']):
        hdr_cells[i].text = heading
        paragraph = hdr_cells[i].paragraphs[0]
        run = paragraph.runs[0]
        run.font.size = Pt(11)
        run.font.name = 'Calibri'
        run.bold = True
        centralizar_celula(hdr_cells[i])

        tc = hdr_cells[i]._tc
        tcPr = tc.get_or_add_tcPr()
        tcW = OxmlElement('w:tcW')
        tcW.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w',
                str(int(larguras_colunas[i] * 567)))
        tcPr.append(tcW)

    for bem in bens:
        row_cells = tabela.add_row().cells
        row_cells[0].text = str(bem["numero"])
        row_cells[1].text = bem["descricao"] or ""
        row_cells[2].text = bem["complemento"] or ""
        row_cells[3].text = formatar_moeda(bem["valor_atual"] or 0)
        for cell in row_cells:
            centralizar_celula(cell)

    valor_total = sum(b["valor_atual"] or 0 for b in bens)
    total_row = tabela.add_row().cells
    total_row[0].merge(total_row[2])
    total_row[0].text = "TOTAL"
    run = total_row[0].paragraphs[0].runs[0]
    run.font.size = Pt(11)
    run.font.name = 'Calibri'
    run.bold = True
    centralizar_celula(total_row[0])
    total_row[3].text = formatar_moeda(valor_total)
    centralizar_celula(total_row[3])

    doc.add_paragraph("")

    doc.add_paragraph("")
    p = doc.add_paragraph(t["individual_compromissos_intro"].format_map(campos))
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for item in textos_mod.linhas(t["individual_compromissos"]):
        doc.add_paragraph("")
        p = doc.add_paragraph(item.format_map(campos))
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.add_paragraph("")
    doc.add_paragraph(t["individual_ciencia"].format_map(campos))
    doc.add_paragraph("")

    p_assinado = doc.add_paragraph()
    p_assinado.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_assinado.add_run(nome).bold = True

    p_assinado = doc.add_paragraph()
    p_assinado.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_assinado.add_run(t["assinatura_eletronica"])

    doc.save(str(destino))
    return Path(destino)
