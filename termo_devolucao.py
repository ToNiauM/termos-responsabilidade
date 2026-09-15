from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Pt
from datetime import date

import config
import textos as textos_mod
from termos_html import data_por_extenso

def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def centralizar_celula(celula):
    for paragraph in celula.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tc = celula._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = OxmlElement("w:vAlign")
    vAlign.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "center")
    tcPr.append(vAlign)

def gerar_termo_devolucao(nome, bens, destino, textos=None):
    """bens: dicts de db.bens. Devolve None se a lista estiver vazia."""
    if not bens:
        return None
    doc = Document(str(config.caminho_timbrado()))
    t = textos or textos_mod.PADRAO
    campos = dict(textos_mod.campos_gerais(t), nome=nome, cidade=t["cidade"], data=data_por_extenso(date.today()))

    # Título
    p = doc.paragraphs[0]
    p.clear()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(t["devolucao_titulo"].format_map(campos))
    run.bold = True
    run.italic = False
    run.font.size = Pt(14)

    doc.add_paragraph()

    # Texto de introdução
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    antes, nome_negrito, depois = textos_mod.com_nome(t["devolucao_abertura"], campos)
    p.add_run(antes).bold = False
    if nome_negrito is not None:
        p.add_run(nome_negrito).bold = True
        p.add_run(depois)

    doc.add_paragraph()

    # Tabela de bens
    tabela = doc.add_table(rows=1, cols=4)
    tabela.style = 'Table Grid'
    tabela.alignment = WD_ALIGN_PARAGRAPH.CENTER
    larguras_colunas = [0.5, 4.5, 10, 4]
    cabecalhos = ['Patrimônio', 'Descrição', 'Complemento', 'Valor Atual']
    hdr_cells = tabela.rows[0].cells

    for i, heading in enumerate(cabecalhos):
        hdr_cells[i].text = heading
        paragraph = hdr_cells[i].paragraphs[0]
        run = paragraph.runs[0]
        run.font.size = Pt(12)
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
        row_cells[0].text = str(bem['numero'])
        row_cells[1].text = bem['descricao'] or ""
        row_cells[2].text = bem['complemento'] or ""
        row_cells[3].text = formatar_moeda(bem['valor_atual'] or 0)
        for cell in row_cells:
            centralizar_celula(cell)

    # Linha total
    total = sum(b["valor_atual"] or 0 for b in bens)
    total_row = tabela.add_row().cells
    total_row[0].merge(total_row[2])
    total_row[0].text = "TOTAL"
    run = total_row[0].paragraphs[0].runs[0]
    run.font.size = Pt(12)
    run.bold = True
    centralizar_celula(total_row[0])
    total_row[3].text = formatar_moeda(total)
    centralizar_celula(total_row[3])

    doc.add_paragraph()

    # Data
    p_data = doc.add_paragraph(t["devolucao_data"].format_map(campos))
    p_data.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_paragraph()
    doc.add_paragraph()

    # Assinaturas
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run(nome).bold = True
    p2 = doc.add_paragraph(t["assinatura_eletronica"])
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    p3 = doc.add_paragraph(t["devolucao_recebimento"].format_map(campos))
    p3.alignment = WD_ALIGN_PARAGRAPH.LEFT

    doc.add_paragraph()
    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.add_run(t["recebedor_nome"]).bold = True
    p5 = doc.add_paragraph(t["recebedor_cargo"])
    p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p6 = doc.add_paragraph(t["assinatura_eletronica"])
    p6.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.save(destino)
    return destino
