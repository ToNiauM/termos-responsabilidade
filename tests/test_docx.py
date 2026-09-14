from docx import Document
from openpyxl import load_workbook

from tests.conftest import semear


def bens():
    return [
        {"numero": 1001, "descricao": "CADEIRA", "complemento": "GIRATÓRIA", "localizacao": "01 - SALA", "valor_atual": 64.54},
        {"numero": 1002, "descricao": "NOTEBOOK", "complemento": None, "localizacao": "01 - SALA", "valor_atual": 1500.0},
    ]


def linhas(caminho):
    return Document(caminho).tables[0].rows


def test_termo_individual(dados, tmp_path):
    from Script_Termo_Individual import criar_termo_responsabilidade
    destino = criar_termo_responsabilidade("ANA SILVA", bens(), tmp_path / "t.docx")
    tab = linhas(destino)
    assert len(tab) == 4  # cabeçalho + 2 + total
    assert tab[3].cells[3].text == "R$ 1.564,54"
    assert tab[2].cells[2].text == ""  # complemento None vira vazio


def test_termo_centro_e_planilha(dados, tmp_path):
    from Termo_de_Responsabilidade import gerar_planilha_centro, gerar_termo_centro
    resp = {"ccustos": "CCI", "responsavel": "JAQUELINE", "matricula": "46", "funcao": "coordenadora"}
    destino = gerar_termo_centro("CCI", resp, bens(), tmp_path / "c.docx")
    tab = linhas(destino)
    assert len(tab) == 4 and tab[3].cells[4].text == "R$ 1.564,54"
    assert "JAQUELINE" in Document(destino).paragraphs[-1].text
    xlsx = gerar_planilha_centro(bens(), tmp_path / "c.xlsx")
    ws = load_workbook(xlsx).active
    assert ws.max_row == 3 and ws["A1"].value == "numero"


def test_termo_devolucao(dados, tmp_path):
    from termo_devolucao import gerar_termo_devolucao
    destino = gerar_termo_devolucao("ANA SILVA", bens(), tmp_path / "d.docx")
    assert len(linhas(destino)) == 4
    assert gerar_termo_devolucao("ANA SILVA", [], tmp_path / "vazio.docx") is None
