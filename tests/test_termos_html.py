from datetime import date

import termos_html as th
from tests.test_docx import bens


def test_individual_tabela_80_por_cento_e_total():
    html = th.corpo_individual("ANA SILVA", bens())
    assert "width:80%" in html and "TERMO DE RESPONSABILIDADE" in html
    assert "<b>ANA SILVA</b>" in html and "R$ 1.564,54" in html
    assert html.count("<tr") == 4  # cabeçalho + 2 + total


def test_ccusto_tabela_100_por_cento_e_texto_do_responsavel():
    resp = {"ccustos": "CCI", "responsavel": "JAQUELINE", "matricula": "46", "funcao": "coordenadora"}
    html = th.corpo_ccusto("CCI", resp, bens())
    assert "width:100%" in html and "Termo de Responsabilidade - CCI" in html
    assert "matrícula n.º 46" in html and "Localização" in html


def test_devolucao_com_data_e_assinaturas():
    html = th.corpo_devolucao("ANA SILVA", bens(), hoje=date(2026, 9, 14))
    assert "width:80%" in html and "TERMO DE DEVOLUÇÃO" in html
    assert "Brasília (DF), 14 de setembro de 2026" in html
    assert "Supervisor de Patrimônio" in html


def test_documento_envelopa_com_titulo_e_escapa():
    html = th.documento("Termo", th.corpo_individual("A <B>", []))
    assert html.startswith("<!DOCTYPE html>") and "<title>Termo</title>" in html
    assert "A &lt;B&gt;" in html
