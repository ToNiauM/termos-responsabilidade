from datetime import date

import termos_html as th
import textos
from tests.test_docx import bens


def test_individual_tabela_80_por_cento_e_total():
    html = th.corpo_individual("ANA SILVA", bens())
    assert "width:90%" in html and "TERMO DE RESPONSABILIDADE" in html
    assert "<b>Ana Silva</b>" in html and "ANA SILVA" not in html and "R$ 1.564,54" in html
    assert html.count("<tr") == 4  # cabeçalho + 2 + total


def test_ccusto_tabela_100_por_cento_e_texto_do_responsavel():
    resp = {"ccustos": "CCI", "responsavel": "JAQUELINE", "matricula": "46", "funcao": "coordenadora"}
    html = th.corpo_ccusto("CCI", resp, bens())
    assert "width:90%" in html and "Termo de Responsabilidade - CCI" in html
    assert "matrícula n.º 46" in html and "Localização" in html
    assert html.count("<b>Jaqueline</b>") == 2  # abertura e assinatura, em nome próprio


def test_devolucao_com_data_e_assinaturas():
    html = th.corpo_devolucao("ANA SILVA", bens(), hoje=date(2026, 9, 14))
    assert "width:90%" in html and "TERMO DE DEVOLUÇÃO" in html
    assert "Brasília (DF), 14 de setembro de 2026" in html
    assert "Bruno de Araujo Gomes" in html and "Gerente de Serviços Administrativos" in html


def test_documento_envelopa_com_titulo_e_escapa():
    html = th.documento("Termo", th.corpo_individual("A <B>", []))
    assert html.startswith("<!DOCTYPE html>") and "<title>Termo</title>" in html
    assert "A &lt;b&gt;" in html  # nome vai para nome próprio e é escapado


def test_individual_com_texto_alterado_e_nome_em_negrito():
    t = dict(textos.PADRAO, individual_abertura="TESTE {nome} do {orgao_sigla}.",
             individual_compromissos="a\nb\nc", orgao_sigla="XYZ")
    html = th.corpo_individual("ANA SILVA", bens(), textos=t)
    assert "TESTE <b>Ana Silva</b> do XYZ." in html
    assert '<p class="semrecuo" style="text-align:justify;text-indent:0;margin:0 0 7pt">a</p>' in html
    assert '<p class="semrecuo" style="text-align:justify;text-indent:0;margin:0 0 7pt">c</p>' in html


def test_ccusto_com_dois_paragrafos_e_sigla():
    t = dict(textos.PADRAO, ccusto_paragrafos="Primeiro {ccustos}.\n\nSegundo do {orgao_sigla}.", orgao_sigla="XYZ")
    resp = {"ccustos": "CCI", "responsavel": "JAQUELINE", "matricula": "46", "funcao": "coordenadora"}
    html = th.corpo_ccusto("CCI", resp, bens(), textos=t)
    assert '<p style="text-align:justify;text-indent:1.25cm;margin:0 0 7pt">Primeiro CCI.</p><p style="text-align:justify;text-indent:1.25cm;margin:0 0 7pt">Segundo do XYZ.</p>' in html
    assert "width:90%" in html
    assert "coordenadora do(a) CCI</p>" in html and "do XYZ</p>" not in html   # assinatura sem a sigla do órgão (3e70a8b)


def test_devolucao_usa_cidade_e_recebedor_dos_textos():
    t = dict(textos.PADRAO, cidade="Goiânia (GO)", recebedor_nome="FULANO", recebedor_cargo="Chefe")
    html = th.corpo_devolucao("ANA SILVA", bens(), hoje=date(2026, 9, 14), textos=t)
    assert "Goiânia (GO), 14 de setembro de 2026" in html and "<b>FULANO</b><br>Chefe<br>" in html


def test_escapa_texto_vindo_do_banco():
    t = dict(textos.PADRAO, individual_ciencia="<script>x</script>")
    assert "&lt;script&gt;" in th.corpo_individual("A", [], textos=t)


def test_unidade_alterada_aparece_nos_tres_termos():
    t = dict(textos.PADRAO, unidade_nome="Setor Novo", unidade_sigla="SN")
    resp = {"ccustos": "CCI", "responsavel": "JAQUELINE", "matricula": "46", "funcao": "coordenadora"}
    assert "Setor Novo (SN)" in th.corpo_individual("ANA", bens(), textos=t)
    assert "Setor Novo (SN)" in th.corpo_ccusto("CCI", resp, bens(), textos=t)
    assert "Setor Novo (SN)" in th.corpo_devolucao("ANA", bens(), hoje=date(2026, 9, 14), textos=t)
    assert "Gersev" not in th.corpo_individual("ANA", bens(), textos=t)


def test_estilo_inline_para_o_editor_do_sei():
    """O editor do SEI descarta CSS e classes; só respeita style=. Título menor e centrado, parágrafos
    justificados com recuo, tabela com 90% (decisão do usuário em 2026-09-20)."""
    resp = {"responsavel": "JAQUELINE PORTELA", "matricula": "46", "funcao": "coordenadora"}
    bens = [{"numero": 1, "descricao": "CADEIRA", "complemento": "", "localizacao": "01", "valor_atual": 10.0}]
    html = th.corpo_ccusto("CCI", resp, bens)
    assert '<h1 style="text-align:center;font-size:14pt;margin:10pt 0 12pt">' in html
    assert '<p style="text-align:justify;text-indent:1.25cm;margin:0 0 7pt">' in html      # parágrafo comum
    assert 'class="assinatura" style="text-align:center;text-indent:0;margin:20pt 0 7pt"' in html
    assert "width:90%" in html and "width:100%" not in html and "width:80%" not in html
    ind = th.corpo_individual("ANA SILVA", bens)
    assert 'class="semrecuo" style="text-align:justify;text-indent:0;margin:0 0 7pt"' in ind
    assert "width:90%" in ind
    dev = th.corpo_devolucao("ANA SILVA", bens)
    assert 'class="direita" style="text-align:right;text-indent:0;margin:0 0 7pt"' in dev
    assert "width:90%" in dev
