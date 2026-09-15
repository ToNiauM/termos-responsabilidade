"""Macro grafico e helper graficos.py."""

from flask import render_template_string

import graficos
from tests.conftest import semear


def test_graficos_helper_urls_e_tabela():
    op = graficos.rosca([("ATIVO", 3), ("BAIXADO", 1)], total=(4, "bens"), urls={"ATIVO": "/recorte?situacao=ATIVO"})
    assert op["series"][0]["data"][0] == {"name": "ATIVO", "value": 3, "url": "/recorte?situacao=ATIVO"}
    assert op["graphic"][0]["style"]["text"] == "4\nbens"
    op = graficos.barras_horizontais(["CCI"], [3], "Bens", escala=True, urls=["/r?ccusto=CCI"])
    assert op["series"][0]["data"] == [{"value": 3, "url": "/r?ccusto=CCI"}] and "visualMap" in op
    t = graficos.tabela_dados(["A", "B"], [[{"valor": "x", "url": "/x"}, 2]])
    assert t["linhas"][0] == [{"valor": "x", "url": "/x"}, {"valor": 2}]


def test_macro_grafico_renderiza_json_e_tabela(dados):
    semear(dados)
    from app import app
    g = {"id": "g1", "titulo": "Teste", "subtitulo": None, "alto": False, "col": None, "resumo": "Teste: 1",
         "opcoes": {"series": [{"type": "pie", "data": [{"name": "<b>", "value": 1}]}]},
         "tabela": graficos.tabela_dados(["Rótulo", "Bens"], [[{"valor": "CCI", "url": "/recorte?ccusto=CCI"}, 1]])}
    with app.test_request_context():
        html = render_template_string('{% from "_macros.html" import grafico %}{{ grafico(g) }}', g=g)
    assert 'data-grafico="g1"' in html and '<script type="application/json" id="g1">' in html
    assert "<b>" not in html.split('id="g1">')[1].split("</script>")[0]      # tojson escapa
    assert 'href="/recorte?ccusto=CCI"' in html and "Ver dados" in html and "dsgov-grafico" in html


def test_url_recorte_mantem_situacao_vazia(dados):
    from app import app
    import painel
    with app.test_request_context():
        assert painel.url_recorte({}, ccusto="CCI") == "/recorte?ccusto=CCI&situacao="
        assert painel.url_recorte({"situacao": "ATIVO"}, ccusto="CCI") == "/recorte?situacao=ATIVO&ccusto=CCI"
        assert painel.url_recorte_xlsx({}) == "/recorte/xlsx?situacao="
        assert painel.url_recorte_xlsx({"situacao": "ATIVO"}) == "/recorte/xlsx?situacao=ATIVO"


def test_cards_graficos_tipos_urls_e_omissao(dados):
    from tests.test_db import _semear_painel
    _semear_painel(dados)
    import db, painel
    from app import app
    f = {"situacao": "ATIVO"}
    with app.test_request_context():
        cards = painel.cards_graficos(db.dimensoes(dados, f), f)
        ids = [c["id"] for c in cards]
        assert ids == ["g-situacao", "g-centro", "g-classificacao", "g-localizacao", "g-idade", "g-ano", "g-faixa", "g-pessoa"]
        por = {c["id"]: c for c in cards}
        assert por["g-situacao"]["opcoes"]["series"][0]["type"] == "pie"
        assert por["g-centro"]["opcoes"]["series"][0]["type"] == "bar" and por["g-centro"]["opcoes"]["yAxis"]["type"] == "category"
        assert por["g-idade"]["opcoes"]["xAxis"]["type"] == "category" and por["g-ano"]["opcoes"]["series"][0]["type"] == "line"
        assert por["g-centro"]["opcoes"]["series"][0]["data"][0]["url"] == "/recorte?situacao=ATIVO&ccusto=-"
        assert por["g-centro"]["opcoes"]["series"][0]["data"][1]["url"] == "/recorte?situacao=ATIVO&ccusto=CCI"
        assert por["g-situacao"]["opcoes"]["series"][0]["data"][0]["url"] == "/recorte?situacao=ATIVO"
        assert por["g-classificacao"]["tabela"]["linhas"][-1][0]["valor"] == "SEDE"          # imóveis por último na tabela
        assert "SEDE" not in [d["name"] for d in por["g-classificacao"]["opcoes"]["series"][0]["data"]]
        assert por["g-ano"]["opcoes"]["series"][0]["data"][-1]["url"] == "/recorte?situacao=ATIVO&ano=2024"
        assert por["g-faixa"]["tabela"]["linhas"][0][2]["valor"] == "R$ 64,54"
        cards = painel.cards_graficos(db.dimensoes(dados, {"situacao": "ATIVO", "ccusto": "CCI"}), {"situacao": "ATIVO", "ccusto": "CCI"}, omitir=("situacao", "ccusto"))
        assert "g-centro" not in [c["id"] for c in cards] and "g-situacao" not in [c["id"] for c in cards]
        assert painel.descrever({"situacao": "ATIVO", "ccusto": "CCI", "entrada_de": "2020-01-01"}, {"ccusto": "CCI – JAQUELINE"}) == \
            "Bens ATIVO · centro de custo CCI – JAQUELINE · entrada a partir de 01/01/2020"
        assert painel.moeda(1234.5) == "R$ 1.234,50"
