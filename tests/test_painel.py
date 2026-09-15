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
    """Tipo do gráfico pela quantidade de itens da dimensão (regra do usuário): com a semente de
    _semear_painel, sob ATIVO, situacao=2, centro=2, classificacao=3, localizacao=4, idade=5 e ano=5
    itens (rosca); faixa=6 itens (barras horizontais)."""
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
        assert por["g-centro"]["opcoes"]["series"][0]["type"] == "pie"          # 2 itens (CCI, sem centro)
        assert por["g-classificacao"]["opcoes"]["series"][0]["type"] == "pie"   # 3 itens
        assert por["g-idade"]["opcoes"]["series"][0]["type"] == "pie"           # 5 itens (faixas fixas)
        assert por["g-ano"]["opcoes"]["series"][0]["type"] == "pie"             # 5 itens
        assert por["g-faixa"]["opcoes"]["series"][0]["type"] == "bar" and por["g-faixa"]["opcoes"]["xAxis"]["type"] == "category"  # 6 itens, rótulo curto → colunas
        assert por["g-situacao"]["subtitulo"] == "Todos os bens" and por["g-centro"]["subtitulo"] == "Bens ativos"
        assert por["g-centro"]["opcoes"]["series"][0]["data"][0]["url"] == "/recorte?situacao=ATIVO&ccusto=-"
        assert por["g-centro"]["opcoes"]["series"][0]["data"][1]["url"] == "/recorte?situacao=ATIVO&ccusto=CCI"
        assert por["g-situacao"]["opcoes"]["series"][0]["data"][0]["url"] == "/recorte?situacao=ATIVO"
        assert por["g-classificacao"]["tabela"]["linhas"][-1][0]["valor"] == "SEDE"          # imóveis por último na tabela
        assert "SEDE" in [d["name"] for d in por["g-classificacao"]["opcoes"]["series"][0]["data"]]  # mas no gráfico, como qualquer classe
        assert por["g-ano"]["opcoes"]["series"][0]["data"][-1]["url"] == "/recorte?situacao=ATIVO&ano=2024"
        assert por["g-faixa"]["tabela"]["linhas"][0][2]["valor"] == "R$ 64,54"
        cards = painel.cards_graficos(db.dimensoes(dados, {"situacao": "ATIVO", "ccusto": "CCI"}), {"situacao": "ATIVO", "ccusto": "CCI"}, omitir=("situacao", "ccusto"))
        assert "g-centro" not in [c["id"] for c in cards] and "g-situacao" not in [c["id"] for c in cards]
        assert painel.descrever({"situacao": "ATIVO", "ccusto": "CCI", "entrada_de": "2020-01-01"}, {"ccusto": "CCI – JAQUELINE"}) == \
            "Bens ATIVO · centro de custo CCI – JAQUELINE · entrada a partir de 01/01/2020"
        assert painel.moeda(1234.5) == "R$ 1.234,50"


def test_grafico_tipos_pela_quantidade_de_itens(dados):
    """painel._grafico: até 5 rosca, 6-10 barras, 11-20 colunas, mais de 20 colunas com os 20 maiores."""
    import painel
    from app import app

    def item(i):
        return {"chave": str(i), "rotulo": str(i), "quantidade": i, "valor": 0}

    with app.test_request_context():
        op, sub, altura, col = painel._grafico([item(i) for i in range(25, 0, -1)], {"situacao": "ATIVO"}, "pessoa")
        assert op["series"][0]["type"] == "bar" and op["yAxis"]["type"] == "category"   # rótulo longo → barras horizontais
        assert len(op["yAxis"]["data"]) == 20 and sub is not None and "20 maiores" in sub and altura == "extra" and col == "col-12"
        assert op["yAxis"]["data"][0] == "25"                       # os 20 maiores: começa pelo maior
        op, sub, altura, col = painel._grafico([item(i) for i in range(8, 0, -1)], {"situacao": "ATIVO"}, "ccusto")
        assert op["yAxis"]["type"] == "category" and altura is None and col is None

        anos = [{"chave": str(a), "rotulo": str(a), "quantidade": 1, "valor": 0} for a in range(1990, 2016)]  # 26 anos, cronológico
        op, sub, altura, col = painel._grafico(anos, {"situacao": "ATIVO"}, "ano")
        assert op["xAxis"]["data"] == [str(a) for a in range(1996, 2016)] and "mais recentes" in sub   # corte = 20 mais recentes
        assert op["series"][0]["data"][-1]["url"].endswith("ano=2015")

        op, sub, altura, col = painel._grafico([item(i) for i in range(15, 0, -1)], {"situacao": "ATIVO"}, "ano")
        assert op["series"][0]["type"] == "bar" and len(op["xAxis"]["data"]) == 15 and sub is None and altura == "alto" and col is None
