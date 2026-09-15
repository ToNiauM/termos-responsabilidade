"""Macro grafico e helper graficos.py."""
import json

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
