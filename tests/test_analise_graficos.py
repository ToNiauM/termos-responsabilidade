"""Gráficos da Análise no Tabler: agregações em SQL (db.cruzamento, db.concentracao, db.cobertura_termos),
cards de painel.cards_analise (rosca, Pareto, explorar com seletores, entradas por ano, comprado × hoje,
cobertura de termos) e a página com as variantes dos seletores."""
import json

import pytest

import db
import graficos
import painel
from tests.conftest import semear
from tests.test_db import _semear_painel

ATIVO = {"situacao": "ATIVO"}


def _cards(conn, f, omitir=None):
    """Como a rota: omitir = filtros fixados num valor (imóveis/sem imóveis não fixam a classificação)."""
    from app import app
    if omitir is None:
        omitir = tuple(k for k in ("situacao", "ccusto", "classificacao", "localizacao", "idade", "ano", "faixa", "pessoa")
                       if f.get(k) and f.get(k) not in ("imoveis", "sem-imoveis"))
    r = db.recorte(conn, f)
    with app.test_request_context():
        cards = painel.cards_analise(r["dimensoes"], painel.dados_analise(conn, f, omitir), f, omitir, r["valor_total"])
    return {c["id"]: c for c in cards}


# ---------------------------------------------------------------- agregações

def test_cruzamento_centro_por_idade_e_ano_por_classificacao(dados):
    _semear_painel(dados)
    linhas = {(x["linha"], x["coluna"]): (x["quantidade"], x["valor"], x["compra"]) for x in db.cruzamento(dados, ATIVO, "ccusto", "idade")}
    assert linhas == {("CCI", "ate5"): (1, 3500, 4000), ("CCI", "10a20"): (1, 1500, 3000), ("CCI", "mais20"): (1, 64.54, 75.94),
                      ("-", "ate5"): (1, 800, 900), ("-", "10a20"): (1, 250.5, 500), ("-", "mais20"): (1, 60000000, 1)}
    anos = {(x["linha"], x["coluna"]): x["quantidade"] for x in db.cruzamento(dados, ATIVO, "ano", "classificacao")}
    assert anos == {("1990", "SEDE"): 1, ("1996", "MÓVEIS"): 1, ("2012", "EQUIPAMENTOS"): 1, ("2012", "MÓVEIS"): 1,
                    ("2023", "EQUIPAMENTOS"): 1, ("2024", "EQUIPAMENTOS"): 1}
    # o filtro vale: só o centro CCI; e sem situação entra o baixado (1003, MÓVEIS 2012)
    assert {x["linha"] for x in db.cruzamento(dados, {**ATIVO, "ccusto": "CCI"}, "ccusto", "idade")} == {"CCI"}
    assert sum(x["quantidade"] for x in db.cruzamento(dados, {}, "ano", "classificacao")) == 7


def test_concentracao_pareto_exato(dados):
    _semear_painel(dados)
    c = db.concentracao(dados, ATIVO)   # 60.000.000 · 3.500 · 1.500 · 800 · 250,50 · 64,54
    total = 60000000 + 3500 + 1500 + 800 + 250.5 + 64.54
    assert c["quantidade"] == 6 and c["total"] == pytest.approx(total)
    assert c["k"] == 1 and c["corte"] == 60000000
    assert [rn for rn, _ in c["pontos"]] == [1, 2, 3, 4, 5, 6]
    assert c["pontos"][1][1] == pytest.approx(60003500) and c["pontos"][-1][1] == pytest.approx(total)
    c = db.concentracao(dados, {**ATIVO, "ccusto": "CCI"})   # 3.500 + 1.500 = 5.000 ≥ 80% de 5.064,54
    assert (c["quantidade"], c["k"], c["corte"]) == (3, 2, 1500)
    assert db.concentracao(dados, {**ATIVO, "valor_status": "zero"}) == {"quantidade": 0, "total": 0, "k": None, "corte": None, "pontos": []}


def _semear_cobertura(conn):
    """CCI vigente (1001), ANA SILVA com termo individual vigente (1002), DES desatualizado (5001 no termo,
    5002 entrou depois), NOV sem termo (6001), 1004 sem centro nem pessoa, 1003 baixado (fora)."""
    semear(conn)
    conn.execute("INSERT INTO responsaveis (ccustos, responsavel, email, matricula, funcao) VALUES ('DES','DENISE','d@x','1','f'), ('NOV','NOVO','n@x','2','f')")
    conn.execute("INSERT INTO localizacoes VALUES ('SALA DES','DES'), ('SALA NOV','NOV')")
    conn.execute("INSERT INTO bens VALUES (5001,'ATIVO','MESA','','MÓVEIS','SALA DES','01/01/2020',100,100)")
    conn.execute("INSERT INTO bens VALUES (6001,'ATIVO','SOFÁ','','MÓVEIS','SALA NOV','01/01/2020',300,300)")
    conn.commit()
    db.incluir_processo(conn, "ccusto", "Termos", "1")
    db.incluir_processo(conn, "individual", "Individuais", "2")
    db.registrar_emissao(conn, "ccusto", "CCI", db.bens_do_centro(conn, "CCI"))
    db.registrar_emissao(conn, "ccusto", "DES", db.bens_do_centro(conn, "DES"))
    db.registrar_emissao(conn, "individual", "ANA SILVA", db.bens_da_pessoa(conn, "ANA SILVA"))
    conn.execute("INSERT INTO bens VALUES (5002,'ATIVO','CADEIRA','','MÓVEIS','SALA DES','01/01/2021',200,200)")
    conn.commit()


def test_cobertura_termos_por_guardiao_e_situacao(dados):
    _semear_cobertura(dados)
    itens = {(i["tipo"], i["chave"]): (i["estado"], i["quantidade"], i["valor"]) for i in db.cobertura_termos(dados, ATIVO)}
    assert itens == {("ccusto", "CCI"): ("vigente", 1, 64.54), ("individual", "ANA SILVA"): ("vigente", 1, 1500),
                     ("ccusto", "DES"): ("desatualizado", 2, 300), ("ccusto", "NOV"): ("sem_termo", 1, 300),
                     ("-", "-"): ("sem_responsavel", 1, 250.5)}
    # a situação do termo é a de todos os bens do guardião, mesmo com o recorte cortando parte deles
    assert db.cobertura_termos(dados, {**ATIVO, "valor_ate": "150"}) == [
        {"tipo": "ccusto", "chave": "CCI", "quantidade": 1, "valor": 64.54, "estado": "vigente"},
        {"tipo": "ccusto", "chave": "DES", "quantidade": 1, "valor": 100, "estado": "desatualizado"}]
    assert [i["chave"] for i in db.cobertura_termos(dados, {"situacao": "", "ccusto": "DES"})] == ["DES"]   # todas → só ativos
    assert db.cobertura_termos(dados, {"situacao": "BAIXADO"}) == []
    # igual às telas de termos
    assert {c["ccustos"]: c["estado"] for c in db.situacoes_centros(dados)} == {"CCI": "vigente", "DES": "desatualizado", "NOV": "sem_termo"}


def test_cabecalhos_pareto_e_cobertura(dados):
    assert painel.cabecalho_pareto(3, 3520) == {"forte": "3 bens", "texto": "(0,09% do recorte) concentram 80% do valor atual"}
    assert painel.cabecalho_pareto(1, 6) == {"forte": "1 bem", "texto": "(17% do recorte) concentra 80% do valor atual"}
    _semear_cobertura(dados)
    # 2 de 6 bens em dia (CCI e ANA); valor 1.564,54 de 2.415,04
    assert painel.cabecalho_cobertura(db.cobertura_termos(dados, ATIVO)) == {"forte": "33% dos bens", "texto": "(65% do valor) com termo em dia"}
    assert painel.cabecalho_cobertura([]) is None
    assert [painel.percentual(p) for p in (58.1, 5.26, 0.0852, 100)] == ["58%", "5,3%", "0,09%", "100%"]
    assert [painel.moeda_curta(v) for v in (103258592.52, 12400, 999)] == ["103,3 mi", "12,4 mil", "999,00"]


# ---------------------------------------------------------------- cards

def test_cards_analise_ordem_rosca_e_cores(dados):
    _semear_painel(dados)
    cards = _cards(dados, ATIVO)
    assert list(cards) == ["g-valor", "g-concentracao", "g-explorar", "g-entradas", "g-compra", "g-cobertura"]
    rosca = cards["g-valor"]["opcoes"]
    fatias = rosca["series"][0]["data"]
    assert [d["name"] for d in fatias] == ["SEDE", "EQUIPAMENTOS", "MÓVEIS"]
    assert fatias[0]["url"] == "/analise?situacao=ATIVO&classificacao=SEDE"
    assert rosca["graphic"][0]["style"]["text"] == "60,0 mi\nem R$"   # formato "curta" que o graficos.js reconhece
    assert sum(d["value"] for d in fatias) == pytest.approx(60006115.04)
    cor = {d["name"]: d["itemStyle"]["color"] for d in fatias}
    assert cor["SEDE"] == graficos.CORES_CATEGORICA[0] and len(set(cor.values())) == 3
    # a mesma classificação tem a mesma cor nas entradas por ano (nas duas medidas)
    for v in cards["g-entradas"]["variantes"]:
        for s in v["opcoes"]["series"]:
            if s["name"] in cor:
                assert s["itemStyle"]["color"] == cor[s["name"]]


def test_cards_analise_explorar_variantes_e_drill_down(dados):
    _semear_painel(dados)
    g = _cards(dados, ATIVO)["g-explorar"]
    assert [s["rotulo"] for s in g["seletores"]] == ["Dimensão", "Medida"]
    assert [v for v, _ in g["seletores"][0]["opcoes"]] == ["ccusto", "localizacao", "classificacao", "pessoa"]   # situação fixada: fora
    ids = [v["id"] for v in g["variantes"]]
    assert ids == [f"g-explorar--{d}--{m}" for d in ("ccusto", "localizacao", "classificacao", "pessoa") for m in ("quantidade", "valor")]
    v = {x["id"]: x for x in g["variantes"]}
    op = v["g-explorar--ccusto--valor"]["opcoes"]
    assert op["yAxis"]["data"] == ["sem centro", "CCI – JAQUELINE PORTELA"] and op["yAxis"]["type"] == "category"
    assert all(s["stack"] == "total" for s in op["series"]) and op["tooltip"]["valueFormatter"] == "moeda"
    serie = {s["name"]: s for s in op["series"]}
    assert list(serie) == ["até 5 anos", "10 a 20 anos", "mais de 20 anos"]      # só as faixas com bens, na ordem
    assert serie["até 5 anos"]["itemStyle"]["color"] == graficos.CORES_SEQUENCIAL[1]
    assert serie["até 5 anos"]["data"][1] == {"value": 3500, "url": "/analise?situacao=ATIVO&ccusto=CCI&idade=ate5"}
    assert serie["mais de 20 anos"]["data"][0] == {"value": 60000000, "url": "/analise?situacao=ATIVO&ccusto=-&idade=mais20"}
    tabela = v["g-explorar--ccusto--quantidade"]["tabela"]
    assert tabela["colunas"] == ["Centro de custo", "até 5 anos", "10 a 20 anos", "mais de 20 anos", "Total"]
    assert tabela["linhas"][0][0] == {"valor": "CCI – JAQUELINE PORTELA", "url": "/analise?situacao=ATIVO&ccusto=CCI"}
    assert [c["valor"] for c in tabela["linhas"][0][1:]] == ["1", "1", "1", "3"]
    pessoa = v["g-explorar--pessoa--quantidade"]["opcoes"]
    assert pessoa["yAxis"]["data"] == ["ANA SILVA"]                                 # bens sem pessoa não viram barra


def test_explorar_dobra_em_outros_e_respeita_filtro(dados):
    semear(dados)
    for i in range(15):
        dados.execute("INSERT INTO bens VALUES (?,?,?,?,?,?,?,?,?)", (7000 + i, "ATIVO", "X", "", f"CLASSE {i:02d}", "01 - SALA CCI", "01/01/2020", 1, i + 1))
    dados.commit()
    g = _cards(dados, ATIVO)["g-explorar"]
    op = {v["id"]: v for v in g["variantes"]}["g-explorar--classificacao--valor"]["opcoes"]
    assert len(op["yAxis"]["data"]) == 13 and op["yAxis"]["data"][:2] == ["EQUIPAMENTOS", "MÓVEIS"] and op["yAxis"]["data"][-1] == "Outros (5)"
    assert all(s["data"][-1]["url"] is None if isinstance(s["data"][-1], dict) else True for s in op["series"])
    g = _cards(dados, {**ATIVO, "ccusto": "CCI"}, omitir=("situacao", "ccusto"))["g-explorar"]
    assert "ccusto" not in [v for v, _ in g["seletores"][0]["opcoes"]]
    assert all("ccusto=CCI" in s["data"][0]["url"] for v in g["variantes"] for s in v["opcoes"]["series"] if isinstance(s["data"][0], dict) and s["data"][0].get("url"))


def test_entradas_por_ano_dobra_anos_antigos(dados):
    semear(dados)
    dados.execute("DELETE FROM atribuicoes"); dados.execute("DELETE FROM bens")
    for i, ano in enumerate(range(2000, 2018)):   # 18 anos → 15 recentes (2003–2017) + "até 2002"
        dados.execute("INSERT INTO bens VALUES (?,?,?,?,?,?,?,?,?)", (8000 + i, "ATIVO", "X", "", "MÓVEIS", "01 - SALA CCI", f"01/01/{ano}", 10 * (i + 1), 1))
    dados.commit()
    g = _cards(dados, ATIVO)["g-entradas"]
    assert [s["rotulo"] for s in g["seletores"]] == ["Medida"] and [v for v, _ in g["seletores"][0]["opcoes"]] == ["quantidade", "compra"]
    v = {x["id"]: x for x in g["variantes"]}
    op = v["g-entradas--quantidade"]["opcoes"]
    assert op["xAxis"]["data"] == ["até 2002"] + [str(a) for a in range(2003, 2018)]
    s = op["series"][0]
    assert s["name"] == "MÓVEIS" and s["stack"] == "total"
    assert s["data"][0] == {"value": 3, "url": "/analise?situacao=ATIVO&entrada_ate=2002-12-31&classificacao=M%C3%93VEIS"}
    assert s["data"][-1] == {"value": 1, "url": "/analise?situacao=ATIVO&ano=2017&classificacao=M%C3%93VEIS"}
    compra = v["g-entradas--compra"]["opcoes"]
    assert compra["series"][0]["data"][0]["value"] == 60 and compra["tooltip"]["valueFormatter"] == "moeda"
    assert v["g-entradas--compra"]["tabela"]["linhas"][0][-1]["valor"] == "R$ 60,00"


def test_pareto_e_comprado_hoje(dados):
    _semear_painel(dados)
    cards = _cards(dados, ATIVO)
    p = cards["g-concentracao"]
    assert p["destaque"] == {"forte": "1 bem", "texto": "(17% do recorte) concentra 80% do valor atual"}
    linha = p["opcoes"]["series"][0]
    assert linha["data"][0] == [0, 0] and linha["data"][-1] == [100.0, 100.0]
    marco = linha["markPoint"]["data"][0]
    assert marco["coord"] == [16.67, 99.99] and marco["url"] == "/analise?situacao=ATIVO&valor_de=60000000.00"
    assert p["opcoes"]["yAxis"]["max"] == 100 and not isinstance(p["opcoes"]["yAxis"], list)   # um eixo de valor só
    c = cards["g-compra"]["opcoes"]
    assert list(s["name"] for s in c["series"]) == ["Valor de compra", "Valor atual"] and c["yAxis"]["data"][0] == "EQUIPAMENTOS"
    assert c["series"][0]["data"][0] == {"value": 7900, "url": "/analise?situacao=ATIVO&classificacao=EQUIPAMENTOS"}
    assert c["series"][1]["data"][0]["value"] == 5800 and not isinstance(c["xAxis"], list)
    # com a classificação fixada, rosca e comprado × hoje passam ao centro de custo
    cards = _cards(dados, {**ATIVO, "classificacao": "EQUIPAMENTOS"}, omitir=("situacao", "classificacao"))
    assert "centro de custo" in cards["g-valor"]["subtitulo"] and cards["g-compra"]["tabela"]["colunas"][0] == "Centro de custo"


def test_card_cobertura(dados):
    _semear_cobertura(dados)
    from app import app
    g = _cards(dados, ATIVO)["g-cobertura"]
    assert g["destaque"] == {"forte": "33% dos bens", "texto": "(65% do valor) com termo em dia"}
    v = {x["id"]: x for x in g["variantes"]}
    op = v["g-cobertura--quantidade"]["opcoes"]
    assert op["yAxis"]["data"] == ["DES – DENISE", "CCI – JAQUELINE PORTELA", "NOV – NOVO", "Termos individuais", "Sem responsável"]
    serie = {s["name"]: s for s in op["series"]}
    assert list(serie) == ["Termo em dia", "Termo desatualizado", "Sem termo", "Sem responsável"]
    assert serie["Termo em dia"]["itemStyle"]["color"] == graficos.CORES_STATUS["sucesso"]
    assert serie["Termo desatualizado"]["itemStyle"]["color"] == graficos.CORES_STATUS["alerta"]
    assert serie["Sem termo"]["itemStyle"]["color"] == graficos.CORES_STATUS["erro"]
    assert serie["Sem responsável"]["itemStyle"]["color"] == graficos.CORES_STATUS["neutro"]
    with app.test_request_context():
        from flask import url_for
        individuais = url_for("termos_individuais")
    assert serie["Termo desatualizado"]["data"][0] == {"value": 2, "url": "/analise?situacao=ATIVO&ccusto=DES"}
    assert serie["Termo em dia"]["data"][3] == {"value": 1, "url": individuais}
    assert serie["Sem responsável"]["data"][4] == {"value": 1, "url": "/analise?situacao=ATIVO&ccusto=-&pessoa=-"}
    assert serie["Termo em dia"]["data"][0] == {"value": None, "url": "/analise?situacao=ATIVO&ccusto=DES"}
    tabela = v["g-cobertura--valor"]["tabela"]
    assert tabela["linhas"][0][1:] == [{"valor": "R$ 0,00"}, {"valor": "R$ 300,00"}, {"valor": "R$ 0,00"}, {"valor": "R$ 0,00"}, {"valor": "R$ 300,00"}]
    assert _cards(dados, {"situacao": "BAIXADO"}).get("g-cobertura") is None


def test_recorte_vazio_sem_graficos(dados):
    semear(dados)
    assert _cards(dados, {**ATIVO, "classificacao": "NADA"}) == {}


# ---------------------------------------------------------------- página no Tabler

@pytest.fixture
def tabler(cliente):
    import app as modulo
    original = modulo.app.jinja_loader
    modulo.app.jinja_loader = modulo.carregador_templates(tabler=True)
    modulo.app.jinja_env.cache.clear()
    yield cliente
    modulo.app.jinja_loader = original
    modulo.app.jinja_env.cache.clear()


def test_pagina_tabler_tem_variantes_seletores_e_tabelas(tabler):
    html = tabler.get("/analise").text
    for id_ in ("g-valor", "g-concentracao", "g-compra", "g-cobertura--quantidade", "g-cobertura--valor", "g-entradas--quantidade", "g-entradas--compra"):
        assert f'<script type="application/json" id="{id_}"' in html, id_
    for d in ("ccusto", "localizacao", "classificacao", "pessoa"):
        for m in ("quantidade", "valor"):
            assert f'id="g-explorar--{d}--{m}"' in html and f'data-grafico-tabela="g-explorar--{d}--{m}"' in html
    assert 'data-grafico="g-explorar--ccusto--quantidade" data-grafico-base="g-explorar"' in html
    assert html.count('data-grafico-seletor="g-explorar"') == 2 and 'class="form-select form-select-sm w-auto"' in html
    assert 'data-grafico-tabela="g-explorar--ccusto--quantidade">' in html          # a primeira visível
    assert 'data-grafico-tabela="g-explorar--ccusto--valor" hidden>' in html        # as outras escondidas
    assert "g-situacao" not in html and "g-faixa" not in html                        # cards repetitivos saíram
    assert "concentram 80% do valor atual" in html or "concentra 80% do valor atual" in html
    assert "com termo em dia" in html and "style=" not in html.split('id="g-valor"')[1].split("Bens do recorte")[0]
    json.loads(html.split('id="g-explorar--ccusto--valor" data-resumo="')[1].split('">', 1)[1].split("</script>")[0])
    # filtros do recorte seguem valendo: centro fixado sai do seletor
    html = tabler.get("/analise?situacao=ATIVO&ccusto=CCI").text
    assert 'id="g-explorar--ccusto--quantidade"' not in html and 'id="g-explorar--localizacao--quantidade"' in html
