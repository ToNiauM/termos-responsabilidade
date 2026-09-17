"""Macro grafico e helper graficos.py."""

from flask import render_template_string

import graficos
from tests.conftest import semear


def test_graficos_helper_urls_e_tabela():
    op = graficos.rosca([("ATIVO", 3), ("BAIXADO", 1)], total=(4, "bens"), urls={"ATIVO": "/analise?situacao=ATIVO"})
    assert op["series"][0]["data"][0] == {"name": "ATIVO", "value": 3, "url": "/analise?situacao=ATIVO"}
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
         "tabela": graficos.tabela_dados(["Rótulo", "Bens"], [[{"valor": "CCI", "url": "/analise?ccusto=CCI"}, 1]])}
    with app.test_request_context():
        html = render_template_string('{% from "_macros.html" import grafico %}{{ grafico(g) }}', g=g)
    assert 'data-grafico="g1"' in html and '<script type="application/json" id="g1">' in html
    assert "<b>" not in html.split('id="g1">')[1].split("</script>")[0]      # tojson escapa
    assert 'href="/analise?ccusto=CCI"' in html and "Ver dados" in html and "dsgov-grafico" in html


def test_url_recorte_mantem_situacao_vazia(dados):
    from app import app
    import painel
    with app.test_request_context():
        assert painel.url_recorte({}, ccusto="CCI") == "/analise?ccusto=CCI&situacao="
        assert painel.url_recorte({"situacao": "ATIVO"}, ccusto="CCI") == "/analise?situacao=ATIVO&ccusto=CCI"
        assert painel.url_recorte_xlsx({}) == "/analise/xlsx?situacao="
        assert painel.url_recorte_xlsx({"situacao": "ATIVO"}) == "/analise/xlsx?situacao=ATIVO"


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
        assert por["g-idade"]["opcoes"]["yAxis"]["type"] == "category" and por["g-idade"]["opcoes"]["yAxis"]["data"][0] == "até 5 anos"   # ordinal → barras na ordem
        assert por["g-ano"]["opcoes"]["series"][0]["type"] == "pie"             # 5 itens
        assert por["g-faixa"]["opcoes"]["series"][0]["type"] == "bar" and por["g-faixa"]["opcoes"]["yAxis"]["type"] == "category"  # 6 itens, rótulo "R$ 1.000 a 5.000" → barras
        assert por["g-faixa"]["opcoes"]["yAxis"]["data"][0] == "até R$ 100"                                                    # faixas mantêm a ordem ordinal
        assert por["g-situacao"]["subtitulo"] == "Todos os bens" and por["g-centro"]["subtitulo"] == "Bens ativos"
        assert por["g-centro"]["opcoes"]["series"][0]["data"][0]["url"] == "/analise?situacao=ATIVO&ccusto=-"
        assert por["g-centro"]["opcoes"]["series"][0]["data"][1]["url"] == "/analise?situacao=ATIVO&ccusto=CCI"
        assert por["g-situacao"]["opcoes"]["series"][0]["data"][0]["url"] == "/analise?situacao=ATIVO"
        assert por["g-classificacao"]["tabela"]["linhas"][-1][0]["valor"] == "SEDE"          # imóveis por último na tabela
        assert "SEDE" in [d["name"] for d in por["g-classificacao"]["opcoes"]["series"][0]["data"]]  # mas no gráfico, como qualquer classe
        assert por["g-ano"]["opcoes"]["series"][0]["data"][-1]["url"] == "/analise?situacao=ATIVO&ano=2024"
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
        assert op["series"][0]["type"] == "bar" and len(op["xAxis"]["data"]) == 15 and sub is None and altura == "alto" and col == "col-12"
        assert op["xAxis"]["axisLabel"]["rotate"] == 45                 # muitas colunas: rótulos inclinados


def test_indicadores_nao_trocam_filtro(dados):
    """Teste de fumaça do brief: todas as contagens são zero, então todo card fica sem link por
    short-circuit de `refinar` (`not contagem`) — não prova por si só o ramo de conflito de filtro
    (ver test_indicadores_bloqueiam_com_filtro_ja_fixado_em_outro_valor, abaixo)."""
    from app import app
    import painel
    r = dict(quantidade=1, valor_total=10, imoveis=0, valor_imoveis=0,
             sem_centro=0, valor_nao_informado=0, valor_zero=0)
    with app.test_request_context():
        cards = painel.indicadores_analise(r, {'ccusto': 'CCI', 'classificacao': 'MÓVEIS', 'valor_de': '1'})
    assert all(c['url'] is None for c in cards)


def test_indicadores_bloqueiam_com_filtro_ja_fixado_em_outro_valor():
    """O primeiro `if` de `refinar` (painel.py) é o ponto central da tarefa ("sem trocar filtros"):
    um filtro já fixado num valor DIFERENTE do que o card ofereceria bloqueia o link mesmo com
    contagem positiva — não só quando a contagem é zero (short-circuit, teste acima) nem quando o
    filtro já está exatamente no valor do card (a outra branch, coberta em
    test_indicadores_analise_clique_completo com classificacao='imoveis' e valor_status='zero'/'nao_informado')."""
    from app import app
    import painel
    r = dict(quantidade=10, valor_total=100, imoveis=3, valor_imoveis=30,
             sem_centro=4, valor_nao_informado=2, valor_zero=2)
    with app.test_request_context():
        por = {c['rotulo']: c for c in painel.indicadores_analise(r, {'classificacao': 'TERRENOS'})}
        assert por['Imóveis']['url'] is None                # classificação já fixada numa classe real diferente
        por = {c['rotulo']: c for c in painel.indicadores_analise(r, {'valor_status': 'zero'})}
        assert por['Valor não informado']['url'] is None    # situação do valor já fixada em 'zero'
        por = {c['rotulo']: c for c in painel.indicadores_analise(r, {'valor_status': 'nao_informado'})}
        assert por['Valor zero']['url'] is None              # situação do valor já fixada em 'não informado'
        por = {c['rotulo']: c for c in painel.indicadores_analise(r, {'ccusto': 'CCI'})}
        assert por['Sem centro nem pessoa']['url'] is None   # centro já fixado num centro real


def test_indicadores_valor_status_bloqueado_por_intervalo_numerico():
    """Valor não informado nunca combina com um intervalo (NULL nunca satisfaz >=/<=); valor zero pode
    coincidir com o intervalo, mas o clique não deve oferecer a combinação (spec: sem contradição)."""
    from app import app
    import painel
    r = dict(quantidade=5, valor_total=100, imoveis=0, valor_imoveis=0,
             sem_centro=0, valor_nao_informado=2, valor_zero=2)
    with app.test_request_context():
        por = {c['rotulo']: c for c in painel.indicadores_analise(r, {'valor_de': '1'})}
        assert por['Valor não informado']['url'] is None and por['Valor zero']['url'] is None
        por = {c['rotulo']: c for c in painel.indicadores_analise(r, {'valor_ate': '100'})}
        assert por['Valor não informado']['url'] is None and por['Valor zero']['url'] is None
        por = {c['rotulo']: c for c in painel.indicadores_analise(r, {})}
        assert por['Valor não informado']['url'] is not None and por['Valor zero']['url'] is not None


def test_indicadores_analise_valores_e_detalhe(dados):
    """Os seis cards, na ordem da spec, com rótulo/valor/detalhe corretos e URLs de imóveis e sem
    centro nem pessoa quando não há conflito de filtro."""
    from app import app
    import painel
    r = dict(quantidade=7, valor_total=1234.5, imoveis=2, valor_imoveis=6000,
             sem_centro=3, valor_nao_informado=1, valor_zero=1)
    with app.test_request_context():
        cards = painel.indicadores_analise(r, {'situacao': 'ATIVO'})
    assert [c['rotulo'] for c in cards] == ['Bens no recorte', 'Valor atual', 'Imóveis',
                                             'Sem centro nem pessoa', 'Valor não informado', 'Valor zero']
    por = {c['rotulo']: c for c in cards}
    assert por['Bens no recorte']['valor'] == 7 and por['Bens no recorte']['url'] is None
    assert por['Valor atual']['valor'] == 'R$ 1.234,50' and por['Valor atual']['url'] is None
    assert por['Imóveis']['valor'] == 2 and por['Imóveis']['detalhe'] == 'R$ 6.000,00'
    assert por['Imóveis']['url'] == '/analise?situacao=ATIVO&classificacao=imoveis'
    assert por['Sem centro nem pessoa']['url'] == '/analise?situacao=ATIVO&ccusto=-&pessoa=-'
    assert por['Valor não informado']['url'] == '/analise?situacao=ATIVO&valor_status=nao_informado'
    assert por['Valor zero']['url'] == '/analise?situacao=ATIVO&valor_status=zero'


def test_indicadores_analise_clique_completo(dados):
    """Clique num indicador positivo: decodifica a URL do card e confirma que db.recorte com esses
    filtros traz a mesma contagem (o laço `clicar` faz isso para todo card positivo, em toda chamada
    abaixo). Cobre situacao=, centro, pessoa, imóvel, intervalo, NULL, zero e filtros já fixados —
    nesses últimos, cada cenário assere explicitamente se o card correspondente ficou sem link (valor
    já fixado, real e diferente do que o card ofereceria) ou se o link continua disponível (só um dos
    dois lados de 'sem centro nem pessoa' está fixado)."""
    from urllib.parse import urlparse, parse_qs
    from app import app
    from tests.conftest import semear
    import db, painel
    semear(dados)
    dados.execute("INSERT INTO bens VALUES (5001,'ATIVO','TERRENO','','TERRENOS','','01/01/2000',1,1000)")
    dados.execute("INSERT INTO bens VALUES (5002,'ATIVO','TV','','EQUIPAMENTOS','01 - SALA CCI','01/01/2020',NULL,NULL)")
    dados.execute("INSERT INTO bens VALUES (5003,'ATIVO','TV2','','EQUIPAMENTOS','01 - SALA CCI','01/01/2020',0,0)")
    dados.commit()

    def clicar(f):
        with app.test_request_context():
            r = db.recorte(dados, f)
            cards = painel.indicadores_analise(r, f)
        for c in cards:
            if c['url'] is None or not isinstance(c['valor'], int):
                continue
            q = parse_qs(urlparse(c['url']).query, keep_blank_values=True)
            f2 = {k: v[0] for k, v in q.items() if k in db.FILTROS}
            assert db.recorte(dados, f2)['quantidade'] == c['valor'], (c['rotulo'], c['url'], f2)
        return r, {c['rotulo']: c for c in cards}

    # situacao= (todas as situações): imóvel e sem centro nem pessoa aparecem
    r, por = clicar({'situacao': ''})
    assert r['imoveis'] == 1 and r['sem_centro'] == 2

    # situacao padrão (ATIVO)
    clicar({'situacao': 'ATIVO'})

    # centro já fixado num centro real (diverge de '-'): "sem centro nem pessoa" fica sem link
    r, por = clicar({'situacao': 'ATIVO', 'ccusto': 'CCI'})
    assert por['Sem centro nem pessoa']['url'] is None

    # pessoa já fixada (diverge de '-'): mesma lógica sobre pessoa
    r, por = clicar({'situacao': 'ATIVO', 'pessoa': 'ANA SILVA'})
    assert por['Sem centro nem pessoa']['url'] is None

    # ccusto já na sentinela '-' e pessoa livre: a spec permite continuar oferecendo o link (falta só
    # fixar pessoa) — o card tem link e o round-trip do laço acima já confere a contagem
    r, por = clicar({'situacao': 'ATIVO', 'ccusto': '-'})
    assert por['Sem centro nem pessoa']['url'] is not None and r['sem_centro'] > 0

    # classificação já EXATAMENTE no valor do card ('imoveis'): sem novo refinamento (branch 2 de
    # refinar — nada sobra para oferecer)
    r, por = clicar({'situacao': 'ATIVO', 'classificacao': 'imoveis'})
    assert por['Imóveis']['url'] is None

    # classificação fixada numa classe real de imóvel (TERRENOS, diverge de 'imoveis'): mesmo com
    # r['imoveis'] positivo (TERRENOS pertence ao grupo), o link fica bloqueado — branch 1 de refinar
    # com dado real, não só sintético (achado do reviewer)
    r, por = clicar({'situacao': 'ATIVO', 'classificacao': 'TERRENOS'})
    assert r['imoveis'] > 0 and por['Imóveis']['url'] is None

    # intervalo numérico ativo: bloqueia os dois cards de situação do valor
    r, por = clicar({'situacao': 'ATIVO', 'valor_de': '0'})
    assert por['Valor não informado']['url'] is None and por['Valor zero']['url'] is None

    # NULL já isolado
    r, por = clicar({'situacao': 'ATIVO', 'valor_status': 'nao_informado'})
    assert por['Valor não informado']['url'] is None

    # zero já isolado
    r, por = clicar({'situacao': 'ATIVO', 'valor_status': 'zero'})
    assert por['Valor zero']['url'] is None

    # sem centro nem pessoa já fixado
    r, por = clicar({'situacao': 'ATIVO', 'ccusto': '-', 'pessoa': '-'})
    assert por['Sem centro nem pessoa']['url'] is None


def test_descrever_valor_status(dados):
    import painel
    assert painel.descrever({'situacao': 'ATIVO', 'valor_status': 'nao_informado'}) == 'Bens ATIVO · valor não informado'
    assert painel.descrever({'situacao': 'ATIVO', 'valor_status': 'zero'}) == 'Bens ATIVO · valor zero'
