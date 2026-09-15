"""Cards de gráfico do painel e do recorte a partir de db.dimensoes(). Só apresentação: escolhe o tipo
de gráfico por dimensão (spec, parte 4) e monta as URLs de drill-down para /recorte."""
from flask import url_for

import db
import graficos

ROTULO_FILTRO = {"situacao": "situação", "ccusto": "centro de custo", "pessoa": "pessoa", "localizacao": "localização",
                 "classificacao": "classificação", "idade": "idade", "faixa": "faixa de valor", "ano": "ano de entrada"}
TOP = 20


def moeda(v) -> str:
    return f"R$ {v or 0:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")


def url_recorte(f: dict, **extra) -> str:
    """URL de /recorte com os filtros atuais mais os de `extra` (drill-down acrescenta, não substitui)."""
    return url_for("recorte", **{k: v for k, v in {**f, **extra}.items() if v})


def _tabela(itens, f, chave, rotulo_col):
    return graficos.tabela_dados([rotulo_col, "Bens", "Valor"], [
        [{"valor": i["rotulo"], "url": url_recorte(f, **{chave: i["chave"]})}, i["quantidade"], moeda(i["valor"])] for i in itens])


def _card(id, titulo, opcoes, itens, f, chave, rotulo_col, alto=False, subtitulo=None):
    resumo = f"{titulo}: " + ", ".join(f"{i['rotulo']} {i['quantidade']}" for i in itens[:6])
    return {"id": id, "titulo": titulo, "subtitulo": subtitulo, "opcoes": opcoes, "resumo": resumo,
            "alto": alto, "col": None, "tabela": _tabela(itens, f, chave, rotulo_col)}


def _urls(itens, f, chave):
    return [url_recorte(f, **{chave: i["chave"]}) for i in itens]


def _barras(id, titulo, itens, f, chave, rotulo_col):
    top = itens[:TOP]
    sub = f"{TOP} maiores no gráfico; todos na tabela" if len(itens) > TOP else None
    op = graficos.barras_horizontais([i["rotulo"] for i in top], [i["quantidade"] for i in top], "Bens",
                                     escala=True, urls=_urls(top, f, chave))
    return _card(id, titulo, op, itens, f, chave, rotulo_col, alto=len(top) > 8, subtitulo=sub)


def _colunas(id, titulo, itens, f, chave, rotulo_col):
    op = graficos.colunas([i["rotulo"] for i in itens], {"Bens": [i["quantidade"] for i in itens]},
                          rotulos=True, urls={"Bens": _urls(itens, f, chave)})
    return _card(id, titulo, op, itens, f, chave, rotulo_col)


def cards_graficos(dim: dict, f: dict, omitir=()) -> list[dict]:
    """Um card por dimensão, na ordem da spec. omitir = filtros já fixados num único valor."""
    cards = []
    if "situacao" not in omitir and dim["situacao"]:
        it = dim["situacao"]
        op = graficos.rosca([(i["rotulo"], i["quantidade"]) for i in it], total=(sum(i["quantidade"] for i in it), "bens"),
                            urls={i["rotulo"]: u for i, u in zip(it, _urls(it, f, "situacao"))})
        cards.append(_card("g-situacao", "Bens por situação", op, it, f, "situacao", "Situação"))
    if "ccusto" not in omitir:
        cards.append(_barras("g-centro", "Bens por centro de custo", dim["centro"], f, "ccusto", "Centro de custo"))
    if "classificacao" not in omitir:
        it = dim["classificacao"]
        comuns = [i for i in it if i["chave"] not in db.IMOVEIS]
        imoveis = [i for i in it if i["chave"] in db.IMOVEIS]
        fatias = [(i["rotulo"], i["quantidade"]) for i in comuns[:5]]
        if len(comuns) > 5:
            fatias.append(("Outras", sum(i["quantidade"] for i in comuns[5:])))
        op = graficos.rosca(fatias, total=(sum(i["quantidade"] for i in comuns), "bens"),
                            urls={i["rotulo"]: u for i, u in zip(comuns[:5], _urls(comuns[:5], f, "classificacao"))})
        cards.append(_card("g-classificacao", "Bens por classificação contábil", op, comuns + imoveis, f, "classificacao",
                           "Classificação", subtitulo="Imóveis (SEDE, TERRENOS) só na tabela" if imoveis else None))
    if "localizacao" not in omitir:
        cards.append(_barras("g-localizacao", "Bens por localização", dim["localizacao"], f, "localizacao", "Localização"))
    if "idade" not in omitir:
        cards.append(_colunas("g-idade", "Bens por idade (data de entrada)", dim["idade"], f, "idade", "Faixa"))
    if "ano" not in omitir and dim["ano"]:
        it = dim["ano"]
        op = graficos.linha([i["rotulo"] for i in it], {"Bens": [i["quantidade"] for i in it]}, urls={"Bens": _urls(it, f, "ano")})
        cards.append(_card("g-ano", "Bens por ano de entrada", op, it, f, "ano", "Ano"))
    if "faixa" not in omitir:
        cards.append(_colunas("g-faixa", "Bens por faixa de valor", dim["faixa"], f, "faixa", "Faixa"))
    if "pessoa" not in omitir and dim["pessoa"]:
        cards.append(_barras("g-pessoa", "Bens atribuídos por pessoa", dim["pessoa"], f, "pessoa", "Pessoa"))
    return cards


def _data_br(iso: str) -> str:
    return f"{iso[8:10]}/{iso[5:7]}/{iso[:4]}"


def descrever(f: dict, nomes: dict | None = None) -> str:
    """Frase do recorte: "Bens ATIVO · centro de custo CCI · entrada a partir de 01/01/2020". nomes = rótulos
    legíveis por filtro (ex.: {"ccusto": "CCI – JAQUELINE", "idade": "mais de 20 anos"})."""
    nomes = nomes or {}
    partes = ["Bens " + f["situacao"] if f.get("situacao") else "Bens (todas as situações)"]
    for k in ("ccusto", "pessoa", "localizacao", "classificacao", "idade", "faixa", "ano"):
        if f.get(k):
            v = nomes.get(k) or {"-": "sem centro", "imoveis": "imóveis", "sem-imoveis": "sem imóveis"}.get(f[k], f[k])
            partes.append(f"{ROTULO_FILTRO[k]} {v}")
    if f.get("valor_de") and f.get("valor_ate"):
        partes.append(f"valor de {moeda(float(f['valor_de']))} a {moeda(float(f['valor_ate']))}")
    elif f.get("valor_de"):
        partes.append(f"valor a partir de {moeda(float(f['valor_de']))}")
    elif f.get("valor_ate"):
        partes.append(f"valor até {moeda(float(f['valor_ate']))}")
    if f.get("entrada_de") and f.get("entrada_ate"):
        partes.append(f"entrada entre {_data_br(f['entrada_de'])} e {_data_br(f['entrada_ate'])}")
    elif f.get("entrada_de"):
        partes.append(f"entrada a partir de {_data_br(f['entrada_de'])}")
    elif f.get("entrada_ate"):
        partes.append(f"entrada até {_data_br(f['entrada_ate'])}")
    return " · ".join(partes)
