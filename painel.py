"""Cards de gráfico do painel e do recorte a partir de db.dimensoes(). Só apresentação: escolhe o tipo
de gráfico por dimensão (spec, parte 4) e monta as URLs de drill-down para /recorte."""
from flask import url_for

import db
import graficos

ROTULO_FILTRO = {"situacao": "situação", "ccusto": "centro de custo", "pessoa": "pessoa", "localizacao": "localização",
                 "classificacao": "classificação", "idade": "idade", "faixa": "faixa de valor", "ano": "ano de entrada"}
ROTULOS_ESPECIAIS = {"ccusto": {"-": "sem centro"}, "pessoa": {"-": "sem pessoa"},
                     "classificacao": {"-": "sem classificação", "imoveis": "imóveis", "sem-imoveis": "sem imóveis"},
                     "localizacao": {"-": "sem localização"}, "ano": {"-": "sem data"}}
TOP = 20


def moeda(v) -> str:
    return f"R$ {v or 0:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")


def url_recorte(f: dict, **extra) -> str:
    """URL de /recorte com os filtros atuais mais `extra` (drill-down acrescenta). `situacao` sempre viaja,
    mesmo vazia: ausente significa ATIVO, vazia significa todas."""
    args = {k: v for k, v in {**f, **extra}.items() if v}
    args.setdefault("situacao", f.get("situacao", ""))
    return url_for("recorte", **args)


def url_recorte_xlsx(f: dict) -> str:
    args = {k: v for k, v in f.items() if v}
    args.setdefault("situacao", f.get("situacao", ""))
    return url_for("recorte_xlsx", **args)


def _tabela(itens, f, chave, rotulo_col):
    return graficos.tabela_dados([rotulo_col, "Bens", "Valor"], [
        [{"valor": i["rotulo"], "url": url_recorte(f, **{chave: i["chave"]})}, i["quantidade"], moeda(i["valor"])] for i in itens])


def _card(id, titulo, opcoes, itens, f, chave, rotulo_col, alto=False, subtitulo=None):
    resumo = f"{titulo}: " + ", ".join(f"{i['rotulo']} {i['quantidade']}" for i in itens[:6])
    return {"id": id, "titulo": titulo, "subtitulo": subtitulo, "opcoes": opcoes, "resumo": resumo,
            "alto": alto, "col": None, "tabela": _tabela(itens, f, chave, rotulo_col)}


def _urls(itens, f, chave):
    return [url_recorte(f, **{chave: i["chave"]}) for i in itens]


def _grafico(itens, f, chave):
    """Tipo pelo número de itens (regra do usuário): ≤5 rosca, 6–10 barras, 11–20 colunas, >20 colunas dos 20 maiores."""
    top = itens[:TOP]
    if len(itens) <= 5:
        op = graficos.rosca([(i["rotulo"], i["quantidade"]) for i in itens], total=(sum(i["quantidade"] for i in itens), "bens"),
                            urls={i["rotulo"]: u for i, u in zip(itens, _urls(itens, f, chave))})
    elif len(itens) <= 10:
        op = graficos.barras_horizontais([i["rotulo"] for i in itens], [i["quantidade"] for i in itens], "Bens", escala=True, urls=_urls(itens, f, chave))
    else:
        op = graficos.colunas([i["rotulo"] for i in top], {"Bens": [i["quantidade"] for i in top]}, rotulos=True, urls={"Bens": _urls(top, f, chave)})
    sub = f"{TOP} maiores no gráfico; todos na tabela" if len(itens) > TOP else None
    return op, sub, len(itens) > 10


def cards_graficos(dim: dict, f: dict, omitir=()) -> list[dict]:
    """Um card por dimensão, na ordem da spec. omitir = filtros já fixados num único valor."""
    cards = []
    if "situacao" not in omitir and dim["situacao"]:
        it = dim["situacao"]
        op, sub, alto = _grafico(it, f, "situacao")
        cards.append(_card("g-situacao", "Bens por situação", op, it, f, "situacao", "Situação", alto=alto, subtitulo=sub))
    if "ccusto" not in omitir:
        it = dim["centro"]
        op, sub, alto = _grafico(it, f, "ccusto")
        cards.append(_card("g-centro", "Bens por centro de custo", op, it, f, "ccusto", "Centro de custo", alto=alto, subtitulo=sub))
    if "classificacao" not in omitir:
        it = dim["classificacao"]
        comuns = [i for i in it if i["chave"] not in db.IMOVEIS]
        imoveis = [i for i in it if i["chave"] in db.IMOVEIS]
        op, sub, alto = _grafico(it, f, "classificacao")   # imóveis entram no gráfico como qualquer classe (é contagem)
        cards.append(_card("g-classificacao", "Bens por classificação contábil", op, comuns + imoveis, f, "classificacao",
                           "Classificação", alto=alto, subtitulo=sub))   # imóveis só ficam por último na tabela
    if "localizacao" not in omitir:
        it = dim["localizacao"]
        op, sub, alto = _grafico(it, f, "localizacao")
        cards.append(_card("g-localizacao", "Bens por localização", op, it, f, "localizacao", "Localização", alto=alto, subtitulo=sub))
    if "idade" not in omitir:
        it = dim["idade"]
        op, sub, alto = _grafico(it, f, "idade")
        cards.append(_card("g-idade", "Bens por idade (data de entrada)", op, it, f, "idade", "Faixa", alto=alto, subtitulo=sub))
    if "ano" not in omitir and dim["ano"]:
        it = dim["ano"]
        op, sub, alto = _grafico(it, f, "ano")
        cards.append(_card("g-ano", "Bens por ano de entrada", op, it, f, "ano", "Ano", alto=alto, subtitulo=sub))
    if "faixa" not in omitir:
        it = dim["faixa"]
        op, sub, alto = _grafico(it, f, "faixa")
        cards.append(_card("g-faixa", "Bens por faixa de valor", op, it, f, "faixa", "Faixa", alto=alto, subtitulo=sub))
    if "pessoa" not in omitir and dim["pessoa"]:
        it = dim["pessoa"]
        op, sub, alto = _grafico(it, f, "pessoa")
        cards.append(_card("g-pessoa", "Bens atribuídos por pessoa", op, it, f, "pessoa", "Pessoa", alto=alto, subtitulo=sub))
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
            v = nomes.get(k) or ROTULOS_ESPECIAIS.get(k, {}).get(f[k], f[k])
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
