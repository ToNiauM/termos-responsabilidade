"""Cards de gráfico do painel e da análise a partir de db.dimensoes(). Só apresentação: escolhe o tipo
de gráfico por dimensão (spec, parte 4) e monta as URLs de drill-down para /analise."""
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
    """URL de /analise com os filtros atuais mais `extra` (drill-down acrescenta). `situacao` sempre viaja,
    mesmo vazia: ausente significa ATIVO, vazia significa todas."""
    args = {k: v for k, v in {**f, **extra}.items() if v}
    args.setdefault("situacao", f.get("situacao", ""))
    return url_for("analise", **args)


def url_recorte_xlsx(f: dict) -> str:
    args = {k: v for k, v in f.items() if v}
    args.setdefault("situacao", f.get("situacao", ""))
    return url_for("analise_xlsx", **args)


def _tabela(itens, f, chave, rotulo_col):
    return graficos.tabela_dados([rotulo_col, "Bens", "Valor"], [
        [{"valor": i["rotulo"], "url": url_recorte(f, **{chave: i["chave"]})}, i["quantidade"], moeda(i["valor"])] for i in itens])


def _card(id, titulo, opcoes, itens, f, chave, rotulo_col, alto=False, subtitulo=None, col=None):
    resumo = f"{titulo}: " + ", ".join(f"{i['rotulo']} {i['quantidade']}" for i in itens[:6])
    return {"id": id, "titulo": titulo, "subtitulo": subtitulo, "opcoes": opcoes, "resumo": resumo,
            "alto": alto == "alto", "altura": alto, "col": col, "tabela": _tabela(itens, f, chave, rotulo_col)}


def _urls(itens, f, chave):
    return [url_recorte(f, **{chave: i["chave"]}) for i in itens]


CURTOS = ("ano",)              # rótulos realmente curtos (anos): colunas servem; o resto → barras horizontais
ORDINAIS = ("idade", "faixa")  # escalas ordenadas: sempre barras horizontais na ordem das faixas (rosca esconde a ordem)


def _grafico(itens, f, chave):
    """Tipo pela natureza e quantidade dos itens: ≤5 rosca; rótulos curtos → colunas; rótulos longos → barras
    horizontais (nome à esquerda, maior para menor). Acima de 20, os 20 maiores no gráfico e todos na tabela;
    'ano' vem em ordem cronológica, então o corte é "20 mais recentes". Devolve (opções, subtítulo, altura, col)."""
    recentes = chave == "ano"
    n = len(itens)
    top = itens[-TOP:] if recentes else itens[:TOP]
    if n <= 5 and chave not in ORDINAIS:
        op = graficos.rosca([(i["rotulo"], i["quantidade"]) for i in itens], total=(sum(i["quantidade"] for i in itens), "bens"),
                            urls={i["rotulo"]: u for i, u in zip(itens, _urls(itens, f, chave))})
    elif chave in CURTOS:
        op = graficos.colunas([i["rotulo"] for i in top], {"Bens": [i["quantidade"] for i in top]}, rotulos=True, urls={"Bens": _urls(top, f, chave)})
    else:
        op = graficos.barras_horizontais([i["rotulo"] for i in top], [i["quantidade"] for i in top], "Bens", escala=False, urls=_urls(top, f, chave))  # azul da marca em todas as barras (a escala clara ficava ilegível)
    sub = (f"{TOP} anos mais recentes no gráfico; todos na tabela" if recentes else f"{TOP} maiores no gráfico; todos na tabela") if n > TOP else None
    barras = (n > 5 or chave in ORDINAIS) and chave not in CURTOS
    colunas = n > 5 and chave in CURTOS
    if colunas and len(top) > 10:
        op["xAxis"]["axisLabel"] = {"interval": 0, "rotate": 45}   # 20 anos lado a lado não cabem na horizontal
    altura = "extra" if barras and len(top) > 15 else ("alto" if len(top) > 8 else None)
    col = "col-12" if (barras or colunas) and len(top) > 10 else None
    return op, sub, altura, col


def _sub(universo, sub):
    return " · ".join(x for x in (universo, sub) if x)


def cards_graficos(dim: dict, f: dict, omitir=()) -> list[dict]:
    """Um card por dimensão, na ordem da spec. omitir = filtros já fixados num único valor."""
    cards = []
    sit = f.get("situacao")
    universo = "Bens ativos" if sit == "ATIVO" else (f"Bens {sit}" if sit else "Todas as situações")
    if "situacao" not in omitir and dim["situacao"]:
        it = dim["situacao"]
        op, sub, alto, col = _grafico(it, f, "situacao")
        cards.append(_card("g-situacao", "Bens por situação", op, it, f, "situacao", "Situação", alto=alto, subtitulo=_sub("Todos os bens", sub), col=col))
    if "ccusto" not in omitir:
        it = dim["centro"]
        op, sub, alto, col = _grafico(it, f, "ccusto")
        cards.append(_card("g-centro", "Bens por centro de custo", op, it, f, "ccusto", "Centro de custo", alto=alto, subtitulo=_sub(universo, sub), col=col))
    if "classificacao" not in omitir:
        it = dim["classificacao"]
        comuns = [i for i in it if i["chave"] not in db.IMOVEIS]
        imoveis = [i for i in it if i["chave"] in db.IMOVEIS]
        op, sub, alto, col = _grafico(it, f, "classificacao")   # imóveis entram no gráfico como qualquer classe (é contagem)
        cards.append(_card("g-classificacao", "Bens por classificação contábil", op, comuns + imoveis, f, "classificacao",
                           "Classificação", alto=alto, subtitulo=_sub(universo, sub), col=col))   # imóveis só ficam por último na tabela
    if "localizacao" not in omitir:
        it = dim["localizacao"]
        op, sub, alto, col = _grafico(it, f, "localizacao")
        cards.append(_card("g-localizacao", "Bens por localização", op, it, f, "localizacao", "Localização", alto=alto, subtitulo=_sub(universo, sub), col=col))
    if "idade" not in omitir:
        it = dim["idade"]
        op, sub, alto, col = _grafico(it, f, "idade")
        cards.append(_card("g-idade", "Bens por idade (data de entrada)", op, it, f, "idade", "Faixa", alto=alto, subtitulo=_sub(universo, sub), col=col))
    if "ano" not in omitir and dim["ano"]:
        it = dim["ano"]
        op, sub, alto, col = _grafico(it, f, "ano")
        cards.append(_card("g-ano", "Bens por ano de entrada", op, it, f, "ano", "Ano", alto=alto, subtitulo=_sub(universo, sub), col=col))
    if "faixa" not in omitir:
        it = dim["faixa"]
        op, sub, alto, col = _grafico(it, f, "faixa")
        cards.append(_card("g-faixa", "Bens por faixa de valor", op, it, f, "faixa", "Faixa", alto=alto, subtitulo=_sub(universo, sub), col=col))
    if "pessoa" not in omitir and dim["pessoa"]:
        it = dim["pessoa"]
        op, sub, alto, col = _grafico(it, f, "pessoa")
        cards.append(_card("g-pessoa", "Bens atribuídos por pessoa", op, it, f, "pessoa", "Pessoa", alto=alto, subtitulo=_sub(universo, sub), col=col))
    return cards


def indicadores_analise(r: dict, f: dict) -> list[dict]:
    """Seis cards do recorte: dois totais (sem link) e quatro contagens que, quando positivas e sem
    conflito com o filtro atual, viram um clique que ACRESCENTA uma restrição sem trocar as demais
    (spec fase5b §3). 'Valor não informado'/'Valor zero' nunca combinam com um intervalo numérico:
    NULL nunca satisfaz >=/<=, então essa combinação é sempre vazia; zero pode até coincidir com o
    intervalo, mas o clique não deve propor uma mistura que pareça contraditória."""
    def refinar(contagem, **novos):
        if not contagem or any(f.get(k) and f[k] != v for k, v in novos.items()):
            return None
        if all(f.get(k) == v for k, v in novos.items()):
            return None
        return url_recorte(f, **novos)

    def refinar_valor_status(contagem, valor):
        if f.get("valor_de") or f.get("valor_ate"):
            return None
        return refinar(contagem, valor_status=valor)

    return [
      dict(rotulo='Bens no recorte', valor=r['quantidade'], detalhe=None, url=None),
      dict(rotulo='Valor atual', valor=moeda(r['valor_total']), detalhe=None, url=None),
      dict(rotulo='Imóveis', valor=r['imoveis'], detalhe=moeda(r['valor_imoveis']),
           url=refinar(r['imoveis'], classificacao='imoveis')),
      dict(rotulo='Sem centro nem pessoa', valor=r['sem_centro'], detalhe=None,
           url=refinar(r['sem_centro'], ccusto='-', pessoa='-')),
      dict(rotulo='Valor não informado', valor=r['valor_nao_informado'], detalhe=None,
           url=refinar_valor_status(r['valor_nao_informado'], 'nao_informado')),
      dict(rotulo='Valor zero', valor=r['valor_zero'], detalhe=None,
           url=refinar_valor_status(r['valor_zero'], 'zero')),
    ]


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
    if f.get("valor_status"):
        partes.append({"nao_informado": "valor não informado", "zero": "valor zero"}[f["valor_status"]])
    return " · ".join(partes)
