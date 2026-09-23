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


# ---------------------------------------------------------------- Análise no Tabler (gráficos redesenhados)
# O DSGov (templates/analise.html) segue com cards_graficos acima; a tela no Tabler usa cards_analise.
DIMENSOES_EXPLORAR = [("ccusto", "Centro de custo", "centro"), ("localizacao", "Localização", "localizacao"),
                      ("classificacao", "Classificação contábil", "classificacao"), ("pessoa", "Pessoa", "pessoa"),
                      ("situacao", "Situação", "situacao")]   # (filtro, rótulo, chave em db.dimensoes)
CORES_IDADE = dict(zip([k for k, _ in db.FAIXAS_IDADE], graficos.CORES_SEQUENCIAL[1:5] + [graficos.COR_OUTROS]))
TOP_EXPLORAR, TOP_ROSCA, TOP_ENTRADAS, TOP_COMPRA, ANOS = 12, 6, 5, 8, 15
DIM_VALOR = ("classificacao", "ccusto", "localizacao")   # a rosca e "comprado × hoje" usam a primeira não fixada


def dados_analise(conn, f: dict, omitir=()) -> dict:
    """Agregações que os gráficos da análise no Tabler precisam, todas em SQL (db.cruzamento, db.concentracao)."""
    dims = [d for d, _, _ in DIMENSOES_EXPLORAR if d not in omitir]
    return {"cruz": {d: db.cruzamento(conn, f, d, "idade") for d in dims},
            "entradas": db.cruzamento(conn, f, "ano", "classificacao") if "ano" not in omitir else [],
            "pareto": db.concentracao(conn, f), "cobertura": db.cobertura_termos(conn, f)}


def moeda_curta(v) -> str:
    """Igual à `curta` de static/tabler/graficos.js (o total da rosca é reescrito pelo JS no mesmo formato)."""
    a, s = abs(v or 0), "-" if (v or 0) < 0 else ""
    if a >= 1e6:
        return f"{s}{a / 1e6:.1f} mi".replace(".", ",")
    if a >= 1e3:
        return f"{s}{a / 1e3:.1f} mil".replace(".", ",")
    return f"{v or 0:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")


def percentual(p: float) -> str:
    """Percentual legível: 58% · 5,3% · 0,09% (mais casas só quando o número é pequeno)."""
    casas = 0 if p >= 10 else (1 if p >= 1 else 2)
    texto = f"{p:.{casas}f}".replace(".", ",")
    if "," in texto:
        texto = texto.rstrip("0").rstrip(",")
    return texto + "%"


def _inteiro(n) -> str:
    return f"{n:,}".replace(",", ".")


def _somar(linhas, chave="linha"):
    """Soma cruzamento por uma das chaves: {chave: {quantidade, valor, compra}}."""
    tot = {}
    for x in linhas:
        t = tot.setdefault(x[chave], {"quantidade": 0, "valor": 0.0, "compra": 0.0})
        for m in ("quantidade", "valor", "compra"):
            t[m] += x[m]
    return tot


def _cores_classificacao(por_valor: list, usadas: set) -> dict:
    """Cor por classificação, igual em todos os gráficos: a paleta categórica segue a ordem de valor atual,
    distribuída só entre as classes que aparecem em algum gráfico (as demais caem em "Outras")."""
    ordem = [c for c in por_valor if c in usadas] + sorted(usadas - set(por_valor))
    paleta = graficos.CORES_CATEGORICA
    return {c: paleta[i % len(paleta)] for i, c in enumerate(ordem)}


def _rotulos(dim: dict, chave_dim: str) -> dict:
    return {x["chave"]: x["rotulo"] for x in dim.get(chave_dim, [])}


def _medida_fmt(medida, v):
    return _inteiro(v) if medida == "quantidade" else moeda(v)


def _card_rosca(tot, rotulos, f, dim, rotulo_dim, cores, universo):
    itens = sorted(((k, t["valor"]) for k, t in tot.items() if t["valor"] > 0), key=lambda x: (-x[1], x[0]))
    if not itens:
        return None
    topo, resto = itens[:TOP_ROSCA], itens[TOP_ROSCA:]
    fatias = [(rotulos.get(k, k), round(v, 2)) for k, v in topo]
    urls = {rotulos.get(k, k): url_recorte(f, **{dim: k}) for k, _ in topo}
    cor = {rotulos.get(k, k): cores[k] for k, _ in topo if k in cores}
    if resto:
        nome = f"Outras ({len(resto)})" if dim == "classificacao" else f"Outros ({len(resto)})"
        fatias.append((nome, round(sum(v for _, v in resto), 2)))
        cor[nome] = graficos.COR_OUTROS
    total = round(sum(v for _, v in fatias), 2)
    op = graficos.rosca(fatias, total=(moeda_curta(total), "em R$"), urls=urls, cores=cor, dica="{b}<br/>R$ {c} ({d}%)")
    linhas = [[{"valor": rotulos.get(k, k), "url": url_recorte(f, **{dim: k})}, moeda(v), percentual(100 * v / total)] for k, v in itens]
    resumo = "Onde está o valor: " + ", ".join(f"{n} {moeda(v)}" for n, v in fatias)
    return {"id": "g-valor", "titulo": "Onde está o valor", "subtitulo": _sub(universo, f"valor atual por {rotulo_dim.lower()}"),
            "opcoes": op, "resumo": resumo, "alto": True, "altura": "alto", "col": "col-12 col-lg-5",
            "tabela": graficos.tabela_dados([rotulo_dim, "Valor atual", "Parcela"], linhas)}


def _card_explorar(cruz, dim_rotulos, f, medidas, faixas, universo):
    """Barras horizontais empilhadas por faixa de idade; uma variante por dimensão × medida (seletores no card)."""
    variantes, opcoes_dim = [], []
    nomes_faixa = dict(db.FAIXAS_IDADE)
    for dim, rotulo_dim, chave_dim in DIMENSOES_EXPLORAR:
        if dim not in cruz:
            continue
        linhas = [x for x in cruz[dim] if not (dim == "pessoa" and x["linha"] == "-")]
        if not linhas:
            continue
        opcoes_dim.append((dim, rotulo_dim))
        rot = dim_rotulos.get(chave_dim, {})
        por = {}
        for x in linhas:
            por.setdefault(x["linha"], {})[x["coluna"]] = x
        for medida, rotulo_medida in medidas:
            total = {k: sum(c[medida] for c in v.values()) for k, v in por.items()}
            ordem = sorted(por, key=lambda k: (-total[k], rot.get(k, k)))
            topo, resto = ordem[:TOP_EXPLORAR], ordem[TOP_EXPLORAR:]
            categorias = [rot.get(k, k) for k in topo] + ([f"Outros ({len(resto)})"] if resto else [])
            series, urls, cores = {}, {}, {}
            for fx in faixas:
                nome = nomes_faixa[fx]
                vals = [round(por[k].get(fx, {}).get(medida, 0), 2) for k in topo]
                if resto:
                    vals.append(round(sum(por[k].get(fx, {}).get(medida, 0) for k in resto), 2))
                series[nome] = vals
                urls[nome] = [url_recorte(f, **{dim: k, "idade": fx}) for k in topo] + ([None] if resto else [])
                cores[nome] = CORES_IDADE[fx]
            op = graficos.colunas(categorias, series, empilhado=True, horizontal=True, urls=urls, cores=cores,
                                  moeda=medida == "valor")
            tabela = graficos.tabela_dados([rotulo_dim] + [nomes_faixa[fx] for fx in faixas] + ["Total"], [
                [{"valor": rot.get(k, k), "url": url_recorte(f, **{dim: k})}]
                + [_medida_fmt(medida, por[k].get(fx, {}).get(medida, 0)) for fx in faixas] + [_medida_fmt(medida, total[k])]
                for k in topo] + ([[f"Outros ({len(resto)})"]
                                   + [_medida_fmt(medida, sum(por[k].get(fx, {}).get(medida, 0) for k in resto)) for fx in faixas]
                                   + [_medida_fmt(medida, sum(total[k] for k in resto))]] if resto else []))
            resumo = f"{rotulo_medida} por {rotulo_dim.lower()}: " + ", ".join(
                f"{rot.get(k, k)} {_medida_fmt(medida, total[k])}" for k in topo[:6])
            variantes.append({"id": f"g-explorar--{dim}--{medida}", "opcoes": op, "resumo": resumo, "tabela": tabela})
    if not variantes:
        return None
    card = {"id": "g-explorar", "titulo": "Explorar a base",
            "subtitulo": _sub(universo, "por faixa de idade (data de entrada)"),
            "alto": True, "altura": "alto", "col": "col-12"}
    return _com_variantes(card, variantes, [("Dimensão", opcoes_dim), ("Medida", medidas)])


def _card_entradas(linhas, f, rotulos_class, cores, universo):
    linhas = [x for x in linhas if x["linha"] != "-"]
    anos = sorted({x["linha"] for x in linhas})
    if len(anos) < 2:
        return None
    recentes = anos[-ANOS:]
    antes = anos[:-ANOS] if len(anos) > ANOS else []
    categorias = ([f"até {int(recentes[0]) - 1}"] if antes else []) + recentes
    limite = f"{int(recentes[0]) - 1}-12-31"
    if f.get("entrada_ate") and f["entrada_ate"] < limite:
        limite = f["entrada_ate"]
    variantes = []
    for medida, rotulo_medida in (("quantidade", "Quantidade"), ("compra", "Valor de compra")):
        tot = _somar(linhas, "coluna")
        classes = [c for c in sorted(tot, key=lambda c: (-tot[c][medida], c)) if tot[c][medida] > 0]
        topo, resto = classes[:TOP_ENTRADAS], classes[TOP_ENTRADAS:]
        if not topo:
            continue
        grupos = [(rotulos_class.get(c, c), {c}, c) for c in topo] + ([(f"Outras ({len(resto)})", set(resto), None)] if resto else [])

        def soma(anos_cat, classes_g):
            return round(sum(x[medida] for x in linhas if x["linha"] in anos_cat and x["coluna"] in classes_g), 2)
        colunas_anos = ([set(antes)] if antes else []) + [{a} for a in recentes]
        extras = ([{"entrada_ate": limite}] if antes else []) + [{"ano": a} for a in recentes]
        series, urls, cor = {}, {}, {}
        for nome, classes_g, chave in grupos:
            series[nome] = [soma(ac, classes_g) for ac in colunas_anos]
            urls[nome] = [url_recorte(f, **e, classificacao=chave) if chave else None for e in extras]
            cor[nome] = cores.get(chave, graficos.COR_OUTROS) if chave else graficos.COR_OUTROS
        op = graficos.colunas(categorias, series, empilhado=True, urls=urls, cores=cor, moeda=medida == "compra")
        op["series"][-1].setdefault("itemStyle", {})["borderRadius"] = [4, 4, 0, 0]   # topo da pilha arredondado
        totais = [sum(series[n][i] for n in series) for i in range(len(categorias))]
        tabela = graficos.tabela_dados(["Ano de entrada"] + list(series) + ["Total"], [
            [{"valor": cat, "url": url_recorte(f, **e)}] + [_medida_fmt(medida, series[n][i]) for n in series] + [_medida_fmt(medida, totais[i])]
            for i, (cat, e) in enumerate(zip(categorias, extras))])
        resumo = f"Entradas por ano ({rotulo_medida.lower()}): " + ", ".join(f"{c} {_medida_fmt(medida, t)}" for c, t in zip(categorias[-6:], totais[-6:]))
        variantes.append({"id": f"g-entradas--{medida}", "opcoes": op, "resumo": resumo, "tabela": tabela})
    if not variantes:
        return None
    sub = f"últimos {len(recentes)} anos com entrada" + (f"; os anteriores somados em “{categorias[0]}”" if antes else "")
    card = {"id": "g-entradas", "titulo": "Entradas por ano", "subtitulo": _sub(universo, sub + "; por classificação"),
            "alto": True, "altura": "alto", "col": "col-12"}
    return _com_variantes(card, variantes, [("Medida", [("quantidade", "Quantidade"), ("compra", "Valor de compra")])])


def _com_variantes(card: dict, variantes: list, seletores: list) -> dict:
    """Card com seletores no cabeçalho: só as opções que têm variante; seletor de uma opção só não aparece
    (e não entra no id). Sem nenhuma escolha, vira card simples com a única variante."""
    ids = {v["id"] for v in variantes}
    partes = [v["id"].split("--")[1:] for v in variantes]
    visiveis = []
    for i, (rotulo, opcoes) in enumerate(seletores):
        existentes = [(val, r) for val, r in opcoes if any(p[i] == val for p in partes)]
        visiveis.append((rotulo, existentes))
    manter = [i for i, (_, op) in enumerate(visiveis) if len(op) > 1]
    for v in variantes:
        p = v["id"].split("--")
        v["id"] = "--".join([p[0]] + [p[1 + i] for i in manter])
    assert len({v["id"] for v in variantes}) == len(ids)
    if not manter:
        return {**card, **{k: variantes[0][k] for k in ("opcoes", "resumo", "tabela")}}
    return {**card, "variantes": variantes, "seletores": [{"rotulo": visiveis[i][0], "opcoes": visiveis[i][1]} for i in manter]}


def cabecalho_pareto(k: int, n: int) -> dict:
    """Frase do card de concentração: {"forte": "3 bens", "texto": "(0,09% do recorte) concentram 80% do valor atual"}."""
    return {"forte": f"{_inteiro(k)} {'bem' if k == 1 else 'bens'}",
            "texto": f"({percentual(100 * k / n)} do recorte) {'concentra' if k == 1 else 'concentram'} 80% do valor atual"}


def _card_pareto(c, f, universo):
    if not c["k"]:
        return None
    n, total = c["quantidade"], c["total"]
    acum = dict(c["pontos"])
    pontos = [(0, 0)] + [(round(100 * rn / n, 2), round(100 * v / total, 2)) for rn, v in c["pontos"]]
    url = url_recorte(f, valor_de=f"{c['corte']:.2f}")
    x80, y80 = round(100 * c["k"] / n, 2), round(100 * acum[c["k"]] / total, 2)
    op = graficos.pareto(pontos, marco=(x80, y80, url))
    op["tooltip"]["axisPointer"] = {"type": "line"}
    linhas = [[{"valor": f"{_inteiro(c['k'])} mais valiosos", "url": url}, percentual(100 * c["k"] / n), percentual(100 * acum[c["k"]] / total)]]
    for i in (1, 5, 10, 20, 50):
        rn = -(-n * i // 100)
        if rn in acum and rn != c["k"]:
            linhas.append([f"{_inteiro(rn)} mais valiosos", percentual(100 * rn / n), percentual(100 * acum[rn] / total)])
    linhas.sort(key=lambda l: float(l[1].rstrip("%").replace(",", ".")))
    destaque = cabecalho_pareto(c["k"], n)
    return {"id": "g-concentracao", "titulo": "Concentração do valor",
            "subtitulo": _sub(universo, "do bem mais valioso ao menos valioso"),
            "destaque": destaque, "opcoes": op, "resumo": f"Concentração do valor: {destaque['forte']} {destaque['texto']}",
            "alto": True, "altura": "alto", "col": "col-12 col-lg-7",
            "tabela": graficos.tabela_dados(["Bens", "Parcela dos bens", "Parcela do valor"], linhas)}


def _card_compra(tot, rotulos, f, dim, rotulo_dim, universo):
    itens = sorted((k for k, t in tot.items() if t["compra"] > 0 or t["valor"] > 0), key=lambda k: (-tot[k]["compra"], -tot[k]["valor"], k))
    topo = itens[:TOP_COMPRA]
    if not topo:
        return None
    cats = [rotulos.get(k, k) for k in topo]
    urls = [url_recorte(f, **{dim: k}) for k in topo]
    op = graficos.colunas(cats, {"Valor de compra": [round(tot[k]["compra"], 2) for k in topo],
                                 "Valor atual": [round(tot[k]["valor"], 2) for k in topo]},
                          horizontal=True, rotulos=True, moeda=True, urls={"Valor de compra": urls, "Valor atual": urls},
                          cores={"Valor de compra": graficos.CORES_SEQUENCIAL[1], "Valor atual": graficos.CORES_SEQUENCIAL[3]})
    linhas = [[{"valor": rotulos.get(k, k), "url": url_recorte(f, **{dim: k})}, moeda(tot[k]["compra"]), moeda(tot[k]["valor"]),
               percentual(100 * tot[k]["valor"] / tot[k]["compra"]) if tot[k]["compra"] else "—"] for k in itens]
    sub = f"{len(topo)} maiores por valor de compra; todos na tabela" if len(itens) > len(topo) else "por valor de compra"
    return {"id": "g-compra", "titulo": "Comprado × vale hoje", "subtitulo": _sub(universo, f"por {rotulo_dim.lower()}, {sub}"),
            "opcoes": op, "resumo": "Comprado × vale hoje: " + ", ".join(f"{rotulos.get(k, k)} {moeda(tot[k]['compra'])} → {moeda(tot[k]['valor'])}" for k in topo[:4]),
            "alto": True, "altura": "alto", "col": "col-12 col-lg-6",
            "tabela": graficos.tabela_dados([rotulo_dim, "Valor de compra", "Valor atual", "Atual / compra"], linhas)}


ESTADOS_TERMO = [("vigente", "Termo em dia", "sucesso"), ("desatualizado", "Termo desatualizado", "alerta"),
                 ("sem_termo", "Sem termo", "erro"), ("sem_responsavel", "Sem responsável", "neutro")]


def cabecalho_cobertura(itens: list) -> dict | None:
    """Frase do card de cobertura: "X% dos bens (Y% do valor) com termo em dia"."""
    n = sum(i["quantidade"] for i in itens)
    if not n:
        return None
    v = sum(i["valor"] for i in itens)
    q_ok = sum(i["quantidade"] for i in itens if i["estado"] == "vigente")
    v_ok = sum(i["valor"] for i in itens if i["estado"] == "vigente")
    return {"forte": f"{percentual(100 * q_ok / n)} dos bens",
            "texto": (f"({percentual(100 * v_ok / v)} do valor) " if v > 0 else "") + "com termo em dia"}


def _card_cobertura(itens, f, rotulos_centro, medidas):
    """Barras horizontais empilhadas pela situação do termo: uma por centro (os maiores), Outros centros,
    Termos individuais (todas as pessoas) e Sem responsável; seletor de medida."""
    if not itens:
        return None
    base = {"situacao": "ATIVO"}
    estados = [e for e in ESTADOS_TERMO if any(i["estado"] == e[0] for i in itens)]
    variantes = []
    for medida, rotulo_medida in medidas:
        centros_ = sorted((i for i in itens if i["tipo"] == "ccusto"), key=lambda i: (-i[medida], i["chave"]))
        topo, resto = centros_[:TOP_EXPLORAR], centros_[TOP_EXPLORAR:]
        barras = [(rotulos_centro.get(i["chave"], i["chave"]), [i], url_recorte(f, **base, ccusto=i["chave"])) for i in topo]
        if resto:
            barras.append((f"Outros centros ({len(resto)})", resto, None))
        pessoas_ = [i for i in itens if i["tipo"] == "individual"]
        if pessoas_:
            barras.append(("Termos individuais", pessoas_, url_for("termos_individuais")))
        sem = [i for i in itens if i["tipo"] == "-"]
        if sem:
            barras.append(("Sem responsável", sem, url_recorte(f, **base, ccusto="-", pessoa="-")))
        series, urls, status = {}, {}, {}
        for chave, nome, st in estados:
            vals = [round(sum(i[medida] for i in grupo if i["estado"] == chave), 2) for _, grupo, _ in barras]
            series[nome] = [v if v else None for v in vals]   # vazio sai como "-" no tooltip, não como segmento zero
            urls[nome] = [u for _, _, u in barras]
            status[nome] = st
        op = graficos.colunas([b[0] for b in barras], series, empilhado=True, horizontal=True, urls=urls, status=status,
                              moeda=medida == "valor")
        tabela = graficos.tabela_dados(["Guarda"] + [e[1] for e in estados] + ["Total"], [
            [{"valor": nome, "url": u} if u else nome]
            + [_medida_fmt(medida, sum(i[medida] for i in grupo if i["estado"] == e[0])) for e in estados]
            + [_medida_fmt(medida, sum(i[medida] for i in grupo))] for nome, grupo, u in barras])
        resumo = f"Cobertura de termos ({rotulo_medida.lower()}): " + ", ".join(
            f"{e[1]} {_medida_fmt(medida, sum(i[medida] for i in itens if i['estado'] == e[0]))}" for e in estados)
        variantes.append({"id": f"g-cobertura--{medida}", "opcoes": op, "resumo": resumo, "tabela": tabela})
    card = {"id": "g-cobertura", "titulo": "Cobertura de termos", "destaque": cabecalho_cobertura(itens),
            "subtitulo": "Bens ativos, pelo termo de quem os guarda (pessoa ou centro de custo)",
            "alto": True, "altura": "alto", "col": "col-12 col-lg-6"}
    return _com_variantes(card, variantes, [("Medida", medidas)])


def cards_analise(dim: dict, dados: dict, f: dict, omitir=(), total_valor: float = 0) -> list[dict]:
    """Cards da análise no Tabler, na ordem da tela: rosca do valor, concentração (Pareto), explorar a base
    (seletores dimensão × medida), entradas por ano (seletor de medida), comprado × vale hoje e cobertura de termos.
    dim = db.dimensoes (rótulos e contagens); dados = dados_analise(); omitir = filtros fixados num valor."""
    if not sum(x["quantidade"] for x in dim["idade"]):
        return []
    sit = f.get("situacao")
    universo = "Bens ativos" if sit == "ATIVO" else (f"Bens {sit}" if sit else "Todas as situações")
    rotulos = {d: _rotulos(dim, c) for d, _, c in DIMENSOES_EXPLORAR}
    cruz = dados["cruz"]
    dim_valor = next((d for d in DIM_VALOR if d in cruz), None)
    rot_dim = dict((d, r) for d, r, _ in DIMENSOES_EXPLORAR)
    tot_valor = _somar(cruz[dim_valor]) if dim_valor else {}

    # cores das classificações: ordem de valor atual; só as que aparecem na rosca ou nas entradas
    por_valor = [x["chave"] for x in sorted(dim["classificacao"], key=lambda x: (-x["valor"], x["chave"]))]
    usadas = set()
    if dim_valor == "classificacao":
        usadas |= set(sorted((k for k in tot_valor if tot_valor[k]["valor"] > 0), key=lambda k: -tot_valor[k]["valor"])[:TOP_ROSCA])
    tot_ent = _somar([x for x in dados["entradas"] if x["linha"] != "-"], "coluna")
    for m in ("quantidade", "compra"):
        usadas |= set(sorted((c for c in tot_ent if tot_ent[c][m] > 0), key=lambda c: (-tot_ent[c][m], c))[:TOP_ENTRADAS])
    cores = _cores_classificacao(por_valor, usadas)

    medidas = [("quantidade", "Quantidade")] + ([("valor", "Valor atual")] if total_valor > 0 else [])
    faixas = [x["chave"] for x in dim["idade"] if x["quantidade"]]
    cards = [
        _card_rosca(tot_valor, rotulos.get(dim_valor, {}), f, dim_valor, rot_dim.get(dim_valor, ""), cores if dim_valor == "classificacao" else {}, universo) if dim_valor else None,
        _card_pareto(dados["pareto"], f, universo),
        _card_explorar(cruz, {c: rotulos[d] for d, _, c in DIMENSOES_EXPLORAR}, f, medidas, faixas, universo),
        _card_entradas(dados["entradas"], f, rotulos["classificacao"], cores, universo),
        _card_compra(tot_valor, rotulos.get(dim_valor, {}), f, dim_valor, rot_dim.get(dim_valor, ""), universo) if dim_valor else None,
        _card_cobertura(dados["cobertura"], f, rotulos["ccusto"], medidas),
    ]
    return [c for c in cards if c]
