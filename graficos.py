"""Opções ECharts prontas, no tema dsgov (copiado da skill /dsgov, sem Django). A rota monta os dados; o JSON vai ao template pela macro grafico.

Cores nunca são escolhidas aqui, salvo os dois espelhos do tema (`echarts-dsgov.js`):
CORES_STATUS (dsgov.graficos.status) para fatias/barras cujo significado é um status, e
CORES_SEQUENCIAL (dsgov.graficos.sequencial) para escala de intensidade por valor (`escala=True`).

Drill-down (references/graficos.md, regra 9): todo montador aceita `urls` — um item de dado com `url`
navega ao clique. `rotulos=True` mostra o valor no próprio gráfico; `tabela` no card ("Ver dados")
é montada pela view com `tabela_dados()`.
"""

CORES_STATUS = {  # mesmas chaves do filtro classe_status (templatetags/dsgov.py) e de dsgov.graficos.status
    "sucesso": "#168821", "concluido": "#168821", "ativo": "#168821",
    "alerta": "#ffcd07", "pendente": "#ffcd07",
    "erro": "#e52207", "atrasado": "#e52207",
    "cancelado": "#757575", "neutro": "#757575", "inativo": "#757575",
    "info": "#155bcb", "andamento": "#155bcb",
}

CORES_SEQUENCIAL = ["#d4e5ff", "#81aefc", "#2670e8", "#1351b4", "#0c326f"]  # espelho de dsgov.graficos.sequencial
# Espelho da paleta categórica do tema Tabler (CATEGORICA em static/tabler/graficos.js): usada só quando a cor
# acompanha uma entidade em vários gráficos (a mesma classificação com a mesma cor na rosca e nas colunas).
CORES_CATEGORICA = ["#066fd1", "#2fb344", "#f76707", "#ae3ec9", "#17a2b8", "#f59f00", "#d6336c", "#4263eb", "#74b816", "#0ca678"]
COR_OUTROS = "#888888"  # cinza neutro do balde "Outros"/"Outras" (o tema troca pelo tom da aparência)
_COR_TEXTO = "#333333"  # gray-80, espelho de dsgov.graficos.cores.texto (só para o total no centro da rosca)


def _dado(valor, url=None):
    """Um item de dado: número puro ou objeto {value, url} quando há drill-down."""
    if url:
        return {"value": valor, "url": url}
    return valor


def rosca(fatias, status=None, total=None, rotulos=False, urls=None, rotulo_formato="{c} ({d}%)", cores=None, dica="{b}: {c} ({d}%)"):
    """fatias = [(rótulo, valor)]. status opcional = {rótulo: chave de CORES_STATUS}.
    total = (número, legenda) exibido no centro da rosca (graficos.md, tipo Rosca).
    rotulos=True escreve o valor e o percentual nas fatias (o nome fica na legenda — rótulo longo
    fora da rosca é truncado em 280px; use rotulo_formato="{b}: {c} ({d}%)" só em cards altos).
    urls = {rótulo: url} para drill-down. cores = {rótulo: cor} quando a fatia é uma entidade com cor própria.
    dica = modelo do tooltip ("R$ {c}" sai em moeda exata no tema Tabler)."""
    dados = []
    for rotulo, valor in fatias:
        item = {"name": rotulo, "value": valor}
        if status and rotulo in status:
            item["itemStyle"] = {"color": CORES_STATUS[status[rotulo]]}
        elif cores and rotulo in cores:
            item["itemStyle"] = {"color": cores[rotulo]}
        if urls and urls.get(rotulo):
            item["url"] = urls[rotulo]
        dados.append(item)
    opcoes = {
        "tooltip": {"trigger": "item", "formatter": dica},
        "legend": {"bottom": 0},
        "series": [{
            "type": "pie", "radius": ["55%", "80%"], "avoidLabelOverlap": True,
            "label": {"show": True, "formatter": rotulo_formato} if rotulos else {"show": False},
            "emphasis": {"label": {"show": True, "fontWeight": "bold"}},
            "data": dados,
        }],
    }
    if total is not None:
        numero, legenda = total if isinstance(total, (tuple, list)) else (total, "")
        opcoes["graphic"] = [{
            "type": "text", "left": "center", "top": "center",
            "style": {"text": f"{numero}\n{legenda}".rstrip(), "textAlign": "center",
                      "fontSize": 20, "fontWeight": "bold", "fill": _COR_TEXTO, "lineHeight": 24},
        }]
    return opcoes


def barras_horizontais(rotulos, valores, nome, escala=False, urls=None):
    """escala=True colore cada barra pela intensidade do valor (paleta sequencial do tema).
    urls = lista paralela a `valores` com o destino do clique (ou None)."""
    dados = [_dado(v, urls[i] if urls else None) for i, v in enumerate(valores)]
    opcoes = {
        "tooltip": {"trigger": "axis", "axisPointer": {"type": "shadow"}},
        "xAxis": {"type": "value"},
        "yAxis": {"type": "category", "data": rotulos, "inverse": True, "axisLabel": {"interval": 0}},
        "series": [{"type": "bar", "name": nome, "data": dados, "label": {"show": True, "position": "right"}}],
    }
    if escala and valores:
        numeros = [v for v in valores if isinstance(v, (int, float))]
        opcoes["visualMap"] = {"show": False, "min": min(numeros or [0]), "max": max(numeros or [1]),
                               "inRange": {"color": CORES_SEQUENCIAL}, "seriesIndex": 0}
    return opcoes


def colunas(categorias, series, empilhado=False, rotulos=False, urls=None, status=None, horizontal=False, cores=None, moeda=False):
    """series = {nome: [valores]}. urls = {nome: [url por categoria]}. status = {nome: chave de CORES_STATUS}
    (só quando a série É um status). rotulos=True escreve o valor no topo/dentro das colunas.
    horizontal=True deita as barras (categorias no eixo Y, na ordem dada, de cima para baixo) — para nomes longos.
    cores = {nome: cor} quando a série é uma entidade com cor própria. moeda=True: tooltip em R$ exato."""
    lista = []
    ultima = list(series)[-1] if series else None
    for nome, dados in series.items():
        u = (urls or {}).get(nome)
        serie = {"type": "bar", "name": nome, "data": [_dado(v, u[i] if u else None) for i, v in enumerate(dados)]}
        if empilhado:
            serie["stack"] = "total"
        if rotulos:
            serie["label"] = {"show": True, "position": "inside" if empilhado else ("right" if horizontal else "top")}
        estilo = {}
        if status and nome in status:
            estilo["color"] = CORES_STATUS[status[nome]]
        elif cores and nome in cores:
            estilo["color"] = cores[nome]
        if horizontal and (not empilhado or nome == ultima):
            estilo["borderRadius"] = [0, 4, 4, 0]   # ponta arredondada do lado do valor (à direita)
        if estilo:
            serie["itemStyle"] = estilo
        lista.append(serie)
    categoria = {"type": "category", "data": categorias, "axisLabel": {"interval": 0}}
    valor = {"type": "value"}
    opcoes = {
        "tooltip": {"trigger": "axis", "axisPointer": {"type": "shadow"}},
        "legend": {"bottom": 0} if len(series) > 1 else {"show": False},
        "xAxis": valor if horizontal else categoria,
        "yAxis": {**categoria, "inverse": True} if horizontal else valor,
        "series": lista,
    }
    if horizontal:
        opcoes["grid"] = {"right": 40}   # folga para o último número do eixo e os rótulos à direita da barra
    if moeda:
        opcoes["tooltip"]["valueFormatter"] = "moeda"
    return opcoes


def linha(categorias, series, urls=None):
    lista = []
    for nome, dados in series.items():
        u = (urls or {}).get(nome)
        lista.append({"type": "line", "name": nome, "data": [_dado(v, u[i] if u else None) for i, v in enumerate(dados)]})
    return {
        "tooltip": {"trigger": "axis"},
        "legend": {"bottom": 0} if len(series) > 1 else {"show": False},
        "xAxis": {"type": "category", "data": categorias, "boundaryGap": False, "axisLabel": {"interval": 0}},
        "yAxis": {"type": "value"},
        "series": lista,
    }


def pareto(pontos, marco=None, nome="Valor acumulado", referencia="Distribuição igual"):
    """Curva de concentração: pontos = [(% dos bens, % do valor acumulado)], de 0 a 100 nos dois eixos.
    marco = (x, y, url) marca o ponto de corte (ex.: 80% do valor) com linha tracejada e clique de drill-down.
    A diagonal tracejada mostra a distribuição igualitária (referência, cinza, sem área)."""
    serie = {"type": "line", "name": nome, "data": [[x, y] for x, y in pontos], "smooth": False, "showSymbol": False}
    if marco:
        x, y, url = marco
        ponto = {"coord": [x, y], "value": y}
        if url:
            ponto["url"] = url
        serie["markPoint"] = {"symbol": "circle", "symbolSize": 10, "label": {"show": False}, "data": [ponto]}
        serie["markLine"] = {"silent": True, "symbol": "none", "label": {"show": False},
                             "lineStyle": {"type": "dashed", "color": COR_OUTROS}, "data": [{"yAxis": y}]}
    return {
        "tooltip": {"trigger": "axis", "valueFormatter": "percentual"},
        "legend": {"bottom": 0},
        "grid": {"right": 24},   # o "100%" do fim do eixo X não é cortado
        "xAxis": {"type": "value", "min": 0, "max": 100, "interval": 20, "axisLabel": {"formatter": "{value}%"}},
        "yAxis": {"type": "value", "min": 0, "max": 100, "interval": 20, "axisLabel": {"formatter": "{value}%"}},
        "series": [serie, {"type": "line", "name": referencia, "data": [[0, 0], [100, 100]], "smooth": False,
                           "showSymbol": False, "silent": True, "areaStyle": {"opacity": 0},
                           "itemStyle": {"color": COR_OUTROS}, "lineStyle": {"type": "dashed", "width": 1.5, "color": COR_OUTROS}}],
    }


def tabela_dados(colunas, linhas):
    """Tabela companheira do card ("Ver dados"): colunas = [rótulo], linhas = [[célula]] onde célula é
    um valor já formatado ou {"valor": ..., "url": ...}. Renderizada por dsgov/_grafico_card.html."""
    return {"colunas": list(colunas),
            "linhas": [[c if isinstance(c, dict) else {"valor": c} for c in linha] for linha in linhas]}
