"""Cards do painel do evento de inventário, no formato da macro `grafico` (templates/_macros.html) e no padrão do
Recorte: tipo de gráfico pelo nº de itens, tudo clicável, tabela "Ver dados". Só monta opções ECharts com
graficos.py; os números vêm de inventario.painel()."""
from flask import url_for

import graficos
from inventario import ANDAR_SEM

TOP = 20
STATUS_SITUACAO = {"Localizado": "sucesso", "Divergente": "alerta", "Não localizado": "neutro"}
STATUS_PROGRESSO = {"Localizados": "sucesso", "Pendentes": "pendente"}


def _card(id, titulo, opcoes, tabela, resumo_itens, subtitulo=None, altura=None, col=None):
    resumo = f"{titulo}: " + ", ".join(f"{r} {q}" for r, q in resumo_itens[:6])
    return {"id": id, "titulo": titulo, "subtitulo": subtitulo, "opcoes": opcoes, "resumo": resumo,
            "alto": altura == "alto", "altura": altura, "col": col, "tabela": tabela}


def _tabela_contagem(itens, urls, rotulo_col, valor_col):
    return graficos.tabela_dados([rotulo_col, valor_col], [[{"valor": i["rotulo"], "url": u}, i["quantidade"]] for i, u in zip(itens, urls)])


def _tabela_progresso(itens, urls, rotulo_col, chave_rotulo):
    return graficos.tabela_dados([rotulo_col, "Bens", "Localizados", "Divergentes", "Pendentes"],
                                 [[{"valor": i[chave_rotulo], "url": u}, i["total"], i["localizados"], i["divergentes"], i["pendentes"]]
                                  for i, u in zip(itens, urls)])


def _grafico(itens, urls, nome):
    """≤ 5 itens → rosca com o total no centro; senão barras horizontais dos TOP maiores.
    Devolve (opções, subtítulo, altura)."""
    if len(itens) <= 5:
        op = graficos.rosca([(i["rotulo"], i["quantidade"]) for i in itens], total=(sum(i["quantidade"] for i in itens), nome.lower()),
                            urls={i["rotulo"]: u for i, u in zip(itens, urls)})
        return op, None, None
    top, top_urls = itens[:TOP], urls[:TOP]
    op = graficos.barras_horizontais([i["rotulo"] for i in top], [i["quantidade"] for i in top], nome, urls=top_urls)
    sub = f"{TOP} maiores no gráfico; todos na tabela" if len(itens) > TOP else None
    return op, sub, "extra" if len(top) > 15 else ("alto" if len(top) > 8 else None)


def _empilhado(rotulos, itens, urls):
    return graficos.colunas(rotulos, {"Localizados": [i["localizados"] for i in itens], "Pendentes": [i["pendentes"] for i in itens]},
                            empilhado=True, rotulos=True, status=STATUS_PROGRESSO, urls={"Localizados": urls, "Pendentes": urls})


def cards(dados: dict, evento_id: int, andar_sel: str | None = None) -> list[dict]:
    def rel(**f):
        return url_for("inventario.relatorio_tela", id=evento_id, **f)
    lista = []
    it = dados["situacao"]
    urls = [rel(situacao=i["chave"]) for i in it]
    op = graficos.rosca([(i["rotulo"], i["quantidade"]) for i in it], status=STATUS_SITUACAO,
                        total=(sum(i["quantidade"] for i in it), "bens"), urls={i["rotulo"]: u for i, u in zip(it, urls)})
    lista.append(_card("g-situacao", "Bens por situação", op, _tabela_contagem(it, urls, "Situação", "Bens"),
                       [(i["rotulo"], i["quantidade"]) for i in it], subtitulo="clique para abrir o relatório"))
    it = dados["integrantes"]
    if it:
        urls = [rel(integrante=i["chave"]) for i in it]
        op, sub, altura = _grafico(it, urls, "Leituras")
        lista.append(_card("g-integrantes", "Leituras por integrante", op, _tabela_contagem(it, urls, "Integrante", "Leituras"),
                           [(i["rotulo"], i["quantidade"]) for i in it], subtitulo=sub, altura=altura))
    it = dados["conservacao"]
    if it:
        urls = [rel(conservacao=i["chave"]) for i in it]
        op, sub, altura = _grafico(it, urls, "Leituras")
        lista.append(_card("g-conservacao", "Conservação informada", op, _tabela_contagem(it, urls, "Conservação", "Leituras"),
                           [(i["rotulo"], i["quantidade"]) for i in it], subtitulo=sub, altura=altura))
    it = dados["andares"]
    if it:
        urls = [url_for("inventario.painel_tela", id=evento_id, andar=i["andar"]) for i in it]
        lista.append(_card("g-andares", "Progresso por andar", _empilhado([i["andar"] for i in it], it, urls),
                           _tabela_progresso(it, urls, "Andar", "andar"), [(i["andar"], i["total"]) for i in it],
                           subtitulo="localizados e pendentes; clique no andar para ver as salas", col="col-12" if len(it) > 10 else None))
    it = dados["salas_do_andar"]
    if andar_sel and it:
        urls = [url_for("inventario.sala_tela", id=evento_id, localizacao=s["localizacao"]) for s in it]
        rotulos = [s["localizacao"] if andar_sel == ANDAR_SEM else (s["localizacao"].partition(" - ")[2] or s["localizacao"]) for s in it]
        op = _empilhado(rotulos, it, urls)
        if len(it) > 10:
            op["xAxis"]["axisLabel"] = {"interval": 0, "rotate": 45}
        lista.append(_card("g-salas", f"Salas do andar {andar_sel}", op, _tabela_progresso(it, urls, "Sala", "localizacao"),
                           [(s["localizacao"], s["total"]) for s in it],
                           subtitulo="clique na sala para abrir a leitura", col="col-12" if len(it) > 10 else None))
    return lista
