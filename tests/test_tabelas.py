"""Tabelas ordenáveis e com largura ajustável (static/js/tabelas.js), nas duas aparências.

O comportamento (arrastar, ordenar no navegador) é JS; aqui se prova o contrato de markup — toda tabela de dados
marcada com data-tabela, o script carregado com a aparência certa, o documento do termo de fora — e a ordenação
no servidor das listas cortadas (?ordem=&dir=, lista branca, vazios por último)."""
import re

import pytest

import db
from tests.test_tabler import tabler  # noqa: F401  (fixture)

TELAS = ["/", "/termos-emitidos", "/analise", "/pesquisa?q=cci", "/pesquisa?ccusto=CCI", "/centro-custos",
         "/termos-individuais", "/cadastros/processos", "/cadastros/responsaveis", "/cadastros/pessoas?nome=ANA%20SILVA",
         "/termo_devolucao?nome=ANA%20SILVA", "/upload", "/usuarios", "/administracao"]


def _tabelas(html):
    return re.findall(r"<table\b[^>]*>", html)


def _ordem_das_linhas(html, numeros):
    corpo = html.split("<tbody>", 1)[1] if "<tbody>" in html else html
    return sorted(numeros, key=lambda n: corpo.find(f">{n}<"))


@pytest.mark.parametrize("aparencia", ["dsgov", "tabler"])
def test_base_carrega_tabelas_js_com_a_aparencia_e_versao(cliente, aparencia, request):
    c = request.getfixturevalue("tabler") if aparencia == "tabler" else cliente
    html = c.get("/termos-emitidos").text
    assert re.search(r'<script src="/static/js/tabelas\.js\?v=[\w-]+" data-aparencia="%s"></script>' % aparencia, html)
    if aparencia == "dsgov":
        assert re.search(r'dsgov/css/dsgov\.css\?v=[\w-]+', html)


@pytest.mark.parametrize("aparencia", ["dsgov", "tabler"])
def test_toda_tabela_de_dados_e_ordenavel_e_ajustavel(cliente, aparencia, request):
    c = request.getfixturevalue("tabler") if aparencia == "tabler" else cliente
    c.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "T", "numero_sei": "2222", "vigente": "1"})
    c.get("/termo/ccusto/CCI/docx")
    vistas = 0
    for url in TELAS:
        r = c.get(url)
        assert r.status_code == 200, url
        for tag in _tabelas(r.text):
            vistas += 1
            assert "data-tabela" in tag, f"{aparencia} {url}: {tag}"
    assert vistas >= 10


@pytest.mark.parametrize("aparencia", ["dsgov", "tabler"])
def test_documento_do_termo_nao_ganha_alcas_nem_ordenacao(cliente, aparencia, request):
    c = request.getfixturevalue("tabler") if aparencia == "tabler" else cliente
    c.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "T", "numero_sei": "2222", "vigente": "1"})
    assert "data-tabela" not in c.get("/termo/ccusto/CCI").text
    doc = c.get("/termo/ccusto/CCI/documento").text
    assert "CADEIRA" in doc and "<table" in doc and "data-tabela" not in doc and "tabelas.js" not in doc


def test_cadastros_mantem_a_ordenacao_por_link_do_servidor(cliente):
    html = cliente.get("/cadastros/processos?ordem=descricao&direcao=desc").text
    assert 'id="tabela-cadastro" data-tabela' in html and "ordem=" in html


@pytest.mark.parametrize("aparencia", ["dsgov", "tabler"])
def test_cabecalhos_do_recorte_levam_o_campo_e_o_estado_da_ordem(cliente, aparencia, request):
    c = request.getfixturevalue("tabler") if aparencia == "tabler" else cliente
    html = c.get("/analise?situacao=&ordem=valor&dir=desc").text
    assert re.search(r'<th scope="col" class="[^"]+" data-ordem="valor" aria-sort="descending">Valor</th>', html)
    assert 'data-ordem="numero">Número</th>' in html
    assert 'data-tabela-ordem="servidor"' not in html   # 4 bens: lista inteira, ordena no navegador
    assert _ordem_das_linhas(html, [1001, 1002, 1003, 1004]) == [1002, 1004, 1001, 1003]


@pytest.mark.parametrize("aparencia", ["dsgov", "tabler"])
def test_lista_cortada_ordena_no_servidor(cliente, aparencia, request, monkeypatch):
    c = request.getfixturevalue("tabler") if aparencia == "tabler" else cliente
    original = db.pesquisar
    monkeypatch.setattr(db, "pesquisar", lambda conn, q, **kw: original(conn, q, limite=2, **kw))
    html = c.get("/pesquisa?q=a&ordem=descricao&dir=desc").text
    assert "Muitos resultados" in html
    assert re.search(r'<table[^>]*data-tabela data-tabela-ordem="servidor"', html)
    assert 'data-ordem="descricao" aria-sort="descending">Descrição</th>' in html
    # ordem no SQL antes do LIMIT: os 2 "maiores" de todos, não os 2 primeiros por número
    assert "NOTEBOOK" in html and "MESA" in html and "CADEIRA" not in html.split("<tbody>")[1]


def test_termos_emitidos_cortados_ordenam_no_servidor(cliente, monkeypatch):
    import app as modulo
    cliente.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "T", "numero_sei": "2222", "vigente": "1"})
    cliente.get("/termo/ccusto/CCI/docx")
    assert 'data-tabela-ordem="servidor"' not in cliente.get("/termos-emitidos").text
    monkeypatch.setattr(modulo, "LIMITE_TERMOS_EMITIDOS", 1)
    html = cliente.get("/termos-emitidos?ordem=valor&dir=asc").text
    assert 'data-tabela-ordem="servidor"' in html and 'data-ordem="valor" aria-sort="ascending">Valor</th>' in html


# ---------------------------------------------------------------- SQL: lista branca, vazios por último, pt-BR
def _numeros(bens):
    return [b["numero"] for b in bens]


def test_recorte_ordena_por_coluna_da_lista_branca(cliente):
    conn = db.conectar()
    todos = {"situacao": ""}
    assert _numeros(db.recorte(conn, todos)["bens"]) == [1001, 1002, 1003, 1004]
    assert _numeros(db.recorte(conn, todos, ordem="valor", direcao="desc")["bens"]) == [1002, 1004, 1001, 1003]
    assert _numeros(db.recorte(conn, todos, ordem="entrada")["bens"]) == [1001, 1002, 1003, 1004]  # 1996 antes de 2012
    # acento não empurra para o fim: ARMÁRIO < CADEIRA < MESA < NOTEBOOK
    assert _numeros(db.recorte(conn, todos, ordem="descricao")["bens"]) == [1004, 1001, 1003, 1002]
    # pessoa: só 1002 tem; vazios por último nos dois sentidos
    assert _numeros(db.recorte(conn, todos, ordem="pessoa")["bens"])[0] == 1002
    assert _numeros(db.recorte(conn, todos, ordem="pessoa", direcao="desc")["bens"])[0] == 1002
    # campo fora da lista (ou tentativa de SQL) cai no padrão, sem erro
    assert _numeros(db.recorte(conn, todos, ordem="b.numero DESC; DROP TABLE bens")["bens"]) == [1001, 1002, 1003, 1004]


def test_ordem_antes_do_limite_no_recorte(cliente):
    conn = db.conectar()
    r = db.recorte(conn, {"situacao": ""}, limite=2, ordem="valor", direcao="desc")
    assert r["truncado"] and _numeros(r["bens"]) == [1002, 1004]
