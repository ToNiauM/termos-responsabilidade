"""Protótipo Tabler: Cadastros, Análise, Atualizar base (upload/importação) e Textos.

Cada tela abre no tema Tabler (sem DSGov nem classes br-*) e mantém ids, textos e fluxos do original.
"""
import io
import re

import pytest

import db
from tests.conftest import confirmar_revisao, logar


@pytest.fixture
def tabler(cliente):
    """Liga o carregador do Tabler no app já criado e limpa o cache do Jinja nas duas pontas."""
    import app as modulo
    original = modulo.app.jinja_loader
    modulo.app.jinja_loader = modulo.carregador_templates(tabler=True)
    modulo.app.jinja_env.cache.clear()
    yield cliente
    modulo.app.jinja_loader = original
    modulo.app.jinja_env.cache.clear()


def _e_tabler(html):
    return ("tabler/vendor/tabler.min.css" in html and "govbr-ds/core.min.css" not in html
            and 'class="br-' not in html and "dsgov-" not in html.replace("dsgov/", ""))


def _abrir(cliente, url):
    r = cliente.get(url)
    assert r.status_code == 200, url
    assert _e_tabler(r.text), url
    return r.text


# ---------------------------------------------------------------- cadastros
@pytest.mark.parametrize("aba", ["responsaveis", "localizacoes", "pessoas", "processos"])
def test_cadastros_abas_no_tabler(tabler, aba):
    html = _abrir(tabler, f"/cadastros/{aba}")
    assert 'id="busca-cadastro"' in html and 'id="filtros-cadastros"' in html
    assert 'aria-label="Áreas de cadastro"' in html and html.count('aria-current="page"') >= 1
    assert "registro(s) encontrado(s)" in html and "Baixar todos os cadastros" in html
    assert 'aria-label="Paginação dos cadastros"' in html and 'class="pagination' in html
    assert 'class="form-select"' in html and 'name="por_pagina"' in html
    assert html.count('<h1 class="page-title">') == 1


def test_cadastros_responsaveis_botao_novo_e_acoes(tabler):
    html = _abrir(tabler, "/cadastros/responsaveis")
    assert "Novo centro de custo" in html and "cadastrar nova pessoa" not in html.lower()
    assert "CCI" in html and "Editar" in html and 'action="/cadastros/responsaveis/excluir"' in html
    assert 'name="ccustos" value="CCI"' in html and 'name="retorno"' in html


def test_cadastros_localizacoes_selecao_em_lote(tabler):
    html = _abrir(tabler, "/cadastros/localizacoes")
    for id_ in ('id="mover-localizacoes"', 'id="contagem-selecao"', 'id="revisar-lote"', 'id="selecionar-pagina"'):
        assert id_ in html
    assert 'data-parent="localizacoes"' in html and 'data-child="localizacoes"' in html
    assert 'form="mover-localizacoes"' in html and 'data-selecao-info="localizacoes"' in html
    assert "99 - SEM MAPA" in html and "Sem centro de custo" in html and "Resolver vínculos" in html
    assert "0 localizações selecionadas" in html
    # o script de seleção não depende mais do .br-select
    assert ".br-select" not in html and "localização(ões) selecionada(s) nesta página" in html


def test_cadastros_busca_e_paginacao_no_tabler(tabler, dados):
    for n in range(25):
        db.incluir_pessoa(dados, f"PESSOA {n:02}")
    html = _abrir(tabler, "/cadastros/pessoas?q=pessoa&por_pagina=10&pagina=2")
    assert "11–20 de 25 registros" in html and "PESSOA 10" in html and "PESSOA 00" not in html
    assert 'aria-label="Página anterior"' in html and 'aria-label="Próxima página"' in html
    assert '<option value="10" selected>' in html


def test_cadastros_processos_cards_e_fluxo(tabler, dados):
    html = _abrir(tabler, "/cadastros/processos")
    assert "Processos SEI" in html and "Sem processo vigente" in html
    r = tabler.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "Termos 2026", "numero_sei": "2222", "vigente": "1"},
                    follow_redirects=True)
    assert _e_tabler(r.text) and "Termos 2026" in r.text and "Vigente" in r.text
    pid = db.processos(dados)[0]["id"]
    r = tabler.post("/cadastros/processos/encerrar", data={"id": pid}, follow_redirects=True)
    assert _e_tabler(r.text) and 'name="revisao"' in r.text and "alert alert-warning" in r.text
    r = confirmar_revisao(tabler, r, "/cadastros/processos/encerrar")
    assert "Encerrado" in r.text


def test_cadastros_pessoa_detalhe_e_atribuicao(tabler, dados):
    html = _abrir(tabler, "/cadastros/pessoas?nome=ANA SILVA")
    assert "ANA SILVA" in html and "Atribuir um patrimônio" in html and 'id="numero-patrimonio"' in html
    assert "Voltar à lista de pessoas" in html and "Ver termo individual" in html and "Excluir pessoa" in html
    assert "NOTEBOOK" in html and 'action="/cadastros/pessoas/desatribuir"' in html
    db.incluir_pessoa(dados, "BRUNO LIMA")
    r = tabler.post("/cadastros/pessoas/atribuir", data={"nome": "BRUNO LIMA", "numero": "1002"})
    assert _e_tabler(r.text) and "Confira os registros afetados" in r.text and "ANA SILVA" in r.text
    confirmar_revisao(tabler, r, "/cadastros/pessoas/atribuir")
    assert db.pessoa_do_bem(dados, 1002) == "BRUNO LIMA"


def test_cadastros_atribuir_patrimonio_invalido(tabler):
    r = tabler.post("/cadastros/pessoas/atribuir", data={"nome": "ANA SILVA", "numero": "9999"})
    assert r.status_code == 200 and _e_tabler(r.text)
    assert 'value="9999"' in r.text and 'id="erro-numero"' in r.text and "não encontrado" in r.text
    assert "is-invalid" in r.text and "Consultar patrimônio" in r.text


def test_cadastros_mover_com_destino_invalido_e_confirmacao(tabler, dados):
    r = tabler.post("/cadastros/localizacoes/mover", data={"localizacoes": ["01 - SALA CCI"], "ccustos_destino": ""})
    assert _e_tabler(r.text)
    assert 'id="erro-ccustos_destino"' in r.text and 'name="localizacoes" value="01 - SALA CCI"' in r.text
    assert "Revisar alteração" in r.text and "Centro atual: CCI" in r.text
    db.incluir_responsavel(dados, {"ccustos": "GEX", "responsavel": "MARIA"})
    r = tabler.post("/cadastros/localizacoes/mover", data={"localizacoes": ["01 - SALA CCI"], "ccustos_destino": "GEX"})
    assert _e_tabler(r.text)
    r = confirmar_revisao(tabler, r, "/cadastros/localizacoes/mover")
    assert "movida(s) para GEX" in r.text
    html = _abrir(tabler, "/cadastros/localizacoes/alterar?localizacao=01 - SALA CCI")
    assert "Alterar centro da localização" in html


def test_cadastros_exclusao_bloqueada(tabler):
    r = tabler.post("/cadastros/responsaveis/excluir", data={"ccustos": "CCI"})
    assert _e_tabler(r.text) and "alert alert-danger" in r.text
    assert "Voltar aos cadastros" in r.text and "Gerenciar localizações" in r.text and 'name="revisao"' not in r.text


# ---------------------------------------------------------------- análise
def test_analise_no_tabler_cards_filtros_e_tabela(tabler, dados):
    dados.execute("INSERT INTO bens VALUES (9001,'ATIVO','SEM VALOR','','MÓVEIS','01 - SALA CCI','01/01/2020',NULL,NULL)")
    dados.execute("INSERT INTO bens VALUES (9002,'ATIVO','VALOR ZERO','','MÓVEIS','01 - SALA CCI','01/01/2020',0,0)")
    dados.commit()
    html = _abrir(tabler, "/analise")
    assert html.count("astra-kpi") == 6
    hrefs = re.findall(r'<a class="card card-link h-100 astra-kpi" href="([^"]*)">', html)
    assert hrefs and all(h.startswith("/analise?") for h in hrefs)
    assert 'name="valor_status"' in html and 'value="nao_informado"' in html and "Valor zero" in html
    assert "Exportar .xlsx" in html and "data-grafico=" in html and "echarts-dsgov.js" in html
    assert 'data-filtro-tabela="tabela-bens"' in html
    assert "Não informado" in re.search(r"<tr>.*?9001.*?</tr>", html).group()
    assert "R$ 0,00" in re.search(r"<tr>.*?9002.*?</tr>", html).group()


def test_analise_termo_completo_no_tabler(tabler):
    html = _abrir(tabler, "/analise?ccusto=CCI&classificacao=MÓVEIS")
    assert "Abrir termo completo" in html and "/termo/ccusto/CCI" in html
    assert "Os filtros desta análise não limitam o termo" in html and "badge" in html


# ---------------------------------------------------------------- atualizar base / importação
def _xlsx(linhas):
    from openpyxl import Workbook
    from tests.test_app import CABECALHO
    wb = Workbook()
    ws = wb.active
    ws.append(CABECALHO)
    for l in linhas:
        ws.append(l)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def test_upload_no_tabler_carrega_e_lista(tabler):
    html = _abrir(tabler, "/upload")
    assert html.count("data-confirmar=") == 2 and 'id="arquivo-bens"' in html and 'id="arquivo-cadastros"' in html
    assert 'type="file"' in html and 'class="form-control"' in html and "Baixar" in html and "Carregar" in html
    buf = _xlsx([[1002, "ATIVO", "NOTEBOOK", "DELL", "EQ", "01 - SALA CCI", "x", 1, 1],
                 [5000, "ATIVO", "TV", "", "EQ", "77 - NOVA SALA", "x", 1, 1]])
    r = tabler.post("/upload", data={"arquivo": (buf, "export.xlsx")}, content_type="multipart/form-data", follow_redirects=True)
    assert _e_tabler(r.text) and "2 bens importados (2 ativos)" in r.text and "77 - NOVA SALA" in r.text
    assert "Últimas cargas de bens" in r.text and "export.xlsx" in r.text


def test_importacao_no_tabler(tabler, dados):
    iid = db.importar_bens(dados, _xlsx([[1001, "ATIVO", "CADEIRA", "", "MÓVEIS", "02 - OUTRA", "x", 1, 1],
                                         [1002, "ATIVO", "NOTEBOOK", "DELL", "EQ", "01 - SALA CCI", "x", 1, 1]]),
                           nome_arquivo="export.xlsx")["importacao_id"]
    html = _abrir(tabler, f"/importacoes/{iid}")
    assert "Importação de " in html and "Mudanças" in html and "02 - OUTRA" in html and "removido" in html
    assert "Situação alterada" in html and "badge bg-red-lt" in html


def test_upload_execucoes_do_robo_no_tabler(tabler, dados):
    resumo = db.importar_bens(dados, _xlsx([[1002, "ATIVO", "NOTEBOOK", "DELL", "EQ", "01 - SALA CCI", "x", 1, 1]]), nome_arquivo="SPW automático")
    db.registrar_execucao_robo(dados, "2026-09-18 04:00:00", "importado", hash="h", importacao_id=resumo["importacao_id"], mensagem="1 bens")
    db.registrar_execucao_robo(dados, "2026-09-20 04:00:00", "erro", mensagem="SPW fora do ar")
    html = _abrir(tabler, "/upload")
    assert "Atualizações com o SPW" in html and "SPW fora do ar" in html and "bg-danger" in html
    assert f'href="/importacoes/{resumo["importacao_id"]}"' in html.split("Atualizações com o SPW")[-1]


def test_upload_atualizar_com_spw_no_tabler(tabler, dados, chave):
    import usuarios
    from tests.conftest import ADMIN_LOGIN
    usuarios.salvar_acesso_spw(dados, usuarios.por_login(dados, ADMIN_LOGIN)["id"], "antonio.junior", "S3nha!")
    html = _abrir(tabler, "/upload")
    assert "Atualizar com SPW" in html and 'action="/atualizar-base/spw"' in html and 'http-equiv="refresh"' not in html
    tabler.post("/atualizar-base/spw")
    html = _abrir(tabler, "/upload")
    assert "Atualizando com o SPW" in html and 'http-equiv="refresh" content="5"' in html and "disabled" in html


def test_upload_sem_atualizar_spw_para_operador(tabler, usuarios_exemplo):
    tabler.post("/sair")
    logar(tabler, *usuarios_exemplo["operador"])
    r = tabler.get("/upload")
    if r.status_code == 200:
        assert "Atualizar com SPW" not in r.text


# ---------------------------------------------------------------- textos
def test_textos_no_tabler_salva_e_restaura(tabler):
    import textos
    html = _abrir(tabler, "/textos")
    assert "Textos dos termos" in html and "Compromissos" in html and 'id="t-individual_abertura"' in html
    assert "<textarea class=\"form-control\"" in html and "Restaurar padrão" not in html
    r = tabler.post("/textos", data=dict(textos.PADRAO, individual_abertura="TESTE {nome}."), follow_redirects=True)
    assert _e_tabler(r.text) and "Textos salvos" in r.text and "Restaurar padrão" in r.text
    assert 'name="restaurar" value="individual_abertura"' in r.text
    r = tabler.post("/textos", data={"restaurar": "individual_abertura"}, follow_redirects=True)
    assert "Padrão restaurado" in r.text and "Restaurar padrão" not in r.text


def test_textos_marcador_invalido_no_tabler(tabler):
    import textos
    r = tabler.post("/textos", data=dict(textos.PADRAO, individual_abertura="Eu {nomee}"), follow_redirects=True)
    assert _e_tabler(r.text) and "nomee" in r.text
