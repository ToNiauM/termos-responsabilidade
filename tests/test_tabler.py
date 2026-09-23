"""Protótipo Tabler (branch tabler): com ASTRA_UI=tabler, templates/tabler/ tem precedência sobre templates/.

As telas migradas herdam tabler_base.html; as demais continuam no base.html do DSGov, intactas.
"""
import pytest

from tests.conftest import logar


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
    return "tabler/vendor/tabler.min.css" in html and "govbr-ds/core.min.css" not in html and 'class="br-' not in html


def test_sem_a_chave_tudo_continua_no_dsgov(cliente):
    html = cliente.get("/").text
    assert "govbr-ds/core.min.css" in html and "tabler.min.css" not in html


def test_inicio_no_tabler_mantem_atalhos_indicadores_e_permissoes(tabler, usuarios_exemplo):
    html = tabler.get("/").text
    assert _e_tabler(html)
    assert html.count("data-atalho") == 6 and 'href="/inventario"' in html
    assert "bens ativos" in html and "sem centro nem pessoa" in html and "Nenhuma importação registrada" in html
    assert "Sobre o patrimônio" in html and "Ver guia" in html
    assert "Fulano" in html and 'action="/sair"' in html and 'name="csrf"' in html   # menu do usuário com o POST de sair

    tabler.post("/sair"); logar(tabler, *usuarios_exemplo["consulta"])
    html = tabler.get("/").text
    assert html.count("data-atalho") == 5 and "Realizar inventário" not in html
    assert "Nenhuma importação registrada" not in html


def test_menu_lateral_do_tabler_marca_a_tela_atual(tabler):
    html = tabler.get("/termos-emitidos").text
    trecho = html.split('id="menu-termos"')[1].split("</ul>")[0]
    assert 'aria-current="page"' in trecho and "Termos emitidos" in trecho
    assert 'class="collapse show" id="menu-termos"' in html     # o grupo da tela atual abre expandido
    assert 'class="collapse" id="menu-cadastros"' in html


def test_termos_emitidos_no_tabler_lista_e_filtra(tabler):
    tabler.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "T", "numero_sei": "2222", "vigente": "1"})
    tabler.get("/termo/ccusto/CCI/docx")
    html = tabler.get("/termos-emitidos").text
    assert _e_tabler(html)
    assert "CCI" in html and "2222" in html and 'class="form-select"' in html and 'data-filtro-tabela' in html
    assert "CCI" not in tabler.get("/termos-emitidos?tipo=individual").text.split("<tbody>")[1]
    assert "Nenhum termo registrado com esse filtro" in tabler.get("/termos-emitidos?tipo=individual").text


def test_formulario_de_cadastro_no_tabler_mostra_erros_e_preserva_valores(tabler):
    html = tabler.get("/cadastros/pessoas/novo").text
    assert _e_tabler(html) and html.count("<h1") == 1 and 'class="form-control' in html
    r = tabler.post("/cadastros/pessoas/incluir", data={"nome": "ana silva"})
    assert "Esta pessoa já está cadastrada" in r.text and 'value="ana silva"' in r.text
    assert "is-invalid" in r.text and "invalid-feedback" in r.text
    html = tabler.get("/cadastros/processos/novo").text
    assert '<select class="form-select' in html and 'name="tipo"' in html and 'type="checkbox"' in html


def test_tela_nao_migrada_continua_no_dsgov_com_a_chave_ligada(tabler):
    html = tabler.get("/ajuda").text
    assert "govbr-ds/core.min.css" in html and "tabler.min.css" not in html
