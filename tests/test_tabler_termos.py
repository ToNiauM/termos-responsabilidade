"""Telas de termos e pesquisa no Tabler (TERMOS_DESIGN=tabler): centro de custo, individual, devolução,
termo, registro emitido, pesquisa e bem."""
import pytest

import db


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


def _processo(cliente, tipo="ccusto", numero="2222"):
    cliente.post("/cadastros/processos/incluir", data={"tipo": tipo, "descricao": "T", "numero_sei": numero, "vigente": "1"})


def test_centro_custos_no_tabler(tabler):
    r = tabler.get("/centro-custos")
    html = r.text
    assert r.status_code == 200 and _e_tabler(html)
    assert "Termo por centro de custo" in html and 'action="/gerar"' in html and 'name="csrf"' in html
    assert '<select class="form-select" id="campo-ccusto" name="ccusto"' in html and '<option value="CCI">CCI</option>' in html
    assert 'id="tabela-centros"' in html and 'data-filtro-tabela="tabela-centros"' in html
    assert "JAQUELINE PORTELA" in html and 'href="/termo/ccusto/CCI"' in html and "sem termo" in html
    assert tabler.post("/gerar", data={"ccusto": "CCI"}).headers["Location"].endswith("/termo/ccusto/CCI")


def test_termos_individuais_no_tabler(tabler):
    html = tabler.get("/termos-individuais").text
    assert _e_tabler(html)
    assert "Termo individual" in html and 'action="/gerar-individual"' in html and 'id="campo-nome"' in html
    assert "ANA SILVA" in html and 'id="tabela-pessoas"' in html and 'aria-label="Termo de ANA SILVA"' in html


def test_termo_devolucao_no_tabler_fluxo(tabler):
    _processo(tabler, "devolucao", "3333")
    html = tabler.get("/termo_devolucao").text
    assert _e_tabler(html)
    assert 'id="form-devolucao"' in html and 'id="numero_bem"' in html and 'select[name="nome"]' in html
    assert 'id="tabela-devolucoes"' in html and 'aria-label="Devolução de ANA SILVA"' in html
    html = tabler.post("/termo_devolucao", data={"nome": "ANA SILVA"}, follow_redirects=True).text
    assert _e_tabler(html) and "Bens com ANA SILVA (sugestão)" in html and "Adicionar todos" in html
    assert '<option value="ANA SILVA" selected>' in html
    html = tabler.post("/termo_devolucao", data={"nome": "ANA SILVA", "todos": "1"}, follow_redirects=True).text
    assert "Bens a devolver" in html and "TOTAL" in html and 'name="gerar"' in html and 'name="limpar"' in html
    assert "(sugestão)" not in html
    html = tabler.post("/termo_devolucao", data={"nome": "ANA SILVA", "numero_bem": "9999"}, follow_redirects=True).text
    assert "Bem 9999 não encontrado" in html and "alert-danger" in html
    r = tabler.post("/termo_devolucao", data={"nome": "ANA SILVA", "gerar": "1"})
    assert r.headers["Location"].endswith("/termo/devolucao/ANA%20SILVA")


def test_termo_no_tabler_sem_e_com_processo(tabler):
    html = tabler.get("/termo/ccusto/CCI").text
    assert _e_tabler(html)
    assert "Sem processo SEI." in html and "alert-danger" in html and 'id="copiar"' not in html
    assert 'id="documento"' in html and 'src="/termo/ccusto/CCI/documento"' in html
    _processo(tabler)
    html = tabler.get("/termo/ccusto/CCI").text
    assert _e_tabler(html)
    assert 'id="copiar"' in html and 'data-registrar="/termo/ccusto/CCI/registrar"' in html
    assert 'href="/termo/ccusto/CCI/docx"' in html and "Baixar planilha" in html and "Emitir Termo no SEI" in html
    assert 'id="aviso-copiado"' in html and 'id="aviso-titulo"' in html and 'id="aviso-texto"' in html
    assert "Nenhum termo registrado para CCI." in html
    assert tabler.post("/termo/ccusto/CCI/registrar").status_code == 200
    assert "ver registro" in tabler.get("/termo/ccusto/CCI").text


def test_termo_emitido_no_tabler_salva_documento(tabler):
    _processo(tabler)
    tabler.get("/termo/ccusto/CCI/docx")
    html = tabler.get("/termos-emitidos/1").text
    assert _e_tabler(html)
    assert "Termos por centro de custo: CCI</h1>" in html
    assert "datagrid" in html and "Processo SEI" in html and "2222" in html
    assert 'action="/termos-emitidos/1/documento"' in html and 'id="numero_termo"' in html and 'id="bloco_sei"' in html
    assert 'id="tabela-foto"' in html and 'href="/bem?numero=1001"' in html
    tabler.post("/termos-emitidos/1/documento", data={"documento_sei": "1557099", "bloco_sei": "69766"})
    html = tabler.get("/termos-emitidos/1").text
    assert "1557099" in html and "bloco 69766" in html and 'id="enviar-email"' in html and "mailto:" in html
    assert 'id="registrar-email"' in html and "E-mail não enviado" in html
    tabler.post("/termos-emitidos/1/email")
    html = tabler.get("/termos-emitidos/1").text
    assert "Enviar email novamente" in html and "E-mail enviado em" in html and "badge bg-green-lt" in html


def test_pesquisa_no_tabler(tabler):
    html = tabler.get("/pesquisa").text
    assert _e_tabler(html) and "Digite no campo de pesquisa" in html
    html = tabler.get("/pesquisa?q=cci").text
    assert _e_tabler(html) and "Centros de custo" in html and 'href="/pesquisa?ccusto=CCI"' in html
    html = tabler.get("/pesquisa?q=ana").text
    assert "ANA SILVA" in html and "/pesquisa?pessoa=ANA" in html and "Termo individual" in html
    html = tabler.get("/pesquisa?q=zzz").text
    assert "Nada encontrado." in html and "alert-info" in html
    html = tabler.get("/pesquisa?ccusto=CCI").text
    assert _e_tabler(html) and "Bens de CCI" in html and "Termo do centro de custo" in html
    assert 'id="tabela-bens"' in html and 'href="/bem?numero=1001"' in html


def test_bem_no_tabler(tabler):
    html = tabler.get("/bem?numero=1001").text
    assert _e_tabler(html)
    assert "Bem 1001" in html and "CADEIRA" in html and "datagrid" in html and "Responsável individual" in html
    assert "Histórico" in html and "Nenhuma mudança registrada." in html and "Nenhum termo registrado com este bem." in html
    assert "Fotos do invent" not in html


def test_bem_no_tabler_mostra_fotos_por_evento(tabler):
    from tests.test_app import _abrir
    import inventario
    eid = _abrir(tabler)
    tabler.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    conn = db.conectar()
    inventario.adicionar_foto(conn, eid, 1001, lambda c: "https://x/a.webp")
    inventario.adicionar_foto(conn, eid, 1001, lambda c: "https://x/b.webp")
    html = tabler.get("/bem?numero=1001").text
    assert _e_tabler(html) and "Fotos do invent" in html and html.count('class="astra-miniatura"') == 2
    assert html.index("https://x/a.webp") < html.index("https://x/b.webp") and "lido em" in html
