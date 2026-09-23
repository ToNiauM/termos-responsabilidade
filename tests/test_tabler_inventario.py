"""Telas do inventário no protótipo Tabler (TERMOS_DESIGN=tabler): mesmos textos, ids e fluxos do DSGov, sem br-*."""

import pytest

from tests.conftest import SENHA_PADRAO, logar


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


def _ids(*nomes):
    import db
    conn = db.conectar()
    return [conn.execute("SELECT id FROM usuarios WHERE nome = ?", (n,)).fetchone()[0] for n in nomes]


def _abrir(cliente, comissao=("Fulano", "Beltrana")):
    cliente.post("/inventario/abrir", data={"nome": "Inv", "usuarios": _ids(*comissao), "escopo": "todas"})
    import db, inventario
    return inventario.evento_aberto(db.conectar())["id"]


def test_eventos_no_tabler(tabler):
    html = tabler.get("/inventario").text
    assert _e_tabler(html) and "Nenhum invent" in html and 'href="/administracao"' in html
    eid = _abrir(tabler)
    html = tabler.get("/inventario").text
    assert _e_tabler(html)
    assert "Salas e leitura" in html and "Criado em" in html and ">aberto<" in html and 'class="progress' in html
    assert f'href="/inventario/{eid}/relatorio"' in html
    tabler.post(f"/inventario/{eid}/fechar")
    html = tabler.get("/inventario").text
    assert ">fechado<" in html and "Inventário fechado." in html and 'class="alert alert-warning' in html


def test_evento_no_tabler(tabler):
    eid = _abrir(tabler)
    html = tabler.get(f"/inventario/{eid}").text
    assert _e_tabler(html)
    corpo = html.split("<tbody>")[1]
    assert "01 - SALA CCI" in corpo and "99 - SEM MAPA" in corpo and "não iniciada" in corpo
    assert "Conferir sala 01 - SALA CCI" in html and f'href="/inventario/{eid}/painel"' in html
    assert 'data-filtro-tabela="tabela-salas"' in html and "table-responsive" in html
    tabler.post(f"/inventario/{eid}/fechar")
    html = tabler.get(f"/inventario/{eid}").text
    assert "Inventário fechado." in html and "Ver sala 01 - SALA CCI" in html and "Conferir sala" not in html


def test_sala_no_tabler_leitura_lote_e_fotos(tabler, monkeypatch):
    import db, fotos, inventario
    for v in fotos.VARIAVEIS:
        monkeypatch.setenv(v, "x")
    eid = _abrir(tabler)
    url = f"/inventario/{eid}/sala/01 - SALA CCI"
    html = tabler.get(url).text
    assert _e_tabler(html)
    assert 'id="leitura" type="text" inputmode="none" autocomplete="off" enterkeyhint="done" placeholder="Aproxime o leitor…" autofocus' in html
    assert 'id="btn-camera"' in html and 'id="btn-digitar"' in html and "html5-qrcode.min.js" in html and 'id="leitor-camera"' in html
    assert 'id="aviso"' in html and 'id="aviso-texto"' in html and 'id="btn-sobra"' in html and 'id="form-sobra"' in html
    assert 'id="n-localizados"' in html and 'id="contadores" data-total="2"' in html
    assert 'id="form-lote"' in html and 'data-papel="selecao"' in html and 'id="selecionar-todos"' in html
    assert 'id="lista-bens"' in html and 'class="list-group list-group-flush"' in html and 'id="grade-bens"' in html
    assert "<style" not in html.split("<body")[1] and ' style="' not in html   # só classes nativas do Tabler
    assert "tabler-icons.min.css" in html and "ti ti-map-pin" in html and "ti ti-dots-vertical" in html
    assert "dsgov.js" not in html and "core.min.js" not in html
    # leitura e foto pela API; a tela volta com a foto no avatar, o selo de fotos extras e o excluir no menu
    assert tabler.post(f"{url}/ler", json={"numero": "1001"}).get_json()["situacao"] == "localizado"
    inventario.adicionar_foto(db.conectar(), eid, 1001, lambda c: "https://x/1.webp")
    inventario.adicionar_foto(db.conectar(), eid, 1001, lambda c: "https://x/2.webp")
    html = tabler.get(url).text
    item = html.split('id="lista-bens"')[1].split('data-bem="1001"')[1].split('data-bem="')[0]
    assert item.count('class="avatar avatar-lg avatar-square object-cover"') == 1 and "+1 foto(s)" in item and 'aria-label="Excluir foto 1"' in item
    assert 'class="foto-input" hidden/>' in item and ">Localizado<" in item and "sem conservação" in item
    assert "JAQUELINE PORTELA" in item and "ti ti-armchair" not in item   # responsável do centro; tem foto, sem ícone
    item2 = html.split('id="lista-bens"')[1].split('data-bem="1002"')[1].split('data-bem="')[0]
    assert 'class="foto-input" hidden disabled' in item2 and ">Pendente<" in item2 and "ti ti-package" in item2   # classificação genérica
    assert 'data-acao="desfazer" hidden' in item2
    r = tabler.post(f"{url}/lote", data={"acao": "marcar", "numeros": ["1002"]}, follow_redirects=True)
    assert "1 bem(ns) marcado(s)" in r.text and r.text.count(">Localizado<") == 4 and _e_tabler(r.text)   # lista + grade
    # sobra
    for v in fotos.VARIAVEIS:
        monkeypatch.delenv(v, raising=False)
    r = tabler.post(f"{url}/sobra", data={"descricao": "VENTILADOR", "observacao": "sem plaqueta"}, follow_redirects=True)
    assert "VENTILADOR" in r.text and 'aria-label="Excluir sobra"' in r.text and "Fotos desativadas" in r.text
    # encerrado: somente consulta
    tabler.post(f"/inventario/{eid}/encerrar", data={"confirmar": "1"})
    html = tabler.get(url).text
    assert "Evento encerrado: somente consulta." in html and 'id="form-lote"' not in html and 'placeholder="Aproxime o leitor…" disabled' in html


def test_sala_somente_consulta_fora_da_comissao_no_tabler(tabler):
    eid = _abrir(tabler, comissao=("Beltrana",))
    html = tabler.get(f"/inventario/{eid}/sala/01 - SALA CCI").text
    assert _e_tabler(html) and "Somente consulta" in html and 'id="form-lote"' not in html and 'id="form-sobra"' not in html
    assert "não faz parte da comissão" in html
    tabler.post("/sair"); logar(tabler, "beltrana", SENHA_PADRAO)
    html = tabler.get(f"/inventario/{eid}/sala/01 - SALA CCI").text
    assert "lendo como <strong>Beltrana</strong>" in html and 'id="form-lote"' in html


def test_painel_no_tabler(tabler):
    eid = _abrir(tabler)
    tabler.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    html = tabler.get(f"/inventario/{eid}/painel").text
    assert _e_tabler(html)
    assert 'id="g-situacao"' in html and "bens no escopo" in html and "1 (33.3%)" in html
    assert "echarts.min.js" in html and "echarts-dsgov.js" in html and "data-grafico" in html
    html = tabler.get(f"/inventario/{eid}/painel?andar=01").text
    assert 'id="g-salas"' in html and "todos os andares" in html


def test_relatorio_no_tabler_filtros_modal_e_xlsx(tabler):
    import db, inventario
    eid = _abrir(tabler)
    tabler.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    inventario.adicionar_foto(db.conectar(), eid, 1001, lambda c: "https://x/1001.webp")
    inventario.adicionar_foto(db.conectar(), eid, 1001, lambda c: "https://x/1001b.webp")
    html = tabler.get(f"/inventario/{eid}/relatorio?ordem=numero&dir=desc").text
    assert _e_tabler(html)
    assert 'class="form-select"' in html and 'id="campo-situacao"' in html and 'name="fotos"' in html and 'id="busca"' in html
    assert 'name="ordem" value="numero"' in html and "fa-sort-down" in html
    assert 'data-foto="https://x/1001.webp"' in html and 'data-bs-toggle="modal" data-bs-target="#scrim-foto"' in html
    assert 'class="modal fade" id="scrim-foto"' in html and 'id="modal-foto-img"' in html and 'id="modal-foto-link"' in html
    assert ">+1<" in html and ">Localizado<" in html
    corpo = tabler.get(f"/inventario/{eid}/relatorio?situacao=localizado").text.split("<tbody>")[1]
    assert ">1001<" in corpo and ">1002<" not in corpo
    r = tabler.get(f"/inventario/{eid}/xlsx?fotos=1")
    assert r.status_code == 200 and r.headers["Content-Disposition"].endswith(".xlsx")


def test_comissao_no_tabler(tabler):
    fulano, beltrana = _ids("Fulano", "Beltrana")
    eid = _abrir(tabler, comissao=("Beltrana",))
    html = tabler.get(f"/inventario/{eid}/comissao").text
    assert _e_tabler(html) and "Comissão de Inv" in html
    assert f'value="{beltrana}" checked' in html and f'value="{fulano}"/' in html and 'class="form-check-input"' in html
    r = tabler.post(f"/inventario/{eid}/comissao", data={"usuarios": [fulano, beltrana]}, follow_redirects=True)
    assert "Comissão atualizada" in r.text


def test_excluir_no_tabler(tabler):
    import db, inventario
    eid = _abrir(tabler)
    tabler.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    html = tabler.get(f"/inventario/{eid}/excluir").text
    assert _e_tabler(html) and "1 leitura" in html and 'name="nome"' in html and "Excluir definitivamente" in html
    assert "Esta ação não tem desfazer." in html and 'class="alert alert-danger' in html
    r = tabler.post(f"/inventario/{eid}/excluir", data={"nome": "Errado"}, follow_redirects=True)
    assert "não confere" in r.text
    r = tabler.post(f"/inventario/{eid}/excluir", data={"nome": "Inv"}, follow_redirects=True)
    assert "Evento Inv excluído" in r.text and inventario.evento(db.conectar(), eid) is None
