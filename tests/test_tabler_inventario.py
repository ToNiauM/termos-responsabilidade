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
    for secao in ("pendentes", "divergentes", "localizados"):   # três seções, cada uma com lista (celular) e grade (desktop)
        assert f'id="lista-{secao}"' in html and f'id="grade-{secao}"' in html and f'id="secao-{secao}"' in html
    assert 'class="list-group list-group-flush"' in html
    assert "<style" not in html.split("<body")[1] and ' style="' not in html   # só classes nativas do Tabler
    assert "tabler-icons.min.css" in html and "ti ti-map-pin" in html and "ti ti-dots-vertical" in html
    assert "dsgov.js" not in html and "core.min.js" not in html
    # leitura e foto pela API; a tela volta com a foto no avatar, o selo de fotos extras e o excluir no menu
    assert tabler.post(f"{url}/ler", json={"numero": "1001"}).get_json()["situacao"] == "localizado"
    inventario.adicionar_foto(db.conectar(), eid, 1001, lambda c: "https://x/1.webp")
    inventario.adicionar_foto(db.conectar(), eid, 1001, lambda c: "https://x/2.webp")
    html = tabler.get(url).text
    item = html.split('id="lista-localizados"')[1].split('data-bem="1001"')[1].split('data-bem="')[0]   # Lista: sem foto
    assert "<img" not in item and "avatar" not in item and "+1 foto(s)" in item   # excluir foto: na ficha (modal)
    assert 'class="foto-input" hidden/>' in item and ">Localizado<" in item and "sem conservação" in item and "JAQUELINE PORTELA" in item
    quadro = html.split('id="grade-localizados"')[1].split('data-bem="1001"')[1].split('data-bem="')[0]   # Quadros: foto no topo
    assert quadro.count('<img class="card-img-top object-cover" src="https://x/1.webp"') == 1 and "+1 foto(s)" in quadro
    assert "JAQUELINE PORTELA" in quadro and "ti ti-armchair" not in quadro   # responsável do centro; tem foto, sem ícone
    item2 = html.split('id="grade-pendentes"')[1].split('data-bem="1002"')[1].split('data-bem="')[0]
    assert 'class="foto-input" hidden disabled' in item2 and ">Pendente<" in item2 and "ti ti-package" in item2   # sem foto: ícone genérico
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
    assert "echarts.min.js" in html and "tabler/graficos.js" in html and "data-grafico" in html
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


def _secao(html, nome):
    """Trecho da seção (aba) `nome` da sala virtual: do painel até o próximo painel ou os modelos."""
    return html.split(f'id="secao-{nome}"')[1].split('data-secao-painel="')[1].split('<template id="modelo-lista"')[0].split('id="secao-')[0]


def _fichas(html):
    import json
    return json.loads(html.split('<script type="application/json" id="fichas">')[1].split("</script>")[0])


def _cenario_sala(tabler, comissao=("Fulano", "Beltrana")):
    """1001 lido aqui (localizado); 1002 lido em outra sala (pendente aqui); 1004 de outra sala lido aqui
    (divergente); 1003 não ativo lido aqui (em Divergentes, fora da contagem)."""
    eid = _abrir(tabler, comissao)
    url = f"/inventario/{eid}/sala/01 - SALA CCI"
    assert tabler.post(f"{url}/ler", json={"numero": "1001"}).get_json()["situacao"] == "localizado"
    assert tabler.post(f"/inventario/{eid}/sala/99 - SEM MAPA/ler", json={"numero": "1002"}).status_code == 200
    assert tabler.post(f"{url}/ler", json={"numero": "1004"}).get_json()["situacao"] == "divergente"
    assert tabler.post(f"{url}/ler", json={"numero": "1003"}).status_code == 200
    return eid, url


def test_sala_tres_secoes_com_contagens(tabler):
    eid, url = _cenario_sala(tabler)
    html = tabler.get(url).text
    assert _e_tabler(html) and ' style="' not in html and "<style" not in html.split("<body")[1]
    # cards-aba com os mesmos números de inventario.salas; Pendentes aberta por padrão; alerta em Divergentes
    assert 'id="n-pendentes">1<' in html and 'id="n-divergentes">1<' in html and 'id="n-localizados">1<' in html
    assert 'role="tablist"' in html and 'aria-controls="secao-pendentes" aria-selected="true"' in html
    assert 'id="secao-divergentes" role="tabpanel" aria-labelledby="aba-divergentes" data-secao-painel="divergentes" hidden' in html
    assert 'class="card-status-top bg-warning" data-papel="alerta-divergentes">' in html
    pend, div, loc = _secao(html, "pendentes"), _secao(html, "divergentes"), _secao(html, "localizados")
    # bem da sala lido em outra sala: continua Pendente, com "Encontrado em" e quem leu
    assert 'data-bem="1002"' in pend and "Encontrado em 99 - SEM MAPA" in pend and "Lido por Fulano em" in pend and ">Pendente<" in pend
    assert 'data-bem="1002"' not in div + loc and "Divergente<" not in pend
    # bem de outra sala lido aqui: Divergentes, com a localização cadastrada (que não muda)
    assert 'data-bem="1004"' in div and "Cadastrado em 99 - SEM MAPA" in div and ">Divergente<" in div and "Lido aqui por Fulano" in div
    assert 'data-bem="1003"' in div and ">Não ativo<" in div and "BAIXADO" in div   # não ativo lido aqui
    assert 'data-bem="1001"' in loc and ">Localizado<" in loc and 'data-bem="1001"' not in pend + div
    # número do bem leva ao cadastro, na mesma aba
    for n in (1001, 1002, 1003, 1004):
        assert f'<a class="fw-bold" href="/bem?numero={n}" data-papel="numero"' in html
    assert 'target="_blank"' not in html.split('data-papel="numero"')[1][:80]
    # a localização cadastrada nunca muda com a leitura
    import db
    assert db.buscar_bem(db.conectar(), 1004)["localizacao"] == "99 - SEM MAPA"


def test_sala_ficha_do_modal_com_fotos(tabler, monkeypatch):
    import db, fotos, inventario
    for v in fotos.VARIAVEIS:
        monkeypatch.setenv(v, "x")
    eid, url = _cenario_sala(tabler)
    inventario.adicionar_foto(db.conectar(), eid, 1001, lambda c: "https://x/1.webp")
    inventario.adicionar_foto(db.conectar(), eid, 1001, lambda c: "https://x/2.webp")
    inventario.adicionar_foto(db.conectar(), eid, 1004, lambda c: "https://x/3.webp")
    html = tabler.get(url).text
    f = _fichas(html)
    assert [(x["nfoto"], x["url"]) for x in f["1001"]["fotos"]] == [(1, "https://x/1.webp"), (2, "https://x/2.webp")]
    assert f["1001"]["url_bem"] == "/bem?numero=1001" and f["1001"]["secao"] == "localizado" and f["1001"]["lido"]
    assert f["1001"]["lido_em_sala"] == "01 - SALA CCI" and f["1001"]["integrante"] == "Fulano"
    assert len(f["1001"]["lido_em"]) == 19 and f["1001"]["lido_em"][2] == "/" and f["1001"]["lido_em"][5] == "/"   # dd/mm/aaaa hh:mm:ss
    assert f["1001"]["ccustos"] == "CCI" and f["1001"]["responsavel_centro"] == "JAQUELINE PORTELA" and f["1001"]["valor_atual"] == "R$ 64,54"
    assert f["1002"]["secao"] == "pendente" and f["1002"]["lido_em_sala"] == "99 - SEM MAPA" and f["1002"]["pessoa"] == "ANA SILVA"
    assert f["1004"]["secao"] == "divergente" and [x["url"] for x in f["1004"]["fotos"]] == ["https://x/3.webp"]
    assert f["1003"]["situacao"] == "BAIXADO" and f["1003"]["fotos"] == []
    # modal: campos da página do bem, bloco da leitura, galeria e ações
    modal = html.split('id="modal-bem"')[1].split("</form>")[0]
    for rotulo in ("Descrição", "Complemento", "Classificação", "Situação", "Localização", "Data de entrada", "Valor atual",
                   "Centro de custo", "Responsável individual", "Sala da leitura", "Lido por", "Data/hora da leitura"):
        assert f'<div class="datagrid-title">{rotulo}</div>' in modal
    assert 'id="det-link"' in modal and 'id="det-fotos"' in modal and 'id="det-foto-input"' in modal
    assert 'id="det-conservacao"' in modal and 'id="det-quem-usa"' in modal and 'id="det-observacao"' in modal
    assert 'id="det-ler"' in modal and 'id="det-desfazer"' in modal
    # /ler devolve a ficha para a tela atualizar sem recarregar
    j = tabler.post(f"{url}/ler", json={"numero": "1002"}).get_json()
    assert j["ficha"]["secao"] == "localizado" and j["ficha"]["lido_em_sala"] == "01 - SALA CCI" and j["ficha"]["url_bem"] == "/bem?numero=1002"


def test_foto_excluir_responde_json(tabler, monkeypatch):
    import db, fotos, inventario
    eid, url = _cenario_sala(tabler, comissao=("Fulano",))
    inventario.adicionar_foto(db.conectar(), eid, 1001, lambda c: "https://x/1.webp")
    inventario.adicionar_foto(db.conectar(), eid, 1001, lambda c: "https://x/2.webp")
    apagadas = []
    monkeypatch.setattr(fotos, "apagar", apagadas.append)   # bucket falso: nada real é apagado
    json_ = {"Accept": "application/json"}
    r = tabler.post(f"/inventario/{eid}/leitura/1001/foto/1/excluir", headers=json_)
    assert r.status_code == 200 and r.get_json() == {"fotos": [{"nfoto": 2, "url": "https://x/2.webp"}]} and apagadas == ["https://x/1.webp"]
    r = tabler.post(f"/inventario/{eid}/leitura/1001/foto/1/excluir", headers=json_)
    assert r.status_code == 404 and r.get_json()["erro"] and apagadas == ["https://x/1.webp"]
    # falha no bucket: erro em JSON, a foto fica
    monkeypatch.setattr(fotos, "apagar", lambda u: (_ for _ in ()).throw(RuntimeError("bucket fora")))
    r = tabler.post(f"/inventario/{eid}/leitura/1001/foto/2/excluir", headers=json_)
    assert r.status_code == 409 and "Não foi possível apagar" in r.get_json()["erro"]
    assert [f["nfoto"] for f in inventario.fotos_do_bem_no_evento(db.conectar(), eid, 1001)] == [2]
    # formulário comum (DSGov) continua com redirect
    monkeypatch.setattr(fotos, "apagar", apagadas.append)
    r = tabler.post(f"/inventario/{eid}/leitura/1001/foto/2/excluir", data={"volta": "01 - SALA CCI"})
    assert r.status_code == 302 and apagadas[-1] == "https://x/2.webp"
    # fora da comissão (Beltrana é inventariante, mas não deste evento): negado em JSON, nada apagado
    inventario.adicionar_foto(db.conectar(), eid, 1001, lambda c: "https://x/4.webp")
    tabler.post("/sair"); logar(tabler, "beltrana", SENHA_PADRAO)
    r = tabler.post(f"/inventario/{eid}/leitura/1001/foto/3/excluir", headers=json_)
    assert r.status_code == 403 and r.get_json()["erro"] and apagadas[-1] == "https://x/2.webp"
    assert len(inventario.fotos_do_bem_no_evento(db.conectar(), eid, 1001)) == 1


def test_sala_visoes_lista_e_quadros(tabler, monkeypatch):
    import db, fotos, inventario
    for v in fotos.VARIAVEIS:
        monkeypatch.setenv(v, "x")
    eid, url = _cenario_sala(tabler)
    inventario.adicionar_foto(db.conectar(), eid, 1001, lambda c: "https://x/1.webp")
    inventario.adicionar_foto(db.conectar(), eid, 1004, lambda c: "https://x/3.webp")
    html = tabler.get(url).text
    assert ' style="' not in html and "<style" not in html.split("<body")[1]
    # botão Lista/Quadros perto da busca, lembrado no aparelho
    assert 'role="group" aria-label="Visualização dos bens"' in html
    assert 'data-visao="lista" aria-pressed="false" aria-label="Ver em lista"' in html and "ti ti-list" in html
    assert 'data-visao="quadros" aria-pressed="false" aria-label="Ver em quadros"' in html and "ti ti-layout-grid" in html
    assert "localStorage" in html
    for secao, numeros in (("pendentes", [1002]), ("divergentes", [1003, 1004]), ("localizados", [1001])):
        lista = html.split(f'id="lista-{secao}"')[1].split('data-papel="caixa-grade"')[0]
        grade = html.split(f'id="grade-{secao}"')[1].split('data-papel="vazio"')[0]
        for n in numeros:   # o mesmo bem nas duas visualizações
            assert f'data-bem="{n}"' in lista and f'data-bem="{n}"' in grade
        assert "<img" not in lista and 'data-papel="avatar"' not in lista and "avatar" not in lista   # Lista nunca tem foto
        assert grade.count('data-papel="avatar"') == len(numeros)   # Quadros: foto ou ícone em todo bem
    assert '<img class="card-img-top object-cover" src="https://x/1.webp"' in html and '<img class="card-img-top object-cover" src="https://x/3.webp"' in html
    # os modelos do bem lido ao vivo também seguem as duas visualizações
    modelos = html.split('<template id="modelo-lista">')[1]
    assert "<img" not in modelos.split("</template>")[0] and 'data-papel="avatar"' in modelos.split('<template id="modelo-grade">')[1].split("</template>")[0]
