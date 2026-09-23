"""Protótipo Tabler — Contas, Administração, Ajuda e erros (login, senha, acessos, 403, administração,
usuários e ajuda): cada tela abre no Tabler, sem DSGov, e o fluxo principal continua funcionando."""
import re

import pytest

import usuarios
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


def test_login_no_tabler_entra_e_recusa_senha_errada(tabler):
    tabler.post("/sair")
    html = tabler.get("/login?proximo=/usuarios").text
    assert _e_tabler(html) and "container-tight" in html and "dsgov.js" not in html
    assert 'name="login"' in html and 'name="senha"' in html and 'name="csrf"' in html and "Usuário ou e-mail" in html
    assert 'action="/login?proximo=/usuarios"' in html
    r = tabler.post("/login", data={"login": "beltrana", "senha": "errada!!"})
    assert r.status_code == 200 and _e_tabler(r.text) and 'class="alert alert-danger"' in r.text and 'value="beltrana"' in r.text
    assert logar(tabler, "beltrana", SENHA_PADRAO).status_code == 302


def test_login_sem_usuarios_orienta_criar_admin(tabler, dados):
    tabler.post("/sair")
    dados.execute("DELETE FROM usuarios"); dados.commit()
    html = tabler.get("/login").text
    assert _e_tabler(html) and "Nenhum usuário cadastrado." in html and "criar-admin" in html and 'name="senha"' not in html


def test_trocar_senha_no_tabler(tabler, dados):
    html = tabler.get("/senha").text
    assert _e_tabler(html) and 'id="atual"' in html and 'id="nova"' in html and 'id="confirmacao"' in html
    assert 'href="/ajuda#conta"' in html and ">Cancelar<" in html
    r = tabler.post("/senha", data={"atual": "errada", "nova": "Nova!2345", "confirmacao": "Nova!2345"})
    assert _e_tabler(r.text) and 'class="alert alert-danger"' in r.text
    usuarios.criar(dados, "temp", "Temporário", SENHA_PADRAO, ["consulta"], trocar_senha=True)
    tabler.post("/sair"); logar(tabler, "temp", SENHA_PADRAO)
    html = tabler.get("/senha").text
    assert "Defina uma nova senha para continuar." in html and ">Cancelar<" not in html
    r = tabler.post("/senha", data={"atual": SENHA_PADRAO, "nova": "Nova!2345", "confirmacao": "Nova!2345"})
    assert r.status_code == 302
    assert logar(tabler, "temp", "Nova!2345").status_code == 302 or tabler.get("/").status_code == 200


def test_meus_acessos_no_tabler(tabler, chave):
    html = tabler.get("/meus-acessos").text
    assert _e_tabler(html) and 'id="sei_login"' in html and 'id="spw_senha"' in html and 'name="sei_unidade"' in html
    assert 'action="/meus-acessos/sei"' in html and 'action="/meus-acessos/spw"' in html and "Apagar" not in html
    r = tabler.post("/meus-acessos/sei", data={"sei_login": "fulano.sei", "sei_senha": "S3nha!", "sei_unidade": "gelic"}, follow_redirects=True)
    assert _e_tabler(r.text) and "Acesso ao SEI salvo." in r.text and "S3nha" not in r.text
    assert "Senha cadastrada em" in r.text and 'value="fulano.sei"' in r.text and 'action="/meus-acessos/sei/apagar"' in r.text
    r = tabler.post("/meus-acessos/sei/apagar", follow_redirects=True)
    assert "Acesso ao SEI apagado." in r.text


def test_403_no_tabler_e_empty_state(tabler, dados):
    tabler.post("/sair"); logar(tabler, "beltrana", SENHA_PADRAO)
    r = tabler.get("/usuarios")
    assert r.status_code == 403 and _e_tabler(r.text) and 'class="empty"' in r.text
    assert "Seu usuário não tem permissão para esta ação." in r.text and "Acesso negado" in r.text and "/ajuda#" not in r.text


def test_administracao_no_tabler_abas_e_criar_inventario(tabler, dados):
    html = tabler.get("/administracao").text
    assert _e_tabler(html) and 'class="nav nav-tabs"' in html
    abas = html.split('aria-label="Administração"')[1].split("</nav>")[0]
    assert 'href="/administracao" aria-current="page"' in abas and 'href="/usuarios"' in abas and ">Usuários<" in abas
    assert 'id="lista-salas"' in html and 'id="escopo-escolher"' in html
    assert "checked" in html.split('name="abrir_agora"')[1][:80]
    beltrana = usuarios.por_login(dados, "beltrana")["id"]
    r = tabler.post("/inventario/abrir", data={"nome": "Inv T", "usuarios": [beltrana], "escopo": "todas", "abrir_agora": "1"}, follow_redirects=True)
    assert r.request.path == "/administracao" and _e_tabler(r.text) and "Inventário Inv T aberto." in r.text
    assert ">aberto<" in r.text and "Fechar" in r.text and 'aria-label="Excluir Inv T"' in r.text and 'id="tabela-inventarios"' in r.text

    html = tabler.get("/usuarios").text                               # a outra aba
    abas = html.split('aria-label="Administração"')[1].split("</nav>")[0]
    assert 'href="/usuarios" aria-current="page"' in abas and 'href="/administracao"' in abas


def test_usuarios_lista_e_criar_no_tabler(tabler, dados):
    html = tabler.get("/usuarios").text
    assert _e_tabler(html) and 'id="tabela-usuarios"' in html and ">admin<" in html and 'class="form-select"' in html
    assert 'value="consulta_inventarios"' in html and 'href="/ajuda#usuarios"' in html
    html = tabler.get("/usuarios/novo").text
    assert _e_tabler(html) and 'id="login"' in html and 'id="funcao-operador"' in html and 'id="trocar_senha"' in html
    r = tabler.post("/usuarios/incluir", data={"login": "novo", "nome": "Novo", "funcoes": ["operador"], "senha": "Senha!234", "confirmacao": "Outra!234"})
    assert _e_tabler(r.text) and "Confira os campos." in r.text and 'value="novo"' in r.text
    r = tabler.post("/usuarios/incluir", data={"login": "novo", "nome": "Novo", "funcoes": ["operador"], "senha": "Senha!234", "confirmacao": "Senha!234"}, follow_redirects=True)
    assert _e_tabler(r.text) and ">novo<" in r.text
    uid = usuarios.por_login(dados, "novo")["id"]
    html = tabler.get(f"/usuarios/{uid}/editar").text
    assert _e_tabler(html) and "Usuário <code class=\"ms-2\">novo</code>" in html and 'id="ativo"' in html and f'action="/usuarios/{uid}/nova-senha' in html
    r = tabler.post(f"/usuarios/{uid}/nova-senha")
    assert _e_tabler(r.text) and 'id="senha-temporaria"' in r.text and ">Voltar<" in r.text


def test_ajuda_no_tabler_preserva_ancoras(tabler):
    html = tabler.get("/ajuda").text
    assert _e_tabler(html) and html.count("<h1") == 1
    secoes = re.findall(r'<section class="[^"]*" id="([^"]+)"', html)
    assert secoes and all(f'href="#{s}"' in html for s in secoes)
    assert {"inicio", "termos", "usuarios", "conta", "perguntas"} <= set(secoes)
    assert 'class="list-group-item list-group-item-action" href="#inicio"' in html
