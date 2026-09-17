"""Entrar, sair, troca obrigatória de senha, sessão e modo desktop (rotas)."""
from datetime import datetime

import usuarios
from tests.conftest import ADMIN_LOGIN, ADMIN_NOME, ADMIN_SENHA, SENHA_PADRAO, logar


def _sem_login(dados, monkeypatch):
    from tests.conftest import _app_de_teste
    monkeypatch.setenv("TERMOS_LOGIN", "1")
    return _app_de_teste().test_client()


def test_sem_sessao_redireciona_para_login_com_proximo(cliente):
    cliente.post("/sair")
    r = cliente.get("/bem?numero=1001")
    assert r.status_code == 302 and r.headers["Location"] == "/login?proximo=%2Fbem%3Fnumero%3D1001"
    assert cliente.post("/textos", data={}).headers["Location"] == "/login"          # POST não carrega proximo
    r = cliente.get("/login")
    assert r.status_code == 200 and b'name="login"' in r.data and b'name="senha"' in r.data
    assert b"main-navigation" not in r.data                                            # tela sem menu


def test_login_acerto_erro_e_proximo(cliente, dados):
    cliente.post("/sair")
    r = cliente.post("/login?proximo=%2Fbem%3Fnumero%3D1001", data={"login": ADMIN_LOGIN, "senha": ADMIN_SENHA})
    assert r.status_code == 302 and r.headers["Location"] == "/bem?numero=1001"
    assert ADMIN_NOME.encode() in cliente.get("/").data and b"Sair" in cliente.get("/").data
    assert usuarios.por_login(dados, ADMIN_LOGIN)["ultimo_acesso"] is not None
    assert cliente.get("/login").headers["Location"] == "/"                           # já logado
    cliente.post("/sair")
    r = cliente.post("/login", data={"login": ADMIN_LOGIN, "senha": "errada"})
    assert r.status_code == 200 and "Usuário ou senha inválidos".encode() in r.data
    for proximo in ("//evil.example", "http://evil.example", "/\\evil"):
        r = cliente.post(f"/login?proximo={proximo}", data={"login": ADMIN_LOGIN, "senha": ADMIN_SENHA})
        assert r.headers["Location"] == "/"
        cliente.post("/sair")
    r = cliente.post("/login?proximo=/bem%0D%0Ainjetado", data={"login": ADMIN_LOGIN, "senha": ADMIN_SENHA})
    assert r.status_code == 302 and r.headers["Location"] == "/"
    cliente.post("/sair")


def test_login_inativo_e_bloqueado(cliente, dados, monkeypatch):
    cliente.post("/sair")
    relogio = {"agora": datetime(2026, 9, 17, 10, 0)}
    monkeypatch.setattr(usuarios, "_agora_dt", lambda: relogio["agora"])
    for _ in range(5):
        cliente.post("/login", data={"login": "beltrana", "senha": "errada"})
    r = cliente.post("/login", data={"login": "beltrana", "senha": SENHA_PADRAO})
    assert b"Muitas tentativas" in r.data
    relogio["agora"] = datetime(2026, 9, 17, 10, 16)
    assert cliente.post("/login", data={"login": "beltrana", "senha": SENHA_PADRAO}).status_code == 302
    cliente.post("/sair")
    usuarios.editar(dados, usuarios.por_login(dados, "beltrana")["id"], "Beltrana", ["inventariante"], ativo=False)
    r = cliente.post("/login", data={"login": "beltrana", "senha": SENHA_PADRAO})
    assert "inválidos".encode() in r.data


def test_usuario_inativado_com_sessao_aberta_cai_no_login(cliente, dados):
    logar(cliente, "beltrana", SENHA_PADRAO)
    assert cliente.get("/").status_code == 200
    usuarios.editar(dados, usuarios.por_login(dados, "beltrana")["id"], "Beltrana", ["inventariante"], ativo=False)
    assert cliente.get("/").headers["Location"].startswith("/login")


def test_sair_limpa_a_sessao(cliente):
    with cliente.session_transaction() as s:
        s["bens_selecionados"] = ["1002"]
    r = cliente.post("/sair")
    assert r.headers["Location"] == "/login"
    with cliente.session_transaction() as s:
        assert "usuario_id" not in s and "bens_selecionados" not in s


def test_troca_obrigatoria_de_senha(cliente, dados):
    uid = usuarios.por_login(dados, "beltrana")["id"]
    temp = usuarios.nova_senha_temporaria(dados, uid)
    cliente.post("/sair")
    logar(cliente, "beltrana", temp)
    assert cliente.get("/").headers["Location"] == "/senha"
    assert cliente.get("/inventario").headers["Location"] == "/senha"
    r = cliente.get("/senha")
    assert r.status_code == 200 and b"Defina uma nova senha" in r.data
    r = cliente.post("/senha", data={"atual": temp, "nova": "curta", "confirmacao": "curta"})
    assert r.status_code == 200 and b"8 caracteres" in r.data
    r = cliente.post("/senha", data={"atual": temp, "nova": "NovaSenha1", "confirmacao": "NovaSenha1"})
    assert r.headers["Location"] == "/"
    assert cliente.get("/").status_code == 200 and usuarios.por_id(dados, uid)["trocar_senha"] == 0
    r = cliente.get("/senha")
    assert r.status_code == 200 and b"Trocar senha" in r.data                          # voluntária também


def test_sessao_e_cookie(cliente):
    from app import app
    assert app.config["PERMANENT_SESSION_LIFETIME"].total_seconds() == 12 * 3600
    assert app.config["SESSION_COOKIE_HTTPONLY"] and app.config["SESSION_COOKIE_SAMESITE"] == "Lax"
    with cliente.session_transaction() as s:
        assert s.permanent


def test_modo_desktop_entra_direto_como_admin_local(cliente_local):
    r = cliente_local.get("/")
    assert r.status_code == 200 and b"Sair" not in r.data and b"Administrador local" not in r.data
    assert cliente_local.get("/login").status_code == 404
    assert cliente_local.get("/senha").status_code == 404
    assert cliente_local.get("/textos").status_code == 200


def test_sem_usuarios_cadastrados_tela_de_login_orienta(dados, monkeypatch):
    c = _sem_login(dados, monkeypatch)
    r = c.get("/login")
    assert r.status_code == 200 and "Nenhum usuário cadastrado".encode() in r.data and b"criar-admin" in r.data
    assert b'name="senha"' not in r.data


def test_login_por_email(cliente, dados):
    usuarios.editar(dados, usuarios.por_login(dados, "beltrana")["id"], "Beltrana", ["inventariante"], ativo=True, email="beltrana@cfc.org.br")
    cliente.post("/sair")
    assert "Usuário ou e-mail".encode() in cliente.get("/login").data
    assert logar(cliente, "Beltrana@CFC.org.br", SENHA_PADRAO).status_code == 302
    assert b"Beltrana" in cliente.get("/").data
