"""Usuários: cadastro, regras, autenticação, senhas, permissões (módulo usuarios.py, sem Flask)."""
import pytest
from datetime import datetime, timedelta

import db
import usuarios


def test_criar_e_buscar(dados):
    uid = usuarios.criar(dados, "Antonio", "Antônio Sousa", "Senha!234", "admin", trocar_senha=False)
    u = usuarios.por_id(dados, uid)
    assert u["login"] == "antonio" and u["nome"] == "Antônio Sousa" and u["perfil"] == "admin"
    assert u["ativo"] == 1 and u["trocar_senha"] == 0 and u["falhas"] == 0 and u["bloqueado_ate"] is None
    assert u["senha_hash"] != "Senha!234" and "$" in u["senha_hash"]
    assert usuarios.por_login(dados, "ANTONIO")["id"] == uid
    assert usuarios.por_login(dados, "ninguem") is None and usuarios.por_id(dados, 999) is None


def test_criar_valida_login_nome_senha_perfil(dados):
    with pytest.raises(db.ErroDeNegocio, match="Login"):
        usuarios.criar(dados, "a", "Nome", "Senha!234", "admin")            # curto
    with pytest.raises(db.ErroDeNegocio, match="Login"):
        usuarios.criar(dados, "com espaço", "Nome", "Senha!234", "admin")
    with pytest.raises(db.ErroDeNegocio, match="Nome"):
        usuarios.criar(dados, "fulano", "  ", "Senha!234", "admin")
    with pytest.raises(db.ErroDeNegocio, match="8 caracteres"):
        usuarios.criar(dados, "fulano", "Fulano", "curta", "admin")
    with pytest.raises(db.ErroDeNegocio, match="Perfil"):
        usuarios.criar(dados, "fulano", "Fulano", "Senha!234", "chefe")
    usuarios.criar(dados, "fulano", "Fulano  de  Tal", "Senha!234", "operador")
    assert usuarios.por_login(dados, "fulano")["nome"] == "Fulano de Tal"      # espaços colapsados
    with pytest.raises(db.ErroDeNegocio, match="já existe"):
        usuarios.criar(dados, "FULANO", "Outro", "Senha!234", "consulta")


def test_listar_filtra_busca_perfil_e_inativos(dados):
    a = usuarios.criar(dados, "admin", "Ana Admin", "Senha!234", "admin")
    b = usuarios.criar(dados, "beto", "Beto Operador", "Senha!234", "operador")
    c = usuarios.criar(dados, "carla", "Carla Consulta", "Senha!234", "consulta")
    usuarios.editar(dados, c, "Carla Consulta", "consulta", ativo=False)
    assert [u["login"] for u in usuarios.listar(dados)] == ["admin", "beto"]                 # inativo some
    assert [u["login"] for u in usuarios.listar(dados, inativos=True)] == ["admin", "beto", "carla"]
    assert [u["login"] for u in usuarios.listar(dados, busca="oper")] == ["beto"]           # busca no nome
    assert [u["login"] for u in usuarios.listar(dados, busca="ADM")] == ["admin"]           # e no login
    assert [u["login"] for u in usuarios.listar(dados, perfil="operador")] == ["beto"]
    assert "senha_hash" not in usuarios.listar(dados)[0]


def test_editar_nome_perfil_ativo_e_travas(dados):
    a = usuarios.criar(dados, "admin", "Ana", "Senha!234", "admin")
    b = usuarios.criar(dados, "beto", "Beto", "Senha!234", "operador")
    usuarios.editar(dados, b, "Beto Silva", "inventariante", ativo=True)
    u = usuarios.por_id(dados, b)
    assert u["nome"] == "Beto Silva" and u["perfil"] == "inventariante" and u["login"] == "beto"
    with pytest.raises(db.ErroDeNegocio, match="último administrador"):
        usuarios.editar(dados, a, "Ana", "operador", ativo=True)
    with pytest.raises(db.ErroDeNegocio, match="último administrador"):
        usuarios.editar(dados, a, "Ana", "admin", ativo=False)
    usuarios.editar(dados, b, "Beto Silva", "admin", ativo=True)         # agora há dois admins
    with pytest.raises(db.ErroDeNegocio, match="própria conta"):
        usuarios.editar(dados, a, "Ana", "consulta", ativo=True, logado_id=a)
    with pytest.raises(db.ErroDeNegocio, match="própria conta"):
        usuarios.editar(dados, a, "Ana", "admin", ativo=False, logado_id=a)
    usuarios.editar(dados, a, "Ana", "consulta", ativo=True, logado_id=b)   # outro admin pode
    assert usuarios.por_id(dados, a)["perfil"] == "consulta"
    with pytest.raises(db.ErroDeNegocio, match="não encontrado"):
        usuarios.editar(dados, 999, "X", "admin", ativo=True)


def test_elegiveis_comissao(dados):
    usuarios.criar(dados, "admin", "Ana", "Senha!234", "admin")
    usuarios.criar(dados, "beto", "Beto", "Senha!234", "operador")
    usuarios.criar(dados, "cris", "Cris", "Senha!234", "inventariante")
    d = usuarios.criar(dados, "dora", "Dora", "Senha!234", "inventariante")
    usuarios.criar(dados, "eva", "Eva", "Senha!234", "consulta")
    usuarios.editar(dados, d, "Dora", "inventariante", ativo=False)
    assert [(u["nome"], u["perfil"]) for u in usuarios.elegiveis_comissao(dados)] == [
        ("Ana", "admin"), ("Beto", "operador"), ("Cris", "inventariante")]


def test_usuario_local():
    assert usuarios.USUARIO_LOCAL["id"] is None and usuarios.USUARIO_LOCAL["perfil"] == "admin"
    assert usuarios.USUARIO_LOCAL["nome"] == "Administrador local"


def test_autenticar_acerto_erro_inativo(dados, monkeypatch):
    uid = usuarios.criar(dados, "ana", "Ana", "Senha!234", "admin")
    monkeypatch.setattr(usuarios, "_agora_dt", lambda: datetime(2026, 9, 17, 10, 0, 0))
    u = usuarios.autenticar(dados, "ANA", "Senha!234")
    assert u["id"] == uid and usuarios.por_id(dados, uid)["ultimo_acesso"] == "2026-09-17 10:00:00"
    with pytest.raises(db.ErroDeNegocio, match="Usuário ou senha inválidos"):
        usuarios.autenticar(dados, "ana", "errada")
    with pytest.raises(db.ErroDeNegocio, match="Usuário ou senha inválidos"):
        usuarios.autenticar(dados, "ninguem", "Senha!234")
    assert usuarios.por_id(dados, uid)["falhas"] == 1
    usuarios.editar(dados, usuarios.criar(dados, "outro", "Outro", "Senha!234", "admin"), "Outro", "admin", ativo=True)
    usuarios.editar(dados, uid, "Ana", "admin", ativo=False)
    with pytest.raises(db.ErroDeNegocio, match="Usuário ou senha inválidos"):
        usuarios.autenticar(dados, "ana", "Senha!234")            # inativo: mesma mensagem


def test_bloqueio_na_quinta_falha_e_liberacao(dados, monkeypatch):
    uid = usuarios.criar(dados, "ana", "Ana", "Senha!234", "admin")
    relogio = {"agora": datetime(2026, 9, 17, 10, 0, 0)}
    monkeypatch.setattr(usuarios, "_agora_dt", lambda: relogio["agora"])
    for _ in range(4):
        with pytest.raises(db.ErroDeNegocio, match="inválidos"):
            usuarios.autenticar(dados, "ana", "errada")
    assert usuarios.por_id(dados, uid)["bloqueado_ate"] is None
    with pytest.raises(db.ErroDeNegocio, match="inválidos"):
        usuarios.autenticar(dados, "ana", "errada")                # 5ª falha bloqueia
    assert usuarios.por_id(dados, uid)["bloqueado_ate"] == "2026-09-17 10:15:00"
    with pytest.raises(db.ErroDeNegocio, match="Muitas tentativas"):
        usuarios.autenticar(dados, "ana", "Senha!234")             # certa, mas bloqueado
    assert usuarios.por_id(dados, uid)["falhas"] == 5              # bloqueado não conta falha
    relogio["agora"] = datetime(2026, 9, 17, 10, 15, 1)
    u = usuarios.autenticar(dados, "ana", "Senha!234")
    assert u["id"] == uid
    u = usuarios.por_id(dados, uid)
    assert u["falhas"] == 0 and u["bloqueado_ate"] is None


def test_senha_temporaria_e_troca_obrigatoria(dados):
    uid = usuarios.criar(dados, "ana", "Ana", "Senha!234", "admin", trocar_senha=False)
    temp = usuarios.nova_senha_temporaria(dados, uid)
    assert len(temp) == 10 and set(temp) <= set(usuarios.ALFABETO_TEMP)
    u = usuarios.por_id(dados, uid)
    assert u["trocar_senha"] == 1 and u["falhas"] == 0 and u["bloqueado_ate"] is None
    assert usuarios.autenticar(dados, "ana", temp)["trocar_senha"] == 1
    with pytest.raises(db.ErroDeNegocio, match="atual"):
        usuarios.trocar_senha(dados, uid, "errada", "NovaSenha1", "NovaSenha1")
    with pytest.raises(db.ErroDeNegocio, match="8 caracteres"):
        usuarios.trocar_senha(dados, uid, temp, "curta", "curta")
    with pytest.raises(db.ErroDeNegocio, match="diferente"):
        usuarios.trocar_senha(dados, uid, temp, temp, temp)
    with pytest.raises(db.ErroDeNegocio, match="confirmação"):
        usuarios.trocar_senha(dados, uid, temp, "NovaSenha1", "NovaSenha2")
    usuarios.trocar_senha(dados, uid, temp, "NovaSenha1", "NovaSenha1")
    assert usuarios.por_id(dados, uid)["trocar_senha"] == 0
    assert usuarios.autenticar(dados, "ana", "NovaSenha1")["id"] == uid
    with pytest.raises(db.ErroDeNegocio, match="não encontrado"):
        usuarios.nova_senha_temporaria(dados, 999)


def test_criar_admin_cria_ou_redefine(dados):
    uid = usuarios.criar_admin(dados, "antonio", "Antônio", "Senha!234")
    assert usuarios.por_id(dados, uid)["perfil"] == "admin" and usuarios.por_id(dados, uid)["trocar_senha"] == 0
    for _ in range(5):
        with pytest.raises(db.ErroDeNegocio):
            usuarios.autenticar(dados, "antonio", "x")
    usuarios.editar(dados, usuarios.criar(dados, "bd", "B", "Senha!234", "admin"), "B", "admin", ativo=True)
    usuarios.editar(dados, uid, "Antônio", "consulta", ativo=False)
    assert usuarios.criar_admin(dados, "antonio", "Antônio R.", "OutraSenha9") == uid   # mesmo id
    u = usuarios.por_id(dados, uid)
    assert u["perfil"] == "admin" and u["ativo"] == 1 and u["falhas"] == 0 and u["bloqueado_ate"] is None
    assert u["nome"] == "Antônio"                      # nome não muda na redefinição
    assert usuarios.autenticar(dados, "antonio", "OutraSenha9")["id"] == uid


def test_main_criar_admin(dados, capsys):
    senhas = iter(["Senha!234", "Senha!234"])
    assert usuarios.main(["criar-admin", "ze", "Zé"], ler_senha=lambda _p: next(senhas)) == 0
    assert usuarios.por_login(dados, "ze")["perfil"] == "admin"
    senhas = iter(["Senha!234", "Diferente1"])
    assert usuarios.main(["criar-admin", "ze", "Zé"], ler_senha=lambda _p: next(senhas)) == 1
    assert "não conferem" in capsys.readouterr().err
    assert usuarios.main(["outra-coisa"], ler_senha=lambda _p: "x") == 2


def test_email_opcional_unico_e_login_por_email(dados):
    a = usuarios.criar(dados, "ana", "Ana", "Senha!234", "admin", email=" Ana@CFC.org.br ")
    assert usuarios.por_id(dados, a)["email"] == "ana@cfc.org.br"
    b = usuarios.criar(dados, "beto", "Beto", "Senha!234", "operador")            # sem e-mail
    assert usuarios.por_id(dados, b)["email"] is None
    with pytest.raises(db.ErroDeNegocio, match="e-mail já"):
        usuarios.criar(dados, "carla", "Carla", "Senha!234", "consulta", email="ANA@cfc.org.br")
    for ruim in ("sem-arroba", "com espaco@x.org", "@x.org", "x@"):
        with pytest.raises(db.ErroDeNegocio, match="E-mail inválido"):
            usuarios.criar(dados, "carla", "Carla", "Senha!234", "consulta", email=ruim)
    assert usuarios.autenticar(dados, "ANA@cfc.org.br", "Senha!234")["id"] == a   # entra pelo e-mail
    assert usuarios.autenticar(dados, "ana", "Senha!234")["id"] == a              # ou pelo login
    with pytest.raises(db.ErroDeNegocio, match="inválidos"):
        usuarios.autenticar(dados, "ninguem@cfc.org.br", "Senha!234")
    usuarios.editar(dados, b, "Beto", "operador", ativo=True, email="beto@cfc.org.br")
    assert usuarios.autenticar(dados, "beto@cfc.org.br", "Senha!234")["id"] == b
    with pytest.raises(db.ErroDeNegocio, match="e-mail já"):
        usuarios.editar(dados, b, "Beto", "operador", ativo=True, email="ana@cfc.org.br")
    usuarios.editar(dados, b, "Beto Silva", "operador", ativo=True)                # sem email: mantém
    assert usuarios.por_id(dados, b)["email"] == "beto@cfc.org.br"
    usuarios.editar(dados, b, "Beto", "operador", ativo=True, email="")            # vazio: apaga
    assert usuarios.por_id(dados, b)["email"] is None
    assert usuarios.listar(dados)[0]["email"] == "ana@cfc.org.br"
    assert [u["login"] for u in usuarios.listar(dados, busca="cfc.org")] == ["ana"]   # busca também no e-mail


def test_criar_admin_com_email(dados):
    uid = usuarios.criar_admin(dados, "antonio", "Antônio", "Senha!234", email="antonio@cfc.org.br")
    assert usuarios.por_id(dados, uid)["email"] == "antonio@cfc.org.br"
    assert usuarios.criar_admin(dados, "antonio", "Antônio", "Outra!2345") == uid          # sem e-mail: mantém
    assert usuarios.por_id(dados, uid)["email"] == "antonio@cfc.org.br"
    usuarios.criar_admin(dados, "antonio", "Antônio", "Outra!2345", email="novo@cfc.org.br")
    assert usuarios.por_id(dados, uid)["email"] == "novo@cfc.org.br"
    senhas = iter(["Senha!234", "Senha!234"])
    assert usuarios.main(["criar-admin", "ze", "Zé", "ze@cfc.org.br"], ler_senha=lambda _p: next(senhas)) == 0
    assert usuarios.por_login(dados, "ze")["email"] == "ze@cfc.org.br"
