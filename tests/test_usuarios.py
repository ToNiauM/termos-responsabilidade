"""Usuários: cadastro, regras, autenticação, senhas, permissões (módulo usuarios.py, sem Flask)."""
import pytest

import db
import usuarios


def test_criar_e_buscar(dados):
    uid = usuarios.criar(dados, "Antonio", "Antônio Sousa", "Senha!234", "admin", trocar_senha=False)
    u = usuarios.por_id(dados, uid)
    assert u["login"] == "antonio" and u["nome"] == "Antônio Sousa" and u["perfil"] == "admin"
    assert u["ativo"] == 1 and u["trocar_senha"] == 0 and u["falhas"] == 0 and u["bloqueado_ate"] is None
    assert u["senha_hash"] != "Senha!234" and u["senha_hash"].startswith("scrypt:")
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
