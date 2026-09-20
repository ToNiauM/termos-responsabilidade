"""Fixtures compartilhadas: banco SQLite temporário com dados de exemplo."""
import pytest
from flask.testing import FlaskClient
from html.parser import HTMLParser
from werkzeug.datastructures import MultiDict

import cofre


@pytest.fixture
def chave(tmp_path, monkeypatch):
    """Chave Fernet isolada em tmp_path, para testes de cofre.py e do acesso ao SEI/SPW."""
    arq = tmp_path / "chaves.env"
    arq.write_text(f"CHAVE_SENHAS={cofre.gerar_chave()}\n")
    monkeypatch.setattr(cofre, "ARQUIVO_CHAVE", arq)
    return arq


class ClienteComCSRF(FlaskClient):
    """Todo POST leva o token da sessão no cabeçalho X-CSRF, como o navegador levaria o campo oculto.
    Assim os testes existentes não precisam mudar; tests/test_csrf.py usa FlaskClient cru para provar o 400."""

    def open(self, *args, **kwargs):
        metodo = kwargs.get("method") or (args[1] if len(args) > 1 and isinstance(args[1], str) else "GET")
        if str(metodo).upper() == "POST":
            with self.session_transaction() as sess:
                token = sess.get("csrf")
                if not token:
                    token = sess["csrf"] = "token-de-teste"
            cabecalhos = dict(kwargs.get("headers") or {})
            cabecalhos.setdefault("X-CSRF", token)
            kwargs["headers"] = cabecalhos
        return super().open(*args, **kwargs)


@pytest.fixture(autouse=True)
def _hash_rapido(monkeypatch):
    """scrypt é lento de propósito; nos testes basta um pbkdf2 curto (check_password_hash lê o método do hash)."""
    from werkzeug.security import generate_password_hash
    import usuarios
    monkeypatch.setattr(usuarios, "generate_password_hash", lambda senha: generate_password_hash(senha, method="pbkdf2:sha256:1000"))


def confirmar_revisao(cliente, resposta, rota):
    """Envia os campos da revisão exibida, como faria o navegador."""
    class Campos(HTMLParser):
        def __init__(self):
            super().__init__()
            self.dados = MultiDict()

        def handle_starttag(self, tag, attrs):
            attrs = dict(attrs)
            if tag == "input" and attrs.get("type") == "hidden" and attrs.get("name"):
                self.dados.add(attrs["name"], attrs.get("value", ""))

    campos = Campos()
    campos.feed(resposta.get_data(as_text=True))
    assert campos.dados.get("revisao"), "A resposta precisa conter uma revisão antes de confirmar"
    return cliente.post(rota, data=campos.dados, follow_redirects=True)


@pytest.fixture
def dados(tmp_path, monkeypatch):
    """Pasta de dados isolada + conexão com esquema criado. Nada de dados ainda."""
    monkeypatch.setenv("TERMOS_DADOS", str(tmp_path))
    import config
    import db
    config.preparar_pastas()
    conn = db.conectar()
    db.criar_esquema(conn)
    yield conn
    conn.close()


ADMIN_LOGIN, ADMIN_NOME, ADMIN_SENHA = "admin", "Fulano", "Senha!234"
SENHA_PADRAO = "Senha!234"


def logar(cliente, login, senha):
    """POST no /login como o formulário faria; devolve a resposta (302 para / no acerto)."""
    return cliente.post("/login", data={"login": login, "senha": senha})


def _app_de_teste():
    from app import app
    app.config["TESTING"] = True
    app.config["SESSION_COOKIE_SECURE"] = False     # o test client fala http
    app.test_client_class = ClienteComCSRF
    return app


@pytest.fixture
def cliente(dados, monkeypatch):
    """Login ligado (como na web), admin 'Fulano' criado e logado, mais a inventariante 'Beltrana'
    (os dois nomes que os testes do inventário sempre usaram como comissão)."""
    import usuarios
    semear(dados)
    monkeypatch.setenv("TERMOS_LOGIN", "1")
    usuarios.criar(dados, ADMIN_LOGIN, ADMIN_NOME, ADMIN_SENHA, ["admin"], trocar_senha=False)
    usuarios.criar(dados, "beltrana", "Beltrana", SENHA_PADRAO, ["inventariante"], trocar_senha=False)
    app = _app_de_teste()
    with app.test_client() as c:
        assert logar(c, ADMIN_LOGIN, ADMIN_SENHA).status_code == 302
        yield c


@pytest.fixture
def cliente_local(dados, monkeypatch):
    """Modo desktop: sem TERMOS_LOGIN, sem usuários; tudo como Administrador local."""
    semear(dados)
    monkeypatch.delenv("TERMOS_LOGIN", raising=False)
    app = _app_de_teste()
    with app.test_client() as c:
        yield c


@pytest.fixture
def usuarios_exemplo(dados):
    """Um usuário de cada função além do admin: login → (login, senha)."""
    import usuarios
    usuarios.criar(dados, "op", "Operador Teste", SENHA_PADRAO, ["operador"], trocar_senha=False)
    usuarios.criar(dados, "leitor", "Consulta Teste", SENHA_PADRAO, ["consulta"], trocar_senha=False)
    return {"admin": (ADMIN_LOGIN, ADMIN_SENHA), "operador": ("op", SENHA_PADRAO),
            "inventariante": ("beltrana", SENHA_PADRAO), "consulta": ("leitor", SENHA_PADRAO)}


def semear(conn):
    """Cenário mínimo: 1 centro (CCI), 1 sala mapeada, 4 bens, 1 pessoa com 1 bem atribuído."""
    conn.execute("INSERT INTO responsaveis (ccustos, responsavel, email, matricula, funcao) VALUES ('CCI','JAQUELINE PORTELA','j@cfc.org.br','46','coordenadora')")
    conn.execute("INSERT INTO localizacoes VALUES ('01 - SALA CCI','CCI')")
    conn.executemany(
        "INSERT INTO bens VALUES (?,?,?,?,?,?,?,?,?)",
        [
            (1001, "ATIVO", "CADEIRA", "GIRATÓRIA", "MÓVEIS", "01 - SALA CCI", "31/12/1996", 75.94, 64.54),
            (1002, "ATIVO", "NOTEBOOK", "DELL", "EQUIPAMENTOS", "01 - SALA CCI", "06/12/2012", 3000.0, 1500.0),
            (1003, "BAIXADO", "MESA", "ANTIGA", "MÓVEIS", "01 - SALA CCI", "06/12/2012", 100.0, 10.0),
            (1004, "ATIVO", "ARMÁRIO", "AÇO", "MÓVEIS", "99 - SEM MAPA", "06/12/2012", 500.0, 250.5),
        ],
    )
    conn.execute("INSERT INTO pessoas (nome, email, matricula) VALUES ('ANA SILVA', NULL, NULL)")
    conn.execute("INSERT INTO atribuicoes VALUES ('ANA SILVA', 1002)")
    conn.commit()
