"""CSRF: token na sessão, campo oculto em todo formulário POST, cabeçalho nos fetch."""
import re
from pathlib import Path

from flask.testing import FlaskClient

from tests.conftest import ADMIN_LOGIN, ADMIN_SENHA, semear

TEMPLATES = Path(__file__).resolve().parent.parent / "templates"


def _cru(dados, monkeypatch):
    import usuarios
    from tests.conftest import _app_de_teste, ADMIN_NOME
    semear(dados)
    monkeypatch.setenv("TERMOS_LOGIN", "1")
    usuarios.criar(dados, ADMIN_LOGIN, ADMIN_NOME, ADMIN_SENHA, "admin", trocar_senha=False)
    return FlaskClient(_app_de_teste())


def _token(html: bytes) -> str:
    return re.search(rb'name="csrf" value="([^"]+)"', html).group(1).decode()


def test_post_sem_token_da_400_e_com_token_passa(dados, monkeypatch):
    c = _cru(dados, monkeypatch)
    r = c.post("/login", data={"login": ADMIN_LOGIN, "senha": ADMIN_SENHA})
    assert r.status_code == 400 and "Sessão expirada ou formulário inválido".encode() in r.data
    token = _token(c.get("/login").data)
    assert c.post("/login", data={"login": ADMIN_LOGIN, "senha": ADMIN_SENHA, "csrf": token}).status_code == 302
    r = c.post("/textos", data={"restaurar": "orgao_nome"})
    assert r.status_code == 400
    r = c.post("/textos", data={"restaurar": "orgao_nome", "csrf": "errado"})
    assert r.status_code == 400
    novo = _token(c.get("/textos").data)                       # login gerou token novo
    assert novo != token
    assert c.post("/textos", data={"restaurar": "orgao_nome", "csrf": novo}).status_code == 302
    r = c.post("/termo/ccusto/CCI/registrar")                  # JSON sem cabeçalho
    assert r.status_code == 400 and r.is_json and r.get_json()["erro"].startswith("Sessão expirada")
    r = c.post("/termo/ccusto/CCI/registrar", headers={"X-CSRF": novo})
    assert r.status_code in (200, 409)                          # passou do CSRF (409 = sem processo vigente)


def test_todo_formulario_post_tem_csrf_campo():
    faltam = []
    for caminho in sorted(TEMPLATES.rglob("*.html")):
        texto = caminho.read_text(encoding="utf-8")
        for m in re.finditer(r"<form\b[^>]*>", texto, re.I):
            if not re.search(r'method\s*=\s*"post"', m.group(0), re.I):
                continue
            fim = texto.find("</form>", m.end())
            if "csrf_campo()" not in texto[m.end():fim if fim > 0 else None]:
                faltam.append(f"{caminho.relative_to(TEMPLATES)} @{texto.count(chr(10), 0, m.start()) + 1}")
    assert not faltam, "Formulários POST sem {{ csrf_campo() }}: " + ", ".join(faltam)


def test_meta_e_cabecalho_nos_fetch(cliente):
    html = cliente.get("/").data
    assert re.search(rb'<meta name="csrf" content="[^"]{20,}"', html)
    faltam = []
    for caminho in sorted(TEMPLATES.rglob("*.html")):
        texto = caminho.read_text(encoding="utf-8")
        if "fetch(" not in texto:
            continue
        if texto.count("fetch(") != texto.count('"X-CSRF"'):
            faltam.append(str(caminho.relative_to(TEMPLATES)))
    assert not faltam, "Nem todo fetch manda X-CSRF: " + ", ".join(faltam)


def test_cliente_de_teste_injeta_token(cliente):
    assert cliente.post("/textos", data={"restaurar": "orgao_nome"}).status_code == 302
