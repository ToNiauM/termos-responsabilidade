# Fase 4 (usuários, senha e perfis) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Login próprio no Flask com quatro perfis (admin, operador, inventariante, consulta), tela de usuários, senha temporária, comissão do inventário formada por usuários, exclusão de evento pelo admin, CSRF, e modo desktop sem login.

**Architecture:** Tabela `usuarios` no SQLite e módulo `usuarios.py` (dados e regras, recebe `conn`, sem Flask), blueprint `app_usuarios.py` (login, sair, senha, telas de usuários), `before_request` em `app.py` que resolve `g.usuario`, checa CSRF, troca obrigatória e `usuarios.permitido(perfil, endpoint, metodo)` contra uma matriz `PERMISSOES` que nega por padrão. Inventário passa a gravar `g.usuario["nome"]` como integrante. Sem `TERMOS_LOGIN=1`, tudo roda como "Administrador local".

**Tech Stack:** Python 3.12, Flask 3, `werkzeug.security` (hash scrypt, já vem com o Flask), SQLite, Jinja2, DSGov 3.7.0 (só visual), pytest. Nenhuma dependência nova.

**Spec:** `docs/superpowers/specs/2026-09-17-fase4-usuarios-design.md`

## Global Constraints

- Nenhuma biblioteca nova: `requirements.txt` e `Dockerfile` não mudam (Flask-Login descartado).
- Todo módulo de dados recebe `conn` como primeiro argumento e não importa Flask (padrão de `db.py` e `inventario.py`).
- Erros de regra são `db.ErroDeNegocio` (viram `flash` pelo handler global de `app.py`; nas rotas JSON viram `{"erro": ...}` 409).
- Textos de interface em português, na voz do sistema atual (ex.: "Evento aberto. Comece pelas salas.").
- Perfis exatamente: `admin`, `operador`, `inventariante`, `consulta` (CHECK na tabela). Rótulos: Administrador, Operador, Inventariante, Consulta.
- Login: `[a-z0-9._-]{2,30}`, normalizado para minúsculas, imutável. Senha mínima: 8 caracteres. Bloqueio: 5 falhas seguidas → 15 minutos. Sessão: 12 horas.
- Senha temporária: 10 caracteres do alfabeto `ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789`, mostrada uma vez, `trocar_senha = 1`.
- Usuário não se exclui, só se inativa. Último admin ativo não é inativado nem rebaixado; o próprio logado também não.
- Rota fora de `PERMISSOES` = 403 para todos, e o teste da matriz falha nomeando a rota.
- Modo desktop: `config.exigir_login()` é `os.environ.get("TERMOS_LOGIN") == "1"`; sem isso `g.usuario = usuarios.USUARIO_LOCAL` (`{"id": None, "login": "local", "nome": "Administrador local", "perfil": "admin", "ativo": 1, "trocar_senha": 0}`).
- Não excluir termos emitidos. Não mudar `inventario_integrantes` (continua `evento_id, nome`).
- Cada tarefa termina com `.venv/bin/pytest -q` verde e um commit. Mensagens de commit em português, como o histórico do repo.
- Rodar testes com `.venv/bin/pytest -q` a partir da raiz do repositório.

---

## Estrutura de arquivos

| Arquivo | Responsabilidade |
|---|---|
| `usuarios.py` (novo) | tabela `usuarios`: criar, listar, editar, autenticar, senhas, elegíveis para comissão, `PERMISSOES`/`permitido`, comando `criar-admin` |
| `app_usuarios.py` (novo) | blueprint `usuarios`: `/login`, `/sair`, `/senha`, `/usuarios*` |
| `app.py` | config de sessão, `before_request` (usuário, CSRF, troca obrigatória, permissão), `csrf_campo`, menu por perfil, 403, registro do blueprint |
| `config.py` | `exigir_login()` |
| `db.py` | `CREATE TABLE usuarios` em `ESQUEMA` |
| `inventario.py` | `editar_comissao`, `contagem_para_exclusao`, `urls_das_fotos`, `excluir_evento`, `abrir_evento(..., elegiveis=)`, mensagem de comissão |
| `app_inventario.py` | integrante = `g.usuario["nome"]`; sai `/integrante`; entram `/comissao` e `/excluir` |
| `templates/login.html`, `senha.html`, `403.html`, `usuarios/lista.html`, `usuarios/formulario.html`, `inventario_comissao.html`, `inventario_excluir.html` (novos) | telas |
| `templates/base.html`, `inventario_eventos.html`, `inventario_evento.html`, `inventario_sala.html`, `termo.html` e todos com `<form method="post">` | cabeçalho com usuário, comissão por caixas, sem seletor de integrante, botões por perfil, `csrf_campo()` |
| `tests/conftest.py` | fixture `cliente` única (login ligado, admin logado, cliente que injeta CSRF), `logar`, `usuarios_exemplo` |
| `tests/test_usuarios.py`, `tests/test_permissoes.py`, `tests/test_csrf.py` (novos); `tests/test_app.py`, `tests/test_cadastros_ux.py` (ajustes) | testes |
| `compose.yml`, `README.md` | `TERMOS_LOGIN=1`; publicação, primeiro admin, remoção do `auth_basic` |

---

### Task 1: Tabela `usuarios` e cadastro básico (`usuarios.py`)

**Files:**
- Modify: `db.py` (bloco `ESQUEMA`, logo após `CREATE TABLE IF NOT EXISTS inventario_fotos (...)`)
- Create: `usuarios.py`
- Create: `tests/test_usuarios.py`

**Interfaces:**
- Produces: `usuarios.PERFIS`, `usuarios.ROTULO_PERFIL`, `usuarios.USUARIO_LOCAL`, `usuarios.criar(conn, login, nome, senha, perfil, trocar_senha=True) -> int`, `usuarios.por_id(conn, id) -> dict | None`, `usuarios.por_login(conn, login) -> dict | None`, `usuarios.listar(conn, busca="", perfil=None, inativos=False) -> list[dict]`, `usuarios.editar(conn, id, nome, perfil, ativo, logado_id=None) -> None`, `usuarios.elegiveis_comissao(conn) -> list[dict]`, `usuarios.SENHA_MINIMA`.
- Dicionário de usuário: chaves `id, login, nome, senha_hash, perfil, ativo, trocar_senha, falhas, bloqueado_ate, criado_em, ultimo_acesso`.

- [ ] **Step 1: Escrever os testes que falham**

Criar `tests/test_usuarios.py`:

```python
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
```

- [ ] **Step 2: Rodar para ver falhar**

Run: `.venv/bin/pytest tests/test_usuarios.py -q`
Expected: erros `ModuleNotFoundError: No module named 'usuarios'`.

- [ ] **Step 3: Tabela no esquema**

Em `db.py`, dentro de `ESQUEMA`, depois do bloco `CREATE TABLE IF NOT EXISTS inventario_fotos (...);`, acrescentar:

```sql
CREATE TABLE IF NOT EXISTS usuarios (
  id            INTEGER PRIMARY KEY,
  login         TEXT NOT NULL UNIQUE,
  nome          TEXT NOT NULL,
  senha_hash    TEXT NOT NULL,
  perfil        TEXT NOT NULL CHECK (perfil IN ('admin','operador','inventariante','consulta')),
  ativo         INTEGER NOT NULL DEFAULT 1,
  trocar_senha  INTEGER NOT NULL DEFAULT 0,
  falhas        INTEGER NOT NULL DEFAULT 0,
  bloqueado_ate TEXT,
  criado_em     TEXT NOT NULL,
  ultimo_acesso TEXT
);
```

- [ ] **Step 4: Criar `usuarios.py`**

```python
"""Usuários, senhas e permissões. Só dados e regras: toda função recebe `conn` primeiro e não importa Flask
(padrão de db.py). Hash de senha com werkzeug.security (scrypt), que já vem com o Flask."""
import re
import secrets
from datetime import datetime, timedelta

from werkzeug.security import check_password_hash, generate_password_hash

from db import ErroDeNegocio, _agora, _obrigatorio, _todos, _um

PERFIS = ("admin", "operador", "inventariante", "consulta")
ROTULO_PERFIL = {"admin": "Administrador", "operador": "Operador", "inventariante": "Inventariante", "consulta": "Consulta"}
PERFIS_COMISSAO = ("admin", "operador", "inventariante")
SENHA_MINIMA = 8
MAX_FALHAS = 5
BLOQUEIO_MINUTOS = 15
ALFABETO_TEMP = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789"
_LOGIN = re.compile(r"[a-z0-9._-]{2,30}")
_FORMATO = "%Y-%m-%d %H:%M:%S"

# Modo desktop (sem TERMOS_LOGIN): quem usa o programa Windows é o administrador da instalação.
USUARIO_LOCAL = {"id": None, "login": "local", "nome": "Administrador local", "perfil": "admin", "ativo": 1, "trocar_senha": 0}

_COLUNAS_LISTA = "id, login, nome, perfil, ativo, trocar_senha, falhas, bloqueado_ate, criado_em, ultimo_acesso"


def _agora_dt() -> datetime:
    """Separado para os testes simularem o relógio (bloqueio de 15 min)."""
    return datetime.now()


def _validar_login(login) -> str:
    v = str(login or "").strip().lower()
    if not _LOGIN.fullmatch(v):
        raise ErroDeNegocio("Login inválido: use de 2 a 30 caracteres entre letras minúsculas, números, ponto, hífen e sublinhado.")
    return v


def _validar_senha(senha) -> str:
    s = str(senha or "")
    if len(s) < SENHA_MINIMA:
        raise ErroDeNegocio(f"A senha precisa ter ao menos {SENHA_MINIMA} caracteres.")
    return s


def _validar_perfil(perfil) -> str:
    if perfil not in PERFIS:
        raise ErroDeNegocio("Perfil inválido.")
    return perfil


def criar(conn, login, nome, senha, perfil, trocar_senha=True) -> int:
    login = _validar_login(login)
    nome = _obrigatorio(nome, "Nome")
    senha = _validar_senha(senha)
    perfil = _validar_perfil(perfil)
    if por_login(conn, login):
        raise ErroDeNegocio(f"O login {login} já existe.")
    cur = conn.execute("INSERT INTO usuarios (login, nome, senha_hash, perfil, trocar_senha, criado_em) VALUES (?,?,?,?,?,?)",
                       (login, nome, generate_password_hash(senha), perfil, 1 if trocar_senha else 0, _agora()))
    conn.commit()
    return cur.lastrowid


def por_id(conn, id) -> dict | None:
    return _um(conn, "SELECT * FROM usuarios WHERE id = ?", id) if id is not None else None


def por_login(conn, login) -> dict | None:
    return _um(conn, "SELECT * FROM usuarios WHERE login = ?", str(login or "").strip().lower())


def listar(conn, busca="", perfil=None, inativos=False) -> list[dict]:
    """Sem senha_hash. busca casa em login e nome (sem caixa); inativos=False esconde os inativos."""
    sql, p = f"SELECT {_COLUNAS_LISTA} FROM usuarios WHERE 1=1", []
    if not inativos:
        sql += " AND ativo = 1"
    if perfil:
        sql += " AND perfil = ?"
        p.append(perfil)
    if busca and busca.strip():
        sql += " AND (lower(login) LIKE ? OR lower(nome) LIKE ?)"
        termo = f"%{busca.strip().lower()}%"
        p += [termo, termo]
    return _todos(conn, sql + " ORDER BY login", *p)


def _admins_ativos(conn) -> int:
    return conn.execute("SELECT count(*) FROM usuarios WHERE perfil = 'admin' AND ativo = 1").fetchone()[0]


def editar(conn, id, nome, perfil, ativo, logado_id=None) -> None:
    """Nome, perfil e ativo. Login não muda. Travas: o próprio logado não se inativa nem se rebaixa; o último
    administrador ativo não é inativado nem rebaixado."""
    u = por_id(conn, id)
    if not u:
        raise ErroDeNegocio("Usuário não encontrado.")
    nome = _obrigatorio(nome, "Nome")
    perfil = _validar_perfil(perfil)
    ativo = 1 if ativo else 0
    perde_admin = u["perfil"] == "admin" and u["ativo"] and (perfil != "admin" or not ativo)
    if perde_admin and logado_id is not None and int(logado_id) == int(id):
        raise ErroDeNegocio("Você não pode rebaixar nem inativar a própria conta.")
    if perde_admin and _admins_ativos(conn) <= 1:
        raise ErroDeNegocio("Este é o último administrador ativo: não pode ser rebaixado nem inativado.")
    conn.execute("UPDATE usuarios SET nome = ?, perfil = ?, ativo = ? WHERE id = ?", (nome, perfil, ativo, id))
    conn.commit()


def elegiveis_comissao(conn) -> list[dict]:
    """Usuários ativos que podem compor a comissão de um inventário (admin, operador, inventariante), por nome."""
    marcas = ",".join("?" * len(PERFIS_COMISSAO))
    return _todos(conn, f"SELECT {_COLUNAS_LISTA} FROM usuarios WHERE ativo = 1 AND perfil IN ({marcas}) ORDER BY nome, login",
                  *PERFIS_COMISSAO)
```

- [ ] **Step 5: Rodar os testes**

Run: `.venv/bin/pytest tests/test_usuarios.py -q`
Expected: 6 passed.

- [ ] **Step 6: Rodar a suíte inteira e commitar**

Run: `.venv/bin/pytest -q`
Expected: tudo verde (216 + 6).

```bash
git add db.py usuarios.py tests/test_usuarios.py
git commit -m "Usuários: tabela usuarios e cadastro básico (criar, listar, editar com travas, elegíveis para comissão)"
```

---

### Task 2: Autenticação, bloqueio, senhas e `criar-admin` (`usuarios.py`)

**Files:**
- Modify: `usuarios.py`
- Modify: `tests/test_usuarios.py`

**Interfaces:**
- Consumes: Task 1.
- Produces: `usuarios.autenticar(conn, login, senha) -> dict` (levanta `ErroDeNegocio` com "Usuário ou senha inválidos." ou "Muitas tentativas. Aguarde 15 minutos."), `usuarios.nova_senha_temporaria(conn, id) -> str`, `usuarios.trocar_senha(conn, id, atual, nova, confirmacao) -> None`, `usuarios.criar_admin(conn, login, nome, senha) -> int`, `usuarios.main(argv, ler_senha) -> int`, `usuarios._agora_dt()`.

- [ ] **Step 1: Escrever os testes que falham**

Acrescentar em `tests/test_usuarios.py`:

```python
from datetime import datetime, timedelta


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
    usuarios.editar(dados, usuarios.criar(dados, "b", "B", "Senha!234", "admin"), "B", "admin", ativo=True)
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
```

- [ ] **Step 2: Rodar para ver falhar**

Run: `.venv/bin/pytest tests/test_usuarios.py -q`
Expected: 5 falhas com `AttributeError: module 'usuarios' has no attribute 'autenticar'` (e afins).

- [ ] **Step 3: Implementar**

Acrescentar ao fim de `usuarios.py`:

```python
# ---------------------------------------------------------------- autenticação e senhas
_INVALIDO = "Usuário ou senha inválidos."


def autenticar(conn, login, senha) -> dict:
    """Devolve o usuário. Mensagem única para inexistente, inativo e senha errada. 5 falhas seguidas bloqueiam
    por 15 minutos (a senha certa também falha nesse período e não conta como falha)."""
    u = por_login(conn, login)
    if not u or not u["ativo"]:
        raise ErroDeNegocio(_INVALIDO)
    agora = _agora_dt()
    if u["bloqueado_ate"] and agora.strftime(_FORMATO) < u["bloqueado_ate"]:
        raise ErroDeNegocio(f"Muitas tentativas. Aguarde {BLOQUEIO_MINUTOS} minutos.")
    if not check_password_hash(u["senha_hash"], str(senha or "")):
        falhas = u["falhas"] + 1
        bloqueado = (agora + timedelta(minutes=BLOQUEIO_MINUTOS)).strftime(_FORMATO) if falhas >= MAX_FALHAS else None
        conn.execute("UPDATE usuarios SET falhas = ?, bloqueado_ate = ? WHERE id = ?", (falhas, bloqueado, u["id"]))
        conn.commit()
        raise ErroDeNegocio(_INVALIDO)
    conn.execute("UPDATE usuarios SET falhas = 0, bloqueado_ate = NULL, ultimo_acesso = ? WHERE id = ?",
                 (agora.strftime(_FORMATO), u["id"]))
    conn.commit()
    return por_id(conn, u["id"])


def nova_senha_temporaria(conn, id) -> str:
    """Gera, grava o hash e devolve a senha em claro UMA vez; obriga a troca e limpa bloqueio."""
    if not por_id(conn, id):
        raise ErroDeNegocio("Usuário não encontrado.")
    senha = "".join(secrets.choice(ALFABETO_TEMP) for _ in range(10))
    conn.execute("UPDATE usuarios SET senha_hash = ?, trocar_senha = 1, falhas = 0, bloqueado_ate = NULL WHERE id = ?",
                 (generate_password_hash(senha), id))
    conn.commit()
    return senha


def trocar_senha(conn, id, atual, nova, confirmacao) -> None:
    u = por_id(conn, id)
    if not u:
        raise ErroDeNegocio("Usuário não encontrado.")
    if not check_password_hash(u["senha_hash"], str(atual or "")):
        raise ErroDeNegocio("A senha atual não confere.")
    nova = _validar_senha(nova)
    if nova == str(atual):
        raise ErroDeNegocio("A nova senha precisa ser diferente da atual.")
    if nova != str(confirmacao or ""):
        raise ErroDeNegocio("A confirmação não confere com a nova senha.")
    conn.execute("UPDATE usuarios SET senha_hash = ?, trocar_senha = 0 WHERE id = ?", (generate_password_hash(nova), id))
    conn.commit()


def criar_admin(conn, login, nome, senha) -> int:
    """Primeiro administrador e socorro: se o login já existe, redefine a senha, volta a admin, reativa e
    desbloqueia (o nome não muda)."""
    u = por_login(conn, login)
    if not u:
        return criar(conn, login, nome, senha, "admin", trocar_senha=False)
    senha = _validar_senha(senha)
    conn.execute("""UPDATE usuarios SET senha_hash = ?, perfil = 'admin', ativo = 1, trocar_senha = 0, falhas = 0,
                    bloqueado_ate = NULL WHERE id = ?""", (generate_password_hash(senha), u["id"]))
    conn.commit()
    return u["id"]


def main(argv, ler_senha=None) -> int:
    """`python usuarios.py criar-admin <login> "<Nome>"`: pede a senha duas vezes no terminal."""
    import getpass
    import sys

    import db
    ler_senha = ler_senha or getpass.getpass
    if len(argv) != 3 or argv[0] != "criar-admin":
        print('Uso: python usuarios.py criar-admin <login> "<Nome completo>"', file=sys.stderr)
        return 2
    senha, repetida = ler_senha("Senha: "), ler_senha("Repita a senha: ")
    if senha != repetida:
        print("As senhas não conferem.", file=sys.stderr)
        return 1
    db.inicializar()
    conn = db.conectar()
    try:
        uid = criar_admin(conn, argv[1], argv[2], senha)
    except ErroDeNegocio as e:
        print(str(e), file=sys.stderr)
        return 1
    finally:
        conn.close()
    print(f"Administrador '{argv[1].lower()}' pronto (id {uid}).")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main(sys.argv[1:]))
```

- [ ] **Step 4: Rodar os testes**

Run: `.venv/bin/pytest tests/test_usuarios.py -q`
Expected: 11 passed.

- [ ] **Step 5: Suíte e commit**

Run: `.venv/bin/pytest -q`

```bash
git add usuarios.py tests/test_usuarios.py
git commit -m "Usuários: autenticação com bloqueio de 15 min na 5ª falha, senha temporária, troca de senha e comando criar-admin"
```

---

### Task 3: Matriz `PERMISSOES` e `permitido`

**Files:**
- Modify: `usuarios.py`
- Create: `tests/test_permissoes.py`

**Interfaces:**
- Produces: `usuarios.PERMISSOES: dict[str, frozenset]`, `usuarios.permitido(perfil, endpoint, metodo="GET") -> bool`, `usuarios.TODOS`, `usuarios.GESTAO`, `usuarios.ADMIN`, `usuarios.LEITURA`, `usuarios.TERMOS_VER`.
- Chave da matriz: nome do endpoint Flask (`"home"`, `"inventario.ler"`); para POST diferente de GET, chave `"<endpoint>:POST"`. `HEAD` conta como `GET`.

- [ ] **Step 1: Escrever os testes que falham**

Criar `tests/test_permissoes.py`:

```python
"""Matriz de permissões: unidade (permitido) e cobertura de todas as rotas do app."""
import pytest

import usuarios


def test_permitido_por_perfil():
    assert usuarios.permitido("consulta", "home") and usuarios.permitido("inventariante", "bem")
    assert usuarios.permitido("consulta", "termo") and usuarios.permitido("consulta", "termo_documento")
    assert not usuarios.permitido("inventariante", "termo") and not usuarios.permitido("inventariante", "centro_custos")
    assert not usuarios.permitido("consulta", "termo_docx") and not usuarios.permitido("consulta", "termo_registrar", "POST")
    assert usuarios.permitido("consulta", "termo_devolucao") and not usuarios.permitido("consulta", "termo_devolucao", "POST")
    assert usuarios.permitido("operador", "termo_devolucao", "POST")
    assert usuarios.permitido("operador", "responsaveis_editar", "POST") and not usuarios.permitido("operador", "responsaveis_excluir", "POST")
    assert usuarios.permitido("operador", "upload", "POST") and not usuarios.permitido("operador", "importar_cadastros", "POST")
    assert not usuarios.permitido("inventariante", "cadastros") and not usuarios.permitido("inventariante", "textos_tela")
    assert usuarios.permitido("inventariante", "inventario.ler", "POST") and not usuarios.permitido("consulta", "inventario.ler", "POST")
    assert not usuarios.permitido("operador", "inventario.abrir", "POST") and usuarios.permitido("admin", "inventario.excluir", "POST")
    assert not usuarios.permitido("operador", "usuarios.lista") and usuarios.permitido("admin", "usuarios.lista")
    for perfil in usuarios.PERFIS:
        assert usuarios.permitido(perfil, "usuarios.senha", "POST") and usuarios.permitido(perfil, "usuarios.sair", "POST")
    assert usuarios.permitido("admin", "termo_docx", "HEAD")                 # HEAD conta como GET
    assert not usuarios.permitido("admin", "rota_que_nao_existe")            # fora da matriz = negado a todos
    assert not usuarios.permitido("chefe", "home")


def test_toda_rota_do_app_esta_na_matriz():
    from app import app
    faltam = []
    for regra in app.url_map.iter_rules():
        if regra.endpoint == "static":
            continue
        for metodo in regra.methods - {"HEAD", "OPTIONS"}:
            chave = regra.endpoint if metodo == "GET" else f"{regra.endpoint}:{metodo}"
            if chave not in usuarios.PERMISSOES and regra.endpoint not in usuarios.PERMISSOES:
                faltam.append(f"{regra.endpoint} [{metodo}]")
    assert not faltam, "Rotas sem regra em usuarios.PERMISSOES: " + ", ".join(sorted(faltam))
```

- [ ] **Step 2: Rodar para ver falhar**

Run: `.venv/bin/pytest tests/test_permissoes.py -q`
Expected: `AttributeError: module 'usuarios' has no attribute 'permitido'`.

- [ ] **Step 3: Implementar a matriz**

Acrescentar em `usuarios.py`, antes de `# ---- autenticação e senhas`:

```python
# ---------------------------------------------------------------- permissões (nega por padrão)
TODOS = frozenset(PERFIS)
GESTAO = frozenset({"admin", "operador"})
ADMIN = frozenset({"admin"})
LEITURA = frozenset({"admin", "operador", "inventariante"})     # ler no inventário: exige ainda estar na comissão
TERMOS_VER = frozenset({"admin", "operador", "consulta"})       # telas de termos: inventariante não vê

# Chave = endpoint Flask; "<endpoint>:POST" quando o POST tem regra diferente do GET. Rota ausente = 403 para todos
# (tests/test_permissoes.py garante que toda rota do app está aqui).
PERMISSOES = {
    # consulta geral
    "home": TODOS, "bem": TODOS, "pesquisa": TODOS, "recorte": TODOS, "recorte_xlsx": TODOS,
    # termos: ver para admin, operador e consulta; emitir/registrar só gestão
    "centro_custos": TERMOS_VER, "termos_individuais": TERMOS_VER, "termo": TERMOS_VER, "termo_documento": TERMOS_VER,
    "termo_devolucao": TERMOS_VER, "termo_devolucao:POST": GESTAO,
    "termos_emitidos_tela": TERMOS_VER, "termo_emitido_tela": TERMOS_VER,
    "gerar": GESTAO, "gerar_individual": GESTAO, "termo_docx": GESTAO, "termo_planilha": GESTAO, "termo_registrar": GESTAO,
    "termo_emitido_documento": GESTAO, "termo_emitido_email": GESTAO,
    # cadastros: gestão inclui/edita; só admin exclui
    "cadastros": GESTAO, "cadastro_novo": GESTAO,
    "responsaveis_incluir": GESTAO, "responsaveis_editar": GESTAO, "responsaveis_excluir": ADMIN,
    "pessoas_incluir": GESTAO, "pessoas_editar": GESTAO, "pessoas_excluir": ADMIN,
    "pessoas_atribuir": GESTAO, "pessoas_desatribuir": GESTAO,
    "localizacoes_incluir": GESTAO, "localizacoes_alterar": GESTAO, "localizacoes_mover": GESTAO, "localizacoes_excluir": ADMIN,
    "processos_incluir": GESTAO, "processos_vigente": GESTAO, "processos_encerrar": GESTAO, "processos_excluir": ADMIN,
    "cadastros_exportar": GESTAO, "importar_cadastros": ADMIN,
    # textos e base
    "textos_tela": GESTAO, "textos_salvar": GESTAO,
    "upload": GESTAO, "bens_exportar": GESTAO, "importacao_tela": GESTAO,
    # inventário
    "inventario.eventos_tela": TODOS, "inventario.evento_tela": TODOS, "inventario.sala_tela": TODOS,
    "inventario.relatorio_tela": TODOS, "inventario.xlsx": TODOS, "inventario.painel_tela": TODOS,
    "inventario.abrir": ADMIN, "inventario.encerrar": ADMIN, "inventario.comissao": ADMIN, "inventario.excluir": ADMIN,
    "inventario.ler": LEITURA, "inventario.atualizar_leitura": LEITURA, "inventario.lote": LEITURA,
    "inventario.foto_leitura": LEITURA, "inventario.foto_excluir": LEITURA, "inventario.sobra": LEITURA, "inventario.sobra_excluir": LEITURA,
    # conta e usuários
    "usuarios.login": TODOS, "usuarios.sair": TODOS, "usuarios.senha": TODOS,
    "usuarios.lista": ADMIN, "usuarios.novo": ADMIN, "usuarios.incluir": ADMIN, "usuarios.editar": ADMIN, "usuarios.nova_senha": ADMIN,
}


def permitido(perfil, endpoint, metodo="GET") -> bool:
    metodo = "GET" if metodo in ("GET", "HEAD") else metodo
    regra = PERMISSOES.get(f"{endpoint}:{metodo}") if metodo != "GET" else None
    if regra is None:
        regra = PERMISSOES.get(endpoint)
    return regra is not None and perfil in regra
```

- [ ] **Step 4: Rodar os testes**

Run: `.venv/bin/pytest tests/test_permissoes.py -q`
Expected: 2 passed. Se `test_toda_rota_do_app_esta_na_matriz` listar `inventario.integrante`, é esperado até a Task 7: acrescente temporariamente `"inventario.integrante": LEITURA,` na matriz e remova na Task 7.

- [ ] **Step 5: Suíte e commit**

Run: `.venv/bin/pytest -q`

```bash
git add usuarios.py tests/test_permissoes.py
git commit -m "Usuários: matriz PERMISSOES por endpoint (nega por padrão) e teste que exige regra para toda rota"
```

---

### Task 4: Login, sair, troca de senha e modo desktop (`app_usuarios.py`, `app.py`, `config.py`, fixtures)

**Files:**
- Modify: `config.py` (fim do arquivo)
- Create: `app_usuarios.py`
- Create: `templates/login.html`, `templates/senha.html`
- Modify: `app.py` (imports, config, `contexto_dsgov`, novo `before_request`, registro do blueprint)
- Modify: `templates/base.html` (bloco `header-actions`)
- Modify: `tests/conftest.py`, `tests/test_app.py:12-19`, `tests/test_cadastros_ux.py:10-17`
- Create: `tests/test_login.py`

**Interfaces:**
- Consumes: Tasks 1–3.
- Produces: `config.exigir_login() -> bool`; `g.usuario` (dict) em toda requisição fora de `static`; endpoints `usuarios.login` (GET/POST `/login`), `usuarios.sair` (POST `/sair`), `usuarios.senha` (GET/POST `/senha`); `app.ROTAS_JSON`; helper de teste `tests.conftest.logar(cliente, login, senha)`; constantes `tests.conftest.ADMIN_LOGIN = "admin"`, `ADMIN_NOME = "Fulano"`, `ADMIN_SENHA = "Senha!234"`; fixture `cliente` (admin "Fulano" logado, login ligado) e `cliente_local` (sem `TERMOS_LOGIN`); fixture `usuarios_exemplo` (dict `{"operador": ("op", "Senha!234"), "inventariante": ("beltrana", ...), "consulta": ("leitor", ...)}`); contexto Jinja `USUARIO`, `ROTULO_PERFIL`.
- Nesta tarefa o `before_request` ainda NÃO checa `permitido` (Task 5) nem CSRF (Task 9).

- [ ] **Step 1: `config.exigir_login`**

Acrescentar ao fim de `config.py`:

```python
def exigir_login() -> bool:
    """Web (compose.yml define TERMOS_LOGIN=1): pede usuário e senha. Desktop e testes por padrão: entra direto
    como administrador local."""
    return os.environ.get("TERMOS_LOGIN") == "1"
```

- [ ] **Step 2: Fixtures compartilhadas**

Substituir a fixture `cliente` de `tests/test_app.py` (linhas 12–19) e de `tests/test_cadastros_ux.py` (linhas 10–17) por uma única em `tests/conftest.py`. Apagar as duas definições locais (e o `import pytest` de `test_cadastros_ux.py` só se ficar sem uso). Acrescentar em `tests/conftest.py`:

```python
ADMIN_LOGIN, ADMIN_NOME, ADMIN_SENHA = "admin", "Fulano", "Senha!234"
SENHA_PADRAO = "Senha!234"


def logar(cliente, login, senha):
    """POST no /login como o formulário faria; devolve a resposta (302 para / no acerto)."""
    return cliente.post("/login", data={"login": login, "senha": senha})


def _app_de_teste():
    from app import app
    app.config["TESTING"] = True
    app.config["SESSION_COOKIE_SECURE"] = False     # o test client fala http
    return app


@pytest.fixture
def cliente(dados, monkeypatch):
    """Login ligado (como na web), admin 'Fulano' criado e logado, mais a inventariante 'Beltrana'
    (os dois nomes que os testes do inventário sempre usaram como comissão)."""
    import usuarios
    semear(dados)
    monkeypatch.setenv("TERMOS_LOGIN", "1")
    usuarios.criar(dados, ADMIN_LOGIN, ADMIN_NOME, ADMIN_SENHA, "admin", trocar_senha=False)
    usuarios.criar(dados, "beltrana", "Beltrana", SENHA_PADRAO, "inventariante", trocar_senha=False)
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
    """Um usuário de cada perfil além do admin: login → (login, senha)."""
    import usuarios
    usuarios.criar(dados, "op", "Operador Teste", SENHA_PADRAO, "operador", trocar_senha=False)
    usuarios.criar(dados, "leitor", "Consulta Teste", SENHA_PADRAO, "consulta", trocar_senha=False)
    return {"admin": (ADMIN_LOGIN, ADMIN_SENHA), "operador": ("op", SENHA_PADRAO),
            "inventariante": ("beltrana", SENHA_PADRAO), "consulta": ("leitor", SENHA_PADRAO)}
```

Em `tests/test_app.py`, trocar `from tests.conftest import semear, confirmar_revisao` por `from tests.conftest import semear, confirmar_revisao, logar, ADMIN_LOGIN, ADMIN_NOME, ADMIN_SENHA`.

- [ ] **Step 3: Testes de login que falham**

Criar `tests/test_login.py`:

```python
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
    usuarios.editar(dados, usuarios.por_login(dados, "beltrana")["id"], "Beltrana", "inventariante", ativo=False)
    r = cliente.post("/login", data={"login": "beltrana", "senha": SENHA_PADRAO})
    assert "inválidos".encode() in r.data


def test_usuario_inativado_com_sessao_aberta_cai_no_login(cliente, dados):
    logar(cliente, "beltrana", SENHA_PADRAO)
    assert cliente.get("/").status_code == 200
    usuarios.editar(dados, usuarios.por_login(dados, "beltrana")["id"], "Beltrana", "inventariante", ativo=False)
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
```

- [ ] **Step 4: Rodar para ver falhar**

Run: `.venv/bin/pytest tests/test_login.py -q`
Expected: falhas (404 em `/login`, `/sair`; `Sair` ausente).

- [ ] **Step 5: Blueprint `app_usuarios.py` (só login, sair, senha nesta tarefa)**

```python
"""Rotas de conta e de usuários (blueprint `usuarios`). Regras em usuarios.py; g.usuario é resolvido em app.py."""
from urllib.parse import urlsplit

from flask import Blueprint, abort, flash, g, redirect, render_template, request, session, url_for

import config
import db
import usuarios

usuarios_bp = Blueprint("usuarios", __name__)


def _conn():
    if "conn" not in g:
        g.conn = db.conectar()
    return g.conn


def _proximo_seguro(valor: str | None) -> str:
    """Só caminhos relativos do próprio site: '/x?y=1'. '//host', 'http://…' e '/\\host' caem em '/'."""
    v = (valor or "").strip()
    if not v.startswith("/") or v.startswith("//") or v.startswith("/\\"):
        return url_for("home")
    partes = urlsplit(v)
    if partes.scheme or partes.netloc:
        return url_for("home")
    return v


@usuarios_bp.before_request
def _so_com_login_ligado():
    """No desktop (sem TERMOS_LOGIN) não há conta: estas telas não existem."""
    if not config.exigir_login():
        abort(404)


@usuarios_bp.route("/login", methods=["GET", "POST"])
def login():
    conn = _conn()
    if g.usuario:
        return redirect(url_for("home"))
    proximo = request.args.get("proximo")
    sem_usuarios = conn.execute("SELECT count(*) FROM usuarios").fetchone()[0] == 0
    if request.method == "POST" and not sem_usuarios:
        try:
            u = usuarios.autenticar(conn, request.form.get("login", ""), request.form.get("senha", ""))
        except db.ErroDeNegocio as e:
            return render_template("login.html", erro=str(e), login=request.form.get("login", ""), proximo=proximo, sem_usuarios=False)
        session.clear()
        session["usuario_id"] = u["id"]
        session.permanent = True
        return redirect(_proximo_seguro(proximo))
    return render_template("login.html", erro=None, login="", proximo=proximo, sem_usuarios=sem_usuarios)


@usuarios_bp.route("/sair", methods=["POST"])
def sair():
    session.clear()
    return redirect(url_for("usuarios.login"))


@usuarios_bp.route("/senha", methods=["GET", "POST"])
def senha():
    if request.method == "POST":
        try:
            usuarios.trocar_senha(_conn(), g.usuario["id"], request.form.get("atual"), request.form.get("nova"), request.form.get("confirmacao"))
        except db.ErroDeNegocio as e:
            return render_template("senha.html", erro=str(e), obrigatoria=bool(g.usuario["trocar_senha"]), trilha=[("Trocar senha", None)])
        flash("Senha alterada.", "success")
        return redirect(url_for("home"))
    return render_template("senha.html", erro=None, obrigatoria=bool(g.usuario["trocar_senha"]), trilha=[("Trocar senha", None)])
```

- [ ] **Step 6: Templates `login.html` e `senha.html`**

`templates/login.html` (página própria, sem menu; mesmos CSS de `base.html`):

```html
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Entrar — {{ DSGOV.SISTEMA }}</title>
  <link rel="stylesheet" href="{{ url_for('static', filename='dsgov/css/fontes.css') }}"/>
  <link rel="stylesheet" href="{{ url_for('static', filename='dsgov/vendor/govbr-ds/core.min.css') }}"/>
  <link rel="stylesheet" href="{{ url_for('static', filename='dsgov/vendor/fontawesome/css/all.min.css') }}"/>
  <link rel="stylesheet" href="{{ url_for('static', filename='dsgov/css/dsgov.css') }}"/>
</head>
<body>
<main class="container-lg d-flex justify-content-center align-items-center" style="min-height: 100vh">
  <div class="br-card" style="width: 100%; max-width: 420px">
    <div class="card-header">
      <div class="d-flex align-items-center"><img src="{{ url_for('static', filename='logo.png') }}" alt="{{ DSGOV.ORGAO }}" style="height: 40px" class="mr-3"/>
        <div><div class="text-weight-semi-bold text-up-01">{{ DSGOV.SISTEMA }}</div><div class="text-down-01 text-gray-70">{{ DSGOV.ORGAO }}</div></div></div>
    </div>
    <div class="card-content">
      {% if sem_usuarios %}
      <div class="br-message warning"><div class="icon"><i class="fas fa-exclamation-triangle fa-lg" aria-hidden="true"></i></div>
        <div class="content" role="alert"><span class="message-title">Nenhum usuário cadastrado.</span><span class="message-body"> Crie o primeiro administrador no servidor:
          <code>python usuarios.py criar-admin &lt;login&gt; "&lt;Nome&gt;"</code> (na VPS: <code>docker compose exec web python usuarios.py criar-admin …</code>).</span></div></div>
      {% else %}
      {% if erro %}<div class="br-message danger mb-3"><div class="icon"><i class="fas fa-times-circle fa-lg" aria-hidden="true"></i></div>
        <div class="content" role="alert"><span class="message-body">{{ erro }}</span></div></div>{% endif %}
      <form method="post" action="{{ url_for('usuarios.login', proximo=proximo) if proximo else url_for('usuarios.login') }}">
        <div class="br-input mb-3"><label for="login">Usuário</label><input id="login" name="login" type="text" autocomplete="username" autocapitalize="none" value="{{ login }}" required autofocus/></div>
        <div class="br-input mb-4"><label for="senha">Senha</label><input id="senha" name="senha" type="password" autocomplete="current-password" required/></div>
        <button class="br-button primary block" type="submit"><i class="fas fa-sign-in-alt mr-1" aria-hidden="true"></i>Entrar</button>
      </form>
      {% endif %}
    </div>
  </div>
</main>
</body>
</html>
```

`templates/senha.html`:

```html
{% extends "base.html" %}
{% block titulo %}Trocar senha{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Trocar senha</h1></div>
{% if obrigatoria %}
<div class="br-message warning mb-3"><div class="icon"><i class="fas fa-key fa-lg" aria-hidden="true"></i></div>
  <div class="content" role="alert"><span class="message-title">Defina uma nova senha para continuar.</span><span class="message-body"> Sua senha é temporária.</span></div></div>
{% endif %}
{% if erro %}<div class="br-message danger mb-3"><div class="icon"><i class="fas fa-times-circle fa-lg" aria-hidden="true"></i></div>
  <div class="content" role="alert"><span class="message-body">{{ erro }}</span></div></div>{% endif %}
<form method="post" action="{{ url_for('usuarios.senha') }}" class="col-md-6">
  <div class="br-input mb-3"><label for="atual">Senha atual</label><input id="atual" name="atual" type="password" autocomplete="current-password" required/></div>
  <div class="br-input mb-3"><label for="nova">Nova senha (mínimo 8 caracteres)</label><input id="nova" name="nova" type="password" autocomplete="new-password" minlength="8" required/></div>
  <div class="br-input mb-4"><label for="confirmacao">Repita a nova senha</label><input id="confirmacao" name="confirmacao" type="password" autocomplete="new-password" minlength="8" required/></div>
  <button class="br-button primary" type="submit"><i class="fas fa-save mr-1" aria-hidden="true"></i>Salvar nova senha</button>
  {% if not obrigatoria %}<a class="br-button ml-2" href="{{ url_for('home') }}">Cancelar</a>{% endif %}
</form>
{% endblock %}
```

- [ ] **Step 7: `app.py`: sessão, `g.usuario`, redirecionamentos, contexto**

Em `app.py`:

1. Imports: acrescentar `from datetime import timedelta` e `import usuarios`, e `from app_usuarios import usuarios_bp`.
2. Logo após `app.secret_key = ...`:

```python
app.config.update(PERMANENT_SESSION_LIFETIME=timedelta(hours=12), SESSION_COOKIE_HTTPONLY=True,
                  SESSION_COOKIE_SAMESITE="Lax", SESSION_COOKIE_SECURE=config.exigir_login())   # site é só https
app.register_blueprint(usuarios_bp)
ROTAS_JSON = {"inventario.ler", "inventario.atualizar_leitura", "inventario.foto_leitura", "termo_registrar"}
```

3. Novo `before_request`, colocado antes de `obter_conn` (que ele usa; a função pode ficar depois porque a resolução é em tempo de chamada):

```python
@app.before_request
def resolver_usuario():
    """Quem está usando: sessão (web, TERMOS_LOGIN=1) ou o administrador local (desktop). Sem sessão válida
    → /login. Com senha temporária → /senha até trocar."""
    ep = request.endpoint
    if ep is None or ep == "static":
        return None
    if not config.exigir_login():
        g.usuario = usuarios.USUARIO_LOCAL
        return None
    u = usuarios.por_id(obter_conn(), session.get("usuario_id")) if session.get("usuario_id") else None
    if u is None or not u["ativo"]:
        session.clear()
        g.usuario = None
        if ep == "usuarios.login":
            return None
        if request.method == "GET":
            return redirect(url_for("usuarios.login", proximo=request.full_path.rstrip("?")))
        return redirect(url_for("usuarios.login"))
    g.usuario = u
    if u["trocar_senha"] and ep not in ("usuarios.senha", "usuarios.sair", "usuarios.login"):
        return redirect(url_for("usuarios.senha"))
    return None
```

4. Em `contexto_dsgov`, no início do corpo: `usuario = getattr(g, "usuario", None)`; devolver também `"USUARIO": usuario, "ROTULO_PERFIL": usuarios.ROTULO_PERFIL`. (O filtro do menu por perfil é a Task 5.)

- [ ] **Step 8: Cabeçalho com usuário e Sair**

Em `templates/base.html`, dentro de `<div class="header-actions">`, antes de `<div class="header-search-trigger">`:

```html
          {% if USUARIO and USUARIO.id %}
          <div class="header-login">
            <div class="header-sign-in d-flex align-items-center">
              <span class="text-down-01 mr-2"><i class="fas fa-user-circle mr-1" aria-hidden="true"></i>{{ USUARIO.nome }} · {{ ROTULO_PERFIL[USUARIO.perfil] }}</span>
              <a class="br-button small" href="{{ url_for('usuarios.senha') }}">Trocar senha</a>
              <form method="post" action="{{ url_for('usuarios.sair') }}" class="d-inline ml-1">
                <button class="br-button small secondary" type="submit"><i class="fas fa-sign-out-alt mr-1" aria-hidden="true"></i>Sair</button>
              </form>
            </div>
          </div>
          {% endif %}
```

- [ ] **Step 9: Rodar os testes**

Run: `.venv/bin/pytest tests/test_login.py tests/test_app.py tests/test_cadastros_ux.py -q`
Expected: tudo verde. Atenção: `test_toda_rota_do_app_esta_na_matriz` passa a exigir `usuarios.login`, `usuarios.sair`, `usuarios.senha`, já na matriz.

- [ ] **Step 10: Suíte e commit**

Run: `.venv/bin/pytest -q`

```bash
git add config.py app_usuarios.py app.py templates/login.html templates/senha.html templates/base.html tests/conftest.py tests/test_app.py tests/test_cadastros_ux.py tests/test_login.py
git commit -m "Login próprio: /login, /sair, /senha, sessão de 12 h, troca obrigatória, cabeçalho com usuário; desktop entra como administrador local; fixtures com admin logado"
```

---

### Task 5: Autorização por perfil (403, menu, botões de emissão)

**Files:**
- Modify: `app.py` (`resolver_usuario`, `contexto_dsgov`, handler 403)
- Create: `templates/403.html`
- Modify: `templates/termo.html:6-12`
- Modify: `tests/test_permissoes.py`

**Interfaces:**
- Consumes: `usuarios.permitido`, `g.usuario`, `app.ROTAS_JSON`.
- Produces: 403 HTML (`403.html`) ou JSON `{"erro": "Seu perfil não tem acesso a isso."}`; menu filtrado; contexto Jinja `pode(endpoint, metodo="GET") -> bool`.

- [ ] **Step 1: Testes que falham**

Acrescentar em `tests/test_permissoes.py`:

```python
from tests.conftest import logar, SENHA_PADRAO

# (endpoint GET exemplar por área, e o que cada perfil deve receber)
_ROTAS_GET = {
    "/": {"admin": 200, "operador": 200, "inventariante": 200, "consulta": 200},
    "/bem?numero=1001": {"admin": 200, "operador": 200, "inventariante": 200, "consulta": 200},
    "/centro-custos": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 200},
    "/termo/ccusto/CCI": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 200},
    "/termo/ccusto/CCI/documento": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 200},
    "/termo/ccusto/CCI/planilha": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 403},
    "/cadastros/responsaveis": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 403},
    "/textos": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 403},
    "/upload": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 403},
    "/inventario": {"admin": 200, "operador": 200, "inventariante": 200, "consulta": 200},
    "/usuarios": {"admin": 200, "operador": 403, "inventariante": 403, "consulta": 403},
}
_ROTAS_POST = {
    "/gerar": {"admin": 302, "operador": 302, "inventariante": 403, "consulta": 403},
    "/termo_devolucao": {"admin": 302, "operador": 302, "inventariante": 403, "consulta": 403},
    "/cadastros/responsaveis/excluir": {"admin": 302, "operador": 403, "inventariante": 403, "consulta": 403},
    "/importar-cadastros": {"admin": 302, "operador": 403, "inventariante": 403, "consulta": 403},
    "/inventario/abrir": {"admin": 302, "operador": 403, "inventariante": 403, "consulta": 403},
}


@pytest.mark.parametrize("perfil", ["admin", "operador", "inventariante", "consulta"])
def test_rotas_por_perfil(cliente, usuarios_exemplo, perfil):
    login, senha = usuarios_exemplo[perfil]
    cliente.post("/sair")
    assert logar(cliente, login, senha).status_code == 302
    for rota, esperado in _ROTAS_GET.items():
        r = cliente.get(rota)
        assert r.status_code == esperado[perfil], f"{perfil} GET {rota}: {r.status_code} != {esperado[perfil]}"
        if esperado[perfil] == 403:
            assert "Seu perfil não tem acesso a isso".encode() in r.data and b"main-navigation" in r.data
    for rota, esperado in _ROTAS_POST.items():
        r = cliente.post(rota, data={"ccusto": "CCI", "nome": "ANA SILVA", "limpar": "1"})
        assert r.status_code == esperado[perfil], f"{perfil} POST {rota}: {r.status_code} != {esperado[perfil]}"


def test_403_em_json_para_rotas_do_leitor(cliente, usuarios_exemplo):
    cliente.post("/sair")
    logar(cliente, *usuarios_exemplo["consulta"])
    r = cliente.post("/inventario/1/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    assert r.status_code == 403 and r.get_json()["erro"].startswith("Seu perfil")
    r = cliente.post("/termo/ccusto/CCI/registrar")
    assert r.status_code == 403 and r.get_json()["erro"].startswith("Seu perfil")


def test_menu_por_perfil(cliente, usuarios_exemplo):
    def menu():
        return cliente.get("/").data.split(b'id="main-navigation"')[1].split(b"menu-footer")[0]
    m = menu()
    assert b">Usu\xc3\xa1rios<" in m and b">Textos<" in m and b">Cadastros<" in m
    cliente.post("/sair"); logar(cliente, *usuarios_exemplo["operador"])
    m = menu()
    assert b">Usu\xc3\xa1rios<" not in m and b">Textos<" in m and b">Cadastros<" in m and b">Atualizar base<" in m
    cliente.post("/sair"); logar(cliente, *usuarios_exemplo["inventariante"])
    m = menu()
    assert b">Cadastros<" not in m and b">Termo por centro de custo<" not in m and b">Invent\xc3\xa1rio<" in m and b">Recorte<" in m
    cliente.post("/sair"); logar(cliente, *usuarios_exemplo["consulta"])
    m = menu()
    assert b">Termo por centro de custo<" in m and b">Textos<" not in m and b">Atualizar base<" not in m


def test_consulta_nao_ve_botoes_de_emissao(cliente, usuarios_exemplo):
    cliente.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "T", "numero_sei": "1111", "vigente": "1"})
    assert b"Copiar para o SEI" in cliente.get("/termo/ccusto/CCI").data
    cliente.post("/sair"); logar(cliente, *usuarios_exemplo["consulta"])
    r = cliente.get("/termo/ccusto/CCI")
    assert r.status_code == 200 and b"Copiar para o SEI" not in r.data and b"Baixar .docx" not in r.data and b"Baixar planilha" not in r.data
    assert b"Processo SEI 1111" in r.data


def test_modo_desktop_sem_tela_de_usuarios(cliente_local):
    assert cliente_local.get("/usuarios").status_code == 404
    m = cliente_local.get("/").data.split(b'id="main-navigation"')[1].split(b"menu-footer")[0]
    assert b">Usu\xc3\xa1rios<" not in m and b">Textos<" in m
```

Nota: `/usuarios` só existe na Task 6; até lá, este teste espera 404 para admin. Ajuste temporário: deixe `"/usuarios"` fora de `_ROTAS_GET` nesta tarefa e acrescente-o na Task 6 (o plano repete a linha lá).

- [ ] **Step 2: Rodar para ver falhar**

Run: `.venv/bin/pytest tests/test_permissoes.py -q`
Expected: falhas com 200 onde se esperava 403.

- [ ] **Step 3: Checagem de permissão e 403**

Em `app.py`, no fim de `resolver_usuario` (antes do `return None` final), acrescentar:

```python
    if not usuarios.permitido(g.usuario["perfil"], ep, request.method):
        return _negado()
```

E, no ramo desktop (`if not config.exigir_login():`), trocar `return None` por:

```python
        g.usuario = usuarios.USUARIO_LOCAL
        return None if usuarios.permitido("admin", ep, request.method) else _negado()
```

Definir, antes de `resolver_usuario`:

```python
NEGADO = "Seu perfil não tem acesso a isso."


def _negado():
    if request.endpoint in ROTAS_JSON or request.is_json:
        return {"erro": NEGADO}, 403
    return render_template("403.html", trilha=[("Acesso negado", None)]), 403
```

`templates/403.html`:

```html
{% extends "base.html" %}
{% block titulo %}Acesso negado{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Acesso negado</h1></div>
<div class="br-message danger"><div class="icon"><i class="fas fa-ban fa-lg" aria-hidden="true"></i></div>
  <div class="content" role="alert"><span class="message-title">Seu perfil não tem acesso a isso.</span><span class="message-body"> Se precisar desta função, peça ao administrador para ajustar seu perfil.</span></div></div>
<a class="br-button mt-3" href="{{ url_for('home') }}"><i class="fas fa-home mr-1" aria-hidden="true"></i>Início</a>
{% endblock %}
```

- [ ] **Step 4: Menu por perfil**

Reescrever `contexto_dsgov` em `app.py`:

```python
@app.context_processor
def contexto_dsgov():
    t = textos.obter(obter_conn())
    dsgov = dict(DSGOV_FIXO, ORGAO=t["orgao_nome"], SUBTITULO=t["unidade_sigla"])
    usuario = getattr(g, "usuario", None)
    contexto = {"DSGOV": dsgov, "USUARIO": usuario, "ROTULO_PERFIL": usuarios.ROTULO_PERFIL, "MENU": []}
    if not usuario:
        return contexto
    perfil = usuario["perfil"]

    def pode(endpoint, metodo="GET"):
        return usuarios.permitido(perfil, endpoint, metodo)

    contexto["pode"] = pode
    inv = [("Eventos", url_for("inventario.eventos_tela"))]
    if (e := inventario.evento_aberto(obter_conn())):
        inv += [(e["nome"], url_for("inventario.evento_tela", id=e["id"])),
                ("Painel", url_for("inventario.painel_tela", id=e["id"])),
                ("Relatório", url_for("inventario.relatorio_tela", id=e["id"]))]
    itens = [
        ("Início", "fa-home", "home", {}, []),
        ("Termo por centro de custo", "fa-building", "centro_custos", {}, []),
        ("Termo individual", "fa-user-check", "termos_individuais", {}, []),
        ("Termo de devolução", "fa-box-open", "termo_devolucao", {}, []),
        ("Termos emitidos", "fa-history", "termos_emitidos_tela", {}, []),
        ("Recorte", "fa-filter", "recorte", {}, []),
        ("Inventário", "fa-clipboard-check", "inventario.eventos_tela", {}, inv),
        ("Cadastros", "fa-address-book", "cadastros", {"aba": "responsaveis"}, []),
        ("Textos", "fa-file-signature", "textos_tela", {}, []),
        ("Atualizar base", "fa-upload", "upload", {}, []),
    ]
    if usuario["id"] is not None:                                   # web: há contas
        itens.append(("Usuários", "fa-users", "usuarios.lista", {}, []))
    contexto["MENU"] = [(rotulo, icone, url_for(endpoint, **kw), filhos) for rotulo, icone, endpoint, kw, filhos in itens
                        if pode(endpoint)]
    return contexto
```

Nota: `url_for("usuarios.lista")` só existe na Task 6. Para esta tarefa, deixe a linha de "Usuários" comentada com `# Task 6` e descomente lá; ou implemente a Task 6 imediatamente a seguir na mesma sessão antes de rodar a suíte. Escolha a primeira opção para manter cada commit verde.

- [ ] **Step 5: Botões de emissão só para quem pode**

Em `templates/termo.html`, trocar `{% if processo %}` (linha 6) por `{% if processo and pode('termo_docx') %}` e, no bloco `{% if not processo %}` (linha 14), manter. Assim consulta vê a página, o parágrafo do processo e o iframe do documento, sem botões.

- [ ] **Step 6: Rodar os testes**

Run: `.venv/bin/pytest tests/test_permissoes.py tests/test_login.py -q`
Expected: verde (com `/usuarios` fora de `_ROTAS_GET` e `test_modo_desktop_sem_tela_de_usuarios` esperando 404, que já é o comportamento sem a rota).

- [ ] **Step 7: Suíte e commit**

Run: `.venv/bin/pytest -q`

```bash
git add app.py templates/403.html templates/termo.html tests/test_permissoes.py
git commit -m "Autorização por perfil: 403 na própria tela (ou JSON), menu filtrado, botões de emissão só para quem emite"
```

---

### Task 6: Telas de usuários (`/usuarios`)

**Files:**
- Modify: `app_usuarios.py`
- Create: `templates/usuarios/lista.html`, `templates/usuarios/formulario.html`
- Modify: `app.py` (descomentar "Usuários" no menu)
- Create: `tests/test_usuarios_telas.py`
- Modify: `tests/test_permissoes.py` (acrescentar `"/usuarios"` em `_ROTAS_GET`)

**Interfaces:**
- Consumes: `usuarios.listar/criar/editar/por_id/nova_senha_temporaria`, `ROTULO_PERFIL`, `PERFIS`.
- Produces: endpoints `usuarios.lista` (GET `/usuarios`), `usuarios.novo` (GET `/usuarios/novo`), `usuarios.incluir` (POST `/usuarios/incluir`), `usuarios.editar` (GET/POST `/usuarios/<int:id>/editar`), `usuarios.nova_senha` (POST `/usuarios/<int:id>/nova-senha`).

- [ ] **Step 1: Testes que falham**

Criar `tests/test_usuarios_telas.py`:

```python
"""Telas de usuários (só admin): lista, novo, editar, nova senha temporária."""
import re

import usuarios
from tests.conftest import ADMIN_LOGIN, SENHA_PADRAO, logar


def test_lista_busca_filtro_e_inativos(cliente, dados, usuarios_exemplo):
    r = cliente.get("/usuarios")
    assert r.status_code == 200 and b">admin<" in r.data and b">op<" in r.data and b"Operador" in r.data
    usuarios.editar(dados, usuarios.por_login(dados, "leitor")["id"], "Consulta Teste", "consulta", ativo=False)
    assert b">leitor<" not in cliente.get("/usuarios").data
    assert b">leitor<" in cliente.get("/usuarios?inativos=1").data
    r = cliente.get("/usuarios?q=oper")
    assert b">op<" in r.data and b">admin<" not in r.data
    r = cliente.get("/usuarios?perfil=inventariante")
    assert b">beltrana<" in r.data and b">op<" not in r.data


def test_novo_usuario_valida_e_cria(cliente, dados):
    r = cliente.get("/usuarios/novo")
    assert r.status_code == 200 and b'name="login"' in r.data and b'name="trocar_senha"' in r.data
    r = cliente.post("/usuarios/incluir", data={"login": "Novo Login", "nome": "Novo", "perfil": "operador", "senha": "Senha!234", "confirmacao": "Senha!234"})
    assert r.status_code == 200 and "Login inválido".encode() in r.data and b'value="Novo"' in r.data      # dados mantidos
    r = cliente.post("/usuarios/incluir", data={"login": "novo", "nome": "Novo", "perfil": "operador", "senha": "Senha!234", "confirmacao": "Outra!234"})
    assert r.status_code == 200 and "confirmação".encode() in r.data
    r = cliente.post("/usuarios/incluir", data={"login": "novo", "nome": "Novo", "perfil": "operador", "senha": "Senha!234", "confirmacao": "Senha!234", "trocar_senha": "1"}, follow_redirects=True)
    assert "Usuário novo criado".encode() in r.data and b">novo<" in r.data
    u = usuarios.por_login(dados, "novo")
    assert u["perfil"] == "operador" and u["trocar_senha"] == 1
    r = cliente.post("/usuarios/incluir", data={"login": "semtroca", "nome": "S", "perfil": "consulta", "senha": "Senha!234", "confirmacao": "Senha!234"}, follow_redirects=True)
    assert usuarios.por_login(dados, "semtroca")["trocar_senha"] == 0


def test_editar_perfil_ativo_e_travas(cliente, dados, usuarios_exemplo):
    op = usuarios.por_login(dados, "op")["id"]
    r = cliente.get(f"/usuarios/{op}/editar")
    assert r.status_code == 200 and b">op<" in r.data and b'name="login"' not in r.data                       # login não editável
    r = cliente.post(f"/usuarios/{op}/editar", data={"nome": "Operador Editado", "perfil": "inventariante", "ativo": "1"}, follow_redirects=True)
    assert b"Operador Editado" in r.data and usuarios.por_id(dados, op)["perfil"] == "inventariante"
    r = cliente.post(f"/usuarios/{op}/editar", data={"nome": "Operador Editado", "perfil": "inventariante"}, follow_redirects=True)
    assert usuarios.por_id(dados, op)["ativo"] == 0 and b"inativo" in r.data.lower()
    eu = usuarios.por_login(dados, ADMIN_LOGIN)["id"]
    r = cliente.post(f"/usuarios/{eu}/editar", data={"nome": "Fulano", "perfil": "consulta", "ativo": "1"})
    assert r.status_code == 200 and "própria conta".encode() in r.data
    assert usuarios.por_id(dados, eu)["perfil"] == "admin"
    assert cliente.get("/usuarios/999/editar").status_code == 404


def test_nova_senha_temporaria_aparece_uma_vez(cliente, dados, usuarios_exemplo):
    op = usuarios.por_login(dados, "op")["id"]
    r = cliente.post(f"/usuarios/{op}/nova-senha")
    assert r.status_code == 200 and b"Anote agora" in r.data
    temp = re.search(rb'id="senha-temporaria">([A-Za-z0-9]{10})<', r.data).group(1).decode()
    assert temp not in cliente.get(f"/usuarios/{op}/editar").data.decode()
    assert usuarios.por_id(dados, op)["trocar_senha"] == 1
    cliente.post("/sair")
    assert logar(cliente, "op", temp).status_code == 302
    assert cliente.get("/").headers["Location"] == "/senha"


def test_lista_preserva_busca_ao_voltar(cliente, dados, usuarios_exemplo):
    op = usuarios.por_login(dados, "op")["id"]
    r = cliente.get(f"/usuarios/{op}/editar?retorno=%2Fusuarios%3Fq%3Doper")
    assert b'href="/usuarios?q=oper"' in r.data
    r = cliente.post(f"/usuarios/{op}/editar?retorno=%2Fusuarios%3Fq%3Doper", data={"nome": "Operador Teste", "perfil": "operador", "ativo": "1"})
    assert r.headers["Location"] == "/usuarios?q=oper"
```

Em `tests/test_permissoes.py`, acrescentar em `_ROTAS_GET`: `"/usuarios": {"admin": 200, "operador": 403, "inventariante": 403, "consulta": 403},`.

- [ ] **Step 2: Rodar para ver falhar**

Run: `.venv/bin/pytest tests/test_usuarios_telas.py -q`
Expected: 404 em `/usuarios`.

- [ ] **Step 3: Rotas**

Acrescentar em `app_usuarios.py`:

```python
def _retorno() -> str:
    return _proximo_seguro(request.args.get("retorno")) if request.args.get("retorno") else url_for("usuarios.lista")


def _form_usuario(u, valores, erro, senha_temporaria=None):
    return render_template("usuarios/formulario.html", u=u, valores=valores, erro=erro, perfis=usuarios.PERFIS,
                           rotulos=usuarios.ROTULO_PERFIL, retorno=_retorno(), senha_temporaria=senha_temporaria,
                           trilha=[("Usuários", url_for("usuarios.lista")), ((u["login"] if u else "Novo usuário"), None)])


@usuarios_bp.route("/usuarios")
def lista():
    q, perfil, inativos = request.args.get("q", ""), request.args.get("perfil") or None, request.args.get("inativos") == "1"
    return render_template("usuarios/lista.html", lista=usuarios.listar(_conn(), q, perfil, inativos), q=q, perfil=perfil,
                           inativos=inativos, perfis=usuarios.PERFIS, rotulos=usuarios.ROTULO_PERFIL, trilha=[("Usuários", None)])


@usuarios_bp.route("/usuarios/novo")
def novo():
    return _form_usuario(None, {"perfil": "consulta", "trocar_senha": "1"}, None)


@usuarios_bp.route("/usuarios/incluir", methods=["POST"])
def incluir():
    f = request.form
    valores = {k: f.get(k, "") for k in ("login", "nome", "perfil", "trocar_senha")}
    try:
        if f.get("senha", "") != f.get("confirmacao", ""):
            raise db.ErroDeNegocio("A confirmação não confere com a senha.")
        usuarios.criar(_conn(), f.get("login"), f.get("nome"), f.get("senha"), f.get("perfil"), trocar_senha=bool(f.get("trocar_senha")))
    except db.ErroDeNegocio as e:
        return _form_usuario(None, valores, str(e))
    flash(f"Usuário {f.get('login', '').strip().lower()} criado.", "success")
    return redirect(_retorno())


@usuarios_bp.route("/usuarios/<int:id>/editar", methods=["GET", "POST"])
def editar(id):
    conn = _conn()
    u = usuarios.por_id(conn, id) or abort(404)
    if request.method == "POST":
        f = request.form
        valores = {"nome": f.get("nome", ""), "perfil": f.get("perfil", ""), "ativo": f.get("ativo")}
        try:
            usuarios.editar(conn, id, f.get("nome"), f.get("perfil"), ativo=bool(f.get("ativo")), logado_id=g.usuario["id"])
        except db.ErroDeNegocio as e:
            return _form_usuario(u, valores, str(e))
        flash(f"Usuário {u['login']} salvo" + ("" if f.get("ativo") else " (inativo)") + ".", "success")
        return redirect(_retorno())
    return _form_usuario(u, {"nome": u["nome"], "perfil": u["perfil"], "ativo": "1" if u["ativo"] else None}, None)


@usuarios_bp.route("/usuarios/<int:id>/nova-senha", methods=["POST"])
def nova_senha(id):
    conn = _conn()
    u = usuarios.por_id(conn, id) or abort(404)
    senha = usuarios.nova_senha_temporaria(conn, id)
    u = usuarios.por_id(conn, id)
    return _form_usuario(u, {"nome": u["nome"], "perfil": u["perfil"], "ativo": "1" if u["ativo"] else None}, None, senha_temporaria=senha)
```

- [ ] **Step 4: Templates**

`templates/usuarios/lista.html`:

```html
{% extends "base.html" %}
{% from "_macros.html" import cabecalho_tabela %}
{% block titulo %}Usuários{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Usuários</h1>
  <a class="br-button primary ml-auto" href="{{ url_for('usuarios.novo', retorno=request.full_path) }}"><i class="fas fa-user-plus mr-1" aria-hidden="true"></i>Novo usuário</a></div>
<form method="get" class="br-card mb-3"><div class="card-content row align-items-end">
  <div class="col-md-4 mb-2"><div class="br-input"><label for="q">Login ou nome</label><input id="q" name="q" type="text" value="{{ q }}"/></div></div>
  <div class="col-md-3 mb-2"><div class="br-select"><div class="br-input"><label for="perfil">Perfil</label>
    <select id="perfil" name="perfil" class="br-input"><option value="">Todos</option>{% for p in perfis %}<option value="{{ p }}"{% if p == perfil %} selected{% endif %}>{{ rotulos[p] }}</option>{% endfor %}</select></div></div></div>
  <div class="col-md-3 mb-2"><div class="br-checkbox"><input id="inativos" name="inativos" type="checkbox" value="1"{% if inativos %} checked{% endif %}/><label for="inativos">Mostrar inativos</label></div></div>
  <div class="col-md-2 mb-2"><button class="br-button secondary" type="submit"><i class="fas fa-search mr-1" aria-hidden="true"></i>Filtrar</button></div>
</div></form>
{{ cabecalho_tabela('Usuários cadastrados', 'usuarios') }}
  <thead><tr><th scope="col">Login</th><th scope="col">Nome</th><th scope="col">Perfil</th><th scope="col">Situação</th><th scope="col">Último acesso</th><th scope="col" class="dsgov-acoes">Ações</th></tr></thead>
  <tbody>{% for u in lista %}
  <tr><td><code>{{ u.login }}</code></td><td>{{ u.nome }}</td><td>{{ rotulos[u.perfil] }}</td>
    <td>{% if u.ativo %}<span class="br-tag bg-success text-pure-0"><span>ativo</span></span>{% else %}<span class="br-tag bg-gray-20"><span>inativo</span></span>{% endif %}
      {% if u.trocar_senha %}<span class="br-tag bg-warning ml-1"><span>senha temporária</span></span>{% endif %}</td>
    <td>{% if u.ultimo_acesso %}{{ u.ultimo_acesso[8:10] }}/{{ u.ultimo_acesso[5:7] }}/{{ u.ultimo_acesso[:4] }} {{ u.ultimo_acesso[11:16] }}{% else %}—{% endif %}</td>
    <td class="dsgov-acoes"><a class="br-button circle small" href="{{ url_for('usuarios.editar', id=u.id, retorno=request.full_path) }}" aria-label="Editar {{ u.login }}"><i class="fas fa-edit" aria-hidden="true"></i></a></td></tr>
  {% else %}<tr><td colspan="6">Nenhum usuário encontrado.</td></tr>{% endfor %}</tbody>
</table></div>
{% endblock %}
```

`templates/usuarios/formulario.html`:

```html
{% extends "base.html" %}
{% block titulo %}{{ 'Usuário ' ~ u.login if u else 'Novo usuário' }}{% endblock %}
{% block conteudo %}
<h1>{% if u %}Usuário <code>{{ u.login }}</code>{% else %}Novo usuário{% endif %}</h1>
<p class="text-gray-70">Campos com <span class="text-danger">*</span> são obrigatórios.</p>
{% if senha_temporaria %}
<div class="br-message warning mb-3"><div class="icon"><i class="fas fa-key fa-lg" aria-hidden="true"></i></div>
  <div class="content" role="alert"><span class="message-title">Senha temporária: <code id="senha-temporaria">{{ senha_temporaria }}</code>.</span>
    <span class="message-body"> Anote agora: ela não será mostrada de novo. Ao entrar, o usuário terá de escolher uma nova senha.</span></div></div>
{% endif %}
{% if erro %}<div class="br-message danger mb-3"><div class="icon"><i class="fas fa-times-circle fa-lg" aria-hidden="true"></i></div>
  <div class="content" role="alert"><span class="message-title">Confira os campos.</span><span class="message-body"> {{ erro }} Seus dados foram mantidos.</span></div></div>{% endif %}
<form method="post" action="{{ url_for('usuarios.editar', id=u.id, retorno=retorno) if u else url_for('usuarios.incluir', retorno=retorno) }}" class="col-md-8">
  {% if not u %}
  <div class="br-input mb-3"><label for="login">Login <span class="text-danger" aria-hidden="true">*</span></label>
    <input id="login" name="login" type="text" autocapitalize="none" value="{{ valores.login or '' }}" required/>
    <p class="text-down-01 text-gray-70 mt-1">Minúsculas, números, ponto, hífen ou sublinhado; de 2 a 30 caracteres. Não muda depois.</p></div>
  {% endif %}
  <div class="br-input mb-3"><label for="nome">Nome <span class="text-danger" aria-hidden="true">*</span></label><input id="nome" name="nome" type="text" value="{{ valores.nome or '' }}" required/>
    <p class="text-down-01 text-gray-70 mt-1">Aparece no cabeçalho e nas leituras do inventário.</p></div>
  <div class="mb-3"><div class="text-weight-semi-bold mb-2">Perfil <span class="text-danger" aria-hidden="true">*</span></div>
    {% for p in perfis %}<div class="br-radio mb-1"><input id="perfil-{{ p }}" name="perfil" type="radio" value="{{ p }}"{% if valores.perfil == p %} checked{% endif %} required/><label for="perfil-{{ p }}">{{ rotulos[p] }}</label></div>{% endfor %}</div>
  {% if u %}
  <div class="br-checkbox mb-3"><input id="ativo" name="ativo" type="checkbox" value="1"{% if valores.ativo %} checked{% endif %}/><label for="ativo">Ativo (desmarque para inativar; usuários não são excluídos)</label></div>
  {% else %}
  <div class="br-input mb-3"><label for="senha">Senha inicial <span class="text-danger" aria-hidden="true">*</span></label><input id="senha" name="senha" type="password" autocomplete="new-password" minlength="8" required/></div>
  <div class="br-input mb-3"><label for="confirmacao">Repita a senha <span class="text-danger" aria-hidden="true">*</span></label><input id="confirmacao" name="confirmacao" type="password" autocomplete="new-password" minlength="8" required/></div>
  <div class="br-checkbox mb-3"><input id="trocar_senha" name="trocar_senha" type="checkbox" value="1"{% if valores.trocar_senha %} checked{% endif %}/><label for="trocar_senha">Obrigar troca de senha no primeiro acesso</label></div>
  {% endif %}
  <div class="dsgov-acoes-formulario mt-4">
    <a class="br-button" href="{{ retorno }}">{{ 'Voltar' if senha_temporaria else 'Cancelar' }}</a>
    <button class="br-button primary" type="submit"><i class="fas fa-save mr-1" aria-hidden="true"></i>{{ 'Salvar alterações' if u else 'Salvar' }}</button>
  </div>
</form>
{% if u %}
<form method="post" action="{{ url_for('usuarios.nova_senha', id=u.id, retorno=retorno) }}" class="mt-4 col-md-8">
  <div class="br-card"><div class="card-content d-flex align-items-center">
    <div class="flex-fill"><div class="text-weight-semi-bold">Nova senha temporária</div><div class="text-down-01 text-gray-70">Gera uma senha aleatória, mostrada uma vez; o usuário troca no primeiro acesso. Também desbloqueia a conta.</div></div>
    <button class="br-button secondary" type="submit"><i class="fas fa-key mr-1" aria-hidden="true"></i>Nova senha</button>
  </div></div>
</form>
{% endif %}
{% endblock %}
```

Em `app.py`, descomentar a linha do item "Usuários" no menu (Task 5, Step 4).

- [ ] **Step 5: Rodar os testes**

Run: `.venv/bin/pytest tests/test_usuarios_telas.py tests/test_permissoes.py -q`
Expected: verde.

- [ ] **Step 6: Suíte e commit**

Run: `.venv/bin/pytest -q`

```bash
git add app_usuarios.py app.py templates/usuarios tests/test_usuarios_telas.py tests/test_permissoes.py
git commit -m "Usuários: lista com busca e filtros, novo, editar (perfil, ativo), nova senha temporária mostrada uma vez"
```

---

### Task 7: Inventário: comissão por usuários e integrante = logado

**Files:**
- Modify: `inventario.py` (`abrir_evento`, `ler`, `registrar_sobra`, novo `editar_comissao`)
- Modify: `app_inventario.py` (`abrir`, `evento_tela`, remover `integrante`, `sala_tela`, `ler`, `lote`, `sobra`, novo `comissao`)
- Modify: `templates/inventario_eventos.html`, `templates/inventario_evento.html`, `templates/inventario_sala.html`
- Create: `templates/inventario_comissao.html`
- Modify: `usuarios.py` (remover `"inventario.integrante"` temporário, se foi acrescentado na Task 3)
- Modify: `tests/test_inventario.py`, `tests/test_app.py`

**Interfaces:**
- Consumes: `usuarios.elegiveis_comissao(conn)`, `g.usuario["nome"]`.
- Produces: `inventario.abrir_evento(conn, nome, descricao, integrantes, salas=None, elegiveis=None) -> int` (com `elegiveis`, cada integrante tem que estar na lista, senão `ErroDeNegocio`); `inventario.editar_comissao(conn, evento_id, nomes, elegiveis=None) -> None`; endpoint `inventario.comissao` (GET/POST `/inventario/<id>/comissao`); mensagem `inventario.FORA_DA_COMISSAO = "Você não faz parte da comissão deste evento."`; contexto de sala `na_comissao: bool`.

- [ ] **Step 1: Testes de dados que falham (`tests/test_inventario.py`)**

Acrescentar:

```python
def test_abrir_evento_com_elegiveis_e_editar_comissao(dados):
    semear(dados)
    with pytest.raises(db.ErroDeNegocio, match="não pode compor"):
        inventario.abrir_evento(dados, "X", "", ["Fulano", "Zé"], elegiveis=["Fulano", "Beltrana"])
    eid = inventario.abrir_evento(dados, "X", "", ["Fulano"], elegiveis=["Fulano", "Beltrana"])
    assert inventario.evento(dados, eid)["integrantes"] == ["Fulano"]
    with pytest.raises(db.ErroDeNegocio, match="não pode compor"):
        inventario.editar_comissao(dados, eid, ["Zé"], elegiveis=["Fulano", "Beltrana"])
    with pytest.raises(db.ErroDeNegocio, match="ao menos um"):
        inventario.editar_comissao(dados, eid, [])
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.editar_comissao(dados, eid, ["Beltrana", " Fulano "], elegiveis=["Fulano", "Beltrana"])
    assert inventario.evento(dados, eid)["integrantes"] == ["Beltrana", "Fulano"]
    inventario.editar_comissao(dados, eid, ["Beltrana"])
    assert inventario.evento(dados, eid)["integrantes"] == ["Beltrana"]
    assert dados.execute("SELECT integrante FROM inventario_leituras WHERE evento_id = ?", (eid,)).fetchone()[0] == "Fulano"  # leitura fica
    with pytest.raises(db.ErroDeNegocio, match="não faz parte da comissão"):
        inventario.ler(dados, eid, "01 - SALA CCI", 1002, "Fulano")
    inventario.encerrar_evento(dados, eid)
    with pytest.raises(db.ErroDeNegocio, match="encerrado"):
        inventario.editar_comissao(dados, eid, ["Fulano"])
```

Ajustar em `test_inventario.py` as asserções que esperam a mensagem antiga: onde houver `match="integrante"` ou texto "Escolha o integrante" para leitura/sobra de quem não está na comissão, trocar por `match="não faz parte da comissão"`. (Procure com `grep -n "integrante" tests/test_inventario.py`.)

- [ ] **Step 2: Rodar para ver falhar**

Run: `.venv/bin/pytest tests/test_inventario.py -q`

- [ ] **Step 3: `inventario.py`**

1. Constante, depois de `ROTULO_FOTO`: `FORA_DA_COMISSAO = "Você não faz parte da comissão deste evento."`
2. Nova função auxiliar e `editar_comissao`, logo após `abrir_evento`:

```python
def _nomes_da_comissao(integrantes, elegiveis) -> list[str]:
    nomes = sorted({" ".join(_texto(n).split()) for n in integrantes if _texto(n).strip()})
    if not nomes:
        raise ErroDeNegocio("Informe ao menos um integrante da comissão.")
    if elegiveis is not None:
        fora = [n for n in nomes if n not in set(elegiveis)]
        if fora:
            raise ErroDeNegocio(f"{', '.join(fora)}: não pode compor a comissão (usuário inexistente, inativo ou de consulta).")
    return nomes


def editar_comissao(conn, evento_id: int, integrantes: list, elegiveis: list | None = None) -> None:
    """Substitui a comissão do evento aberto. Leituras já feitas não mudam: quem sai só deixa de poder ler."""
    _evento_aberto_ou_erro(conn, evento_id)
    nomes = _nomes_da_comissao(integrantes, elegiveis)
    conn.execute("DELETE FROM inventario_integrantes WHERE evento_id = ?", (evento_id,))
    conn.executemany("INSERT INTO inventario_integrantes VALUES (?,?)", [(evento_id, n) for n in nomes])
    conn.commit()
```

3. `abrir_evento`: assinatura `def abrir_evento(conn, nome: str, descricao, integrantes: list, salas: list | None = None, elegiveis: list | None = None) -> int`; substituir as três linhas `nomes = sorted(...)` / `if not nomes:` / `raise ...` por `nomes = _nomes_da_comissao(integrantes, elegiveis)` (mantendo-a depois da checagem de evento aberto, como hoje).
4. Em `ler` e `registrar_sobra`, trocar `raise ErroDeNegocio("Escolha o integrante da comissão antes de ler.")` por `raise ErroDeNegocio(FORA_DA_COMISSAO)`.

- [ ] **Step 4: Rodar testes de dados**

Run: `.venv/bin/pytest tests/test_inventario.py -q`
Expected: verde.

- [ ] **Step 5: Testes de rota (`tests/test_app.py`)**

Trocar o helper `_abrir` (linhas 431–437) por:

```python
def _abrir(cliente, comissao=("Fulano", "Beltrana")):
    """Abre evento com a comissão dada (nomes de usuários existentes). O admin logado é 'Fulano'."""
    cliente.post("/inventario/abrir", data={"nome": "Inv", "integrantes": list(comissao), "escopo": "todas"})
    import db, inventario
    return inventario.evento_aberto(db.conectar())["id"]
```

Ajustes pontuais (linhas do arquivo antes desta tarefa):

- 39: `data={"nome": "Inv 2", "integrantes": ["Fulano"], "escopo": "todas"}`; apagar a linha 41 (POST `/integrante`).
- 443: `"integrantes": ["Fulano", "Beltrana"]`; 446: `"integrantes": ["Fulano"]` (a mensagem "já existe" continua).
- 450–454: apagar os quatro POSTs em `/integrante` e a asserção `b"Fulano" in r.data`; no lugar, `assert b"Comiss" in cliente.get(f"/inventario/{eid}").data`.
- 464: `"integrantes": ["Beltrana"]`.
- 469–474: `eid = _abrir(cliente, comissao=["Beltrana"])`; a asserção da tela passa a ser `"não faz parte da comissão".encode() in r.data`; o 409 passa a checar `"comissão" in r.get_json()["erro"].lower()`; a linha 474 vira `cliente.post(f"/inventario/{eid}/comissao", data={"integrantes": ["Fulano", "Beltrana"]})`.
- 487: continua `== "Fulano"` (o admin logado chama-se Fulano).
- 671–674: substituir `with cliente.session_transaction()... sess.pop("integrante")` por `cliente.post(f"/inventario/{eid}/comissao", data={"integrantes": ["Beltrana"]})`; a asserção `b"integrante" in r.data.lower()` vira `"comissão".encode() in r.data.lower()`; a linha 674 vira `cliente.post(f"/inventario/{eid}/comissao", data={"integrantes": ["Fulano", "Beltrana"]})`.
- 756, 771: `_abrir(cliente)`; 783: `_abrir(cliente)`.

Acrescentar:

```python
def test_inventario_comissao_por_usuarios(cliente, dados, usuarios_exemplo):
    r = cliente.get("/inventario")
    assert b'name="integrantes"' in r.data and b'value="Fulano"' in r.data and b'value="Beltrana"' in r.data
    assert b'value="Operador Teste"' in r.data and b'value="Consulta Teste"' not in r.data          # consulta não compõe
    r = cliente.post("/inventario/abrir", data={"nome": "Inv", "integrantes": ["Consulta Teste"], "escopo": "todas"}, follow_redirects=True)
    assert "não pode compor".encode() in r.data
    eid = _abrir(cliente, comissao=["Beltrana"])
    r = cliente.get(f"/inventario/{eid}")
    assert b"Comiss" in r.data and b'href="/inventario/%d/comissao"' % eid in r.data and b"Quem est" not in r.data
    r = cliente.get(f"/inventario/{eid}/sala/01 - SALA CCI")
    assert "não faz parte da comissão".encode() in r.data and b'id="leitura" type="text" inputmode="none" autocomplete="off" enterkeyhint="done" placeholder="Aproxime o leitor\xe2\x80\xa6" disabled' in r.data
    r = cliente.get(f"/inventario/{eid}/comissao")
    assert r.status_code == 200 and b'value="Beltrana" checked' in r.data and b'value="Fulano"' in r.data
    r = cliente.post(f"/inventario/{eid}/comissao", data={"integrantes": ["Fulano", "Beltrana"]}, follow_redirects=True)
    assert "Comissão atualizada".encode() in r.data
    j = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"}).get_json()
    assert j["situacao"] == "localizado" and j["integrante"] == "Fulano"
    assert b"lendo como <strong>Fulano</strong>" in cliente.get(f"/inventario/{eid}/sala/01 - SALA CCI").data
    cliente.post("/sair"); logar(cliente, *usuarios_exemplo["operador"])
    assert cliente.get(f"/inventario/{eid}/comissao").status_code == 403                         # só admin
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1002"})
    assert r.status_code == 409 and "comissão" in r.get_json()["erro"].lower()                   # operador fora da comissão
    cliente.post("/sair"); logar(cliente, *usuarios_exemplo["inventariante"])
    j = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1002"}).get_json()
    assert j["integrante"] == "Beltrana"
    assert cliente.get("/inventario/%d/integrante" % eid).status_code in (404, 405)


def test_inventario_desktop_admin_local_entra_na_comissao(cliente_local):
    import db, inventario, usuarios
    conn = db.conectar()
    usuarios.criar(conn, "x", "Xis", "Senha!234", "inventariante")
    r = cliente_local.post("/inventario/abrir", data={"nome": "Inv", "integrantes": ["Xis"], "escopo": "todas"}, follow_redirects=True)
    assert r.status_code == 200
    eid = inventario.evento_aberto(conn)["id"]
    assert inventario.evento(conn, eid)["integrantes"] == ["Administrador local", "Xis"]
    j = cliente_local.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"}).get_json()
    assert j["integrante"] == "Administrador local"
```

- [ ] **Step 6: `app_inventario.py`**

1. Import: `import usuarios`.
2. Helper, depois de `_trilha`:

```python
def _elegiveis(conn) -> list[dict]:
    """Usuários que podem compor a comissão. No desktop o administrador local entra sempre."""
    lista = usuarios.elegiveis_comissao(conn)
    if g.usuario["id"] is None:
        lista = [dict(g.usuario)] + lista
    return lista


def _nomes_para_comissao(conn, marcados: list) -> list[str]:
    nomes = list(marcados)
    if g.usuario["id"] is None and g.usuario["nome"] not in nomes:
        nomes.append(g.usuario["nome"])
    return nomes


def _na_comissao(e) -> bool:
    return g.usuario["nome"] in e["integrantes"]
```

3. `eventos_tela`: passar `elegiveis=_elegiveis(conn)` ao template.
4. `abrir`:

```python
@inventario_bp.route("/abrir", methods=["POST"])
def abrir():
    conn = _conn()
    f = request.form
    salas = None if f.get("escopo", "todas") == "todas" else f.getlist("salas")
    eid = inventario.abrir_evento(conn, f.get("nome", ""), f.get("descricao", ""), _nomes_para_comissao(conn, f.getlist("integrantes")),
                                  salas, elegiveis=[u["nome"] for u in _elegiveis(conn)])
    flash("Evento aberto. Comece pelas salas.", "success")
    return redirect(url_for("inventario.evento_tela", id=eid))
```

5. `evento_tela`: trocar `integrante=session.get("integrante")` por `na_comissao=_na_comissao(e)`.
6. Apagar a rota `integrante` inteira (linhas 71–80).
7. Nova rota, depois de `encerrar`:

```python
@inventario_bp.route("/<int:id>/comissao", methods=["GET", "POST"])
def comissao(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    if request.method == "POST":
        # Nomes já na comissão continuam aceitos (integrantes migrados sem usuário); nome novo só se for usuário elegível.
        inventario.editar_comissao(conn, id, _nomes_para_comissao(conn, request.form.getlist("integrantes")),
                                   elegiveis=[u["nome"] for u in _elegiveis(conn)] + e["integrantes"])
        flash("Comissão atualizada.", "success")
        return redirect(url_for("inventario.evento_tela", id=id))
    com_leituras = {r[0] for r in conn.execute("SELECT DISTINCT integrante FROM inventario_leituras WHERE evento_id = ?", (id,))}
    return render_template("inventario_comissao.html", e=e, elegiveis=_elegiveis(conn), com_leituras=com_leituras,
                           trilha=_trilha(e, ("Comissão", None)))
```

8. `sala_tela`: `integrante=session.get("integrante")` → `na_comissao=_na_comissao(e), integrante=g.usuario["nome"]`.
9. `ler`: `session.get("integrante") or ""` → `g.usuario["nome"]`. Idem em `lote` (`ler_lote`) e em `sobra` (`integrante = g.usuario["nome"]`).
10. Remover `session` do import do Flask se ficar sem uso.
11. Em `usuarios.py`, remover a chave temporária `"inventario.integrante"` se existir.

- [ ] **Step 7: Templates**

`templates/inventario_eventos.html`: substituir o `<div class="col-md-6 mb-3"><div class="br-textarea">…integrantes…</div></div>` por:

```html
    <div class="col-md-6 mb-3">
      <div class="text-weight-semi-bold mb-2">Comissão (quem poderá ler)</div>
      {% for u in elegiveis %}<div class="br-checkbox mb-1"><input id="int-{{ loop.index }}" name="integrantes" type="checkbox" value="{{ u.nome }}"{% if u.id is none %} checked disabled{% endif %}/><label for="int-{{ loop.index }}">{{ u.nome }} <span class="text-gray-70 text-down-01">· {{ ROTULO_PERFIL[u.perfil] }}</span></label></div>
      {% else %}<p class="text-gray-70">Nenhum usuário elegível. Cadastre usuários com perfil inventariante, operador ou administrador.</p>{% endfor %}
    </div>
```

`templates/inventario_evento.html`: apagar o `<form method="post" action="{{ url_for('inventario.integrante', …) }}">…</form>` (bloco `{% else %}` do encerrado) e, no lugar, dentro do mesmo `{% else %}`:

```html
{% if not na_comissao %}
<div class="br-message info mb-3"><div class="icon"><i class="fas fa-user-slash fa-lg" aria-hidden="true"></i></div>
  <div class="content"><span class="message-body">Você não faz parte da comissão deste evento: pode consultar, mas não ler bens.</span></div></div>
{% endif %}
```

E no `<div class="ml-auto">`, antes do formulário de encerrar, dentro de `{% if not e.encerrado_em %}`:

```html
    {% if pode('inventario.comissao') %}<a class="br-button secondary ml-2" href="{{ url_for('inventario.comissao', id=e.id) }}"><i class="fas fa-users mr-1" aria-hidden="true"></i>Comissão</a>{% endif %}
```

Envolver o formulário de encerrar com `{% if pode('inventario.encerrar', 'POST') %}…{% endif %}`. Acrescentar ao parágrafo do resumo: ` · Comissão: {{ e.integrantes|join(', ') }}`.

`templates/inventario_comissao.html` (novo):

```html
{% extends "base.html" %}
{% block titulo %}Comissão · {{ e.nome }}{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Comissão de {{ e.nome }}</h1></div>
<p class="text-gray-70">Marque quem pode ler bens neste evento. Quem sair da comissão deixa de ler; as leituras já feitas ficam.</p>
<form method="post" action="{{ url_for('inventario.comissao', id=e.id) }}" class="col-md-6">
  {% for u in elegiveis %}
  <div class="br-checkbox mb-2"><input id="int-{{ loop.index }}" name="integrantes" type="checkbox" value="{{ u.nome }}"{% if u.nome in e.integrantes or u.id is none %} checked{% endif %}{% if u.id is none %} disabled{% endif %}/>
    <label for="int-{{ loop.index }}">{{ u.nome }} <span class="text-gray-70 text-down-01">· {{ ROTULO_PERFIL[u.perfil] }}{% if u.nome in com_leituras %} · já tem leituras neste evento{% endif %}</span></label></div>
  {% endfor %}
  {% for nome in e.integrantes if nome not in elegiveis|map(attribute='nome')|list %}
  <div class="br-checkbox mb-2"><input id="ant-{{ loop.index }}" name="integrantes" type="checkbox" value="{{ nome }}" checked/>
    <label for="ant-{{ loop.index }}">{{ nome }} <span class="text-gray-70 text-down-01">· sem usuário ativo com este nome{% if nome in com_leituras %} · já tem leituras{% endif %}</span></label></div>
  {% endfor %}
  <div class="dsgov-acoes-formulario mt-4">
    <a class="br-button" href="{{ url_for('inventario.evento_tela', id=e.id) }}">Cancelar</a>
    <button class="br-button primary" type="submit"><i class="fas fa-save mr-1" aria-hidden="true"></i>Salvar comissão</button>
  </div>
</form>
{% endblock %}
```

Observação: a segunda lista mantém marcados os nomes antigos (ex.: integrantes migrados sem usuário). A rota `comissao`
(Step 6) já aceita esses nomes porque passa `elegiveis = elegíveis + e["integrantes"]`: nomes antigos podem ficar ou sair,
mas nenhum nome novo fora dos usuários entra.

`templates/inventario_sala.html`:

- Linha 8: `{% if integrante %}lendo como <strong>{{ integrante }}</strong>{% else %}<strong>Escolha o integrante</strong>{% endif %}` → `{% if na_comissao %}lendo como <strong>{{ integrante }}</strong>{% else %}<strong>Você não faz parte da comissão deste evento</strong>{% endif %}`; e o `(<a …>trocar</a>)` vira `(<a href="{{ url_for('inventario.evento_tela', id=e.id) }}">evento</a>)`.
- Linhas 28, 30, 31, 40, 41: `not integrante` → `not na_comissao`.

- [ ] **Step 8: Rodar os testes**

Run: `.venv/bin/pytest tests/test_app.py tests/test_inventario.py tests/test_permissoes.py -q`
Expected: verde. Se `test_toda_rota_do_app_esta_na_matriz` acusar `inventario.comissao`, ela já está na matriz (Task 3); se acusar `inventario.integrante`, a rota não foi removida.

- [ ] **Step 9: Suíte e commit**

Run: `.venv/bin/pytest -q`

```bash
git add inventario.py app_inventario.py usuarios.py templates/inventario_eventos.html templates/inventario_evento.html templates/inventario_sala.html templates/inventario_comissao.html tests/test_inventario.py tests/test_app.py
git commit -m "Inventário: comissão escolhida entre usuários, integrante da leitura é o usuário logado, botão Comissão (só admin); sai o passo de escolher integrante"
```

---

### Task 8: Excluir evento de inventário (só admin, com nome digitado)

**Files:**
- Modify: `inventario.py`
- Modify: `app_inventario.py`
- Create: `templates/inventario_excluir.html`
- Modify: `templates/inventario_evento.html`
- Modify: `tests/test_inventario.py`, `tests/test_app.py`

**Interfaces:**
- Produces: `inventario.contagem_para_exclusao(conn, evento_id) -> dict` (chaves `leituras, fotos, sobras, sobras_com_foto, integrantes, bens_encerrados, salas`); `inventario.urls_das_fotos(conn, evento_id) -> list[str]`; `inventario.excluir_evento(conn, evento_id, nome_confirmacao, apagar=None) -> dict` (devolve o evento apagado); endpoint `inventario.excluir` (GET/POST `/inventario/<id>/excluir`).

- [ ] **Step 1: Testes de dados**

Acrescentar em `tests/test_inventario.py`:

```python
def test_excluir_evento_apaga_tudo_com_fotos_primeiro(dados):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    foto_falsa(dados, eid, 1001, "https://x/1.webp")
    foto_falsa(dados, eid, 1001, "https://x/2.webp")
    sid = inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "VENT", "", "obs", "https://x/s.webp", "Fulano")
    inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "SEM FOTO", "", "obs", "", "Fulano", exigir_foto=False)
    c = inventario.contagem_para_exclusao(dados, eid)
    assert c == {"leituras": 1, "fotos": 2, "sobras": 2, "sobras_com_foto": 1, "integrantes": 2, "bens_encerrados": 0, "salas": 3}
    assert inventario.urls_das_fotos(dados, eid) == ["https://x/1.webp", "https://x/2.webp", "https://x/s.webp"]
    with pytest.raises(db.ErroDeNegocio, match="não confere"):
        inventario.excluir_evento(dados, eid, "Inventario 2026")
    apagadas = []
    def apagar_falha(url):
        apagadas.append(url)
        if url.endswith("2.webp"):
            raise db.ErroDeNegocio("bucket fora")
    with pytest.raises(db.ErroDeNegocio, match="bucket fora"):
        inventario.excluir_evento(dados, eid, "Inventário 2026", apagar=apagar_falha)
    assert inventario.evento(dados, eid) and inventario.contagem_para_exclusao(dados, eid)["fotos"] == 2   # nada mudou
    apagadas.clear()
    inventario.encerrar_evento(dados, eid)                                       # encerrado também pode ser excluído
    assert inventario.contagem_para_exclusao(dados, eid)["bens_encerrados"] == 5
    e = inventario.excluir_evento(dados, eid, "  Inventário  2026 ", apagar=apagadas.append)
    assert e["id"] == eid and apagadas == ["https://x/1.webp", "https://x/2.webp", "https://x/s.webp"]
    assert inventario.evento(dados, eid) is None and inventario.eventos(dados) == []
    for t in ("inventario_leituras", "inventario_fotos", "inventario_sobras", "inventario_integrantes", "inventario_bens_encerrados", "inventario_salas"):
        assert dados.execute(f"SELECT count(*) FROM {t} WHERE evento_id = ?", (eid,)).fetchone()[0] == 0, t
    with pytest.raises(db.ErroDeNegocio, match="não encontrado"):
        inventario.excluir_evento(dados, eid, "x")
```

- [ ] **Step 2: Rodar para ver falhar**

Run: `.venv/bin/pytest tests/test_inventario.py -q -k excluir_evento`

- [ ] **Step 3: `inventario.py`**

Depois de `encerrar_evento`:

```python
_TABELAS_DO_EVENTO = ("inventario_fotos", "inventario_leituras", "inventario_sobras", "inventario_integrantes",
                      "inventario_bens_encerrados", "inventario_salas")


def contagem_para_exclusao(conn, evento_id: int) -> dict:
    """O que a exclusão do evento apaga, para a tela de confirmação."""
    def n(sql):
        return conn.execute(sql, (evento_id,)).fetchone()[0]
    return {"leituras": n("SELECT count(*) FROM inventario_leituras WHERE evento_id = ?"),
            "fotos": n("SELECT count(*) FROM inventario_fotos WHERE evento_id = ?"),
            "sobras": n("SELECT count(*) FROM inventario_sobras WHERE evento_id = ?"),
            "sobras_com_foto": n("SELECT count(*) FROM inventario_sobras WHERE evento_id = ? AND foto_url IS NOT NULL AND foto_url <> ''"),
            "integrantes": n("SELECT count(*) FROM inventario_integrantes WHERE evento_id = ?"),
            "bens_encerrados": n("SELECT count(*) FROM inventario_bens_encerrados WHERE evento_id = ?"),
            "salas": n("SELECT count(*) FROM inventario_salas WHERE evento_id = ?")}


def urls_das_fotos(conn, evento_id: int) -> list[str]:
    """Fotos dos bens (por número e nfoto) e depois das sobras (por id)."""
    bens = [r[0] for r in conn.execute("SELECT url FROM inventario_fotos WHERE evento_id = ? ORDER BY numero, nfoto", (evento_id,))]
    sobras = [r[0] for r in conn.execute("SELECT foto_url FROM inventario_sobras WHERE evento_id = ? AND foto_url IS NOT NULL AND foto_url <> '' ORDER BY id", (evento_id,))]
    return bens + sobras


def excluir_evento(conn, evento_id: int, nome_confirmacao, apagar=None) -> dict:
    """Apaga o evento inteiro (aberto ou encerrado). Exige o nome digitado igual ao do evento. `apagar(url)` roda
    para cada foto ANTES de tocar no banco: se falhar, nada é apagado. Não há desfazer."""
    e = _um(conn, "SELECT * FROM inventario_eventos WHERE id = ?", evento_id)
    if not e:
        raise ErroDeNegocio("Evento de inventário não encontrado.")
    if " ".join(_texto(nome_confirmacao).split()) != " ".join(e["nome"].split()):
        raise ErroDeNegocio("O nome digitado não confere com o nome do evento; nada foi excluído.")
    if apagar:
        for url in urls_das_fotos(conn, evento_id):
            apagar(url)
    for tabela in _TABELAS_DO_EVENTO:
        conn.execute(f"DELETE FROM {tabela} WHERE evento_id = ?", (evento_id,))
    conn.execute("DELETE FROM inventario_eventos WHERE id = ?", (evento_id,))
    conn.commit()
    return e
```

- [ ] **Step 4: Rodar testes de dados**

Run: `.venv/bin/pytest tests/test_inventario.py -q`

- [ ] **Step 5: Teste de rota (`tests/test_app.py`)**

```python
def test_inventario_excluir_evento(cliente, dados, monkeypatch, usuarios_exemplo):
    import fotos, inventario
    eid = _abrir(cliente)
    cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    inventario.adicionar_foto(dados, eid, 1001, lambda c: "https://x/a.webp")
    r = cliente.get(f"/inventario/{eid}")
    assert b"Excluir evento" in r.data
    r = cliente.get(f"/inventario/{eid}/excluir")
    assert r.status_code == 200 and b"1 leitura" in r.data and b"1 foto" in r.data and b'name="nome"' in r.data
    monkeypatch.setattr(fotos, "apagar", lambda url: (_ for _ in ()).throw(RuntimeError("bucket fora")))
    r = cliente.post(f"/inventario/{eid}/excluir", data={"nome": "Inv"}, follow_redirects=True)
    assert "Não foi possível apagar as fotos no bucket; o evento foi mantido".encode() in r.data
    assert inventario.evento(dados, eid) is not None
    apagadas = []
    monkeypatch.setattr(fotos, "apagar", apagadas.append)
    r = cliente.post(f"/inventario/{eid}/excluir", data={"nome": "Errado"}, follow_redirects=True)
    assert "não confere".encode() in r.data and inventario.evento(dados, eid) is not None
    r = cliente.post(f"/inventario/{eid}/excluir", data={"nome": "Inv"}, follow_redirects=True)
    assert "Evento Inv excluído".encode() in r.data and apagadas == ["https://x/a.webp"]
    assert inventario.evento(dados, eid) is None and cliente.get(f"/inventario/{eid}").status_code == 404
    eid = _abrir(cliente)
    cliente.post("/sair"); logar(cliente, *usuarios_exemplo["operador"])
    assert cliente.get(f"/inventario/{eid}/excluir").status_code == 403
    assert cliente.post(f"/inventario/{eid}/excluir", data={"nome": "Inv"}).status_code == 403
    assert b"Excluir evento" not in cliente.get(f"/inventario/{eid}").data
```

- [ ] **Step 6: Rota e templates**

Em `app_inventario.py`, depois de `comissao`:

```python
def _apagar_fotos_do_evento(url):
    try:
        fotos.apagar(url)
    except Exception:
        raise db.ErroDeNegocio("Não foi possível apagar as fotos no bucket; o evento foi mantido. Tente de novo.")


@inventario_bp.route("/<int:id>/excluir", methods=["GET", "POST"])
def excluir(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    if request.method == "POST":
        try:
            inventario.excluir_evento(conn, id, request.form.get("nome", ""), apagar=_apagar_fotos_do_evento)
        except db.ErroDeNegocio as erro:
            flash(str(erro), "error")
            return redirect(url_for("inventario.excluir", id=id))
        flash(f"Evento {e['nome']} excluído.", "success")
        return redirect(url_for("inventario.eventos_tela"))
    return render_template("inventario_excluir.html", e=e, c=inventario.contagem_para_exclusao(conn, id), trilha=_trilha(e, ("Excluir", None)))
```

`templates/inventario_excluir.html`:

```html
{% extends "base.html" %}
{% block titulo %}Excluir · {{ e.nome }}{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Excluir o evento {{ e.nome }}</h1></div>
<div class="br-message danger mb-3"><div class="icon"><i class="fas fa-exclamation-triangle fa-lg" aria-hidden="true"></i></div>
  <div class="content" role="alert"><span class="message-title">Esta ação não tem desfazer.</span><span class="message-body"> O evento {{ 'está aberto' if not e.encerrado_em else 'está encerrado' }}. Serão apagados:
    {{ c.leituras }} leitura(s), {{ c.fotos }} foto(s) de bens{% if c.sobras_com_foto %} e {{ c.sobras_com_foto }} de sobras{% endif %} (também no bucket), {{ c.sobras }} sobra(s), {{ c.integrantes }} integrante(s) da comissão, {{ c.salas }} sala(s) do escopo{% if c.bens_encerrados %} e o retrato de {{ c.bens_encerrados }} bem(ns) congelado(s) no encerramento{% endif %}. O caminho de volta é o backup diário.</span></div></div>
<form method="post" action="{{ url_for('inventario.excluir', id=e.id) }}" class="col-md-6">
  <div class="br-input mb-3"><label for="nome">Digite o nome do evento para confirmar: <strong>{{ e.nome }}</strong></label><input id="nome" name="nome" type="text" autocomplete="off" required/></div>
  <a class="br-button" href="{{ url_for('inventario.evento_tela', id=e.id) }}">Cancelar</a>
  <button class="br-button danger ml-2" type="submit"><i class="fas fa-trash mr-1" aria-hidden="true"></i>Excluir definitivamente</button>
</form>
{% endblock %}
```

`templates/inventario_evento.html`, no `<div class="ml-auto">`, ao final (fora do `{% if not e.encerrado_em %}`):

```html
    {% if pode('inventario.excluir', 'POST') %}<a class="br-button danger ml-2" href="{{ url_for('inventario.excluir', id=e.id) }}"><i class="fas fa-trash mr-1" aria-hidden="true"></i>Excluir evento</a>{% endif %}
```

- [ ] **Step 7: Rodar e commitar**

Run: `.venv/bin/pytest -q`

```bash
git add inventario.py app_inventario.py templates/inventario_excluir.html templates/inventario_evento.html tests/test_inventario.py tests/test_app.py
git commit -m "Inventário: excluir evento (só admin) com nome digitado; fotos apagadas no bucket antes do banco, falha mantém tudo"
```

---

### Task 9: CSRF

**Files:**
- Modify: `app.py` (`resolver_usuario`, global Jinja `csrf_campo`)
- Modify: `templates/base.html` (`<meta name="csrf">` e o form de Sair), `templates/inventario_sala.html` (3 `fetch`), `templates/termo.html` (1 `fetch`), e todo template com `<form method="post">`: `cadastros.html`, `inventario_evento.html`, `inventario_sala.html`, `termo_emitido.html`, `termos_individuais.html`, `upload.html`, `cadastros/_pessoa.html`, `cadastros/confirmar.html`, `cadastros/mover.html`, `centro_custos.html`, `inventario_eventos.html`, `termo_devolucao.html`, `textos.html`, `cadastros/_campos.html`, `cadastros/atribuir.html`, `cadastros/formulario.html`, `login.html`, `senha.html`, `usuarios/formulario.html`, `inventario_comissao.html`, `inventario_excluir.html`
- Modify: `tests/conftest.py` (cliente que injeta o token)
- Create: `tests/test_csrf.py`

**Interfaces:**
- Produces: `session["csrf"]`; função Jinja `csrf_campo()`; cabeçalho `X-CSRF`; 400 com "Sessão expirada ou formulário inválido. Recarregue a página e tente de novo."; `tests.conftest.ClienteComCSRF`.

- [ ] **Step 1: Cliente de teste que injeta o token**

Em `tests/conftest.py`:

```python
from flask.testing import FlaskClient


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
```

Em `_app_de_teste()`, acrescentar `app.test_client_class = ClienteComCSRF`.

Nota: `FlaskClient.post(path, **kw)` chama `open(path, method="POST", **kw)`, então `kwargs["method"]` é o caminho normal.

- [ ] **Step 2: Testes que falham**

Criar `tests/test_csrf.py`:

```python
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
    assert r.status_code == 400 and "Sessão expirada".encode() in r.data and r.is_json
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
    sala = (TEMPLATES / "inventario_sala.html").read_text(encoding="utf-8")
    assert sala.count("fetch(") == sala.count('"X-CSRF"'), "todo fetch da sala manda X-CSRF"
    termo = (TEMPLATES / "termo.html").read_text(encoding="utf-8")
    assert '"X-CSRF"' in termo


def test_cliente_de_teste_injeta_token(cliente):
    assert cliente.post("/textos", data={"restaurar": "orgao_nome"}).status_code == 302
```

- [ ] **Step 3: Rodar para ver falhar**

Run: `.venv/bin/pytest tests/test_csrf.py -q`

- [ ] **Step 4: `app.py`**

1. Imports: `import hmac`, `import secrets`, `from markupsafe import Markup`.
2. Depois de `NEGADO = ...`:

```python
CSRF_INVALIDO = "Sessão expirada ou formulário inválido. Recarregue a página e tente de novo."


def _csrf_token() -> str:
    if "csrf" not in session:
        session["csrf"] = secrets.token_urlsafe(32)
    return session["csrf"]


@app.template_global("csrf_campo")
def csrf_campo():
    return Markup(f'<input type="hidden" name="csrf" value="{_csrf_token()}"/>')


def _csrf_invalido():
    if request.endpoint in ROTAS_JSON or request.is_json:
        return {"erro": CSRF_INVALIDO}, 400
    return render_template("403.html", trilha=[("Sessão expirada", None)], csrf=True), 400
```

3. Em `resolver_usuario`, logo depois de `if ep is None or ep == "static": return None` (antes de resolver o usuário, para valer também no `/login`):

```python
    if request.method == "POST":
        enviado = request.form.get("csrf") or request.headers.get("X-CSRF") or ""
        if not enviado or not hmac.compare_digest(enviado, session.get("csrf", "")):
            return _csrf_invalido()
```

4. No `login()` de `app_usuarios.py`, depois de `session.clear()` e `session["usuario_id"] = u["id"]`, acrescentar `session["csrf"] = secrets.token_urlsafe(32)` (import `secrets`). Em `sair()`, o `session.clear()` basta: o próximo GET gera outro.
5. Em `contexto_dsgov`, incluir `"CSRF": _csrf_token()` no contexto (também quando não há usuário, para a página de login).
6. `templates/403.html`: envolver o `br-message` atual em `{% if not csrf %}…{% else %}` com a mensagem `<span class="message-title">Sessão expirada ou formulário inválido.</span><span class="message-body"> Recarregue a página e tente de novo.</span>` e título `Sessão expirada`.

- [ ] **Step 5: Templates**

- `base.html`: no `<head>`, `<meta name="csrf" content="{{ CSRF }}"/>`; dentro do form de Sair, `{{ csrf_campo() }}`.
- Em cada `<form method="post" …>` dos templates listados em **Files**, inserir `{{ csrf_campo() }}` na linha seguinte à abertura. Em `login.html` (que não estende `base.html`) também: `{{ csrf_campo() }}` logo após o `<form>`.
- `inventario_sala.html`: no topo do script, `var CSRF = document.querySelector('meta[name="csrf"]').content;` e nos três `fetch`:
  - `headers: {"Content-Type": "application/json"}` → `headers: {"Content-Type": "application/json", "X-CSRF": CSRF}` (dois lugares);
  - `fetch(URL_LEITURA... + "/foto", {method: "POST", body: fd})` → `fetch(..., {method: "POST", headers: {"X-CSRF": CSRF}, body: fd})`.
- `termo.html`: `fetch(this.dataset.registrar, {method: "POST"})` → `fetch(this.dataset.registrar, {method: "POST", headers: {"X-CSRF": document.querySelector('meta[name="csrf"]').content}})`.

- [ ] **Step 6: Rodar**

Run: `.venv/bin/pytest tests/test_csrf.py -q` e depois `.venv/bin/pytest -q`. O teste de varredura lista qualquer formulário esquecido; corrija até passar.

- [ ] **Step 7: Commit**

```bash
git add app.py app_usuarios.py templates tests/conftest.py tests/test_csrf.py
git commit -m "CSRF: token na sessão, csrf_campo() em todo formulário POST, X-CSRF nos fetch; 400 sem token; cliente de teste injeta o token"
```

---

### Task 10: Publicação: `compose.yml`, README, spec e revisão final

**Files:**
- Modify: `compose.yml`
- Modify: `README.md`
- Modify: `docs/superpowers/specs/2026-09-17-fase4-usuarios-design.md` (estado)

- [ ] **Step 1: `compose.yml`**

Dentro de `services.web`, depois de `env_file`, acrescentar:

```yaml
    environment:
      TERMOS_LOGIN: "1"     # site pede usuário e senha (o programa Windows não define e entra direto)
```

- [ ] **Step 2: README**

1. Na seção **Uso**, item 1, depois do parágrafo do Inventário, acrescentar:

```
   **Usuários e perfis** (site): entrar com login e senha. Perfis: *administrador* (tudo: usuários, abrir/encerrar/
   excluir inventário, exclusões e importação de cadastros), *operador* (termos, cadastros, textos, atualizar base),
   *inventariante* (lê bens nos eventos em que está na comissão) e *consulta* (só vê; não emite termo). A comissão
   do inventário é escolhida pelo administrador entre os usuários; a leitura grava o nome de quem está logado.
   Usuário não é excluído, só inativado. *Nova senha* gera uma senha temporária mostrada uma vez, com troca
   obrigatória no primeiro acesso. Cinco senhas erradas seguidas bloqueiam o login por 15 minutos. O programa
   Windows não pede senha: entra como "Administrador local".
```

2. Na seção **Servir na web (Docker)**, substituir o texto sobre `auth_basic` e o comando `htpasswd` por:

```
O mesmo código roda em `https://patrimonio.sistemascfc.org`, num container nesta máquina, atrás do nginx do host
(container em `127.0.0.1:12012`, Certbot). O login é do próprio sistema (`TERMOS_LOGIN=1` no `compose.yml`).

    docker compose up -d --build   # (re)constrói e sobe; dados em ./dados (termos.db, timbrado.docx)
    docker compose logs -f         # acompanhar
    docker compose exec web python usuarios.py criar-admin antonio "Antônio Sousa"   # primeiro administrador (ou redefinir a senha de um admin)

Publicação da Fase 4 (uma vez): subir o container; criar o administrador pelo comando acima; entrar e criar os
usuários da comissão do evento aberto com **exatamente** os nomes já gravados nas leituras (Inventário → evento →
*Comissão* mostra quem já tem leituras); remover `auth_basic` e `auth_basic_user_file` do vhost
`/etc/nginx/conf.d/patrimonio.sistemascfc.org.conf` e recarregar (`nginx -t && systemctl reload nginx`).
Enquanto o `auth_basic` ficar, o site pede as duas senhas, sem prejuízo.
```

3. Na tabela **Arquivos**, acrescentar `| \`usuarios.py\`, \`app_usuarios.py\` | usuários, senhas, matriz de permissões e telas de login/usuários |`.

- [ ] **Step 3: Spec**

Em `docs/superpowers/specs/2026-09-17-fase4-usuarios-design.md`, linha `**Estado:**`, acrescentar ` Implementada em <commit final da branch> (2026-09-17).` com o hash do último commit.

- [ ] **Step 4: Revisão final**

Run: `.venv/bin/pytest -q` (esperado: 216 + ~40 novos, todos verdes) e `.venv/bin/python -c "import app"`.

Checagem manual rápida (`TERMOS_LOGIN=1 TERMOS_DADOS=/tmp/t4 .venv/bin/python app.py` numa pasta vazia): `/login` mostra "Nenhum usuário cadastrado"; `TERMOS_DADOS=/tmp/t4 .venv/bin/python usuarios.py criar-admin a "A"`; entrar; menu com Usuários; `/usuarios/novo`; abrir evento com comissão; ler um bem; sair. Sem `TERMOS_LOGIN`: `/` abre direto e não há "Sair".

- [ ] **Step 5: Commit**

```bash
git add compose.yml README.md docs/superpowers/specs/2026-09-17-fase4-usuarios-design.md
git commit -m "Fase 4: TERMOS_LOGIN no compose, README com perfis, primeiro admin e passos de publicação"
```

---

## Ordem e dependências

1 → 2 → 3 → 4 → 5 → 6 → 7 → 8 → 9 → 10. A suíte fica verde ao fim de cada tarefa. As tarefas 5 e 6 têm um item de menu ("Usuários") que só existe após a 6: siga a nota da Task 5, Step 4 (linha comentada até a Task 6).

## Fora deste plano

Fase 5 (menu em árvore, Início, Análise, Ajuda): spec própria, depois desta. Passos manuais de publicação (criar admin na VPS, remover `auth_basic`): do usuário, descritos no README.
