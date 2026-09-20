# Acesso ao SEI por usuário — plano de implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Cada emissão no SEI entra com a credencial (login, senha cifrada, unidade) de quem clicou, e cada "Atualizar com SPW" pelo site entra no SPW com a credencial de quem clicou — cadastradas pelo próprio usuário em "Meus acessos"; sem acesso cadastrado não se emite/atualiza. O cron da madrugada continua com `spw.env`.

**Architecture:** `cofre.py` (Fernet, chave em `secrets/chaves.env`) cifra a senha; `usuarios.py` guarda/lê o acesso; telas em `app_usuarios.py`; `app._enfileirar_emissao` exige acesso antes de registrar; `atender_pedidos.executar_sei` monta o `env` do pedido a partir de `criado_por` e chama `robo_sei.enviar_termo(env=...)`, que deixa de ler `sei.env` sozinho.

**Tech Stack:** Flask + Jinja, SQLite, `cryptography` (Fernet), pytest, DSGov (só visual).

**Spec:** `docs/superpowers/specs/2026-09-20-acesso-sei-por-usuario-design.md`

## Global Constraints

- Senha do SEI **nunca** em claro no banco, em log, em HTML ou em mensagem; nunca devolvida por `acesso_sei`/`listar`/`por_id` para templates (só `credencial_sei`, usada pelo trabalhador).
- Chave Fernet só em `secrets/chaves.env` (`CHAVE_SENHAS=…`), lida por `segredos.ler_env`; testes usam uma chave gerada em `tmp_path` via `monkeypatch.setattr(cofre, "ARQUIVO_CHAVE", …)`. Nunca commitar chaves.
- Mensagens exatas (§6 da spec): "Cadastre seu acesso ao SEI em Meus acessos para emitir."; "Quem pediu a emissão não tem acesso ao SEI cadastrado; cadastre em Meus acessos e emita de novo."; "O SEI recusou seu usuário ou senha; atualize em Meus acessos."; "Informe a senha do SEI."; "Informe a sigla da unidade no SEI."
- Rótulos: link do cabeçalho **"Meus acessos"**; título da tela **"Meus acessos"** com dois blocos, **"SEI"** e **"SPW"**; em cada bloco botões **"Salvar"** e **"Apagar"**; na lista de usuários, coluna **"Acessos"** ("SEI", "SPW", "SEI · SPW" ou "—") e botão **"Apagar acessos"** (apaga os dois). Interface sem a palavra "robô".
- Mensagens SPW: "Cadastre seu acesso ao SPW em Meus acessos para atualizar."; "Quem pediu a atualização não tem acesso ao SPW cadastrado; cadastre em Meus acessos e peça de novo."; "Informe a senha do SPW."; "Informe o usuário do SPW."
- Permissões: `usuarios.acessos GET`, `usuarios.salvar_acesso_sei POST`, `usuarios.apagar_acesso_sei POST`, `usuarios.salvar_acesso_spw POST`, `usuarios.apagar_acesso_spw POST` em TODAS; `usuarios.apagar_acessos POST` (admin, `/usuarios/<id>/apagar-acessos`) em ADMIN. Todas dão 404 sem login ativo (`_so_com_login_ligado`).
- `secrets/sei.env` só precisa de `SEI_LOGIN_URL` e `SEI_ORGAO` (`robo_sei.CHAVES_ENV`).
- `cryptography` em `requirements.txt` e no `pip install` do alvo `web` do Dockerfile (o `robo` herda).
- Testes sem rede/SEI. `.venv/bin/pytest` da raiz. Commits em português. Ramo `acesso-sei` a partir de `main`; merge ff ao final (modo automático autorizado); publicação pelo controlador (gerar chave, ajustar `sei.env`, `docker compose up -d --build`).

---

### Task 1: `cofre.py` — cifra Fernet com chave em `secrets/chaves.env`

**Files:**
- Create: `cofre.py`, `tests/test_cofre.py`
- Modify: `requirements.txt` (acrescentar `cryptography`), `Dockerfile:15` (acrescentar `cryptography` ao `pip install` do alvo `web`)

**Interfaces:**
- Produces: `cofre.ARQUIVO_CHAVE: Path`; `cofre.gerar_chave() -> str`; `cofre.cifrar(texto: str) -> str`; `cofre.decifrar(token: str) -> str`. Erros: `db.ErroDeNegocio` com a mensagem de `segredos.SegredoAusente` (chave ausente), `"CHAVE_SENHAS inválida em secrets/chaves.env."` (chave malformada), `"Senha cifrada com outra chave; cadastre a senha do SEI de novo."` (token não decifra).

- [ ] **Step 1: Testes que falham** — `tests/test_cofre.py`:

```python
import pytest

import cofre
import db


@pytest.fixture
def chave(tmp_path, monkeypatch):
    arq = tmp_path / "chaves.env"
    arq.write_text(f"CHAVE_SENHAS={cofre.gerar_chave()}\n")
    monkeypatch.setattr(cofre, "ARQUIVO_CHAVE", arq)
    return arq


def test_ida_e_volta_e_tokens_diferentes(chave):
    a, b = cofre.cifrar("Segredo!1"), cofre.cifrar("Segredo!1")
    assert a != b and "Segredo" not in a and cofre.decifrar(a) == cofre.decifrar(b) == "Segredo!1"


def test_sem_chave_e_chave_invalida(tmp_path, monkeypatch):
    monkeypatch.setattr(cofre, "ARQUIVO_CHAVE", tmp_path / "chaves.env")
    with pytest.raises(db.ErroDeNegocio, match="secrets/chaves.env não encontrado ou incompleto"):
        cofre.cifrar("x")
    (tmp_path / "chaves.env").write_text("CHAVE_SENHAS=nao-e-uma-chave\n")
    with pytest.raises(db.ErroDeNegocio, match="CHAVE_SENHAS inválida"):
        cofre.cifrar("x")


def test_token_de_outra_chave(chave, tmp_path, monkeypatch):
    token = cofre.cifrar("x")
    outra = tmp_path / "outra.env"
    outra.write_text(f"CHAVE_SENHAS={cofre.gerar_chave()}\n")
    monkeypatch.setattr(cofre, "ARQUIVO_CHAVE", outra)
    with pytest.raises(db.ErroDeNegocio, match="cifrada com outra chave"):
        cofre.decifrar(token)
```

- [ ] **Step 2: Rodar e ver falhar** — `.venv/bin/pytest tests/test_cofre.py -q` → `ModuleNotFoundError`.

- [ ] **Step 3: Implementar** — `cofre.py`:

```python
"""Cifra das senhas do SEI guardadas em usuarios.sei_senha. Fernet (cryptography) com a chave em
secrets/chaves.env (CHAVE_SENHAS=...). A mesma chave serve ao site (cifra ao salvar) e ao serviço robo (decifra
ao atender o pedido); os dois montam secrets/. Sem a chave, nada é salvo nem lido — e um backup do banco sozinho
não revela senha nenhuma. Gerar uma vez: python -c "import cofre; print(cofre.gerar_chave())"."""
from cryptography.fernet import Fernet, InvalidToken

import db
import segredos

ARQUIVO_CHAVE = segredos.PASTA / "chaves.env"


def gerar_chave() -> str:
    return Fernet.generate_key().decode()


def _fernet() -> Fernet:
    try:
        env = segredos.ler_env(ARQUIVO_CHAVE, ("CHAVE_SENHAS",))
    except segredos.SegredoAusente as e:
        raise db.ErroDeNegocio(str(e))
    try:
        return Fernet(env["CHAVE_SENHAS"].encode())
    except (ValueError, TypeError):
        raise db.ErroDeNegocio("CHAVE_SENHAS inválida em secrets/chaves.env.")


def cifrar(texto: str) -> str:
    return _fernet().encrypt(texto.encode("utf-8")).decode()


def decifrar(token: str) -> str:
    try:
        return _fernet().decrypt(token.encode()).decode("utf-8")
    except InvalidToken:
        raise db.ErroDeNegocio("Senha cifrada com outra chave; cadastre a senha do SEI de novo.")
```

`requirements.txt`: linha `cryptography`. `Dockerfile` linha 15: `RUN pip install --no-cache-dir flask openpyxl python-docx waitress Pillow boto3 cryptography`.

- [ ] **Step 4: Rodar e ver passar** — `.venv/bin/pytest tests/test_cofre.py -q` → 3 passed.
- [ ] **Step 5: Commit** — `git add cofre.py tests/test_cofre.py requirements.txt Dockerfile && git commit -m "feat: cofre — cifra Fernet das senhas do SEI com chave em secrets/chaves.env"`

---

### Task 2: Acesso ao SEI em `usuarios.py` (colunas, salvar, apagar, credencial)

**Files:**
- Modify: `db.py` (`criar_esquema`: colunas `sei_login`, `sei_senha`, `sei_unidade`, `sei_atualizado_em` em `usuarios`, guardadas por `_colunas`), `usuarios.py` (`_COLUNAS_LISTA` ganha `sei_login, sei_atualizado_em`; funções novas; `por_id`/`por_login`/`por_email` deixam de trazer `sei_senha`)
- Test: `tests/test_usuarios.py`

**Interfaces:**
- Consumes: `cofre.cifrar/decifrar`.
- Produces: `usuarios.acesso_sei(conn, id) -> dict | None` (`{"login", "unidade", "atualizado_em"}`); `usuarios.salvar_acesso_sei(conn, id, login, senha, unidade) -> None`; `usuarios.apagar_acesso_sei(conn, id) -> None`; `usuarios.credencial_sei(conn, login_sistema) -> dict | None` (`{"SEI_USUARIO", "SEI_SENHA", "SEI_UNIDADE"}`); **idem para o SPW**: `acesso_spw(conn, id) -> {"login", "atualizado_em"} | None`, `salvar_acesso_spw(conn, id, login, senha)` (mensagens "Informe o usuário do SPW." / "Informe a senha do SPW."), `apagar_acesso_spw(conn, id)`, `credencial_spw(conn, login_sistema) -> {"SPW_USUARIO", "SPW_SENHA"} | None`, colunas `spw_login`, `spw_senha`, `spw_atualizado_em`; `listar` traz `sei_login`, `sei_atualizado_em`, `spw_login`, `spw_atualizado_em`. Implementar as funções SPW com a mesma estrutura das do SEI (sem unidade) e um teste espelho `test_acesso_spw_salvar_manter_apagar`.

- [ ] **Step 1: Testes que falham** — acrescentar a `tests/test_usuarios.py` (reaproveitar a fixture `chave` copiando-a de `tests/test_cofre.py` para `tests/conftest.py` — mover a fixture para o conftest e apagá-la do test_cofre):

```python
def test_acesso_sei_salvar_manter_apagar(dados, chave):
    uid = usuarios.criar(dados, "maria", "Maria", SENHA_PADRAO, ["operador"])
    assert usuarios.acesso_sei(dados, uid) is None and usuarios.credencial_sei(dados, "maria") is None
    with pytest.raises(db.ErroDeNegocio, match="Informe a senha do SEI."):
        usuarios.salvar_acesso_sei(dados, uid, "maria.silva", "", "gecont")
    with pytest.raises(db.ErroDeNegocio, match="Informe a sigla da unidade no SEI."):
        usuarios.salvar_acesso_sei(dados, uid, "maria.silva", "S3nha", " ")
    with pytest.raises(db.ErroDeNegocio, match="Informe o usuário do SEI."):
        usuarios.salvar_acesso_sei(dados, uid, "", "S3nha", "GECONT")
    usuarios.salvar_acesso_sei(dados, uid, " maria.silva ", "S3nha", " gecont ")
    a = usuarios.acesso_sei(dados, uid)
    assert a["login"] == "maria.silva" and a["unidade"] == "GECONT" and a["atualizado_em"] and "senha" not in a
    linha = dados.execute("SELECT sei_senha FROM usuarios WHERE id=?", (uid,)).fetchone()[0]
    assert linha and "S3nha" not in linha
    assert usuarios.credencial_sei(dados, "maria") == {"SEI_USUARIO": "maria.silva", "SEI_SENHA": "S3nha", "SEI_UNIDADE": "GECONT"}
    usuarios.salvar_acesso_sei(dados, uid, "maria.silva", "", "GESERV")            # senha vazia mantém
    assert usuarios.credencial_sei(dados, "maria")["SEI_SENHA"] == "S3nha" and usuarios.acesso_sei(dados, uid)["unidade"] == "GESERV"
    assert "sei_senha" not in usuarios.por_id(dados, uid) and "sei_senha" not in usuarios.por_login(dados, "maria")
    lista = [u for u in usuarios.listar(dados) if u["login"] == "maria"][0]
    assert lista["sei_login"] == "maria.silva" and lista["sei_atualizado_em"] and "sei_senha" not in lista
    usuarios.apagar_acesso_sei(dados, uid)
    assert usuarios.acesso_sei(dados, uid) is None and usuarios.credencial_sei(dados, "maria") is None
    assert usuarios.credencial_sei(dados, "nao-existe") is None
```

- [ ] **Step 2: Rodar e ver falhar** — `.venv/bin/pytest tests/test_usuarios.py -q -k acesso_sei` → FAIL.

- [ ] **Step 3: Implementar**

`db.py`, em `criar_esquema` (junto das outras migrações guardadas):

```python
    # Acesso ao SEI por usuário (2026-09-20): login, senha cifrada (cofre.py) e unidade de quem emite.
    for coluna in ("sei_login", "sei_senha", "sei_unidade", "sei_atualizado_em"):
        if coluna not in _colunas(conn, "usuarios"):
            conn.execute(f"ALTER TABLE usuarios ADD COLUMN {coluna} TEXT")
```

`usuarios.py`: `_COLUNAS_LISTA = "id, login, email, nome, ativo, trocar_senha, falhas, bloqueado_ate, criado_em, ultimo_acesso, sei_login, sei_atualizado_em"`; `_COLUNAS_CONTA = _COLUNAS_LISTA + ", senha_hash, sei_unidade"` e `por_id`/`por_login`/`por_email` passam a `SELECT {_COLUNAS_CONTA}` (conferir que `autenticar`/`trocar_senha` usam `senha_hash` desses dicts — continuam funcionando). Funções novas, após `trocar_senha`:

```python
# ---------------------------------------------------------------- acesso ao SEI
def acesso_sei(conn, id) -> dict | None:
    """Login, unidade e data do acesso ao SEI de um usuário — nunca a senha."""
    r = _um(conn, "SELECT sei_login, sei_unidade, sei_atualizado_em FROM usuarios WHERE id=?", id)
    if not r or not (r["sei_login"] and r["sei_unidade"]):
        return None
    return {"login": r["sei_login"], "unidade": r["sei_unidade"], "atualizado_em": r["sei_atualizado_em"]}


def salvar_acesso_sei(conn, id, login, senha, unidade) -> None:
    """Senha vazia mantém a atual (erro se não há atual). Cifra com cofre.py; grava a data."""
    import cofre
    login = " ".join(str(login or "").split())
    unidade = " ".join(str(unidade or "").split()).upper()
    senha = str(senha or "")
    if not login:
        raise ErroDeNegocio("Informe o usuário do SEI.")
    if not unidade:
        raise ErroDeNegocio("Informe a sigla da unidade no SEI.")
    atual = _um(conn, "SELECT sei_senha FROM usuarios WHERE id=?", id)
    if atual is None:
        raise ErroDeNegocio("Usuário não encontrado.")
    if not senha and not atual["sei_senha"]:
        raise ErroDeNegocio("Informe a senha do SEI.")
    cifrada = cofre.cifrar(senha) if senha else atual["sei_senha"]
    with conn:
        conn.execute("UPDATE usuarios SET sei_login=?, sei_senha=?, sei_unidade=?, sei_atualizado_em=? WHERE id=?",
                     (login, cifrada, unidade, _agora(), id))


def apagar_acesso_sei(conn, id) -> None:
    with conn:
        conn.execute("UPDATE usuarios SET sei_login=NULL, sei_senha=NULL, sei_unidade=NULL, sei_atualizado_em=NULL WHERE id=?", (id,))


def credencial_sei(conn, login_sistema) -> dict | None:
    """Só para o trabalhador: credencial decifrada de quem pediu a emissão (None se não há acesso)."""
    import cofre
    r = _um(conn, "SELECT sei_login, sei_senha, sei_unidade FROM usuarios WHERE login=?", str(login_sistema or "").strip().lower())
    if not r or not (r["sei_login"] and r["sei_senha"] and r["sei_unidade"]):
        return None
    return {"SEI_USUARIO": r["sei_login"], "SEI_SENHA": cofre.decifrar(r["sei_senha"]), "SEI_UNIDADE": r["sei_unidade"]}
```

(`import cofre` dentro das funções evita importar `cryptography` no desktop/testes que não usam.)

- [ ] **Step 4: Rodar e ver passar** — `.venv/bin/pytest tests/test_usuarios.py tests/test_login.py tests/test_usuarios_telas.py tests/test_cofre.py -q` → PASS.
- [ ] **Step 5: Commit** — `git add db.py usuarios.py tests/test_usuarios.py tests/conftest.py tests/test_cofre.py && git commit -m "feat: acesso ao SEI por usuário (login, senha cifrada e unidade)"`

---

### Task 3: Telas — "Meus acessos" (SEI e SPW), cabeçalho, lista de usuários, permissões, Ajuda

**Files:**
- Modify: `app_usuarios.py` (rotas: `acessos` GET `/meus-acessos`; `salvar_acesso_sei` POST `/meus-acessos/sei`; `apagar_acesso_sei` POST `/meus-acessos/sei/apagar`; `salvar_acesso_spw` POST `/meus-acessos/spw`; `apagar_acesso_spw` POST `/meus-acessos/spw/apagar`; `apagar_acessos` POST `/usuarios/<int:id>/apagar-acessos` — admin, apaga SEI e SPW), `permissoes.py`, `menu.py` (`AJUDA['usuarios.acesso_sei'] = 'conta'`; `_ORIGEM`/redirecionamentos se houver), `templates/base.html:33` (link "Acesso ao SEI"), `templates/usuarios/lista.html` (coluna + botão), `templates/ajuda.html` (seção `conta`)
- Create: `templates/acessos.html` (dois `br-card`: "SEI" e "SPW", cada um com seu form de salvar e, se houver acesso, seu form de apagar)
- Test: `tests/test_usuarios_telas.py`, `tests/test_permissoes.py`, `tests/test_ajuda.py`

**Interfaces:**
- Consumes: Task 2.
- Produces: os seis endpoints acima. Os testes do Step 1 devem ser adaptados aos caminhos novos (`/meus-acessos`, `/meus-acessos/sei`, `/meus-acessos/sei/apagar`, `/meus-acessos/spw`, `/meus-acessos/spw/apagar`, `/usuarios/<id>/apagar-acessos`) e aos rótulos ("Meus acessos", "Acessos", "Apagar acessos"); acrescentar ao teste da tela o bloco SPW (salvar `spw_login`/`spw_senha`, senha nunca no HTML, apagar) e à lista a coluna com "SEI · SPW". O código do Step 3 é o modelo: escrever as rotas/template com os nomes novos, replicando o bloco SEI para o SPW (sem unidade).

- [ ] **Step 1: Testes que falham** — em `tests/test_usuarios_telas.py`:

```python
def test_meu_acesso_sei_salvar_manter_e_apagar(cliente, dados, chave):
    r = cliente.get("/meu-acesso-sei")
    assert r.status_code == 200 and b"Meu acesso ao SEI" in r.data and b'name="sei_login"' in r.data and b"Apagar meu acesso" not in r.data
    r = cliente.post("/meu-acesso-sei", data={"sei_login": "antonio.junior", "sei_senha": "", "sei_unidade": "GELIC"}, follow_redirects=True)
    assert "Informe a senha do SEI.".encode() in r.data
    r = cliente.post("/meu-acesso-sei", data={"sei_login": "antonio.junior", "sei_senha": "S3nha!", "sei_unidade": "gelic"}, follow_redirects=True)
    assert b"Acesso ao SEI salvo." in r.data
    uid = usuarios.por_login(dados, ADMIN_LOGIN)["id"]
    assert usuarios.credencial_sei(dados, ADMIN_LOGIN)["SEI_SENHA"] == "S3nha!" and usuarios.acesso_sei(dados, uid)["unidade"] == "GELIC"
    r = cliente.get("/meu-acesso-sei")
    assert b"S3nha" not in r.data and b"Senha cadastrada em" in r.data and b"Apagar meu acesso" in r.data and b'value="antonio.junior"' in r.data
    cliente.post("/meu-acesso-sei", data={"sei_login": "antonio.junior", "sei_senha": "", "sei_unidade": "GESERV"})
    assert usuarios.credencial_sei(dados, ADMIN_LOGIN) == {"SEI_USUARIO": "antonio.junior", "SEI_SENHA": "S3nha!", "SEI_UNIDADE": "GESERV"}
    r = cliente.post("/meu-acesso-sei/apagar", follow_redirects=True)
    assert b"Acesso ao SEI apagado." in r.data and usuarios.acesso_sei(dados, uid) is None
    assert b"Acesso ao SEI" in cliente.get("/").data                                  # link no cabeçalho


def test_lista_de_usuarios_mostra_acesso_e_admin_apaga_sem_ver_senha(cliente, dados, chave, usuarios_exemplo):
    uid = usuarios.por_login(dados, "op")["id"]
    usuarios.salvar_acesso_sei(dados, uid, "op.sei", "Outra!", "GECONT")
    r = cliente.get("/usuarios")
    html = r.data.decode()
    assert "Acesso ao SEI" in html and "op.sei" in html and "Outra!" not in html and "Apagar acesso" in html
    r = cliente.post(f"/usuarios/{uid}/apagar-acesso-sei", follow_redirects=True)
    assert "Acesso ao SEI de op apagado.".encode() in r.data and usuarios.acesso_sei(dados, uid) is None
    assert cliente.get("/usuarios").data.count(b"Apagar acesso") == 0


def test_acesso_sei_por_funcao_e_desktop(cliente, dados, usuarios_exemplo, cliente_local):
    cliente.post("/sair")
    assert logar(cliente, *usuarios_exemplo["consulta"]).status_code == 302
    assert cliente.get("/meu-acesso-sei").status_code == 200                          # qualquer função
    assert cliente.post("/usuarios/1/apagar-acesso-sei").status_code == 403           # só admin
    assert cliente_local.get("/meu-acesso-sei").status_code == 404                    # desktop: não existe
```

(`cliente` e `cliente_local` juntos: cada um cria seu app; se conflitar, separar o caso desktop num teste próprio.) Em `tests/test_permissoes.py::_ROTAS_GET`: `"/meus-acessos": {"admin": 200, "operador": 200, "inventariante": 200, "consulta": 200}`; em `_ROTAS_POST`: `"/usuarios/1/apagar-acessos": {"admin": 302, "operador": 403, "inventariante": 403, "consulta": 403}` (o admin cai em 302 mesmo sem acesso cadastrado: a rota só apaga e redireciona). Em `tests/test_ajuda.py`, junto de `('/senha', 'usuarios.senha', {})`: `('/meus-acessos', 'usuarios.acessos', {})` e `assert menu.ancora_ajuda({'admin'}, 'usuarios.acessos', {}, True) == 'conta'`.

- [ ] **Step 2: Rodar e ver falhar** — `.venv/bin/pytest tests/test_usuarios_telas.py tests/test_permissoes.py tests/test_ajuda.py -q` → FAIL.

- [ ] **Step 3: Implementar**

`app_usuarios.py`, após `senha()`:

```python
@usuarios_bp.route("/meu-acesso-sei", methods=["GET", "POST"])
def acesso_sei():
    """Credencial do SEI com que este usuário emite termos; a senha nunca volta para a tela."""
    _so_com_login_ligado()
    conn = _conn()
    if request.method == "POST":
        usuarios.salvar_acesso_sei(conn, g.usuario["id"], request.form.get("sei_login"), request.form.get("sei_senha"),
                                   request.form.get("sei_unidade"))
        flash("Acesso ao SEI salvo.", "success")
        return redirect(url_for("usuarios.acesso_sei"))
    return render_template("acesso_sei.html", acesso=usuarios.acesso_sei(conn, g.usuario["id"]), trilha=[("Meu acesso ao SEI", None)])


@usuarios_bp.route("/meu-acesso-sei/apagar", methods=["POST"])
def apagar_meu_acesso_sei():
    _so_com_login_ligado()
    usuarios.apagar_acesso_sei(_conn(), g.usuario["id"])
    flash("Acesso ao SEI apagado.", "success")
    return redirect(url_for("usuarios.acesso_sei"))


@usuarios_bp.route("/usuarios/<int:id>/apagar-acesso-sei", methods=["POST"])
def apagar_acesso_sei(id):
    """Admin remove o acesso de alguém (desligamento); nunca vê nem define a senha."""
    _so_com_login_ligado()
    u = usuarios.por_id(_conn(), id) or abort(404)
    usuarios.apagar_acesso_sei(_conn(), id)
    flash(f"Acesso ao SEI de {u['login']} apagado.", "success")
    return redirect(_retorno())
```

(`ErroDeNegocio` do `salvar_acesso_sei` cai no handler global: flash + volta à tela.) `permissoes.py`: em TODAS `usuarios.acesso_sei GET POST` e `usuarios.apagar_meu_acesso_sei POST`; em ADMIN `usuarios.apagar_acesso_sei POST`. `menu.py`: `AJUDA['usuarios.acessos'] = 'conta'` e, no bloco que redireciona endpoints de usuários para `usuarios.lista` (linha ~39), acrescentar `usuarios.apagar_acesso_sei`.

`templates/acesso_sei.html`:

```jinja
{% extends "base.html" %}
{% from '_macros.html' import ajuda_titulo %}
{% block titulo %}Meu acesso ao SEI{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Meu acesso ao SEI</h1>{{ ajuda_titulo(AJUDA_ANCORA) }}</div>
<p class="text-gray-70">O termo é criado no SEI com este usuário e nesta unidade, em seu nome. Os blocos de assinatura "Termos {sigla do centro}" precisam existir nessa unidade.</p>
<form method="post" action="{{ url_for('usuarios.acesso_sei') }}" class="col-md-6">{{ csrf_campo() }}
  <div class="br-input mb-3"><label for="sei_login">Usuário do SEI</label><input id="sei_login" name="sei_login" type="text" value="{{ acesso.login if acesso else '' }}" autocomplete="off" required/></div>
  <div class="br-input mb-3"><label for="sei_senha">Senha do SEI{% if acesso %} (deixe em branco para manter){% endif %}</label><input id="sei_senha" name="sei_senha" type="password" autocomplete="new-password" {% if not acesso %}required{% endif %}/>
    {% if acesso and acesso.atualizado_em %}<p class="text-down-01 text-gray-70 mt-1 mb-0">Senha cadastrada em {{ acesso.atualizado_em[8:10] }}/{{ acesso.atualizado_em[5:7] }}/{{ acesso.atualizado_em[:4] }}.</p>{% endif %}</div>
  <div class="br-input mb-4"><label for="sei_unidade">Unidade no SEI (sigla)</label><input id="sei_unidade" name="sei_unidade" type="text" value="{{ acesso.unidade if acesso else '' }}" required/></div>
  <button class="br-button primary" type="submit"><i class="fas fa-save mr-1" aria-hidden="true"></i>Salvar</button>
  <a class="br-button ml-2" href="{{ URL_INICIAL }}">Cancelar</a>
</form>
{% if acesso %}
<form method="post" action="{{ url_for('usuarios.apagar_meu_acesso_sei') }}" class="mt-4" onsubmit="return confirm('Apagar seu acesso ao SEI? Você deixará de conseguir emitir termos até cadastrar de novo.');">{{ csrf_campo() }}
  <button class="br-button secondary" type="submit"><i class="fas fa-trash mr-1" aria-hidden="true"></i>Apagar meu acesso</button>
</form>
{% endif %}
{% endblock %}
```

`templates/base.html` linha 33: antes de *Trocar senha*, `<a class="br-button small" href="{{ url_for('usuarios.acesso_sei') }}">Acesso ao SEI</a>`. `templates/usuarios/lista.html`: `<th scope="col">Acesso ao SEI</th>` antes de "Último acesso"; célula `{% if u.sei_login %}<code>{{ u.sei_login }}</code> · {{ u.sei_atualizado_em[8:10] }}/{{ u.sei_atualizado_em[5:7] }}{% else %}—{% endif %}`; nas ações, quando `u.sei_login`, um `form` POST para `usuarios.apagar_acesso_sei` com `retorno` e botão `br-button circle small` (ícone `fa-user-slash`, `aria-label="Apagar acesso ao SEI de {{ u.login }}"`, texto visualmente oculto "Apagar acesso" — use `<span class="sr-only">Apagar acesso</span>`), com `onsubmit="return confirm(...)"`. `templates/ajuda.html` seção `conta`: acrescentar "Em Acesso ao SEI, cadastre seu usuário, senha e unidade do SEI: os termos que você emitir nascem no SEI em seu nome. A senha fica cifrada e nunca é exibida; o administrador só pode apagar o acesso."

- [ ] **Step 4: Rodar e ver passar** — `.venv/bin/pytest tests/test_usuarios_telas.py tests/test_permissoes.py tests/test_ajuda.py tests/test_menu.py -q` → PASS.
- [ ] **Step 5: Commit** — `git add app_usuarios.py permissoes.py menu.py templates/acesso_sei.html templates/base.html templates/usuarios/lista.html templates/ajuda.html tests/ && git commit -m "feat: tela Meu acesso ao SEI, coluna e apagar acesso na lista de usuários"`

---

### Task 4: Fluxo — exigir acesso ao emitir/atualizar; trabalhador monta o env de quem pediu (SEI e SPW)

**Files:**
- Modify: `app.py` (`_enfileirar_emissao` e `termo_enviar_sei`: checar acesso antes de registrar; `base_atualizar_spw`: checar `usuarios.acesso_spw` antes de enfileirar → "Cadastre seu acesso ao SPW em Meus acessos para atualizar."), `atender_pedidos.py` (`executar_sei`; `executar_spw` passa a montar `env` = `spw.env` com `SPW_USUARIO`/`SPW_SENHA` substituídos por `credencial_spw(criado_por)` quando o pedido tem `criado_por`; sem credencial → `erro` "Quem pediu a atualização não tem acesso ao SPW cadastrado; cadastre em Meus acessos e peça de novo."; `EXECUTORES`), `importar_spw.py` (`executar(conn, baixar=None, agora=None, env=None)` e `baixar_e_ler(env=None)` → `baixar_export(env or ler_env(), ...)`), `robo_sei.py` (`CHAVES_ENV`, `enviar_termo` exige `env`, mensagem de login recusado)
- Test: `tests/test_app.py`, `tests/test_atender_pedidos.py`, `tests/test_robo_sei.py`

**Interfaces:**
- Consumes: `usuarios.acesso_sei`, `usuarios.credencial_sei`, `segredos.ler_env`.
- Produces: `atender_pedidos.executar_spw(conn, pedido, executar=None) -> dict` (`executar` injetável = `importar_spw.executar`, recebe `env=`); `atender_pedidos.executar_sei(conn, pedido, enviar=None) -> dict` (executor do tipo `sei`; `enviar` injetável = `robo_sei.enviar_termo`); `robo_sei.CHAVES_ENV = ("SEI_LOGIN_URL", "SEI_ORGAO")`; `robo_sei.enviar_termo(conn, pedido, abrir=None, env=None)` com `env` obrigatório (`None` → pedido em `erro` "Credencial do SEI não informada ao robô." — mensagem interna, não deve acontecer).

- [ ] **Step 1: Testes que falham**

`tests/test_app.py`: nos testes de emissão existentes que criam pedido (`test_emitir_registra_numera_e_enfileira`, `test_estados_da_pagina_do_termo_emitido`, `test_pagina_sem_pedido…`, `test_erro_*`), cadastrar o acesso do admin antes: helper no topo do arquivo

```python
def _acesso_admin(dados, chave):
    import usuarios
    usuarios.salvar_acesso_sei(dados, usuarios.por_login(dados, ADMIN_LOGIN)["id"], "antonio.junior", "S3nha!", "GELIC")
```

(adicionar `chave` à assinatura desses testes). Teste novo:

```python
def test_emitir_sem_acesso_ao_sei_nao_registra(cliente, dados, chave):
    _processo_ccusto(cliente)
    r = cliente.post("/termo/ccusto/CCI/enviar-sei", follow_redirects=True)
    assert "Cadastre seu acesso ao SEI em Meus acessos para emitir.".encode() in r.data
    assert db.termos_emitidos(dados) == [] and db.pedido_do_termo(dados, 1) is None
    _acesso_admin(dados, chave)
    assert cliente.post("/termo/ccusto/CCI/enviar-sei").status_code == 302
    assert db.pedido_do_termo(dados, 1)["criado_por"] == ADMIN_LOGIN
```

`tests/test_atender_pedidos.py` — além do teste do SEI abaixo, um espelho para o SPW: `test_executar_spw_usa_credencial_de_quem_pediu` (usuário com `salvar_acesso_spw(dados, uid, "maria.spw", "S3nha", )`; `spw.env` em tmp com as quatro chaves e `importar_spw.ARQUIVO_ENV` monkeypatched; `ap.executar_spw(dados, pedido, executar=falso)` recebe `env` com `SPW_USUARIO == "maria.spw"`, `SPW_SENHA == "S3nha"` e as URLs do arquivo; pedido de `criado_por` sem acesso → `erro` com a mensagem exata; pedido **sem** `criado_por` → `env` é o `spw.env` completo, sem erro). Em `tests/test_app.py`: `test_atualizar_com_spw_sem_acesso_avisa` (POST sem acesso → flash "Cadastre seu acesso ao SPW em Meus acessos para atualizar." e `pedido_spw_ativo` None; com acesso → 302 e pedido). Ajustar `test_atualizar_com_spw_enfileira_e_mostra_andamento` para cadastrar o acesso SPW do admin antes (helper `_acesso_spw_admin`). Em `tests/test_importar_spw.py`, um teste de que `executar(conn, baixar=..., env={...})` repassa `env` a `baixar` (tornar `baixar` chamado como `baixar(env)` quando `env` é dado; manter compatível com os testes que passam `baixar=lambda: linhas`: use `inspect.signature` ou simplesmente `baixar(env) if env is not None else baixar()`).

```python
def test_executar_sei_monta_env_de_quem_pediu(dados, chave, tmp_path, monkeypatch):
    import robo_sei, segredos, usuarios
    semear(dados)
    (tmp_path / "sei.env").write_text("SEI_LOGIN_URL=https://sei.cfc.org.br/sei/\nSEI_ORGAO=CFC\n")
    monkeypatch.setattr(robo_sei, "ARQUIVO_ENV", tmp_path / "sei.env")
    uid = usuarios.criar(dados, "maria", "Maria", "Senha!234", ["operador"])
    usuarios.salvar_acesso_sei(dados, uid, "maria.silva", "S3nha", "GECONT")
    db.incluir_processo(dados, "ccusto", "T", "1111")
    t = db.preparar_envio_sei(dados, db.registrar_emissao(dados, "ccusto", "CCI", db.bens_do_centro(dados, "CCI"))["id"])
    pid = db.enfileirar_pedido(dados, "sei", termo_id=t["id"], html="<p>x</p>", criado_por="maria")
    recebido = {}
    def enviar(conn, pedido, env=None):
        recebido.update(env); db.marcar_passo(conn, pedido["id"], "concluido", "ok"); return {"passo": "concluido"}
    ap.executar_sei(dados, db.pedido(dados, pid), enviar=enviar)
    assert recebido == {"SEI_LOGIN_URL": "https://sei.cfc.org.br/sei/", "SEI_ORGAO": "CFC", "SEI_USUARIO": "maria.silva", "SEI_SENHA": "S3nha", "SEI_UNIDADE": "GECONT"}
    # sem credencial (acesso apagado depois de clicar)
    usuarios.apagar_acesso_sei(dados, uid)
    pid2 = db.enfileirar_pedido(dados, "sei", termo_id=t["id"], html="<p>x</p>", criado_por="maria")
    chamado = []
    ap.executar_sei(dados, db.pedido(dados, pid2), enviar=lambda *a, **k: chamado.append(1))
    p = db.pedido(dados, pid2)
    assert p["passo"] == "erro" and p["mensagem"] == "Quem pediu a emissão não tem acesso ao SEI cadastrado; cadastre em Meus acessos e emita de novo." and not chamado
    # sem chave do cofre
    usuarios.salvar_acesso_sei(dados, uid, "maria.silva", "S3nha", "GECONT")
    import cofre
    monkeypatch.setattr(cofre, "ARQUIVO_CHAVE", tmp_path / "nao-existe.env")
    pid3 = db.enfileirar_pedido(dados, "sei", termo_id=t["id"], html="<p>x</p>", criado_por="maria")
    ap.executar_sei(dados, db.pedido(dados, pid3), enviar=lambda *a, **k: chamado.append(1))
    assert db.pedido(dados, pid3)["mensagem"].startswith("secrets/chaves.env não encontrado") and not chamado
```

`tests/test_robo_sei.py`: em `test_login_recusado_tipo_inexistente_e_timeout`, a mensagem do login recusado passa a `"O SEI recusou seu usuário ou senha; atualize em Meus acessos."`; `test_env_ausente_vira_erro_legivel` passa a chamar `enviar_termo(dados, p, abrir=...)` sem `env` e esperar `mensagem == "Credencial do SEI não informada ao robô."` e `passo == "erro"`. `ENV` do arquivo ganha `"SEI_UNIDADE": "GELIC"` só onde os testes de troca de unidade já o definem (não mudar os demais).

- [ ] **Step 2: Rodar e ver falhar** — `.venv/bin/pytest tests/test_app.py tests/test_atender_pedidos.py tests/test_robo_sei.py -q` → FAIL.

- [ ] **Step 3: Implementar**

`app.py`: em `_enfileirar_emissao`, antes de `preparar_envio_sei`:

```python
    import usuarios as usuarios_mod
    if not usuarios_mod.acesso_sei(conn, g.usuario["id"]):
        raise db.ErroDeNegocio("Cadastre seu acesso ao SEI em Meus acessos para emitir.")
```

e em `termo_enviar_sei`, a mesma checagem logo após `db.unidade_sei(...)` (antes de `registrar_emissao`) — extrair um helper `_exigir_acesso_sei(conn)` usado nos dois lugares (`termo_emitido_enviar_sei` passa por `_enfileirar_emissao`).

`atender_pedidos.py`:

```python
def executar_sei(conn, pedido: dict, enviar=None) -> dict:
    """Monta a credencial de quem pediu (usuarios.credencial_sei) + URL/órgão do sei.env e chama o robô."""
    enviar = enviar or robo_sei.enviar_termo
    try:
        env = segredos.ler_env(robo_sei.ARQUIVO_ENV, robo_sei.CHAVES_ENV)
        credencial = usuarios.credencial_sei(conn, pedido.get("criado_por"))
    except (segredos.SegredoAusente, db.ErroDeNegocio) as e:
        db.marcar_passo(conn, pedido["id"], "erro", str(e)[:500])
        return {"passo": "erro", "mensagem": str(e)}
    if not credencial:
        mensagem = "Quem pediu a emissão não tem acesso ao SEI cadastrado; cadastre em Meus acessos e emita de novo."
        db.marcar_passo(conn, pedido["id"], "erro", mensagem)
        return {"passo": "erro", "mensagem": mensagem}
    return enviar(conn, pedido, env={**env, **credencial})


EXECUTORES = {"sei": executar_sei, "spw": executar_spw}
```

(`import segredos`, `import usuarios` no topo.) `robo_sei.py`: `CHAVES_ENV = ("SEI_LOGIN_URL", "SEI_ORGAO")`; em `enviar_termo`, trocar `env = env or segredos.ler_env(...)` por `if not env: raise RoboErro("Credencial do SEI não informada ao robô.")`; a mensagem de login recusado → `"O SEI recusou seu usuário ou senha; atualize em Meus acessos."`. Docstring do módulo: credencial vem do pedido (usuário que clicou), `sei.env` só tem URL e órgão.

- [ ] **Step 4: Rodar e ver passar** — `.venv/bin/pytest -q` (suíte completa) → PASS.
- [ ] **Step 5: Commit** — `git add app.py atender_pedidos.py robo_sei.py importar_spw.py tests/ && git commit -m "feat: emissão no SEI e atualização com SPW pelo site usam a credencial de quem clicou"`

---

### Task 5: README, spec (§10), merge e publicação

- [ ] **Step 1: README** — seção "Acessos por usuário (SEI e SPW)" (gerar `secrets/chaves.env` com `python -c "import cofre; print(cofre.gerar_chave())"`, chmod 600; **guardar a chave junto com o backup do banco** — sem ela as senhas cifradas são inúteis; `sei.env` só com `SEI_LOGIN_URL` e `SEI_ORGAO`; cada operador cadastra em *Meus acessos* (SEI e SPW); ao trocar a senha no SEI/SPW, atualizar na tela; admin só apaga; o cron da madrugada continua com `spw.env`). Ajustar a seção "Emissão no SEI" (não cita mais SEI_USUARIO/SEI_SENHA/SEI_UNIDADE). Tabela de arquivos: `cofre.py`. Commit `docs: README do acesso ao SEI por usuário`.
- [ ] **Step 2: Suíte completa e merge** — `.venv/bin/pytest -q`; `git checkout main && git merge --ff-only acesso-sei && git branch -d acesso-sei`.
- [ ] **Step 3: Publicação (controlador)** — gerar `secrets/chaves.env` (chmod 600); editar `secrets/sei.env` deixando só `SEI_LOGIN_URL` e `SEI_ORGAO`; `./backup.sh`; `docker compose up -d --build`; conferir `docker compose logs robo` ("trabalhador iniciado") e `curl /login` 200; cadastrar o acesso do usuário `antonioj` **não** (é dele: ele cadastra pela tela — mas o controlador pode registrar em §10 como pendência); evidência: `usuarios.acesso_sei` de um usuário de teste do controlador? Não criar usuários de produção — só registrar na spec §10 o que foi publicado e o que fica para o usuário.
- [ ] **Step 4: Spec §10** — evidências (commits, suíte, publicação) e pendência: cada operador cadastra seu acesso; blocos por unidade. Commit `docs: evidências do acesso ao SEI por usuário`.
