# Processos SEI, histórico e painel — plano de implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Só emitir termos com processo SEI vigente, registrando uma foto de cada emissão; rastrear o que muda a cada importação e apontar termos desatualizados; painel com gráficos clicáveis e tela de recorte com drill-down e exportação.

**Architecture:** Só acréscimos: cinco tabelas novas em `db.py` (processos, termos emitidos + foto, importações + mudanças), funções puras de consulta em `db.py`, rotas em `app.py`, e um módulo novo `painel.py` que transforma os agrupamentos de `db.dimensoes()` em cards de gráfico (ECharts, tema DSGov, helper `graficos.py` copiado da skill `/dsgov`). Templates Jinja com uma macro `grafico` reaproveitada no painel e no recorte.

**Tech Stack:** Python 3.12, Flask, SQLite (stdlib), openpyxl, ECharts 5.5.0 (embutido, offline), DSGov 3.7.0, pytest.

**Spec:** `docs/superpowers/specs/2026-09-14-processos-sei-historico-painel-design.md`

## Global Constraints

- Nada do esquema existente muda; só `CREATE TABLE IF NOT EXISTS` novos em `ESQUEMA`.
- `termos_html.py`, `textos.py` e os três geradores `.docx` não são tocados.
- Funções de `db.py` recebem `conn` como 1º argumento e não importam Flask.
- Erros para o usuário são `db.ErroDeNegocio` (viram flash pelo handler existente em `app.py`).
- Datas/horas gravadas em ISO `YYYY-MM-DD HH:MM:SS` (hora local), via `db._agora()`.
- Só inline `style` permitido: `termos_html.py`. Gráficos usam as classes `dsgov-grafico`/`dsgov-grafico-alto`.
- Textos de tela em português, sem gíria; botões DSGov (`br-button primary/secondary`), mensagens `br-message`.
- Rodar a suíte inteira antes de cada commit: `.venv/bin/pytest -q` (92 testes passam na base `574f468`).
- Commits em português, um por tarefa, com o rodapé:
  ```
  Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>
  Claude-Session: https://claude.ai/code/session_01Xz5P5QQJyvfjPDkCPVSG18
  ```
- Fixtures: `dados` (conexão com esquema, `tests/conftest.py`), `semear(conn)` (1 centro CCI com sala "01 - SALA CCI", bens 1001 CADEIRA, 1002 NOTEBOOK/DELL atribuído a ANA SILVA, 1003 MESA BAIXADO, 1004 ARMÁRIO em "99 - SEM MAPA"), `cliente` (Flask test client já semeado, em `tests/test_app.py`), `xlsx(tmp_path, linhas)` e `CABECALHO` em `tests/test_db.py`.

---

## Mapa de arquivos

| arquivo | papel nesta rodada |
|---|---|
| `db.py` | esquema novo; processos; emissões e situação do termo; diff de importação; `dimensoes`, `painel`, `recorte`, `exportar_recorte` |
| `app.py` | rotas novas: processos, registrar, termos emitidos, importações, recorte; menu; gate do termo |
| `painel.py` (novo) | cards de gráfico e descrição de filtros a partir de `db.dimensoes()` (usa `url_for`) |
| `graficos.py` (novo, copiado da skill) | opções ECharts: `rosca`, `barras_horizontais`, `colunas`, `linha`, `tabela_dados` |
| `static/dsgov/vendor/echarts/echarts.min.js`, `static/dsgov/js/echarts-dsgov.js` (copiados) | biblioteca e tema |
| `static/dsgov/css/dsgov.css` | classes `dsgov-grafico`, `dsgov-grafico-alto`, `dsgov-kpi` |
| `templates/_macros.html` | `select` aceita pares (valor, rótulo); macros `tag_termo` e `grafico` |
| `templates/cadastros.html` | aba Processos SEI |
| `templates/termo.html` | gate, situação, registro ao copiar |
| `templates/termos_emitidos.html`, `templates/termo_emitido.html` (novos) | histórico e detalhe |
| `templates/centro_custos.html`, `templates/termos_individuais.html` | tabela com situação do termo |
| `templates/upload.html`, `templates/importacao.html` (novo), `templates/bem.html` | importações e histórico do bem |
| `templates/index.html` | painel |
| `templates/recorte.html` (novo) | recorte |
| `tests/test_db.py`, `tests/test_app.py`, `tests/test_painel.py` (novo) | testes |

---

### Task 1: Esquema novo e cadastro de processos SEI

**Files:**
- Modify: `db.py` (`ESQUEMA`, após a tabela `textos`; funções novas ao final do bloco de cadastros, antes de `def exportar_cadastros`)
- Test: `tests/test_db.py`

**Interfaces:**
- Produces: `db.TIPOS_TERMO`, `db.ROTULO_TIPO`, `db._agora()`, `db.processos(conn)`, `db.processo_vigente(conn, tipo)`, `db.incluir_processo(conn, tipo, descricao, numero_sei, vigente=True) -> int`, `db.marcar_vigente(conn, id)`, `db.encerrar_processo(conn, id)`, `db.excluir_processo(conn, id)`.

- [ ] **Step 1: Testes que falham**

Acrescente ao final de `tests/test_db.py`:

```python
# ---------------------------------------------------------------- processos SEI
def test_esquema_cria_tabelas_novas(dados):
    nomes = {r[0] for r in dados.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    assert {"processos_sei", "termos_emitidos", "termos_emitidos_bens", "importacoes", "importacoes_mudancas"} <= nomes


def test_processos_um_vigente_por_tipo(dados):
    a = db.incluir_processo(dados, "ccusto", "Termos 2025", "1111")
    b = db.incluir_processo(dados, "ccusto", "Termos 2026", "2222")
    c = db.incluir_processo(dados, "individual", "Individuais 2026", "3333")
    assert db.processo_vigente(dados, "ccusto")["id"] == b
    assert db.processo_vigente(dados, "individual")["id"] == c
    assert db.processo_vigente(dados, "devolucao") is None
    db.marcar_vigente(dados, a)
    assert db.processo_vigente(dados, "ccusto")["id"] == a
    db.encerrar_processo(dados, a)
    assert db.processo_vigente(dados, "ccusto") is None
    assert [p["id"] for p in db.processos(dados)][0] == c          # vigentes primeiro
    with pytest.raises(sqlite3.IntegrityError):
        dados.execute("INSERT INTO processos_sei (tipo, descricao, numero_sei, vigente, criado_em) VALUES ('individual','x','9',1,'2026-01-01 00:00:00')")


def test_processos_validacao_e_exclusao(dados):
    with pytest.raises(db.ErroDeNegocio):
        db.incluir_processo(dados, "outro", "x", "1")
    with pytest.raises(db.ErroDeNegocio):
        db.incluir_processo(dados, "ccusto", "", "1")
    with pytest.raises(db.ErroDeNegocio):
        db.incluir_processo(dados, "ccusto", "x", "  ")
    i = db.incluir_processo(dados, "ccusto", "x", "1", vigente=False)
    assert db.processo_vigente(dados, "ccusto") is None
    db.excluir_processo(dados, i)
    assert db.processos(dados) == []
```

Garanta `import sqlite3` no topo de `tests/test_db.py` (adicione se não houver).

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest -q tests/test_db.py -k "processos or tabelas_novas"`
Expected: 3 falhas (`AttributeError: module 'db' has no attribute 'incluir_processo'`; tabelas ausentes).

- [ ] **Step 3: Esquema e funções**

Em `db.py`, dentro de `ESQUEMA`, após `CREATE TABLE IF NOT EXISTS textos (...);`:

```sql
CREATE TABLE IF NOT EXISTS processos_sei (
  id         INTEGER PRIMARY KEY,
  tipo       TEXT NOT NULL CHECK (tipo IN ('ccusto','individual','devolucao')),
  descricao  TEXT NOT NULL,
  numero_sei TEXT NOT NULL,
  vigente    INTEGER NOT NULL DEFAULT 0,
  criado_em  TEXT NOT NULL
);
CREATE UNIQUE INDEX IF NOT EXISTS processos_sei_vigente ON processos_sei(tipo) WHERE vigente = 1;
CREATE TABLE IF NOT EXISTS termos_emitidos (
  id            INTEGER PRIMARY KEY,
  tipo          TEXT NOT NULL,
  chave         TEXT NOT NULL,
  processo_id   INTEGER NOT NULL REFERENCES processos_sei(id),
  documento_sei TEXT,
  emitido_em    TEXT NOT NULL,
  quantidade    INTEGER NOT NULL,
  valor_total   REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS termos_emitidos_bens (
  termo_id    INTEGER NOT NULL REFERENCES termos_emitidos(id) ON DELETE CASCADE,
  numero      INTEGER NOT NULL,
  descricao   TEXT, complemento TEXT, localizacao TEXT, valor_atual REAL,
  PRIMARY KEY (termo_id, numero)
);
CREATE TABLE IF NOT EXISTS importacoes (
  id            INTEGER PRIMARY KEY,
  importado_em  TEXT NOT NULL,
  arquivo       TEXT,
  total INTEGER NOT NULL, ativos INTEGER NOT NULL,
  novos INTEGER NOT NULL, removidos INTEGER NOT NULL, movidos INTEGER NOT NULL, situacao INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS importacoes_mudancas (
  importacao_id INTEGER NOT NULL REFERENCES importacoes(id) ON DELETE CASCADE,
  numero        INTEGER NOT NULL,
  tipo          TEXT NOT NULL CHECK (tipo IN ('novo','removido','movido','situacao')),
  de            TEXT, para TEXT,
  descricao     TEXT
);
```

Logo após `class ErroDeNegocio`:

```python
TIPOS_TERMO = ("ccusto", "individual", "devolucao")
ROTULO_TIPO = {"ccusto": "termos por centro de custo", "individual": "termos individuais",
               "devolucao": "termos de devolução"}


def _agora() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
```

Antes de `def exportar_cadastros`:

```python
# ---------------------------------------------------------------- processos SEI
def processos(conn) -> list[dict]:
    return _todos(conn, "SELECT * FROM processos_sei ORDER BY vigente DESC, criado_em DESC, id DESC")


def processo_vigente(conn, tipo: str) -> dict | None:
    return _um(conn, "SELECT * FROM processos_sei WHERE tipo = ? AND vigente = 1", tipo)


def incluir_processo(conn, tipo: str, descricao: str, numero_sei: str, vigente: bool = True) -> int:
    if tipo not in TIPOS_TERMO:
        raise ErroDeNegocio("Tipo de termo inválido.")
    descricao = _obrigatorio(descricao, "Descrição")
    numero_sei = _obrigatorio(numero_sei, "Número SEI")
    if vigente:
        conn.execute("UPDATE processos_sei SET vigente = 0 WHERE tipo = ?", (tipo,))
    cur = conn.execute("INSERT INTO processos_sei (tipo, descricao, numero_sei, vigente, criado_em) VALUES (?,?,?,?,?)",
                       (tipo, descricao, numero_sei, 1 if vigente else 0, _agora()))
    conn.commit()
    return cur.lastrowid


def _processo(conn, id: int) -> dict:
    return _um(conn, "SELECT * FROM processos_sei WHERE id = ?", id) or _erro("Processo não encontrado.")


def _erro(msg):
    raise ErroDeNegocio(msg)


def marcar_vigente(conn, id: int) -> None:
    p = _processo(conn, id)
    conn.execute("UPDATE processos_sei SET vigente = 0 WHERE tipo = ?", (p["tipo"],))
    conn.execute("UPDATE processos_sei SET vigente = 1 WHERE id = ?", (id,))
    conn.commit()


def encerrar_processo(conn, id: int) -> None:
    _processo(conn, id)
    conn.execute("UPDATE processos_sei SET vigente = 0 WHERE id = ?", (id,))
    conn.commit()


def excluir_processo(conn, id: int) -> None:
    _processo(conn, id)
    if conn.execute("SELECT 1 FROM termos_emitidos WHERE processo_id = ? LIMIT 1", (id,)).fetchone():
        raise ErroDeNegocio("Este processo tem termos registrados; encerre-o em vez de excluir.")
    conn.execute("DELETE FROM processos_sei WHERE id = ?", (id,))
    conn.commit()
```

`_obrigatorio(valor, rotulo)` já existe em `db.py` (levanta `ErroDeNegocio` se vazio e devolve o texto sem espaços nas pontas).

- [ ] **Step 4: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam (95).

- [ ] **Step 5: Commit**

```bash
git add db.py tests/test_db.py
git commit -m "Processos SEI: esquema novo (processos, termos emitidos, importações) e cadastro com um vigente por tipo"
```

---

### Task 2: Registro de emissão, situação do termo e listas com situação

**Files:**
- Modify: `db.py` (após as funções da Task 1; `renomear_centro` e `renomear_pessoa`)
- Test: `tests/test_db.py`

**Interfaces:**
- Consumes: Task 1.
- Produces: `db.registrar_emissao(conn, tipo, chave, bens) -> dict`, `db.termos_emitidos(conn, tipo=None, chave=None, limite=200)`, `db.termo_emitido(conn, id) -> dict | None` (com chave `"bens"`), `db.salvar_documento_sei(conn, id, documento)`, `db.ultimo_termo(conn, tipo, chave)`, `db.situacao_termo(conn, tipo, chave, bens_atuais) -> {"estado","ultimo","entraram","sairam"}`, `db.situacoes_centros(conn)`, `db.situacoes_pessoas(conn)`.

- [ ] **Step 1: Testes que falham**

```python
# ---------------------------------------------------------------- termos emitidos
def test_registrar_emissao_exige_processo_e_grava_foto(dados):
    semear(dados)
    bens = db.bens_do_centro(dados, "CCI")
    with pytest.raises(db.ErroDeNegocio):
        db.registrar_emissao(dados, "ccusto", "CCI", bens)
    db.incluir_processo(dados, "ccusto", "Termos 2026", "2222")
    t = db.registrar_emissao(dados, "ccusto", "CCI", bens)
    assert t["quantidade"] == 1 and t["valor_total"] == 64.54 and t["numero_sei"] == "2222"
    assert [b["numero"] for b in t["bens"]] == [1001] and t["bens"][0]["descricao"] == "CADEIRA"
    assert db.ultimo_termo(dados, "ccusto", "CCI")["id"] == t["id"]
    assert db.termos_emitidos(dados)[0]["id"] == t["id"]
    assert db.termos_emitidos(dados, tipo="individual") == []
    assert db.termos_emitidos(dados, chave="cc")[0]["id"] == t["id"]
    db.salvar_documento_sei(dados, t["id"], " 0451234 ")
    assert db.termo_emitido(dados, t["id"])["documento_sei"] == "0451234"


def test_registrar_emissao_mesmo_dia_mesma_lista_nao_duplica(dados):
    semear(dados)
    db.incluir_processo(dados, "ccusto", "Termos 2026", "2222")
    a = db.registrar_emissao(dados, "ccusto", "CCI", db.bens_do_centro(dados, "CCI"))
    b = db.registrar_emissao(dados, "ccusto", "CCI", db.bens_do_centro(dados, "CCI"))
    assert a["id"] == b["id"] and len(db.termos_emitidos(dados)) == 1
    dados.execute("INSERT INTO bens VALUES (1005,'ATIVO','LUMINÁRIA','','MÓVEIS','01 - SALA CCI','01/01/2020',10,9)")
    c = db.registrar_emissao(dados, "ccusto", "CCI", db.bens_do_centro(dados, "CCI"))
    assert c["id"] != a["id"] and c["quantidade"] == 2


def test_situacao_termo(dados):
    semear(dados)
    bens = db.bens_do_centro(dados, "CCI")
    assert db.situacao_termo(dados, "ccusto", "CCI", bens)["estado"] == "sem_termo"
    db.incluir_processo(dados, "ccusto", "Termos 2026", "2222")
    db.registrar_emissao(dados, "ccusto", "CCI", bens)
    assert db.situacao_termo(dados, "ccusto", "CCI", db.bens_do_centro(dados, "CCI"))["estado"] == "vigente"
    dados.execute("INSERT INTO bens VALUES (1005,'ATIVO','LUMINÁRIA','','MÓVEIS','01 - SALA CCI','01/01/2020',10,9)")
    dados.execute("UPDATE bens SET situacao = 'BAIXADO' WHERE numero = 1001")
    s = db.situacao_termo(dados, "ccusto", "CCI", db.bens_do_centro(dados, "CCI"))
    assert (s["estado"], s["entraram"], s["sairam"]) == ("desatualizado", 1, 1)
    centros = db.situacoes_centros(dados)
    assert centros[0]["ccustos"] == "CCI" and centros[0]["estado"] == "desatualizado" and centros[0]["quantidade"] == 1
    pessoas = db.situacoes_pessoas(dados)
    assert pessoas == [{"nome": "ANA SILVA", "quantidade": 1, "valor": 1500.0, "estado": "sem_termo",
                        "ultimo": None, "entraram": 0, "sairam": 0}]


def test_renomear_leva_historico_junto(dados):
    semear(dados)
    db.incluir_processo(dados, "ccusto", "T", "1")
    db.incluir_processo(dados, "individual", "I", "2")
    db.registrar_emissao(dados, "ccusto", "CCI", db.bens_do_centro(dados, "CCI"))
    db.registrar_emissao(dados, "individual", "ANA SILVA", db.bens_da_pessoa(dados, "ANA SILVA"))
    db.renomear_centro(dados, "CCI", "GEX")
    db.renomear_pessoa(dados, "ANA SILVA", "ANA SOUZA")
    assert db.ultimo_termo(dados, "ccusto", "GEX") and db.ultimo_termo(dados, "ccusto", "CCI") is None
    assert db.ultimo_termo(dados, "individual", "ANA SOUZA") and db.ultimo_termo(dados, "individual", "ANA SILVA") is None
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest -q tests/test_db.py -k "emissao or situacao_termo or historico_junto"`
Expected: 4 falhas por atributo inexistente.

- [ ] **Step 3: Implementar**

Após `excluir_processo` em `db.py`:

```python
# ---------------------------------------------------------------- termos emitidos (foto)
def _numeros_do_termo(conn, termo_id: int) -> list[int]:
    return [r[0] for r in conn.execute("SELECT numero FROM termos_emitidos_bens WHERE termo_id = ? ORDER BY numero", (termo_id,))]


def ultimo_termo(conn, tipo: str, chave: str) -> dict | None:
    return _um(conn, """
        SELECT t.*, p.descricao AS processo, p.numero_sei FROM termos_emitidos t JOIN processos_sei p ON p.id = t.processo_id
        WHERE t.tipo = ? AND t.chave = ? ORDER BY t.emitido_em DESC, t.id DESC LIMIT 1""", tipo, chave)


def registrar_emissao(conn, tipo: str, chave: str, bens: list) -> dict:
    """Foto do termo. Sem processo vigente do tipo → ErroDeNegocio. No mesmo dia, com a mesma lista de
    bens, só atualiza a hora do registro existente."""
    proc = processo_vigente(conn, tipo)
    if not proc:
        raise ErroDeNegocio(f"Cadastre um processo SEI vigente para {ROTULO_TIPO[tipo]} em Cadastros → Processos SEI.")
    agora = _agora()
    numeros = sorted(int(b["numero"]) for b in bens)
    ultimo = ultimo_termo(conn, tipo, chave)
    if ultimo and ultimo["emitido_em"][:10] == agora[:10] and _numeros_do_termo(conn, ultimo["id"]) == numeros:
        conn.execute("UPDATE termos_emitidos SET emitido_em = ? WHERE id = ?", (agora, ultimo["id"]))
        conn.commit()
        return termo_emitido(conn, ultimo["id"])
    cur = conn.execute(
        "INSERT INTO termos_emitidos (tipo, chave, processo_id, emitido_em, quantidade, valor_total) VALUES (?,?,?,?,?,?)",
        (tipo, chave, proc["id"], agora, len(bens), sum(b["valor_atual"] or 0 for b in bens)))
    conn.executemany("INSERT INTO termos_emitidos_bens VALUES (?,?,?,?,?,?)", [
        (cur.lastrowid, b["numero"], b["descricao"], b["complemento"], b["localizacao"], b["valor_atual"]) for b in bens])
    conn.commit()
    return termo_emitido(conn, cur.lastrowid)


def termos_emitidos(conn, tipo: str | None = None, chave: str | None = None, limite: int = 200) -> list[dict]:
    sql = """SELECT t.*, p.descricao AS processo, p.numero_sei FROM termos_emitidos t
             JOIN processos_sei p ON p.id = t.processo_id WHERE 1"""
    params: list = []
    if tipo:
        sql += " AND t.tipo = ?"
        params.append(tipo)
    if chave:
        sql += " AND MAIUSC(t.chave) LIKE ?"
        params.append(f"%{chave.upper()}%")
    conn.create_function("MAIUSC", 1, lambda v: v.upper() if isinstance(v, str) else v)
    return _todos(conn, sql + " ORDER BY t.emitido_em DESC, t.id DESC LIMIT ?", *params, limite)


def termo_emitido(conn, id: int) -> dict | None:
    t = _um(conn, """
        SELECT t.*, p.descricao AS processo, p.numero_sei FROM termos_emitidos t
        JOIN processos_sei p ON p.id = t.processo_id WHERE t.id = ?""", id)
    if not t:
        return None
    t["bens"] = _todos(conn, "SELECT * FROM termos_emitidos_bens WHERE termo_id = ? ORDER BY numero", id)
    return t


def salvar_documento_sei(conn, id: int, documento: str) -> None:
    conn.execute("UPDATE termos_emitidos SET documento_sei = ? WHERE id = ?", (_texto(documento) or None, id))
    conn.commit()


def situacao_termo(conn, tipo: str, chave: str, bens_atuais: list) -> dict:
    """Compara a foto do último termo com os bens de hoje. Só entrada/saída conta."""
    ultimo = ultimo_termo(conn, tipo, chave)
    if not ultimo:
        return {"estado": "sem_termo", "ultimo": None, "entraram": 0, "sairam": 0}
    foto = set(_numeros_do_termo(conn, ultimo["id"]))
    atuais = {int(b["numero"]) for b in bens_atuais}
    entraram, sairam = len(atuais - foto), len(foto - atuais)
    return {"estado": "desatualizado" if entraram or sairam else "vigente",
            "ultimo": ultimo, "entraram": entraram, "sairam": sairam}


def situacoes_centros(conn) -> list[dict]:
    """Cada centro com quantidade/valor dos bens sob guarda e a situação do termo."""
    out = []
    for c in centros(conn):
        bens = bens_do_centro(conn, c["ccustos"])
        out.append({**c, "quantidade": len(bens), "valor": sum(b["valor_atual"] or 0 for b in bens),
                    **situacao_termo(conn, "ccusto", c["ccustos"], bens)})
    return out


def situacoes_pessoas(conn) -> list[dict]:
    out = []
    for nome in pessoas(conn):
        bens = bens_da_pessoa(conn, nome)
        out.append({"nome": nome, "quantidade": len(bens), "valor": sum(b["valor_atual"] or 0 for b in bens),
                    **situacao_termo(conn, "individual", nome, bens)})
    return out
```

`_texto` já existe (`db.py`, converte para str sem espaços; `None` → `""`).

Em `renomear_centro`, logo após o `UPDATE responsaveis`:

```python
    conn.execute("UPDATE termos_emitidos SET chave = ? WHERE tipo = 'ccusto' AND chave = ?", (novo, antigo))
```

Em `renomear_pessoa`, logo após o `UPDATE pessoas`:

```python
    conn.execute("UPDATE termos_emitidos SET chave = ? WHERE tipo IN ('individual','devolucao') AND chave = ?", (novo, antigo))
```

- [ ] **Step 4: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam (99).

- [ ] **Step 5: Commit**

```bash
git add db.py tests/test_db.py
git commit -m "Registro de emissão com foto dos bens; situação do termo (sem termo/vigente/desatualizado)"
```

---

### Task 3: Aba "Processos SEI" em Cadastros

**Files:**
- Modify: `app.py` (`ABAS`, `cadastros()`, rotas novas após `pessoas_desatribuir`)
- Modify: `templates/_macros.html` (macro `select` aceita pares)
- Modify: `templates/cadastros.html`
- Test: `tests/test_app.py`

**Interfaces:**
- Consumes: Task 1.
- Produces: rotas `cadastros(aba='processos')`, `processos_incluir`, `processos_vigente`, `processos_encerrar`, `processos_excluir`; macro `select(nome, rotulo, opcoes, selecionado=None, obrigatorio=True)` onde cada opção é `str` ou `(valor, rotulo)`.

- [ ] **Step 1: Teste que falha**

```python
def test_cadastro_de_processos_sei(cliente):
    r = cliente.get("/cadastros/processos")
    assert r.status_code == 200 and b"Processos SEI" in r.data
    r = cliente.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "Termos 2026", "numero_sei": "2222", "vigente": "1"}, follow_redirects=True)
    assert b"Termos 2026" in r.data and b"2222" in r.data and b"vigente" in r.data
    r = cliente.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "", "numero_sei": "1"}, follow_redirects=True)
    assert "Descrição".encode() in r.data
    import db
    pid = db.processos(db.conectar())[0]["id"]
    r = cliente.post("/cadastros/processos/encerrar", data={"id": pid}, follow_redirects=True)
    assert b"encerrado" in r.data
    r = cliente.post("/cadastros/processos/vigente", data={"id": pid}, follow_redirects=True)
    assert b"vigente" in r.data
    r = cliente.post("/cadastros/processos/excluir", data={"id": pid}, follow_redirects=True)
    assert "exclu".encode() in r.data and b"Termos 2026" not in r.data
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest -q tests/test_app.py -k processos_sei`
Expected: FAIL (404 em `/cadastros/processos`).

- [ ] **Step 3: Macro `select` com pares**

Em `templates/_macros.html`, dentro do `{% for opcao in opcoes %}` da macro `select`, troque as duas referências a `opcao` por valor/rótulo:

```jinja
    {% for opcao in opcoes %}
    {% set valor, rotulo = (opcao, opcao) if opcao is string else opcao %}
    <div class="br-item" tabindex="-1">
      <div class="br-radio">
        <input id="{{ nome }}_{{ loop.index }}" name="{{ nome }}" type="radio" value="{{ valor }}"{% if valor == selecionado %} checked="checked"{% endif %}{% if obrigatorio %} required{% endif %}/>
        <label for="{{ nome }}_{{ loop.index }}">{{ rotulo }}</label>
      </div>
    </div>
    {% endfor %}
```

Acrescente ao final de `_macros.html` a tag de situação do termo (usada nas Tasks 6, 10 e 12):

```jinja
{% macro tag_termo(s) %}
{# s = dict de db.situacao_termo #}
{% if s.estado == 'vigente' %}<span class="br-tag bg-success text-pure-0"><span>vigente</span></span>
{% elif s.estado == 'desatualizado' %}<span class="br-tag bg-warning"><span>desatualizado +{{ s.entraram }} −{{ s.sairam }}</span></span>
{% else %}<span class="br-tag bg-gray-20"><span>sem termo</span></span>{% endif %}
{% endmacro %}
```

- [ ] **Step 4: Rotas**

Em `app.py`: `ABAS = ("responsaveis", "localizacoes", "pessoas", "processos")`. Em `cadastros()`, acrescente ao `render_template`: `processos=db.processos(conn), tipos=list(db.ROTULO_TIPO.items())`.

Após `pessoas_desatribuir`:

```python
@app.route("/cadastros/processos/incluir", methods=["POST"])
def processos_incluir():
    db.incluir_processo(obter_conn(), request.form.get("tipo", ""), request.form.get("descricao", ""),
                        request.form.get("numero_sei", ""), vigente=bool(request.form.get("vigente")))
    flash("Processo incluído.", "success")
    return _volta("processos")


@app.route("/cadastros/processos/vigente", methods=["POST"])
def processos_vigente():
    db.marcar_vigente(obter_conn(), int(request.form["id"]))
    flash("Processo marcado como vigente.", "success")
    return _volta("processos")


@app.route("/cadastros/processos/encerrar", methods=["POST"])
def processos_encerrar():
    db.encerrar_processo(obter_conn(), int(request.form["id"]))
    flash("Processo encerrado; nenhum termo desse tipo será emitido até marcar outro como vigente.", "success")
    return _volta("processos")


@app.route("/cadastros/processos/excluir", methods=["POST"])
def processos_excluir():
    db.excluir_processo(obter_conn(), int(request.form["id"]))
    flash("Processo excluído.", "success")
    return _volta("processos")
```

- [ ] **Step 5: Template**

Em `templates/cadastros.html`, na lista de abas, acrescente `('processos', 'Processos SEI')` ao final da lista do `for`. Antes do `</div>` que fecha `tab-content`, o painel novo:

```jinja
    <div class="tab-panel{% if aba == 'processos' %} active{% endif %}" id="painel-processos" role="tabpanel" aria-labelledby="tab-processos">
      <p class="text-gray-70">Um processo vigente por tipo de termo. Sem processo vigente, o termo daquele tipo não pode ser copiado nem baixado.</p>
      <form method="post" action="{{ url_for('processos_incluir') }}" class="row mb-4">
        <div class="col-md-3 mb-3">{{ select('tipo', 'Tipo de termo', tipos) }}</div>
        <div class="col-md-4 mb-3"><div class="br-input"><label for="p-descricao">Descrição</label><input id="p-descricao" name="descricao" type="text" placeholder="Ex.: Termos de responsabilidade 2026" required/></div></div>
        <div class="col-md-3 mb-3"><div class="br-input"><label for="p-numero">Número do processo SEI</label><input id="p-numero" name="numero_sei" type="text" required/></div></div>
        <div class="col-md-2 mb-3 d-flex align-items-end">
          <div class="br-checkbox mr-2"><input id="p-vigente" name="vigente" type="checkbox" value="1" checked/><label for="p-vigente">Vigente</label></div>
          <button class="br-button primary" type="submit"><i class="fas fa-plus mr-1" aria-hidden="true"></i>Incluir</button>
        </div>
      </form>
      {{ cabecalho_tabela('Processos SEI', 'proc') }}
        <thead><tr><th scope="col">Tipo</th><th scope="col">Descrição</th><th scope="col">Número SEI</th><th scope="col">Situação</th><th scope="col">Cadastrado em</th><th scope="col" class="dsgov-acoes">Ações</th></tr></thead>
        <tbody>
        {% for p in processos %}
        <tr>
          <td>{{ dict(tipos)[p.tipo] }}</td><td>{{ p.descricao }}</td><td>{{ p.numero_sei }}</td>
          <td>{% if p.vigente %}<span class="br-tag bg-success text-pure-0"><span>vigente</span></span>{% else %}<span class="br-tag bg-gray-20"><span>encerrado</span></span>{% endif %}</td>
          <td>{{ p.criado_em[:10] }}</td>
          <td class="dsgov-acoes">
            {% if p.vigente %}
            <form method="post" action="{{ url_for('processos_encerrar') }}" class="d-inline"><input type="hidden" name="id" value="{{ p.id }}"/><button class="br-button secondary small" type="submit">Encerrar</button></form>
            {% else %}
            <form method="post" action="{{ url_for('processos_vigente') }}" class="d-inline"><input type="hidden" name="id" value="{{ p.id }}"/><button class="br-button secondary small" type="submit">Marcar vigente</button></form>
            <form method="post" action="{{ url_for('processos_excluir') }}" class="d-inline"><input type="hidden" name="id" value="{{ p.id }}"/><button class="br-button circle small" type="submit" aria-label="Excluir processo {{ p.numero_sei }}"><i class="fas fa-trash" aria-hidden="true"></i></button></form>
            {% endif %}
          </td>
        </tr>
        {% endfor %}
        </tbody>
      </table></div>
    </div>
```

- [ ] **Step 6: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam (100).

- [ ] **Step 7: Commit**

```bash
git add app.py templates/_macros.html templates/cadastros.html tests/test_app.py
git commit -m "Cadastros: aba Processos SEI (incluir, marcar vigente, encerrar, excluir)"
```

---

### Task 4: Tela do termo — gate, registro ao copiar/baixar, linha de situação

**Files:**
- Modify: `app.py` (`termo`, `termo_docx`; rota nova `termo_registrar` antes de `termo_planilha`)
- Modify: `templates/termo.html`
- Test: `tests/test_app.py`

**Interfaces:**
- Consumes: Task 2 (`registrar_emissao`, `situacao_termo`, `ultimo_termo`, `processo_vigente`, `ROTULO_TIPO`).
- Produces: `POST /termo/<tipo>/<chave>/registrar` → JSON `{"id", "emitido_em"}` (200) ou `{"erro"}` (409).

- [ ] **Step 1: Testes que falham**

```python
def test_termo_sem_processo_vigente_nao_emite(cliente):
    r = cliente.get("/termo/ccusto/CCI")
    assert b"Cadastre um processo SEI vigente" in r.data and b'id="copiar"' not in r.data
    r = cliente.get("/termo/ccusto/CCI/docx", follow_redirects=True)
    assert b"Cadastre um processo SEI vigente" in r.data
    assert cliente.get("/termo/ccusto/CCI/documento").status_code == 200    # prévia continua
    r = cliente.post("/termo/ccusto/CCI/registrar")
    assert r.status_code == 409 and "erro" in r.get_json()


def test_termo_com_processo_registra_ao_baixar_e_ao_copiar(cliente):
    cliente.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "T", "numero_sei": "2222", "vigente": "1"})
    r = cliente.get("/termo/ccusto/CCI")
    assert b'id="copiar"' in r.data and b"Nenhum termo registrado" in r.data
    assert cliente.get("/termo/ccusto/CCI/docx").status_code == 200
    r = cliente.get("/termo/ccusto/CCI")
    assert "Último termo registrado".encode() in r.data and b"Bens iguais aos de hoje" in r.data
    j = cliente.post("/termo/ccusto/CCI/registrar").get_json()
    assert j["id"] and j["emitido_em"][:4] == "2026"
    assert cliente.get("/termo/ccusto/CCI/planilha").status_code == 200
    import db
    assert len(db.termos_emitidos(db.conectar())) == 1     # docx + registrar no mesmo dia = 1 registro; planilha não registra


def test_termo_devolucao_registra_no_processo_de_devolucao(cliente):
    cliente.post("/cadastros/processos/incluir", data={"tipo": "devolucao", "descricao": "D", "numero_sei": "3333", "vigente": "1"})
    cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "numero_bem": "1001"})
    assert cliente.get("/termo/devolucao/ANA SILVA/docx").status_code == 200
    import db
    t = db.termos_emitidos(db.conectar(), tipo="devolucao")[0]
    assert t["chave"] == "ANA SILVA" and t["quantidade"] == 1
```

Atenção ao teste existente `test_termo_ccusto_documento_docx_planilha` e `test_termo_individual` e `test_termo_devolucao_fluxo`: eles chamam `/docx` sem processo. Acrescente no início de cada um a inclusão do processo do tipo certo via `cliente.post("/cadastros/processos/incluir", data={...})` (tipos `ccusto`, `individual`, `devolucao`).

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest -q tests/test_app.py`
Expected: falhas nos 3 novos e nos 3 antigos ajustados (redirect em vez de 200 no docx antes do ajuste; 404 em `/registrar`).

- [ ] **Step 3: Rotas**

Substitua `termo()` e o começo de `termo_docx()` em `app.py`:

```python
def _exigir_processo(conn, tipo, chave):
    """Sem processo SEI vigente do tipo não há emissão: flash + volta à tela do termo."""
    if db.processo_vigente(conn, tipo):
        return None
    flash(f"Cadastre um processo SEI vigente para {db.ROTULO_TIPO[tipo]} em Cadastros → Processos SEI.", "error")
    return redirect(url_for("termo", tipo=tipo, chave=chave))


@app.route("/termo/<tipo>/<chave>")
def termo(tipo, chave):
    conn = obter_conn()
    titulo, _, bens, _ = _bens_do_termo(conn, tipo, chave)
    if tipo == "devolucao":
        situacao = {"estado": None, "ultimo": db.ultimo_termo(conn, tipo, chave), "entraram": 0, "sairam": 0}
    else:
        situacao = db.situacao_termo(conn, tipo, chave, bens)
    return render_template("termo.html", tipo=tipo, chave=chave, titulo=titulo, quantidade=len(bens),
                           processo=db.processo_vigente(conn, tipo), situacao=situacao,
                           rotulo_tipo=db.ROTULO_TIPO[tipo], trilha=[(titulo, None)])
```

Em `termo_docx`, logo após `conn = obter_conn()`:

```python
    if (volta := _exigir_processo(conn, tipo, chave)):
        return volta
```

e antes do `return _baixar(arquivo, nome)` final:

```python
    db.registrar_emissao(conn, tipo, chave, bens)
```

Nova rota, antes de `termo_planilha`:

```python
@app.route("/termo/<tipo>/<chave>/registrar", methods=["POST"])
def termo_registrar(tipo, chave):
    """Chamado pelo botão Copiar depois da cópia dar certo. Responde JSON."""
    conn = obter_conn()
    _, _, bens, _ = _bens_do_termo(conn, tipo, chave)
    if not db.processo_vigente(conn, tipo):
        return {"erro": f"Cadastre um processo SEI vigente para {db.ROTULO_TIPO[tipo]}."}, 409
    t = db.registrar_emissao(conn, tipo, chave, bens)
    return {"id": t["id"], "emitido_em": t["emitido_em"]}
```

- [ ] **Step 4: Template**

Substitua `templates/termo.html` inteiro:

```jinja
{% extends "base.html" %}
{% block titulo %}{{ titulo }}{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4">
  <h1 class="mb-0">{{ titulo }}</h1>
  {% if processo %}
  <div class="ml-auto">
    <button class="br-button primary" type="button" id="copiar" data-registrar="{{ url_for('termo_registrar', tipo=tipo, chave=chave) }}"><i class="fas fa-copy mr-1" aria-hidden="true"></i>Copiar para o SEI</button>
    <a class="br-button secondary ml-2" href="{{ url_for('termo_docx', tipo=tipo, chave=chave) }}"><i class="fas fa-download mr-1" aria-hidden="true"></i>Baixar .docx</a>
    {% if tipo == 'ccusto' %}<a class="br-button ml-2" href="{{ url_for('termo_planilha', chave=chave) }}">Baixar planilha</a>{% endif %}
  </div>
  {% endif %}
</div>
{% if not processo %}
<div class="br-message danger"><div class="icon"><i class="fas fa-times-circle fa-lg" aria-hidden="true"></i></div>
  <div class="content" role="alert"><span class="message-title">Sem processo SEI.</span><span class="message-body"> Cadastre um processo SEI vigente para {{ rotulo_tipo }} em <a href="{{ url_for('cadastros', aba='processos') }}">Cadastros → Processos SEI</a> para copiar ou baixar este termo.</span></div></div>
{% else %}
<p class="text-gray-70 mb-3">Processo SEI {{ processo.numero_sei }} ({{ processo.descricao }}).
  {% if situacao.ultimo %}Último termo registrado em {{ situacao.ultimo.emitido_em[8:10] }}/{{ situacao.ultimo.emitido_em[5:7] }}/{{ situacao.ultimo.emitido_em[:4] }} {{ situacao.ultimo.emitido_em[11:16] }}{% if situacao.ultimo.documento_sei %} (SEI {{ situacao.ultimo.documento_sei }}){% endif %}
    (<a href="{{ url_for('termo_emitido_tela', id=situacao.ultimo.id) }}">ver registro</a>).
    {% if situacao.estado == 'vigente' %}Bens iguais aos de hoje.{% elif situacao.estado == 'desatualizado' %}Desde então entraram {{ situacao.entraram }} e saíram {{ situacao.sairam }} bem(ns): emita de novo.{% endif %}
  {% else %}Nenhum termo registrado para {{ chave }}.{% endif %}</p>
{% endif %}
<div id="aviso-copiado" class="br-message success" hidden>
  <div class="icon"><i class="fas fa-check-circle fa-lg" aria-hidden="true"></i></div>
  <div class="content" role="alert"><span class="message-title">Copiado.</span><span class="message-body" id="aviso-texto"> Cole no editor do SEI (Ctrl+V).</span></div>
</div>
{% if quantidade == 0 %}
<div class="br-message warning"><div class="icon"><i class="fas fa-exclamation-triangle fa-lg" aria-hidden="true"></i></div>
  <div class="content" role="alert"><span class="message-title">Atenção.</span><span class="message-body"> Nenhum bem neste termo.</span></div></div>
{% endif %}
<div class="br-card"><div class="card-content">
  <iframe id="documento" title="{{ titulo }}" src="{{ url_for('termo_documento', tipo=tipo, chave=chave) }}" width="100%" height="900"></iframe>
</div></div>
{% endblock %}
{% block scripts %}
{% if processo %}
<script>
document.getElementById("copiar").addEventListener("click", async function () {
  var corpo = document.getElementById("documento").contentDocument.body;
  var html = corpo.innerHTML, texto = corpo.innerText;
  try {
    await navigator.clipboard.write([new ClipboardItem({
      "text/html": new Blob([html], {type: "text/html"}),
      "text/plain": new Blob([texto], {type: "text/plain"})})]);
  } catch (e) {
    var docIframe = corpo.ownerDocument, sel = docIframe.defaultView.getSelection(), range = docIframe.createRange();
    range.selectNodeContents(corpo); sel.removeAllRanges(); sel.addRange(range);
    docIframe.execCommand("copy"); sel.removeAllRanges();
  }
  var aviso = document.getElementById("aviso-copiado"), textoAviso = document.getElementById("aviso-texto");
  textoAviso.textContent = " Cole no editor do SEI (Ctrl+V).";
  try {
    var r = await fetch(this.dataset.registrar, {method: "POST"});
    var j = await r.json();
    textoAviso.textContent = r.ok ? " Cole no editor do SEI (Ctrl+V). Emissão registrada às " + j.emitido_em.slice(11, 16) + "." : " " + j.erro;
  } catch (e) { textoAviso.textContent += " (não foi possível registrar a emissão)"; }
  aviso.hidden = false;
  setTimeout(function () { aviso.hidden = true; }, 5000);
});
</script>
{% endif %}
{% endblock %}
```

`termo_emitido_tela` é a rota de detalhe criada na Task 5; até lá, para os testes desta task passarem, crie a rota mínima já nesta task (a Task 5 a completa):

```python
@app.route("/termos-emitidos/<int:id>")
def termo_emitido_tela(id):
    t = db.termo_emitido(obter_conn(), id) or abort(404)
    return render_template("termo_emitido.html", t=t, rotulos=db.ROTULO_TIPO, trilha=[("Termos emitidos", url_for("termos_emitidos_tela")), (f"Registro {id}", None)])


@app.route("/termos-emitidos")
def termos_emitidos_tela():
    return redirect(url_for("home"))   # completada na Task 5
```

e um `templates/termo_emitido.html` mínimo (`{% extends "base.html" %}{% block conteudo %}<h1>Registro {{ t.id }}</h1>{% endblock %}`), substituído na Task 5.

- [ ] **Step 5: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam (103).

- [ ] **Step 6: Commit**

```bash
git add app.py templates/termo.html templates/termo_emitido.html tests/test_app.py
git commit -m "Termo: exige processo SEI vigente; copiar e baixar registram a emissão; linha de situação"
```

---

### Task 5: Tela "Termos emitidos" (lista e detalhe) e menu

**Files:**
- Modify: `app.py` (`contexto_dsgov` MENU; substituir `termos_emitidos_tela`; rota `termo_emitido_documento`)
- Create: `templates/termos_emitidos.html`; substituir `templates/termo_emitido.html`
- Test: `tests/test_app.py`

**Interfaces:**
- Consumes: Task 2 (`termos_emitidos`, `termo_emitido`, `salvar_documento_sei`).
- Produces: `GET /termos-emitidos?tipo=&chave=`, `GET /termos-emitidos/<id>`, `POST /termos-emitidos/<id>/documento`.

- [ ] **Step 1: Teste que falha**

```python
def test_termos_emitidos_lista_detalhe_e_documento_sei(cliente):
    cliente.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "T", "numero_sei": "2222", "vigente": "1"})
    cliente.get("/termo/ccusto/CCI/docx")
    r = cliente.get("/termos-emitidos")
    assert r.status_code == 200 and b"CCI" in r.data and b"2222" in r.data
    assert b"CCI" not in cliente.get("/termos-emitidos?tipo=individual").data.split(b"<tbody>")[1]
    assert b"CCI" in cliente.get("/termos-emitidos?chave=cc").data
    import db
    tid = db.termos_emitidos(db.conectar())[0]["id"]
    r = cliente.get(f"/termos-emitidos/{tid}")
    assert b"CADEIRA" in r.data and b"1001" in r.data and b'name="documento_sei"' in r.data
    r = cliente.post(f"/termos-emitidos/{tid}/documento", data={"documento_sei": "0459999"}, follow_redirects=True)
    assert b"0459999" in r.data
    assert cliente.get("/termos-emitidos/999").status_code == 404
    assert b"Termos emitidos" in cliente.get("/").data    # menu
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest -q tests/test_app.py -k termos_emitidos_lista`
Expected: FAIL (302 na lista).

- [ ] **Step 3: Rotas e menu**

No `MENU` de `contexto_dsgov`, após `("Termo de devolução", ...)`:

```python
        ("Termos emitidos", "fa-history", url_for("termos_emitidos_tela")),
```

Substitua a rota provisória `termos_emitidos_tela` e acrescente a de documento:

```python
@app.route("/termos-emitidos")
def termos_emitidos_tela():
    tipo, chave = request.args.get("tipo") or None, request.args.get("chave", "").strip() or None
    return render_template("termos_emitidos.html", termos=db.termos_emitidos(obter_conn(), tipo, chave),
                           tipo=tipo, chave=chave, rotulos=db.ROTULO_TIPO, trilha=[("Termos emitidos", None)])


@app.route("/termos-emitidos/<int:id>/documento", methods=["POST"])
def termo_emitido_documento(id):
    conn = obter_conn()
    db.termo_emitido(conn, id) or abort(404)
    db.salvar_documento_sei(conn, id, request.form.get("documento_sei", ""))
    flash("Documento SEI salvo.", "success")
    return redirect(url_for("termo_emitido_tela", id=id))
```

- [ ] **Step 4: Templates**

`templates/termos_emitidos.html`:

```jinja
{% extends "base.html" %}
{% from "_macros.html" import select, cabecalho_tabela %}
{% block titulo %}Termos emitidos{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Termos emitidos</h1></div>
<form method="get" class="row mb-4">
  <div class="col-md-3 mb-3">{{ select('tipo', 'Tipo', [('', 'Todos')] + rotulos.items()|list, selecionado=tipo or '', obrigatorio=False) }}</div>
  <div class="col-md-4 mb-3"><div class="br-input"><label for="chave">Centro de custo ou pessoa</label><input id="chave" name="chave" type="text" value="{{ chave or '' }}" placeholder="contém..."/></div></div>
  <div class="col-md-2 mb-3 d-flex align-items-end"><button class="br-button secondary" type="submit"><i class="fas fa-filter mr-1" aria-hidden="true"></i>Filtrar</button></div>
</form>
{{ cabecalho_tabela('Registros de emissão', 'emitidos') }}
  <thead><tr><th scope="col">Emitido em</th><th scope="col">Tipo</th><th scope="col">Centro / pessoa</th><th scope="col" class="dsgov-numero">Bens</th><th scope="col" class="dsgov-numero">Valor</th><th scope="col">Processo SEI</th><th scope="col">Documento SEI</th></tr></thead>
  <tbody>
  {% for t in termos %}
  <tr>
    <td><a href="{{ url_for('termo_emitido_tela', id=t.id) }}">{{ t.emitido_em[8:10] }}/{{ t.emitido_em[5:7] }}/{{ t.emitido_em[:4] }} {{ t.emitido_em[11:16] }}</a></td>
    <td>{{ rotulos[t.tipo] }}</td><td>{{ t.chave }}</td>
    <td class="dsgov-numero">{{ t.quantidade }}</td><td class="dsgov-numero">R$ {{ '%.2f'|format(t.valor_total) }}</td>
    <td>{{ t.numero_sei }} · {{ t.processo }}</td><td>{{ t.documento_sei or '—' }}</td>
  </tr>
  {% else %}
  <tr><td colspan="7">Nenhum termo registrado{% if tipo or chave %} com esse filtro{% endif %}.</td></tr>
  {% endfor %}
  </tbody>
</table></div>
{% endblock %}
```

`templates/termo_emitido.html` (substitui o provisório):

```jinja
{% extends "base.html" %}
{% from "_macros.html" import cabecalho_tabela %}
{% block titulo %}Registro {{ t.id }}{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4">
  <h1 class="mb-0">{{ rotulos[t.tipo]|capitalize }}: {{ t.chave }}</h1>
  <div class="ml-auto"><a class="br-button secondary" href="{{ url_for('termo', tipo=t.tipo, chave=t.chave) }}"><i class="fas fa-file-alt mr-1" aria-hidden="true"></i>Abrir termo atual</a></div>
</div>
<div class="br-card mb-4"><div class="card-content">
  <dl class="dsgov-detalhe row">
    <div class="col-md-3"><dt>Emitido em</dt><dd>{{ t.emitido_em[8:10] }}/{{ t.emitido_em[5:7] }}/{{ t.emitido_em[:4] }} {{ t.emitido_em[11:16] }}</dd></div>
    <div class="col-md-3"><dt>Processo SEI</dt><dd>{{ t.numero_sei }}<br/><span class="text-gray-70">{{ t.processo }}</span></dd></div>
    <div class="col-md-2"><dt>Bens</dt><dd>{{ t.quantidade }}</dd></div>
    <div class="col-md-2"><dt>Valor total</dt><dd>R$ {{ '%.2f'|format(t.valor_total) }}</dd></div>
  </dl>
  <form method="post" action="{{ url_for('termo_emitido_documento', id=t.id) }}" class="d-flex align-items-end">
    <div class="br-input mr-2"><label for="documento_sei">Número do documento no SEI</label><input id="documento_sei" name="documento_sei" type="text" value="{{ t.documento_sei or '' }}"/></div>
    <button class="br-button secondary" type="submit"><i class="fas fa-save mr-1" aria-hidden="true"></i>Salvar</button>
  </form>
</div></div>
{{ cabecalho_tabela('Bens no momento da emissão', 'foto') }}
  <thead><tr><th scope="col">Número</th><th scope="col">Descrição</th><th scope="col">Complemento</th><th scope="col">Localização</th><th scope="col" class="dsgov-numero">Valor</th></tr></thead>
  <tbody>
  {% for b in t.bens %}
  <tr><td><a href="{{ url_for('bem', numero=b.numero) }}">{{ b.numero }}</a></td><td>{{ b.descricao }}</td><td>{{ b.complemento or '' }}</td><td>{{ b.localizacao or '' }}</td><td class="dsgov-numero">R$ {{ '%.2f'|format(b.valor_atual or 0) }}</td></tr>
  {% endfor %}
  </tbody>
</table></div>
{% endblock %}
```

- [ ] **Step 5: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam (104).

- [ ] **Step 6: Commit**

```bash
git add app.py templates/termos_emitidos.html templates/termo_emitido.html tests/test_app.py
git commit -m "Tela Termos emitidos: histórico com filtros, detalhe com a foto e documento SEI"
```

---

### Task 6: Listas de centros e de pessoas com situação do termo

**Files:**
- Modify: `app.py` (`centro_custos`, `termos_individuais`)
- Modify: `templates/centro_custos.html`, `templates/termos_individuais.html`
- Test: `tests/test_app.py`

**Interfaces:**
- Consumes: Task 2 (`situacoes_centros`, `situacoes_pessoas`), macro `tag_termo` (Task 3).

- [ ] **Step 1: Teste que falha**

```python
def test_listas_mostram_situacao_do_termo(cliente):
    r = cliente.get("/centro-custos")
    assert b"sem termo" in r.data and b"/termo/ccusto/CCI" in r.data
    r = cliente.get("/termos-individuais")
    assert b"ANA SILVA" in r.data and b"sem termo" in r.data and b"/termo/individual/ANA" in r.data
    cliente.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "T", "numero_sei": "2", "vigente": "1"})
    cliente.get("/termo/ccusto/CCI/docx")
    assert b"vigente" in cliente.get("/centro-custos").data
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest -q tests/test_app.py -k listas_mostram`
Expected: FAIL.

- [ ] **Step 3: Rotas**

```python
@app.route("/centro-custos")
def centro_custos():
    conn = obter_conn()
    return render_template("centro_custos.html", centros=db.situacoes_centros(conn), trilha=[("Termo por centro de custo", None)])


@app.route("/termos-individuais")
def termos_individuais():
    conn = obter_conn()
    return render_template("termos_individuais.html", nomes=db.pessoas(conn), pessoas=db.situacoes_pessoas(conn),
                           trilha=[("Termo individual", None)])
```

- [ ] **Step 4: Templates**

Em `templates/centro_custos.html`, troque o import por `{% from "_macros.html" import select, cabecalho_tabela, tag_termo %}` e acrescente após o `</form>`:

```jinja
<div class="mt-5">
{{ cabecalho_tabela('Situação dos termos por centro de custo', 'centros') }}
  <thead><tr><th scope="col">Centro</th><th scope="col">Responsável</th><th scope="col" class="dsgov-numero">Bens</th><th scope="col" class="dsgov-numero">Valor</th><th scope="col">Último termo</th><th scope="col">Situação</th><th scope="col" class="dsgov-acoes">Termo</th></tr></thead>
  <tbody>
  {% for c in centros %}
  <tr>
    <td>{{ c.ccustos }}</td><td>{{ c.responsavel }}</td>
    <td class="dsgov-numero">{{ c.quantidade }}</td><td class="dsgov-numero">R$ {{ '%.2f'|format(c.valor) }}</td>
    <td>{% if c.ultimo %}{{ c.ultimo.emitido_em[8:10] }}/{{ c.ultimo.emitido_em[5:7] }}/{{ c.ultimo.emitido_em[:4] }}{% else %}—{% endif %}</td>
    <td>{{ tag_termo(c) }}</td>
    <td class="dsgov-acoes"><a class="br-button circle small" href="{{ url_for('termo', tipo='ccusto', chave=c.ccustos) }}" aria-label="Termo de {{ c.ccustos }}"><i class="fas fa-file-alt" aria-hidden="true"></i></a></td>
  </tr>
  {% endfor %}
  </tbody>
</table></div>
</div>
```

Em `templates/termos_individuais.html`, mesmo import e, após o `</form>`:

```jinja
<div class="mt-5">
{{ cabecalho_tabela('Situação dos termos individuais', 'pessoas') }}
  <thead><tr><th scope="col">Pessoa</th><th scope="col" class="dsgov-numero">Bens</th><th scope="col" class="dsgov-numero">Valor</th><th scope="col">Último termo</th><th scope="col">Situação</th><th scope="col" class="dsgov-acoes">Termo</th></tr></thead>
  <tbody>
  {% for p in pessoas %}
  <tr>
    <td>{{ p.nome }}</td><td class="dsgov-numero">{{ p.quantidade }}</td><td class="dsgov-numero">R$ {{ '%.2f'|format(p.valor) }}</td>
    <td>{% if p.ultimo %}{{ p.ultimo.emitido_em[8:10] }}/{{ p.ultimo.emitido_em[5:7] }}/{{ p.ultimo.emitido_em[:4] }}{% else %}—{% endif %}</td>
    <td>{{ tag_termo(p) }}</td>
    <td class="dsgov-acoes"><a class="br-button circle small" href="{{ url_for('termo', tipo='individual', chave=p.nome) }}" aria-label="Termo de {{ p.nome }}"><i class="fas fa-file-alt" aria-hidden="true"></i></a></td>
  </tr>
  {% endfor %}
  </tbody>
</table></div>
</div>
```

- [ ] **Step 5: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam (105).

- [ ] **Step 6: Commit**

```bash
git add app.py templates/centro_custos.html templates/termos_individuais.html tests/test_app.py
git commit -m "Listas de centros e pessoas com situação do termo e link direto"
```

---

### Task 7: Diff de importação e histórico do bem (`db.py`)

**Files:**
- Modify: `db.py` (`importar_bens`; funções novas após `localizacoes_sem_centro`)
- Test: `tests/test_db.py`

**Interfaces:**
- Consumes: esquema da Task 1.
- Produces: `db.importar_bens(conn, arquivo, nome_arquivo=None)` devolve também `novos, removidos, movidos, situacao, importacao_id`; `db.importacoes(conn, limite=20)`, `db.importacao(conn, id) -> dict | None` (com `"mudancas"`), `db.historico_do_bem(conn, numero) -> {"mudancas": [...], "termos": [...]}`.

- [ ] **Step 1: Testes que falham**

```python
# ---------------------------------------------------------------- importações
def test_importar_registra_mudancas(dados, tmp_path):
    semear(dados)
    arq = xlsx(tmp_path, [
        [1001, "ATIVO", "CADEIRA", "GIRATÓRIA", "MÓVEIS", "02 - OUTRA SALA", "31/12/1996", 75.94, 64.54],   # movido
        [1002, "ATIVO", "NOTEBOOK", "DELL", "EQUIPAMENTOS", "01 - SALA CCI", "06/12/2012", 3000, 1500],       # igual
        [1003, "ATIVO", "MESA", "ANTIGA", "MÓVEIS", "03 - DEPÓSITO", "06/12/2012", 100, 10],                  # situação + movido
        [5000, "ATIVO", "LUMINÁRIA", "", "MÓVEIS", "01 - SALA CCI", "01/01/2020", 10, 9],                     # novo; 1004 some
    ])
    r = db.importar_bens(dados, arq, nome_arquivo="export.xlsx")
    assert (r["novos"], r["removidos"], r["movidos"], r["situacao"]) == (1, 1, 2, 1)
    imp = db.importacao(dados, r["importacao_id"])
    assert imp["arquivo"] == "export.xlsx" and imp["total"] == 4 and imp["novos"] == 1
    tipos = sorted((m["numero"], m["tipo"], m["de"], m["para"]) for m in imp["mudancas"])
    assert tipos == [(1001, "movido", "01 - SALA CCI", "02 - OUTRA SALA"), (1003, "movido", "01 - SALA CCI", "03 - DEPÓSITO"),
                     (1003, "situacao", "BAIXADO", "ATIVO"), (1004, "removido", "99 - SEM MAPA", None), (5000, "novo", None, "01 - SALA CCI")]
    assert [i["id"] for i in db.importacoes(dados)] == [r["importacao_id"]]
    h = db.historico_do_bem(dados, 1003)
    assert [m["tipo"] for m in h["mudancas"]] == ["movido", "situacao"] and h["mudancas"][0]["importado_em"] == imp["importado_em"]
    assert h["termos"] == []


def test_importar_com_falha_nao_registra_importacao(dados, tmp_path):
    semear(dados)
    arq = xlsx(tmp_path, [[1001, "ATIVO", "CADEIRA", "", "MÓVEIS", "01 - SALA CCI", "x", 1, 1]])   # some o 1002 (atribuído)
    with pytest.raises(db.ImportacaoInvalida):
        db.importar_bens(dados, arq)
    assert db.importacoes(dados) == []


def test_historico_do_bem_lista_termos(dados):
    semear(dados)
    db.incluir_processo(dados, "ccusto", "T", "1")
    db.registrar_emissao(dados, "ccusto", "CCI", db.bens_do_centro(dados, "CCI"))
    h = db.historico_do_bem(dados, 1001)
    assert len(h["termos"]) == 1 and h["termos"][0]["chave"] == "CCI" and h["termos"][0]["tipo"] == "ccusto"
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest -q tests/test_db.py -k "registra_mudancas or nao_registra_importacao or historico_do_bem"`
Expected: 3 falhas.

- [ ] **Step 3: Implementar**

Assinatura: `def importar_bens(conn: sqlite3.Connection, arquivo, nome_arquivo: str | None = None) -> dict:`. Dentro, antes do bloco `try:` que faz o `DELETE FROM bens`:

```python
    antes = {r["numero"]: (r["situacao"], r["localizacao"], r["descricao"])
             for r in conn.execute("SELECT numero, situacao, localizacao, descricao FROM bens")}
```

Dentro desse `try`, depois do `if orfaos: raise ...` e **antes** de `conn.commit()`:

```python
        mudancas = _mudancas(antes, linhas)
        contagem = {t: sum(1 for m in mudancas if m[1] == t) for t in ("novo", "removido", "movido", "situacao")}
        cur = conn.execute(
            "INSERT INTO importacoes (importado_em, arquivo, total, ativos, novos, removidos, movidos, situacao) VALUES (?,?,?,?,?,?,?,?)",
            (_agora(), nome_arquivo, len(linhas), sum(1 for l in linhas if l[1] == "ATIVO"),
             contagem["novo"], contagem["removido"], contagem["movido"], contagem["situacao"]))
        conn.executemany("INSERT INTO importacoes_mudancas VALUES (?,?,?,?,?,?)",
                         [(cur.lastrowid, n, t, de, para, desc) for n, t, de, para, desc in mudancas])
        importacao_id = cur.lastrowid
```

E no `return` final: `return {"total": total, "ativos": ativos, "sem_centro": localizacoes_sem_centro(conn), "importacao_id": importacao_id, "novos": contagem["novo"], "removidos": contagem["removido"], "movidos": contagem["movido"], "situacao": contagem["situacao"]}`.

Funções novas após `localizacoes_sem_centro`:

```python
def _mudancas(antes: dict, linhas: list[tuple]) -> list[tuple]:
    """[(numero, tipo, de, para, descricao)] comparando a tabela anterior com o export.
    linhas = tuplas na ordem de INSERT INTO bens (numero, situacao, descricao, complemento, classificacao, localizacao, ...)."""
    depois = {l[0]: (l[1], l[5], l[2]) for l in linhas}
    m = []
    for n, (sit, loc, desc) in depois.items():
        if n not in antes:
            m.append((n, "novo", None, loc, desc))
            continue
        sit0, loc0, _ = antes[n]
        if sit != sit0:
            m.append((n, "situacao", sit0, sit, desc))
        if loc != loc0:
            m.append((n, "movido", loc0, loc, desc))
    for n, (sit0, loc0, desc0) in antes.items():
        if n not in depois:
            m.append((n, "removido", loc0, None, desc0))
    return sorted(m, key=lambda x: (x[0], x[1]))


def importacoes(conn, limite: int = 20) -> list[dict]:
    return _todos(conn, "SELECT * FROM importacoes ORDER BY importado_em DESC, id DESC LIMIT ?", limite)


def importacao(conn, id: int) -> dict | None:
    i = _um(conn, "SELECT * FROM importacoes WHERE id = ?", id)
    if i:
        i["mudancas"] = _todos(conn, "SELECT * FROM importacoes_mudancas WHERE importacao_id = ? ORDER BY tipo, numero", id)
    return i


def historico_do_bem(conn, numero: int) -> dict:
    return {
        "mudancas": _todos(conn, """
            SELECT m.*, i.importado_em FROM importacoes_mudancas m JOIN importacoes i ON i.id = m.importacao_id
            WHERE m.numero = ? ORDER BY i.importado_em DESC, m.tipo""", numero),
        "termos": _todos(conn, """
            SELECT t.id, t.tipo, t.chave, t.emitido_em, t.documento_sei FROM termos_emitidos_bens b
            JOIN termos_emitidos t ON t.id = b.termo_id WHERE b.numero = ? ORDER BY t.emitido_em DESC""", numero),
    }
```

Os `except` de `importar_bens` já fazem `conn.rollback()`: a importação e o log caem juntos.

- [ ] **Step 4: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam (108). Se `test_importar_substitui_bens_e_conta` quebrar por comparar o dicionário inteiro, ajuste-o para checar só as chaves `total`, `ativos`, `sem_centro`.

- [ ] **Step 5: Commit**

```bash
git add db.py tests/test_db.py
git commit -m "Importação registra o que mudou (novos, removidos, movidos, situação); histórico do bem"
```

---

### Task 8: Telas de importação e histórico na ficha do bem

**Files:**
- Modify: `app.py` (`upload`, `bem`; rota `importacao_tela`)
- Modify: `templates/upload.html`, `templates/bem.html`; Create: `templates/importacao.html`
- Test: `tests/test_app.py`

**Interfaces:**
- Consumes: Task 7.

- [ ] **Step 1: Teste que falha**

```python
def test_upload_mostra_mudancas_e_detalhe(cliente, tmp_path):
    from tests.test_db import xlsx
    arq = xlsx(tmp_path, [
        [1001, "ATIVO", "CADEIRA", "GIRATÓRIA", "MÓVEIS", "02 - OUTRA", "31/12/1996", 75.94, 64.54],
        [1002, "ATIVO", "NOTEBOOK", "DELL", "EQUIPAMENTOS", "01 - SALA CCI", "06/12/2012", 3000, 1500],
    ])
    with open(arq, "rb") as f:
        r = cliente.post("/upload", data={"arquivo": (f, "export.xlsx")}, content_type="multipart/form-data", follow_redirects=True)
    assert b"2 bens importados" in r.data and b"1 movido" in r.data and b"2 removido" in r.data and b"export.xlsx" in r.data
    import db
    iid = db.importacoes(db.conectar())[0]["id"]
    r = cliente.get(f"/importacoes/{iid}")
    assert r.status_code == 200 and b"02 - OUTRA" in r.data and b"1004" in r.data and b"removido" in r.data
    r = cliente.get("/bem?numero=1001")
    assert b"02 - OUTRA" in r.data and b"movido" in r.data
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest -q tests/test_app.py -k upload_mostra`
Expected: FAIL.

- [ ] **Step 3: Rotas**

Em `upload()`, no POST:

```python
        resumo = db.importar_bens(obter_conn(), arquivo.stream, nome_arquivo=arquivo.filename)
        flash(f"{resumo['total']} bens importados ({resumo['ativos']} ativos): {resumo['novos']} novo(s), "
              f"{resumo['removidos']} removido(s), {resumo['movidos']} movido(s), {resumo['situacao']} com situação alterada.", "success")
```

e no GET: `return render_template("upload.html", sem_centro=..., importacoes=db.importacoes(obter_conn()), trilha=...)`.

Nova rota após `bens_exportar`:

```python
@app.route("/importacoes/<int:id>")
def importacao_tela(id):
    i = db.importacao(obter_conn(), id) or abort(404)
    return render_template("importacao.html", i=i, trilha=[("Atualizar base", url_for("upload")), (f"Importação de {i['importado_em'][:10]}", None)])
```

Em `bem()`: `return render_template("bem.html", bem=ficha, historico=db.historico_do_bem(obter_conn(), int(numero)), rotulos=db.ROTULO_TIPO, trilha=...)`.

Cuidado com a mensagem do teste: "1 movido" e "2 removidos" batem com "1 movido(s)" e "2 removido(s)". Mantenha esse formato.

- [ ] **Step 4: Templates**

`templates/upload.html`, ao final (antes de `{% endblock %}`):

```jinja
{% if importacoes %}
{% from "_macros.html" import cabecalho_tabela %}
<div class="mt-4">
{{ cabecalho_tabela('Últimas importações', 'imp') }}
  <thead><tr><th scope="col">Data</th><th scope="col">Arquivo</th><th scope="col" class="dsgov-numero">Total</th><th scope="col" class="dsgov-numero">Ativos</th><th scope="col" class="dsgov-numero">Novos</th><th scope="col" class="dsgov-numero">Removidos</th><th scope="col" class="dsgov-numero">Movidos</th><th scope="col" class="dsgov-numero">Situação</th></tr></thead>
  <tbody>
  {% for i in importacoes %}
  <tr><td><a href="{{ url_for('importacao_tela', id=i.id) }}">{{ i.importado_em[8:10] }}/{{ i.importado_em[5:7] }}/{{ i.importado_em[:4] }} {{ i.importado_em[11:16] }}</a></td><td>{{ i.arquivo or '—' }}</td>
    <td class="dsgov-numero">{{ i.total }}</td><td class="dsgov-numero">{{ i.ativos }}</td><td class="dsgov-numero">{{ i.novos }}</td><td class="dsgov-numero">{{ i.removidos }}</td><td class="dsgov-numero">{{ i.movidos }}</td><td class="dsgov-numero">{{ i.situacao }}</td></tr>
  {% endfor %}
  </tbody>
</table></div>
</div>
{% endif %}
```

`templates/importacao.html`:

```jinja
{% extends "base.html" %}
{% from "_macros.html" import cabecalho_tabela %}
{% block titulo %}Importação{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Importação de {{ i.importado_em[8:10] }}/{{ i.importado_em[5:7] }}/{{ i.importado_em[:4] }} {{ i.importado_em[11:16] }}</h1></div>
<p class="text-gray-70">{{ i.arquivo or 'arquivo não informado' }} · {{ i.total }} bens ({{ i.ativos }} ativos)</p>
<div class="row mb-4">
  {% for rotulo, n in [('Novos', i.novos), ('Removidos', i.removidos), ('Movidos', i.movidos), ('Situação alterada', i.situacao)] %}
  <div class="col-sm-6 col-md-3 mb-3"><div class="br-card h-100"><div class="card-content"><div class="text-up-04 text-weight-bold">{{ n }}</div><div class="text-down-01 text-gray-70">{{ rotulo }}</div></div></div></div>
  {% endfor %}
</div>
{{ cabecalho_tabela('Mudanças', 'mud') }}
  <thead><tr><th scope="col">Número</th><th scope="col">Descrição</th><th scope="col">Mudança</th><th scope="col">De</th><th scope="col">Para</th></tr></thead>
  <tbody>
  {% for m in i.mudancas %}
  <tr><td><a href="{{ url_for('bem', numero=m.numero) }}">{{ m.numero }}</a></td><td>{{ m.descricao or '' }}</td>
    <td><span class="br-tag {{ {'novo': 'bg-success text-pure-0', 'removido': 'bg-danger text-pure-0', 'movido': 'bg-info text-pure-0', 'situacao': 'bg-warning'}[m.tipo] }}"><span>{{ m.tipo }}</span></span></td>
    <td>{{ m.de or '—' }}</td><td>{{ m.para or '—' }}</td></tr>
  {% else %}<tr><td colspan="5">Nenhuma mudança nesta importação.</td></tr>
  {% endfor %}
  </tbody>
</table></div>
{% endblock %}
```

`templates/bem.html`, antes de `{% endblock %}`:

```jinja
<h2 class="text-up-01 mt-4 mb-2">Histórico</h2>
<div class="row">
  <div class="col-md-6 mb-3"><div class="br-card h-100"><div class="card-header"><div class="text-weight-semi-bold">Movimentações (importações)</div></div><div class="card-content">
    {% for m in historico.mudancas %}
    <div class="mb-2"><span class="text-gray-70">{{ m.importado_em[8:10] }}/{{ m.importado_em[5:7] }}/{{ m.importado_em[:4] }}</span> · <strong>{{ m.tipo }}</strong>{% if m.de or m.para %}: {{ m.de or '—' }} → {{ m.para or '—' }}{% endif %}</div>
    {% else %}<p class="text-gray-70 mb-0">Nenhuma mudança registrada.</p>{% endfor %}
  </div></div></div>
  <div class="col-md-6 mb-3"><div class="br-card h-100"><div class="card-header"><div class="text-weight-semi-bold">Termos em que apareceu</div></div><div class="card-content">
    {% for t in historico.termos %}
    <div class="mb-2"><a href="{{ url_for('termo_emitido_tela', id=t.id) }}">{{ t.emitido_em[8:10] }}/{{ t.emitido_em[5:7] }}/{{ t.emitido_em[:4] }}</a> · {{ rotulos[t.tipo] }} · {{ t.chave }}{% if t.documento_sei %} · SEI {{ t.documento_sei }}{% endif %}</div>
    {% else %}<p class="text-gray-70 mb-0">Nenhum termo registrado com este bem.</p>{% endfor %}
  </div></div></div>
</div>
```

- [ ] **Step 5: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam (109).

- [ ] **Step 6: Commit**

```bash
git add app.py templates/upload.html templates/importacao.html templates/bem.html tests/test_app.py
git commit -m "Atualizar base mostra o que mudou; detalhe da importação; histórico na ficha do bem"
```

---

### Task 9: Infra de gráficos (ECharts, tema, helper, macro, CSS)

**Files:**
- Create (cópias): `static/dsgov/vendor/echarts/echarts.min.js`, `static/dsgov/vendor/echarts/LICENSE`, `static/dsgov/js/echarts-dsgov.js`, `graficos.py`
- Modify: `templates/_macros.html` (macro `grafico`), `static/dsgov/css/dsgov.css`
- Test: `tests/test_painel.py` (novo)

**Interfaces:**
- Produces: `graficos.rosca(fatias, total=None, rotulos=False, urls=None)`, `graficos.barras_horizontais(rotulos, valores, nome, escala=False, urls=None)`, `graficos.colunas(categorias, series, rotulos=False, urls=None)`, `graficos.linha(categorias, series, urls=None)`, `graficos.tabela_dados(colunas, linhas)`; macro `grafico(g)` com `g = {id, titulo, subtitulo, opcoes, resumo, alto, col, tabela}`.

- [ ] **Step 1: Copiar arquivos**

```bash
S=/home/ToNiauM/.claude/skills/dsgov/assets
mkdir -p static/dsgov/vendor/echarts
cp $S/vendor/echarts/5.5.0/echarts.min.js $S/vendor/echarts/5.5.0/LICENSE static/dsgov/vendor/echarts/
cp $S/projeto/core/static/dsgov/js/echarts-dsgov.js static/dsgov/js/
cp $S/projeto/core/graficos.py graficos.py
```

Em `graficos.py`, troque a primeira linha do docstring por: `"""Opções ECharts prontas, no tema dsgov (copiado da skill /dsgov, sem Django). A rota monta os dados; o JSON vai ao template pela macro grafico."""`. Nada mais muda.

- [ ] **Step 2: Teste que falha**

`tests/test_painel.py`:

```python
"""Macro grafico e helper graficos.py."""
import json

from flask import render_template_string

import graficos
from tests.conftest import semear


def test_graficos_helper_urls_e_tabela():
    op = graficos.rosca([("ATIVO", 3), ("BAIXADO", 1)], total=(4, "bens"), urls={"ATIVO": "/recorte?situacao=ATIVO"})
    assert op["series"][0]["data"][0] == {"name": "ATIVO", "value": 3, "url": "/recorte?situacao=ATIVO"}
    assert op["graphic"][0]["style"]["text"] == "4\nbens"
    op = graficos.barras_horizontais(["CCI"], [3], "Bens", escala=True, urls=["/r?ccusto=CCI"])
    assert op["series"][0]["data"] == [{"value": 3, "url": "/r?ccusto=CCI"}] and "visualMap" in op
    t = graficos.tabela_dados(["A", "B"], [[{"valor": "x", "url": "/x"}, 2]])
    assert t["linhas"][0] == [{"valor": "x", "url": "/x"}, {"valor": 2}]


def test_macro_grafico_renderiza_json_e_tabela(dados):
    semear(dados)
    from app import app
    g = {"id": "g1", "titulo": "Teste", "subtitulo": None, "alto": False, "col": None, "resumo": "Teste: 1",
         "opcoes": {"series": [{"type": "pie", "data": [{"name": "<b>", "value": 1}]}]},
         "tabela": graficos.tabela_dados(["Rótulo", "Bens"], [[{"valor": "CCI", "url": "/recorte?ccusto=CCI"}, 1]])}
    with app.test_request_context():
        html = render_template_string('{% from "_macros.html" import grafico %}{{ grafico(g) }}', g=g)
    assert 'data-grafico="g1"' in html and '<script type="application/json" id="g1">' in html
    assert "<b>" not in html.split('id="g1">')[1].split("</script>")[0]      # tojson escapa
    assert 'href="/recorte?ccusto=CCI"' in html and "Ver dados" in html and "dsgov-grafico" in html
```

- [ ] **Step 3: Rodar e ver falhar**

Run: `.venv/bin/pytest -q tests/test_painel.py`
Expected: o 1º passa (helper copiado); o 2º falha (macro inexistente).

- [ ] **Step 4: Macro e CSS**

Ao final de `templates/_macros.html`:

```jinja
{% macro grafico(g) %}
{# g = {id, titulo, subtitulo, opcoes (dict ECharts), resumo (aria-label), alto (bool), col, tabela (graficos.tabela_dados)} #}
<div class="{{ g.col or 'col-sm-12 col-md-6' }} mb-3">
  <div class="br-card h-100">
    <div class="card-header">
      <div class="text-weight-semi-bold text-up-01">{{ g.titulo }}</div>
      {% if g.subtitulo %}<div class="text-down-01 text-gray-70">{{ g.subtitulo }}</div>{% endif %}
    </div>
    <div class="card-content">
      <script type="application/json" id="{{ g.id }}">{{ g.opcoes|tojson }}</script>
      <div class="{% if g.alto %}dsgov-grafico-alto{% else %}dsgov-grafico{% endif %}" data-grafico="{{ g.id }}" role="img" aria-label="{{ g.resumo }}"></div>
      {% if g.tabela %}
      <div class="br-accordion mt-2" id="acc-{{ g.id }}">
        <div class="item">
          <button class="header" type="button" aria-controls="acc-{{ g.id }}-dados" aria-expanded="false">
            <span class="icon"><i class="fas fa-angle-down" aria-hidden="true"></i></span><span class="title">Ver dados</span>
          </button>
        </div>
        <div class="content" id="acc-{{ g.id }}-dados">
          <div class="br-table small"><div class="table-header"></div>
            <table>
              <caption class="sr-only">{{ g.titulo }} — dados</caption>
              <thead><tr>{% for c in g.tabela.colunas %}<th scope="col"{% if not loop.first %} class="dsgov-numero"{% endif %}>{{ c }}</th>{% endfor %}</tr></thead>
              <tbody>
              {% for linha in g.tabela.linhas %}
              <tr>{% for cel in linha %}<td{% if not loop.first %} class="dsgov-numero"{% endif %}>{% if cel.url %}<a href="{{ cel.url }}">{{ cel.valor }}</a>{% else %}{{ cel.valor }}{% endif %}</td>{% endfor %}</tr>
              {% endfor %}
              </tbody>
            </table>
          </div>
        </div>
      </div>
      {% endif %}
    </div>
  </div>
</div>
{% endmacro %}
```

Ao final de `static/dsgov/css/dsgov.css`:

```css
/* Gráficos ECharts: altura fixa pela classe, nunca por style inline. */
.dsgov-grafico { width: 100%; height: 280px; }
.dsgov-grafico-alto { width: 100%; height: 400px; }

/* Card-indicador clicável do painel. */
a.dsgov-kpi { display: block; color: inherit; text-decoration: none; }
a.dsgov-kpi:hover { box-shadow: var(--surface-shadow-md); }
.dsgov-kpi .card-content { min-height: 96px; }
```

- [ ] **Step 5: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam (111).

- [ ] **Step 6: Commit**

```bash
git add static/dsgov/vendor/echarts static/dsgov/js/echarts-dsgov.js graficos.py templates/_macros.html static/dsgov/css/dsgov.css tests/test_painel.py
git commit -m "Gráficos: ECharts 5.5.0 embutido com tema DSGov, helper graficos.py e macro grafico"
```

---

### Task 10: `db.dimensoes`, `db.painel`, `db.recorte`, `db.exportar_recorte`

**Files:**
- Modify: `db.py` (bloco novo antes de `def exportar_bens`)
- Test: `tests/test_db.py`

**Interfaces:**
- Consumes: Tasks 2 e 7 (`situacoes_centros`, `situacoes_pessoas`, `importacoes`).
- Produces: `db.FILTROS`, `db.IMOVEIS`, `db.FAIXAS_IDADE`, `db.FAIXAS_VALOR`, `db.dimensoes(conn, f) -> dict` (chaves `situacao, centro, classificacao, localizacao, idade, ano, faixa, pessoa`; itens `{chave, rotulo, quantidade, valor}`), `db.painel(conn) -> dict`, `db.recorte(conn, f, limite=1000) -> dict`, `db.exportar_recorte(conn, f, destino)`.

- [ ] **Step 1: Testes que falham**

```python
# ---------------------------------------------------------------- painel e recorte
def _semear_painel(conn):
    semear(conn)
    conn.execute("INSERT INTO bens VALUES (2001,'ATIVO','SEDE','','SEDE','','01/01/1990',1,60000000)")
    conn.execute("INSERT INTO bens VALUES (2002,'ATIVO','MICRO','HP','EQUIPAMENTOS','01 - SALA CCI','15/06/2024',4000,3500)")
    conn.commit()


def test_dimensoes_sobre_ativos(dados):
    _semear_painel(dados)
    d = db.dimensoes(dados, {"situacao": "ATIVO"})
    assert {x["chave"]: x["quantidade"] for x in d["situacao"]} == {"ATIVO": 5, "BAIXADO": 1}   # situação ignora o próprio filtro
    centro = {x["chave"]: (x["quantidade"], x["rotulo"]) for x in d["centro"]}
    assert centro["CCI"] == (3, "CCI – JAQUELINE PORTELA") and centro["-"] == (2, "sem centro")
    assert {x["chave"]: x["quantidade"] for x in d["classificacao"]} == {"MÓVEIS": 2, "EQUIPAMENTOS": 2, "SEDE": 1}
    assert [x["chave"] for x in d["idade"]] == ["ate5", "5a10", "10a20", "mais20", "semdata"]
    assert {x["chave"]: x["quantidade"] for x in d["idade"] if x["quantidade"]} == {"ate5": 1, "10a20": 2, "mais20": 2}
    assert [x["chave"] for x in d["ano"]] == ["1990", "1996", "2012", "2024"]
    assert {x["chave"]: x["quantidade"] for x in d["faixa"] if x["quantidade"]} == {"ate100": 1, "100a500": 1, "1000a5000": 2, "mais20000": 1}
    assert d["pessoa"] == [{"chave": "ANA SILVA", "rotulo": "ANA SILVA", "quantidade": 1, "valor": 1500.0}]
    assert any(x["chave"] == "01 - SALA CCI" and "(CCI)" in x["rotulo"] and x["quantidade"] == 3 for x in d["localizacao"])


def test_painel_cards(dados):
    _semear_painel(dados)
    p = db.painel(dados)
    assert p["ativos"] == 5 and p["imoveis"] == 1 and p["valor_imoveis"] == 60000000
    assert round(p["valor_sem_imoveis"], 2) == 64.54 + 1500 + 250.5 + 3500
    assert p["sem_centro"] == 2 and p["sem_valor"] == 0 and p["ultima_importacao"] is None
    assert p["a_emitir_centros"] == 1 and p["a_emitir_pessoas"] == 1
    assert p["centros"][0]["ccustos"] == "CCI" and "dimensoes" in p


def test_recorte_filtros_e_drill_down(dados):
    _semear_painel(dados)
    r = db.recorte(dados, {"situacao": "ATIVO", "ccusto": "CCI"})
    assert [b["numero"] for b in r["bens"]] == [1001, 1002, 2002] and r["quantidade"] == 3 and r["bens"][1]["pessoa"] == "ANA SILVA"
    assert [b["numero"] for b in db.recorte(dados, {"situacao": "ATIVO", "ccusto": "CCI", "faixa": "1000a5000"})["bens"]] == [1002, 2002]
    assert [b["numero"] for b in db.recorte(dados, {"ccusto": "-"})["bens"]] == [1004, 2001]
    assert [b["numero"] for b in db.recorte(dados, {"classificacao": "imoveis"})["bens"]] == [2001]
    assert 2001 not in [b["numero"] for b in db.recorte(dados, {"classificacao": "sem-imoveis"})["bens"]]
    assert [b["numero"] for b in db.recorte(dados, {"valor_de": "1000", "valor_ate": "2000"})["bens"]] == [1002]
    assert [b["numero"] for b in db.recorte(dados, {"entrada_de": "2012-01-01", "entrada_ate": "2012-12-31"})["bens"]] == [1002, 1003, 1004]
    assert [b["numero"] for b in db.recorte(dados, {"ano": "2024"})["bens"]] == [2002]
    assert [b["numero"] for b in db.recorte(dados, {"idade": "mais20"})["bens"]] == [1001, 2001]
    assert [b["numero"] for b in db.recorte(dados, {"pessoa": "ANA SILVA"})["bens"]] == [1002]
    assert [b["numero"] for b in db.recorte(dados, {"localizacao": "99 - SEM MAPA"})["bens"]] == [1004]
    assert db.recorte(dados, {})["quantidade"] == 6          # sem filtro = tudo (a rota põe ATIVO por padrão)
    r = db.recorte(dados, {}, limite=2)
    assert len(r["bens"]) == 2 and r["truncado"] and r["quantidade"] == 6


def test_exportar_recorte_xlsx(dados, tmp_path):
    _semear_painel(dados)
    from openpyxl import load_workbook
    ws = load_workbook(db.exportar_recorte(dados, {"ccusto": "CCI"}, tmp_path / "r.xlsx")).active
    linhas = list(ws.iter_rows(values_only=True))
    assert linhas[0][:3] == ("Número", "Descrição", "Complemento") and len(linhas) == 5   # cabeçalho + 4 bens (inclui 1003 BAIXADO)
    assert linhas[1][5] == "CCI"
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest -q tests/test_db.py -k "dimensoes or painel_cards or recorte"`
Expected: 4 falhas.

- [ ] **Step 3: Implementar**

Antes de `def exportar_bens` em `db.py`:

```python
# ---------------------------------------------------------------- painel e recorte
IMOVEIS = ("SEDE", "TERRENOS")
FAIXAS_IDADE = [("ate5", "até 5 anos"), ("5a10", "5 a 10 anos"), ("10a20", "10 a 20 anos"),
                ("mais20", "mais de 20 anos"), ("semdata", "sem data")]
FAIXAS_VALOR = [("ate100", "até R$ 100"), ("100a500", "R$ 100 a 500"), ("500a1000", "R$ 500 a 1.000"),
                ("1000a5000", "R$ 1.000 a 5.000"), ("5000a20000", "R$ 5.000 a 20.000"), ("mais20000", "acima de R$ 20.000")]
FILTROS = ("situacao", "ccusto", "pessoa", "localizacao", "classificacao", "valor_de", "valor_ate",
           "entrada_de", "entrada_ate", "idade", "faixa", "ano")
_DE = ("FROM bens b LEFT JOIN localizacoes l ON l.localizacao = b.localizacao "
       "LEFT JOIN atribuicoes a ON a.numero = b.numero")
_DATA_ISO = ("CASE WHEN b.data_entrada LIKE '__/__/____' THEN substr(b.data_entrada,7,4)||'-'||"
             "substr(b.data_entrada,4,2)||'-'||substr(b.data_entrada,1,2) END")
_IDADE = f"(julianday('now') - julianday({_DATA_ISO})) / 365.25"
_FAIXA_IDADE = (f"CASE WHEN {_DATA_ISO} IS NULL THEN 'semdata' WHEN {_IDADE} < 5 THEN 'ate5' "
                f"WHEN {_IDADE} < 10 THEN '5a10' WHEN {_IDADE} < 20 THEN '10a20' ELSE 'mais20' END")
_V = "coalesce(b.valor_atual, 0)"
_FAIXA_VALOR = (f"CASE WHEN {_V} < 100 THEN 'ate100' WHEN {_V} < 500 THEN '100a500' WHEN {_V} < 1000 THEN '500a1000' "
                f"WHEN {_V} < 5000 THEN '1000a5000' WHEN {_V} < 20000 THEN '5000a20000' ELSE 'mais20000' END")
_IMOVEIS_SQL = "coalesce(b.classificacao, '') IN ('SEDE', 'TERRENOS')"


def _where(f: dict) -> tuple[str, list]:
    """WHERE só com os filtros presentes em f (chaves de FILTROS; vazio = sem filtro)."""
    cl, p = [], []
    if f.get("situacao"):
        cl.append("b.situacao = ?"); p.append(f["situacao"])
    if f.get("ccusto") == "-":
        cl.append("l.ccustos IS NULL")
    elif f.get("ccusto"):
        cl.append("l.ccustos = ?"); p.append(f["ccusto"])
    if f.get("pessoa"):
        cl.append("a.nome = ?"); p.append(f["pessoa"])
    if f.get("localizacao"):
        cl.append("b.localizacao = ?"); p.append(f["localizacao"])
    c = f.get("classificacao")
    if c == "imoveis":
        cl.append(_IMOVEIS_SQL)
    elif c == "sem-imoveis":
        cl.append(f"NOT {_IMOVEIS_SQL}")
    elif c:
        cl.append("b.classificacao = ?"); p.append(c)
    if f.get("valor_de"):
        cl.append(f"{_V} >= ?"); p.append(float(f["valor_de"]))
    if f.get("valor_ate"):
        cl.append(f"{_V} <= ?"); p.append(float(f["valor_ate"]))
    if f.get("entrada_de"):
        cl.append(f"{_DATA_ISO} >= ?"); p.append(f["entrada_de"])
    if f.get("entrada_ate"):
        cl.append(f"{_DATA_ISO} <= ?"); p.append(f["entrada_ate"])
    if f.get("idade"):
        cl.append(f"{_FAIXA_IDADE} = ?"); p.append(f["idade"])
    if f.get("faixa"):
        cl.append(f"{_FAIXA_VALOR} = ?"); p.append(f["faixa"])
    if f.get("ano"):
        cl.append("substr(b.data_entrada, 7, 4) = ?"); p.append(f["ano"])
    return (" AND ".join(cl) or "1"), p


def _agrupar(conn, chave: str, f: dict, ordem: str = "quantidade DESC, chave") -> list[dict]:
    where, p = _where(f)
    return _todos(conn, f"SELECT {chave} AS chave, count(*) AS quantidade, coalesce(sum(b.valor_atual), 0) AS valor "
                        f"{_DE} WHERE {where} GROUP BY chave ORDER BY {ordem}", *p)


def dimensoes(conn, f: dict) -> dict:
    """Agrupamentos do conjunto filtrado por f. Item: {chave, rotulo, quantidade, valor}.
    'situacao' ignora o filtro de situação (mostra a composição inteira)."""
    resp = {c["ccustos"]: c["responsavel"] for c in centros(conn)}
    mapa = {l["localizacao"]: l["ccustos"] for l in localizacoes_mapeadas(conn)}

    def rot(lista, fn):
        return [{**x, "rotulo": fn(x["chave"])} for x in lista]

    def fixas(faixas, lista):
        por = {x["chave"]: x for x in lista}
        return [{**por.get(k, {"chave": k, "quantidade": 0, "valor": 0}), "rotulo": r} for k, r in faixas]

    return {
        "situacao": rot(_agrupar(conn, "b.situacao", {k: v for k, v in f.items() if k != "situacao"}), str),
        "centro": rot(_agrupar(conn, "coalesce(l.ccustos, '-')", f),
                      lambda k: "sem centro" if k == "-" else f"{k} – {resp.get(k, '')}"),
        "classificacao": rot(_agrupar(conn, "coalesce(b.classificacao, '')", f), lambda k: k or "sem classificação"),
        "localizacao": rot(_agrupar(conn, "coalesce(b.localizacao, '')", f),
                           lambda k: (k or "sem localização") + (f" ({mapa[k]})" if k in mapa else "")),
        "idade": fixas(FAIXAS_IDADE, _agrupar(conn, _FAIXA_IDADE, f)),
        "ano": rot(_agrupar(conn, "coalesce(substr(b.data_entrada, 7, 4), '')", f, ordem="chave"), lambda k: k or "sem data"),
        "faixa": fixas(FAIXAS_VALOR, _agrupar(conn, _FAIXA_VALOR, f)),
        "pessoa": rot([x for x in _agrupar(conn, "a.nome", f) if x["chave"]], str),
    }


def painel(conn) -> dict:
    def um(sql, *p):
        return conn.execute(sql, p).fetchone()[0]
    ativo = "b.situacao = 'ATIVO'"
    d = {
        "ativos": um(f"SELECT count(*) {_DE} WHERE {ativo}"),
        "valor_sem_imoveis": um(f"SELECT coalesce(sum(b.valor_atual), 0) {_DE} WHERE {ativo} AND NOT {_IMOVEIS_SQL}"),
        "imoveis": um(f"SELECT count(*) {_DE} WHERE {ativo} AND {_IMOVEIS_SQL}"),
        "valor_imoveis": um(f"SELECT coalesce(sum(b.valor_atual), 0) {_DE} WHERE {ativo} AND {_IMOVEIS_SQL}"),
        "sem_centro": um(f"SELECT count(*) {_DE} WHERE {ativo} AND l.ccustos IS NULL"),
        "sem_valor": um(f"SELECT count(*) FROM bens b WHERE {ativo} AND {_V} = 0"),
        "ultima_importacao": (importacoes(conn, 1) or [None])[0],
        "centros": situacoes_centros(conn),
        "pessoas": situacoes_pessoas(conn),
        "dimensoes": dimensoes(conn, {"situacao": "ATIVO"}),
    }
    d["a_emitir_centros"] = sum(1 for c in d["centros"] if c["estado"] != "vigente")
    d["a_emitir_pessoas"] = sum(1 for p in d["pessoas"] if p["estado"] != "vigente" and p["quantidade"])
    return d


def recorte(conn, f: dict, limite: int | None = 1000) -> dict:
    where, p = _where(f)
    sql = f"SELECT b.*, l.ccustos AS ccustos, a.nome AS pessoa {_DE} WHERE {where} ORDER BY b.numero"
    bens = _todos(conn, sql + (f" LIMIT {limite + 1}" if limite else ""), *p)
    tot = conn.execute(f"SELECT count(*), coalesce(sum(b.valor_atual), 0) {_DE} WHERE {where}", p).fetchone()
    return {"bens": bens[:limite] if limite else bens, "truncado": bool(limite) and len(bens) > limite,
            "quantidade": tot[0], "valor_total": tot[1], "dimensoes": dimensoes(conn, f)}


def exportar_recorte(conn, f: dict, destino):
    wb = Workbook()
    ws = wb.active
    ws.title = "recorte"
    ws.append(["Número", "Descrição", "Complemento", "Classificação", "Localização", "Centro de custo", "Pessoa",
               "Situação", "Data entrada", "Valor compra", "Valor atual"])
    for b in recorte(conn, f, limite=None)["bens"]:
        ws.append([b["numero"], b["descricao"], b["complemento"], b["classificacao"], b["localizacao"], b["ccustos"],
                   b["pessoa"], b["situacao"], b["data_entrada"], b["valor_compra"], b["valor_atual"]])
    wb.save(destino)
    return destino
```

- [ ] **Step 4: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam (115). Se `idade` divergir por causa da data de hoje (os testes rodam em 2026: 1996 e 1990 → mais de 20; 2012 → 10 a 20; 2024 → até 5), confira o cálculo, não os dados.

- [ ] **Step 5: Commit**

```bash
git add db.py tests/test_db.py
git commit -m "db: dimensões, painel, recorte com filtros combináveis e exportação"
```

---

### Task 11: `painel.py` (cards de gráfico e descrição de filtros)

**Files:**
- Create: `painel.py`
- Test: `tests/test_painel.py`

**Interfaces:**
- Consumes: `graficos.py` (Task 9), `db.dimensoes` (Task 10), `url_for("recorte")` (rota criada na Task 13; nos testes, registrada por `app`).
- Produces: `painel.moeda(v) -> str`, `painel.url_recorte(f, **extra) -> str`, `painel.cards_graficos(dim, f, omitir=()) -> list[dict]` (dicts para a macro `grafico`), `painel.descrever(f, nomes) -> str`.

- [ ] **Step 1: Teste que falha**

Acrescente em `tests/test_painel.py`:

```python
def test_cards_graficos_tipos_urls_e_omissao(dados):
    from tests.test_db import _semear_painel
    _semear_painel(dados)
    import db, painel
    from app import app
    f = {"situacao": "ATIVO"}
    with app.test_request_context():
        cards = painel.cards_graficos(db.dimensoes(dados, f), f)
        ids = [c["id"] for c in cards]
        assert ids == ["g-situacao", "g-centro", "g-classificacao", "g-localizacao", "g-idade", "g-ano", "g-faixa", "g-pessoa"]
        por = {c["id"]: c for c in cards}
        assert por["g-situacao"]["opcoes"]["series"][0]["type"] == "pie"
        assert por["g-centro"]["opcoes"]["series"][0]["type"] == "bar" and por["g-centro"]["opcoes"]["yAxis"]["type"] == "category"
        assert por["g-idade"]["opcoes"]["xAxis"]["type"] == "category" and por["g-ano"]["opcoes"]["series"][0]["type"] == "line"
        assert por["g-centro"]["opcoes"]["series"][0]["data"][0]["url"] == "/recorte?situacao=ATIVO&ccusto=CCI"
        assert por["g-situacao"]["opcoes"]["series"][0]["data"][0]["url"] == "/recorte?situacao=ATIVO"
        assert por["g-classificacao"]["tabela"]["linhas"][-1][0]["valor"] == "SEDE"          # imóveis por último na tabela
        assert "SEDE" not in [d["name"] for d in por["g-classificacao"]["opcoes"]["series"][0]["data"]]
        assert por["g-ano"]["opcoes"]["series"][0]["data"][-1]["url"] == "/recorte?situacao=ATIVO&ano=2024"
        assert por["g-faixa"]["tabela"]["linhas"][0][2]["valor"] == "R$ 64,54"
        cards = painel.cards_graficos(db.dimensoes(dados, {"situacao": "ATIVO", "ccusto": "CCI"}), {"situacao": "ATIVO", "ccusto": "CCI"}, omitir=("situacao", "ccusto"))
        assert "g-centro" not in [c["id"] for c in cards] and "g-situacao" not in [c["id"] for c in cards]
        assert painel.descrever({"situacao": "ATIVO", "ccusto": "CCI", "entrada_de": "2020-01-01"}, {"ccusto": "CCI – JAQUELINE"}) == \
            "Bens ATIVO · centro de custo CCI – JAQUELINE · entrada a partir de 01/01/2020"
        assert painel.moeda(1234.5) == "R$ 1.234,50"
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest -q tests/test_painel.py -k cards_graficos`
Expected: FAIL (`ModuleNotFoundError: painel`).

- [ ] **Step 3: Implementar `painel.py`**

```python
"""Cards de gráfico do painel e do recorte a partir de db.dimensoes(). Só apresentação: escolhe o tipo
de gráfico por dimensão (spec, parte 4) e monta as URLs de drill-down para /recorte."""
from flask import url_for

import db
import graficos

ROTULO_FILTRO = {"situacao": "situação", "ccusto": "centro de custo", "pessoa": "pessoa", "localizacao": "localização",
                 "classificacao": "classificação", "idade": "idade", "faixa": "faixa de valor", "ano": "ano de entrada"}
TOP = 20


def moeda(v) -> str:
    return f"R$ {v or 0:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")


def url_recorte(f: dict, **extra) -> str:
    """URL de /recorte com os filtros atuais mais os de `extra` (drill-down acrescenta, não substitui)."""
    return url_for("recorte", **{k: v for k, v in {**f, **extra}.items() if v})


def _tabela(itens, f, chave, rotulo_col):
    return graficos.tabela_dados([rotulo_col, "Bens", "Valor"], [
        [{"valor": i["rotulo"], "url": url_recorte(f, **{chave: i["chave"]})}, i["quantidade"], moeda(i["valor"])] for i in itens])


def _card(id, titulo, opcoes, itens, f, chave, rotulo_col, alto=False, subtitulo=None):
    resumo = f"{titulo}: " + ", ".join(f"{i['rotulo']} {i['quantidade']}" for i in itens[:6])
    return {"id": id, "titulo": titulo, "subtitulo": subtitulo, "opcoes": opcoes, "resumo": resumo,
            "alto": alto, "col": None, "tabela": _tabela(itens, f, chave, rotulo_col)}


def _urls(itens, f, chave):
    return [url_recorte(f, **{chave: i["chave"]}) for i in itens]


def _barras(id, titulo, itens, f, chave, rotulo_col):
    top = itens[:TOP]
    sub = f"{TOP} maiores no gráfico; todos na tabela" if len(itens) > TOP else None
    op = graficos.barras_horizontais([i["rotulo"] for i in top], [i["quantidade"] for i in top], "Bens",
                                     escala=True, urls=_urls(top, f, chave))
    return _card(id, titulo, op, itens, f, chave, rotulo_col, alto=len(top) > 8, subtitulo=sub)


def _colunas(id, titulo, itens, f, chave, rotulo_col):
    op = graficos.colunas([i["rotulo"] for i in itens], {"Bens": [i["quantidade"] for i in itens]},
                          rotulos=True, urls={"Bens": _urls(itens, f, chave)})
    return _card(id, titulo, op, itens, f, chave, rotulo_col)


def cards_graficos(dim: dict, f: dict, omitir=()) -> list[dict]:
    """Um card por dimensão, na ordem da spec. omitir = filtros já fixados num único valor."""
    cards = []
    if "situacao" not in omitir and dim["situacao"]:
        it = dim["situacao"]
        op = graficos.rosca([(i["rotulo"], i["quantidade"]) for i in it], total=(sum(i["quantidade"] for i in it), "bens"),
                            urls={i["rotulo"]: u for i, u in zip(it, _urls(it, f, "situacao"))})
        cards.append(_card("g-situacao", "Bens por situação", op, it, f, "situacao", "Situação"))
    if "ccusto" not in omitir:
        cards.append(_barras("g-centro", "Bens por centro de custo", dim["centro"], f, "ccusto", "Centro de custo"))
    if "classificacao" not in omitir:
        it = dim["classificacao"]
        comuns = [i for i in it if i["chave"] not in db.IMOVEIS]
        imoveis = [i for i in it if i["chave"] in db.IMOVEIS]
        fatias = [(i["rotulo"], i["quantidade"]) for i in comuns[:5]]
        if len(comuns) > 5:
            fatias.append(("Outras", sum(i["quantidade"] for i in comuns[5:])))
        op = graficos.rosca(fatias, total=(sum(i["quantidade"] for i in comuns), "bens"),
                            urls={i["rotulo"]: u for i, u in zip(comuns[:5], _urls(comuns[:5], f, "classificacao"))})
        cards.append(_card("g-classificacao", "Bens por classificação contábil", op, comuns + imoveis, f, "classificacao",
                           "Classificação", subtitulo="Imóveis (SEDE, TERRENOS) só na tabela" if imoveis else None))
    if "localizacao" not in omitir:
        cards.append(_barras("g-localizacao", "Bens por localização", dim["localizacao"], f, "localizacao", "Localização"))
    if "idade" not in omitir:
        cards.append(_colunas("g-idade", "Bens por idade (data de entrada)", dim["idade"], f, "idade", "Faixa"))
    if "ano" not in omitir and dim["ano"]:
        it = dim["ano"]
        op = graficos.linha([i["rotulo"] for i in it], {"Bens": [i["quantidade"] for i in it]}, urls={"Bens": _urls(it, f, "ano")})
        cards.append(_card("g-ano", "Bens por ano de entrada", op, it, f, "ano", "Ano"))
    if "faixa" not in omitir:
        cards.append(_colunas("g-faixa", "Bens por faixa de valor", dim["faixa"], f, "faixa", "Faixa"))
    if "pessoa" not in omitir and dim["pessoa"]:
        cards.append(_barras("g-pessoa", "Bens atribuídos por pessoa", dim["pessoa"], f, "pessoa", "Pessoa"))
    return cards


def _data_br(iso: str) -> str:
    return f"{iso[8:10]}/{iso[5:7]}/{iso[:4]}"


def descrever(f: dict, nomes: dict | None = None) -> str:
    """Frase do recorte: "Bens ATIVO · centro de custo CCI · entrada a partir de 01/01/2020". nomes = rótulos
    legíveis por filtro (ex.: {"ccusto": "CCI – JAQUELINE", "idade": "mais de 20 anos"})."""
    nomes = nomes or {}
    partes = ["Bens " + f["situacao"] if f.get("situacao") else "Bens (todas as situações)"]
    for k in ("ccusto", "pessoa", "localizacao", "classificacao", "idade", "faixa", "ano"):
        if f.get(k):
            v = nomes.get(k) or {"-": "sem centro", "imoveis": "imóveis", "sem-imoveis": "sem imóveis"}.get(f[k], f[k])
            partes.append(f"{ROTULO_FILTRO[k]} {v}")
    if f.get("valor_de") and f.get("valor_ate"):
        partes.append(f"valor de {moeda(float(f['valor_de']))} a {moeda(float(f['valor_ate']))}")
    elif f.get("valor_de"):
        partes.append(f"valor a partir de {moeda(float(f['valor_de']))}")
    elif f.get("valor_ate"):
        partes.append(f"valor até {moeda(float(f['valor_ate']))}")
    if f.get("entrada_de") and f.get("entrada_ate"):
        partes.append(f"entrada entre {_data_br(f['entrada_de'])} e {_data_br(f['entrada_ate'])}")
    elif f.get("entrada_de"):
        partes.append(f"entrada a partir de {_data_br(f['entrada_de'])}")
    elif f.get("entrada_ate"):
        partes.append(f"entrada até {_data_br(f['entrada_ate'])}")
    return " · ".join(partes)
```

Para o teste passar antes da Task 13, `url_for("recorte")` precisa existir: crie em `app.py` uma rota provisória (substituída na Task 13):

```python
@app.route("/recorte")
def recorte():
    return redirect(url_for("home"))   # completada na Task 13
```

- [ ] **Step 4: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam (116).

- [ ] **Step 5: Commit**

```bash
git add painel.py app.py tests/test_painel.py
git commit -m "painel.py: cards de gráfico por dimensão com drill-down e descrição do recorte"
```

---

### Task 12: Painel na tela inicial

**Files:**
- Modify: `app.py` (`home`), `templates/index.html`
- Test: `tests/test_app.py`

**Interfaces:**
- Consumes: `db.painel`, `painel.cards_graficos`, `painel.moeda`, `painel.url_recorte`, macros `grafico` e `tag_termo`.

- [ ] **Step 1: Teste que falha**

```python
def test_painel_na_tela_inicial(cliente):
    r = cliente.get("/")
    assert r.status_code == 200
    assert b"bens ativos" in r.data and b'data-grafico="g-centro"' in r.data and b'data-grafico="g-ano"' in r.data
    assert b"echarts.min.js" in r.data and b"echarts-dsgov.js" in r.data
    assert b"/recorte?situacao=ATIVO&amp;ccusto=CCI" in r.data or b"/recorte?situacao=ATIVO&ccusto=CCI" in r.data
    assert b"/termo/ccusto/CCI" in r.data and b"sem termo" in r.data
    assert b"Nenhuma" in r.data       # última importação: nenhuma
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest -q tests/test_app.py -k painel_na_tela`
Expected: FAIL.

- [ ] **Step 3: Rota**

No topo de `app.py`: `import painel`. Substitua `home()`:

```python
@app.route("/")
def home():
    p = db.painel(obter_conn())
    f = {"situacao": "ATIVO"}
    return render_template("index.html", p=p, cards=painel.cards_graficos(p["dimensoes"], f), f=f,
                           moeda=painel.moeda, url_recorte=painel.url_recorte, trilha=[])
```

- [ ] **Step 4: Template**

Substitua `templates/index.html` inteiro:

```jinja
{% extends "base.html" %}
{% from "_macros.html" import grafico, tag_termo, cabecalho_tabela %}
{% block titulo %}Início{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Início</h1></div>
<div class="br-card mb-4">
  <div class="card-header"><div class="text-weight-semi-bold text-up-01">Pesquisar</div></div>
  <div class="card-content">
    <form method="get" action="{{ url_for('pesquisa') }}" class="d-flex align-items-end">
      <div class="br-input mr-3 flex-fill">
        <label for="q">Número do patrimônio, descrição, nome ou centro de custo</label>
        <input id="q" name="q" type="text" placeholder="Ex.: 14359, computador, GEX* ou Antônio" required/>
      </div>
      <button class="br-button primary" type="submit"><i class="fas fa-search mr-1" aria-hidden="true"></i>Pesquisar</button>
    </form>
  </div>
</div>

{% if p.sem_valor %}
<div class="br-message warning mb-3"><div class="icon"><i class="fas fa-exclamation-triangle fa-lg" aria-hidden="true"></i></div>
  <div class="content" role="alert"><span class="message-title">{{ p.sem_valor }} bem(ns) ativo(s) sem valor.</span><span class="message-body"> Não entram nas somas. <a href="{{ url_recorte(f, valor_ate='0') }}">Ver quais</a>.</span></div></div>
{% endif %}

<div class="row mb-2">
  {% set kpis = [
    (p.ativos, 'bens ativos', url_recorte(f)),
    (moeda(p.valor_sem_imoveis), 'valor dos ativos (sem imóveis)', url_recorte(f, classificacao='sem-imoveis')),
    (p.imoveis ~ ' · ' ~ moeda(p.valor_imoveis), 'imóveis (SEDE, TERRENOS)', url_recorte(f, classificacao='imoveis')),
    (p.sem_centro, 'ativos sem centro de custo', url_recorte(f, ccusto='-')),
    (p.a_emitir_centros ~ ' centro(s) · ' ~ p.a_emitir_pessoas ~ ' pessoa(s)', 'termos a emitir ou reemitir', url_for('centro_custos')),
  ] %}
  {% for valor, legenda, url in kpis %}
  <div class="col-sm-6 col-md-4 col-lg mb-3"><a class="br-card h-100 dsgov-kpi" href="{{ url }}"><div class="card-content">
    <div class="text-up-03 text-weight-bold">{{ valor }}</div><div class="text-down-01 text-gray-70">{{ legenda }}</div>
  </div></a></div>
  {% endfor %}
  <div class="col-sm-6 col-md-4 col-lg mb-3">
    {% if p.ultima_importacao %}<a class="br-card h-100 dsgov-kpi" href="{{ url_for('importacao_tela', id=p.ultima_importacao.id) }}"><div class="card-content">
      <div class="text-up-03 text-weight-bold">{{ p.ultima_importacao.importado_em[8:10] }}/{{ p.ultima_importacao.importado_em[5:7] }}</div>
      <div class="text-down-01 text-gray-70">última importação · {{ p.ultima_importacao.novos }} novos, {{ p.ultima_importacao.removidos }} removidos</div>
    </div></a>
    {% else %}<div class="br-card h-100"><div class="card-content"><div class="text-up-03 text-weight-bold">—</div><div class="text-down-01 text-gray-70">Nenhuma importação registrada</div></div></div>{% endif %}
  </div>
</div>

<div class="row">
  {% for g in cards %}{{ grafico(g) }}{% endfor %}
</div>

<div class="mt-3">
{{ cabecalho_tabela('Termos por centro de custo', 'centros') }}
  <thead><tr><th scope="col">Centro</th><th scope="col">Responsável</th><th scope="col" class="dsgov-numero">Bens</th><th scope="col" class="dsgov-numero">Valor</th><th scope="col">Situação</th><th scope="col" class="dsgov-acoes">Ações</th></tr></thead>
  <tbody>
  {% for c in p.centros %}
  <tr>
    <td><a href="{{ url_recorte(f, ccusto=c.ccustos) }}">{{ c.ccustos }}</a></td><td>{{ c.responsavel }}</td>
    <td class="dsgov-numero">{{ c.quantidade }}</td><td class="dsgov-numero">{{ moeda(c.valor) }}</td>
    <td>{{ tag_termo(c) }}</td>
    <td class="dsgov-acoes"><a class="br-button circle small" href="{{ url_for('termo', tipo='ccusto', chave=c.ccustos) }}" aria-label="Termo de {{ c.ccustos }}"><i class="fas fa-file-alt" aria-hidden="true"></i></a></td>
  </tr>
  {% endfor %}
  </tbody>
</table></div>
</div>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='dsgov/vendor/echarts/echarts.min.js') }}"></script>
<script src="{{ url_for('static', filename='dsgov/js/echarts-dsgov.js') }}"></script>
{% endblock %}
```

- [ ] **Step 5: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam (117). O teste `test_home_e_busca_de_bem` continua válido.

- [ ] **Step 6: Commit**

```bash
git add app.py templates/index.html tests/test_app.py
git commit -m "Tela inicial vira painel: cards clicáveis, gráficos por dimensão e situação dos termos"
```

---

### Task 13: Tela de recorte, exportação e menu

**Files:**
- Modify: `app.py` (substituir `recorte` provisória; rota `recorte_xlsx`; MENU)
- Create: `templates/recorte.html`
- Test: `tests/test_app.py`

**Interfaces:**
- Consumes: `db.recorte`, `db.exportar_recorte`, `db.FILTROS`, `db.FAIXAS_IDADE`, `db.situacao_termo`, `painel.*`.
- Produces: `GET /recorte?<filtros>`, `GET /recorte/xlsx?<filtros>`.

- [ ] **Step 1: Teste que falha**

```python
def test_recorte_tela_filtros_termo_e_xlsx(cliente):
    r = cliente.get("/recorte")
    assert r.status_code == 200 and b"Bens ATIVO" in r.data and b"1001" in r.data and b"1003" not in r.data
    assert b'data-grafico="g-situacao"' not in r.data and b'data-grafico="g-centro"' in r.data
    r = cliente.get("/recorte?situacao=ATIVO&ccusto=CCI")
    assert b"centro de custo CCI" in r.data and b"/termo/ccusto/CCI" in r.data and b"sem termo" in r.data
    assert b'data-grafico="g-centro"' not in r.data and b"/recorte?situacao=ATIVO&amp;ccusto=CCI&amp;faixa=" in r.data
    r = cliente.get("/recorte?pessoa=ANA SILVA")
    assert b"/termo/individual/ANA" in r.data and b"NOTEBOOK" in r.data and b"CADEIRA" not in r.data
    r = cliente.get("/recorte?situacao=&valor_de=1.000,00&valor_ate=2000")
    assert b"NOTEBOOK" in r.data and b"CADEIRA" not in r.data and b"todas as situa" in r.data
    r = cliente.get("/recorte/xlsx?ccusto=CCI")
    assert r.status_code == 200 and r.headers["Content-Disposition"].endswith("recorte.xlsx")
    assert b"Recorte" in cliente.get("/").data       # menu
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest -q tests/test_app.py -k recorte_tela`
Expected: FAIL (302).

- [ ] **Step 3: Rotas e menu**

No `MENU`, após `("Termos emitidos", ...)`: `("Recorte", "fa-filter", url_for("recorte")),`.

Substitua a rota provisória:

```python
def _filtros_recorte() -> dict:
    """Filtros da query string. situacao ausente = ATIVO; situacao vazia (campo enviado em branco) = todas.
    Valores em R$ aceitam vírgula decimal e ponto de milhar."""
    f = {k: request.args.get(k, "").strip() for k in db.FILTROS}
    if "situacao" not in request.args:
        f["situacao"] = "ATIVO"
    for k in ("valor_de", "valor_ate"):
        if f[k]:
            v = f[k].replace("R$", "").strip()
            f[k] = v.replace(".", "").replace(",", ".") if "," in v else v
            try:
                float(f[k])
            except ValueError:
                raise db.ErroDeNegocio(f"Valor inválido: {v}")
    return {k: v for k, v in f.items() if v}


@app.route("/recorte")
def recorte():
    conn = obter_conn()
    f = _filtros_recorte()
    r = db.recorte(conn, f)
    omitir = tuple(k for k in ("situacao", "ccusto", "classificacao", "localizacao", "idade", "ano", "faixa", "pessoa")
                   if f.get(k) and f.get(k) not in ("imoveis", "sem-imoveis"))
    nomes = {"idade": dict(db.FAIXAS_IDADE).get(f.get("idade")), "faixa": dict(db.FAIXAS_VALOR).get(f.get("faixa"))}
    termo_de = None
    if f.get("pessoa"):
        termo_de = ("individual", f["pessoa"], db.situacao_termo(conn, "individual", f["pessoa"], db.bens_da_pessoa(conn, f["pessoa"])))
    elif f.get("ccusto") and f["ccusto"] != "-" and db.responsavel(conn, f["ccusto"]):
        termo_de = ("ccusto", f["ccusto"], db.situacao_termo(conn, "ccusto", f["ccusto"], db.bens_do_centro(conn, f["ccusto"])))
    opcoes = {
        "situacoes": [r[0] for r in conn.execute("SELECT DISTINCT situacao FROM bens ORDER BY 1")],
        "classificacoes": [r[0] for r in conn.execute("SELECT DISTINCT classificacao FROM bens WHERE classificacao <> '' ORDER BY 1")],
        "localizacoes": [r[0] for r in conn.execute("SELECT DISTINCT localizacao FROM bens WHERE localizacao <> '' ORDER BY 1")],
        "centros": [c["ccustos"] for c in db.centros(conn)], "pessoas": db.pessoas(conn), "idades": db.FAIXAS_IDADE,
    }
    return render_template("recorte.html", f=f, r=r, cards=painel.cards_graficos(r["dimensoes"], f, omitir), termo_de=termo_de,
                           descricao=painel.descrever(f, nomes), opcoes=opcoes, moeda=painel.moeda,
                           trilha=[("Recorte", None)])


@app.route("/recorte/xlsx")
def recorte_xlsx():
    return _baixar(db.exportar_recorte(obter_conn(), _filtros_recorte(), io.BytesIO()), "recorte.xlsx")
```

- [ ] **Step 4: Template**

`templates/recorte.html`:

```jinja
{% extends "base.html" %}
{% from "_macros.html" import select, grafico, tag_termo, cabecalho_tabela %}
{% block titulo %}Recorte{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-3">
  <h1 class="mb-0">Recorte</h1>
  <div class="ml-auto">
    {% if termo_de %}<a class="br-button primary" href="{{ url_for('termo', tipo=termo_de[0], chave=termo_de[1]) }}"><i class="fas fa-file-alt mr-1" aria-hidden="true"></i>Termo de {{ termo_de[1] }}</a> {{ tag_termo(termo_de[2]) }}{% endif %}
    <a class="br-button secondary ml-2" href="{{ url_for('recorte_xlsx', **f) }}"><i class="fas fa-file-excel mr-1" aria-hidden="true"></i>Exportar .xlsx</a>
  </div>
</div>
<p class="text-gray-70">{{ descricao }}</p>

<form method="get" class="br-card mb-4"><div class="card-content row">
  <div class="col-md-3 mb-3">{{ select('situacao', 'Situação', [('', 'Todas')] + opcoes.situacoes, selecionado=f.get('situacao', ''), obrigatorio=False) }}</div>
  <div class="col-md-3 mb-3">{{ select('ccusto', 'Centro de custo', [('', 'Todos'), ('-', 'Sem centro')] + opcoes.centros, selecionado=f.get('ccusto', ''), obrigatorio=False) }}</div>
  <div class="col-md-3 mb-3">{{ select('pessoa', 'Pessoa', [('', 'Todas')] + opcoes.pessoas, selecionado=f.get('pessoa', ''), obrigatorio=False) }}</div>
  <div class="col-md-3 mb-3">{{ select('localizacao', 'Localização', [('', 'Todas')] + opcoes.localizacoes, selecionado=f.get('localizacao', ''), obrigatorio=False) }}</div>
  <div class="col-md-3 mb-3">{{ select('classificacao', 'Classificação', [('', 'Todas'), ('imoveis', 'Imóveis'), ('sem-imoveis', 'Sem imóveis')] + opcoes.classificacoes, selecionado=f.get('classificacao', ''), obrigatorio=False) }}</div>
  <div class="col-md-3 mb-3">{{ select('idade', 'Idade', [('', 'Todas')] + opcoes.idades, selecionado=f.get('idade', ''), obrigatorio=False) }}</div>
  <div class="col-md-3 mb-3"><div class="br-input"><label for="valor_de">Valor de (R$)</label><input id="valor_de" name="valor_de" type="text" inputmode="decimal" value="{{ f.get('valor_de', '') }}"/></div></div>
  <div class="col-md-3 mb-3"><div class="br-input"><label for="valor_ate">Valor até (R$)</label><input id="valor_ate" name="valor_ate" type="text" inputmode="decimal" value="{{ f.get('valor_ate', '') }}"/></div></div>
  <div class="col-md-3 mb-3"><div class="br-input"><label for="entrada_de">Entrada de</label><input id="entrada_de" name="entrada_de" type="date" value="{{ f.get('entrada_de', '') }}"/></div></div>
  <div class="col-md-3 mb-3"><div class="br-input"><label for="entrada_ate">Entrada até</label><input id="entrada_ate" name="entrada_ate" type="date" value="{{ f.get('entrada_ate', '') }}"/></div></div>
  {% if f.get('ano') %}<input type="hidden" name="ano" value="{{ f.ano }}"/>{% endif %}
  {% if f.get('faixa') %}<input type="hidden" name="faixa" value="{{ f.faixa }}"/>{% endif %}
  <div class="col-md-6 mb-3 d-flex align-items-end">
    <button class="br-button primary" type="submit"><i class="fas fa-filter mr-1" aria-hidden="true"></i>Aplicar</button>
    <a class="br-button ml-2" href="{{ url_for('recorte') }}">Limpar</a>
  </div>
</div></form>

<div class="row mb-2">
  <div class="col-sm-6 col-md-3 mb-3"><div class="br-card h-100"><div class="card-content"><div class="text-up-03 text-weight-bold">{{ r.quantidade }}</div><div class="text-down-01 text-gray-70">bens no recorte</div></div></div></div>
  <div class="col-sm-6 col-md-3 mb-3"><div class="br-card h-100"><div class="card-content"><div class="text-up-03 text-weight-bold">{{ moeda(r.valor_total) }}</div><div class="text-down-01 text-gray-70">valor atual</div></div></div></div>
</div>

<div class="row">{% for g in cards %}{{ grafico(g) }}{% endfor %}</div>

{% if r.truncado %}
<div class="br-message warning mb-3"><div class="icon"><i class="fas fa-exclamation-triangle fa-lg" aria-hidden="true"></i></div>
  <div class="content" role="alert"><span class="message-title">Muitos bens.</span><span class="message-body"> A tabela mostra os {{ r.bens|length }} primeiros de {{ r.quantidade }}; refine o recorte ou exporte o .xlsx completo.</span></div></div>
{% endif %}
{{ cabecalho_tabela('Bens do recorte', 'bens') }}
  <thead><tr><th scope="col">Número</th><th scope="col">Descrição</th><th scope="col">Complemento</th><th scope="col">Localização</th><th scope="col">Centro</th><th scope="col">Pessoa</th><th scope="col">Classificação</th><th scope="col">Entrada</th><th scope="col" class="dsgov-numero">Valor</th></tr></thead>
  <tbody>
  {% for b in r.bens %}
  <tr><td><a href="{{ url_for('bem', numero=b.numero) }}">{{ b.numero }}</a></td><td>{{ b.descricao }}</td><td>{{ b.complemento or '' }}</td><td>{{ b.localizacao or '' }}</td><td>{{ b.ccustos or '' }}</td><td>{{ b.pessoa or '' }}</td><td>{{ b.classificacao or '' }}</td><td>{{ b.data_entrada or '' }}</td><td class="dsgov-numero">{{ moeda(b.valor_atual) }}</td></tr>
  {% else %}<tr><td colspan="9">Nenhum bem neste recorte.</td></tr>
  {% endfor %}
  </tbody>
</table></div>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='dsgov/vendor/echarts/echarts.min.js') }}"></script>
<script src="{{ url_for('static', filename='dsgov/js/echarts-dsgov.js') }}"></script>
{% endblock %}
```

Observação: o `select` de situação envia `situacao=` vazio quando "Todas" está marcada, e `_filtros_recorte` trata "campo presente e vazio" como todas. Quando nenhum radio está marcado (primeira visita), o campo não vai na query e o padrão ATIVO vale.

- [ ] **Step 5: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam (118).

- [ ] **Step 6: Commit**

```bash
git add app.py templates/recorte.html tests/test_app.py
git commit -m "Tela Recorte: filtros combináveis, gráficos com drill-down, termo do centro/pessoa e exportação"
```

---

### Task 14: README, conferência no navegador e publicação

**Files:**
- Modify: `README.md`
- Verify: site em https://patrimonio.sistemascfc.org

- [ ] **Step 1: README**

Na seção "Uso", acrescente após o item da Pesquisa:

```markdown
   **Processos SEI** (Cadastros → Processos SEI): um vigente por tipo de termo (centro de custo,
   individual, devolução). Sem processo vigente o termo não pode ser copiado nem baixado.
   **Termos emitidos**: cada cópia ou download registra data, processo e a lista de bens daquele
   momento (foto). Na tela do termo aparece o último registro e se entraram/saíram bens desde então
   (termo *desatualizado*). O número do documento SEI pode ser anotado depois, no registro.
   **Atualizar base** guarda o que mudou a cada importação (novos, removidos, movidos, situação) e a
   ficha do bem mostra o histórico dele.
   **Início** é o painel: cards e gráficos por situação, centro, classificação, localização, idade, ano e
   faixa de valor, todos clicáveis. **Recorte** filtra bens por qualquer combinação, com os mesmos
   gráficos, o termo do centro/pessoa quando couber e *Exportar .xlsx*.
```

Na tabela "Arquivos", acrescente:

```markdown
| `painel.py`, `graficos.py` | cards de gráfico (ECharts embutido, tema DSGov) |
```

- [ ] **Step 2: Suíte inteira e site**

```bash
.venv/bin/pytest -q
docker compose up -d --build
S=/tmp/claude-1003/-opt-web-termos-responsabilidade/4931bceb-8d69-41e8-8c4d-c68a9319e2fa/scratchpad
SENHA=$(cat $S/senha.txt); U=https://patrimonio.sistemascfc.org
for p in / /recorte "/recorte?ccusto=CFC" /termos-emitidos /cadastros/processos /upload; do
  echo "$p $(curl -s -o /dev/null -w '%{http_code}' -u patrimonio:$SENHA "$U$p")"; done
curl -s -u patrimonio:$SENHA "$U/" | grep -c 'data-grafico='
```

Expected: 118 testes passam; todas as rotas 200; 8 gráficos na tela inicial. Abra o site no navegador e confira os gráficos desenhados, clique numa barra de centro e veja o recorte, copie um termo (com processo cadastrado) e veja o registro em Termos emitidos.

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "README: processos SEI, termos emitidos, importações, painel e recorte"
```

---

## Self-review (feito ao escrever)

- **Cobertura da spec**: 2.1–2.4 → Tasks 1–6; 3.1–3.3 → Tasks 7–8; 4 (gráficos, painel, recorte) → Tasks 9–13; 5 (menu) → Tasks 5 e 13; 6 (testes) → cada task. "Baixar planilha não registra" → Task 4. Devolução registra → Task 4. Renomear leva histórico → Task 2.
- **Nomes consistentes**: rotas `termos_emitidos_tela`, `termo_emitido_tela`, `termo_emitido_documento`, `importacao_tela`, `recorte`, `recorte_xlsx`, `termo_registrar`; funções `db.situacoes_centros/pessoas`, `db.dimensoes`, `db.painel`, `db.recorte`, `db.exportar_recorte`, `painel.cards_graficos`, `painel.url_recorte`, `painel.descrever`, `painel.moeda`; macros `select`, `tag_termo`, `grafico`, `cabecalho_tabela`.
- **Placeholders**: nenhum; as rotas provisórias (Tasks 4 e 11) são explicitamente substituídas (Tasks 5 e 13).
