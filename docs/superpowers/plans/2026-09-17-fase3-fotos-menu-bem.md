# Fase 3 do inventário: várias fotos por bem, chaves simples no R2, submenus e fotos no cadastro do bem — plano

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** um bem pode ter várias fotos por evento de inventário, gravadas no R2 com a chave
`{pasta}/{nfoto}-{numero}.webp` (pasta = nome do evento minúsculo sem acento), com as fotos visíveis na sala,
no relatório (primeira + "+N"), no cadastro do bem (agrupadas por evento) e com o menu "Inventário" virando um
grupo com as telas do evento aberto.

**Architecture:** tabela nova `inventario_fotos` (evento, número, nfoto, url) com chave estrangeira composta
para `inventario_leituras` e cascata; a coluna `inventario_leituras.foto_url` migra para a tabela nova e some.
As consultas existentes continuam devolvendo `foto_url` (primeira foto) mais `n_fotos` via subconsultas
correlacionadas, então relatório, xlsx e filtros não mudam. `fotos.py` só monta chaves e fala com o bucket;
`inventario.py` recebe um callable `enviar` para continuar sem rede e testável.

**Tech Stack:** Python 3.12, Flask, SQLite (≥ 3.35 para `DROP COLUMN`), openpyxl, boto3 (só em produção),
Jinja2 + DSGov 3.7 (core.min.js), pytest.

**Spec:** `docs/superpowers/specs/2026-09-17-fase3-fotos-menu-bem-design.md` (ler antes de qualquer tarefa).

## Global Constraints

- Tudo é acréscimo: `bens` continua espelho do SPW e nunca muda aqui. Nenhum passo manual no bucket.
- Chaves no bucket: `{pasta}/{nfoto}-{numero}.webp` e `{pasta}/sobra-{id}.webp`; `pasta` = NFD, sem marcas
  combinantes, casefold, só `[a-z0-9]`; vazia → `evento{id}`. O prefixo `inventario/` deixa de existir para
  fotos novas; URLs antigas continuam válidas e `fotos.apagar` reconhece os dois formatos.
- `nfoto` = `MAX(nfoto) + 1` por `(evento, numero)`, nunca reaproveitado.
- Módulo de dados (`inventario.py`, `fotos.py`) não importa Flask. Testes sem rede: `fotos.enviar`/`fotos.apagar`/
  `fotos._cliente` sempre substituídos por monkeypatch ou callable falso.
- Textos de interface em português, no tom das telas existentes. Mensagens de erro via `db.ErroDeNegocio`.
- Rodar `.venv/bin/pytest -q` ao fim de cada tarefa; a suíte inteira precisa passar antes do commit.
- Branch `fase3-fotos` a partir de `main` (`94c7a48` ou posterior). Commits atômicos por tarefa, mensagem em
  português, terminando com `Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>`.
- Desvio consciente da spec §6.1: a tabela "Trazidos de outras salas" não tem coluna Foto hoje e continua sem
  (não é criada nesta fase). Só "Bens da sala" mostra as miniaturas.

---

## Mapa de arquivos

| arquivo | responsabilidade nesta fase |
|---|---|
| `fotos.py` | `pasta`, `chave_bem`, `chave_sobra`, `enviar(chave, dados)`, `apagar(url)` pelos dois formatos |
| `db.py` | `ESQUEMA` (tabela `inventario_fotos`; `inventario_leituras` sem `foto_url`), migração em `criar_esquema`, `_ler_aba_cadastro(opcionais=)`, `importar_cadastros` (aba `inv_fotos` opcional) |
| `inventario.py` | `pasta_do_evento`, `fotos_do_bem_no_evento`, `adicionar_foto`, `apagar_foto`, `fotos_do_bem`, `desfazer_leituras` via `inventario_fotos`, `_LEITURA`/`_CAMPOS_REL` com subconsultas, `bens_da_sala` com `fotos`, `abrir_evento` com pasta única, `ABAS`/`validar_abas`/`substituir_tabelas` com `inv_fotos` |
| `app_inventario.py` | rotas `foto_leitura` (devolve lista), `foto_excluir(id, numero, nfoto)`, `sobra` com `chave_sobra` |
| `app.py` | `MENU` com filhos; rota `bem` passa `fotos` |
| `templates/base.html` | grupo do menu |
| `templates/inventario_sala.html` | várias miniaturas + câmera sempre; JS `preencherFoto(td, fotos, numero)` |
| `templates/inventario_relatorio.html` | "+N" ao lado da miniatura |
| `templates/bem.html` | card "Fotos do inventário" |
| `tests/test_fotos.py`, `tests/test_inventario.py`, `tests/test_app.py` | cobertura de tudo acima |
| `README.md` | chave das fotos, 7 abas, grupo do menu |

---

### Task 1: chaves simples no bucket (`fotos.py`)

**Files:**
- Modify: `fotos.py` (funções `enviar`, `apagar`, `_carimbo`, `nome_bem`, `nome_sobra`, constante `PREFIXO`)
- Test: `tests/test_fotos.py`

**Interfaces:**
- Produces: `fotos.pasta(nome: str, evento_id: int) -> str`; `fotos.chave_bem(pasta: str, nfoto: int, numero: int) -> str`;
  `fotos.chave_sobra(pasta: str, sobra_id: int) -> str`; `fotos.enviar(chave: str, dados: bytes) -> str` (grava a
  chave exatamente como recebida); `fotos.apagar(url: str | None) -> None`.
- Removes: `fotos.PREFIXO`, `fotos._carimbo`, `fotos.nome_bem`, `fotos.nome_sobra` (as rotas que os usam são
  trocadas na Task 5; até lá `tests/test_app.py::test_inventario_fotos_e_sobras` fica vermelho — esperado e
  documentado no commit desta tarefa; a Task 5 o corrige).

- [ ] **Step 1: Escrever os testes novos e ajustar o existente**

Em `tests/test_fotos.py`, substituir `test_enviar_apagar_e_configuracao` por:

```python
def test_pasta_e_chaves():
    assert fotos.pasta("Inventário 2026", 1) == "inventario2026"
    assert fotos.pasta("  INVENTÁRIO – Sede / Anexo 2 ", 1) == "inventariosedeanexo2"
    assert fotos.pasta("Ação Çedilha", 1) == "acaocedilha"
    assert fotos.pasta("???", 7) == "evento7"
    assert fotos.pasta("", 7) == "evento7"
    assert fotos.chave_bem("inventario2026", 1, 12334) == "inventario2026/1-12334.webp"
    assert fotos.chave_sobra("inventario2026", 7) == "inventario2026/sobra-7.webp"


def test_enviar_apagar_e_configuracao(monkeypatch):
    for v in fotos.VARIAVEIS:
        monkeypatch.delenv(v, raising=False)
    assert fotos.configurado() is False
    with pytest.raises(db.ErroDeNegocio):
        fotos.enviar("x.webp", b"...")
    monkeypatch.setenv("R2_ACCESS_KEY_ID", "k"); monkeypatch.setenv("R2_SECRET_ACCESS_KEY", "s")
    monkeypatch.setenv("R2_ENDPOINT_URL", "https://acc.r2.cloudflarestorage.com"); monkeypatch.setenv("R2_BUCKET_NAME", "fotos")
    monkeypatch.setenv("R2_PUBLIC_URL", "https://fotos.exemplo.org/")
    falso = ClienteFalso()
    monkeypatch.setattr(fotos, "_cliente", lambda: falso)
    assert fotos.configurado()
    url = fotos.enviar("inventario2026/1-1001.webp", b"webp")
    assert url == "https://fotos.exemplo.org/inventario2026/1-1001.webp"
    assert falso.enviados == [("fotos", "inventario2026/1-1001.webp", 4, "image/webp")]
    fotos.apagar(url)
    assert falso.apagados == [("fotos", "inventario2026/1-1001.webp")]
    # formato antigo (fotos gravadas antes da Fase 3) continua reconhecido
    fotos.apagar("https://fotos.exemplo.org/inventario/INV1_BEM_1001_20260915120000.webp")
    assert falso.apagados[-1] == ("fotos", "inventario/INV1_BEM_1001_20260915120000.webp")
    fotos.apagar("https://outro/sem-prefixo.webp")                       # ignora
    fotos.apagar(None)
    assert len(falso.apagados) == 2
    monkeypatch.delenv("R2_PUBLIC_URL")
    assert fotos.enviar("a/1-2.webp", b"1") == "https://acc.r2.cloudflarestorage.com/fotos/a/1-2.webp"
    fotos.apagar("https://acc.r2.cloudflarestorage.com/fotos/a/1-2.webp")
    assert falso.apagados[-1] == ("fotos", "a/1-2.webp")
    assert not hasattr(fotos, "nome_bem") and not hasattr(fotos, "PREFIXO")
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_fotos.py -q`
Expected: FAIL (`AttributeError: module 'fotos' has no attribute 'pasta'`).

- [ ] **Step 3: Implementar**

Em `fotos.py`: apagar `PREFIXO`, `_carimbo`, `nome_bem`, `nome_sobra` e o `from datetime import datetime`;
acrescentar `import re` e `import unicodedata` no topo; trocar `enviar`/`apagar` e acrescentar as três funções:

```python
def pasta(nome: str, evento_id: int) -> str:
    """Pasta das fotos do evento no bucket: nome em minúsculas, sem acento, só [a-z0-9]. Vazio → evento<id>."""
    base = unicodedata.normalize("NFD", nome or "")
    limpo = re.sub(r"[^a-z0-9]", "", "".join(c for c in base if not unicodedata.combining(c)).casefold())
    return limpo or f"evento{evento_id}"


def chave_bem(pasta: str, nfoto: int, numero: int) -> str:
    return f"{pasta}/{nfoto}-{numero}.webp"


def chave_sobra(pasta: str, sobra_id: int) -> str:
    return f"{pasta}/sobra-{sobra_id}.webp"


def enviar(chave: str, dados: bytes) -> str:
    """Grava `chave` no bucket exatamente como recebida e devolve a URL pública."""
    if not configurado():
        raise ErroDeNegocio("Fotos desativadas: bucket não configurado.")
    _cliente().put_object(Bucket=os.environ["R2_BUCKET_NAME"], Key=chave, Body=dados, ContentType="image/webp")
    return _url(chave)


_PREFIXO_ANTIGO = "inventario/"


def apagar(url: str | None) -> None:
    """Apaga o objeto pela chave contida na URL: URL nova = base pública + chave; URL antiga (antes da
    Fase 3) tem "inventario/" no meio. Erro do bucket é ignorado (a URL some do banco de qualquer jeito)."""
    if not url or not configurado():
        return
    base = _url("")
    if url.startswith(base):
        chave = url[len(base):]
    else:
        pos = url.find(_PREFIXO_ANTIGO)
        if pos < 0:
            return
        chave = url[pos:]
    if not chave:
        return
    try:
        _cliente().delete_object(Bucket=os.environ["R2_BUCKET_NAME"], Key=chave)
    except Exception:
        pass
```

Atualizar a docstring do módulo: trocar a frase sobre `inventario/<nome>` por "Chave `{pasta}/{nfoto}-{numero}.webp`
(pasta = nome do evento normalizado); fotos anteriores à Fase 3 ficam em `inventario/INV...`".

- [ ] **Step 4: Rodar**

Run: `.venv/bin/pytest tests/test_fotos.py -q`
Expected: PASS. Rodar também `.venv/bin/pytest -q`: só `tests/test_app.py::test_inventario_fotos_e_sobras`
falha (usa `fotos.nome_bem`); é corrigido na Task 5.

- [ ] **Step 5: Commit**

```bash
git add fotos.py tests/test_fotos.py
git commit -m "Fotos: chave {pasta}/{nfoto}-{numero}.webp com pasta pelo nome do evento; apagar reconhece URL nova e antiga

test_app::test_inventario_fotos_e_sobras fica vermelho até a rota mudar (Task 5).

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 2: tabela `inventario_fotos` e migração de `foto_url` (`db.py`)

**Files:**
- Modify: `db.py` (`ESQUEMA` — bloco de `inventario_leituras` e após `inventario_bens_encerrados`; `criar_esquema`)
- Test: `tests/test_inventario.py`

**Interfaces:**
- Produces: tabela `inventario_fotos(evento_id, numero, nfoto, url, criado_em)` PK `(evento_id, numero, nfoto)`,
  FK composta para `inventario_leituras(evento_id, numero)` ON DELETE CASCADE; `inventario_leituras` sem `foto_url`.
- Nota: até a Task 3, `inventario.atualizar_leitura(foto_url=...)`, `_LEITURA` e `substituir_tabelas` ainda
  referenciam `foto_url` → vários testes ficam vermelhos entre esta tarefa e a Task 4. Por isso **Tasks 2, 3 e 4
  formam um commit só** (ver Step 5 da Task 4). Nesta tarefa só se roda o teste novo.

- [ ] **Step 1: Teste da migração**

Acrescentar ao fim de `tests/test_inventario.py`:

```python
def test_migracao_foto_url_para_inventario_fotos(dados):
    """Banco anterior à Fase 3: inventario_leituras tinha foto_url. criar_esquema move para inventario_fotos
    (nfoto 1), apaga a coluna e é idempotente."""
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.ler(dados, eid, "01 - SALA CCI", 1002, "Fulano")
    dados.execute("ALTER TABLE inventario_leituras ADD COLUMN foto_url TEXT")
    dados.execute("UPDATE inventario_leituras SET foto_url = 'https://x/inventario/INV1_BEM_1001_1.webp' WHERE numero = 1001")
    dados.commit()
    db.criar_esquema(dados)
    assert "foto_url" not in db._colunas(dados, "inventario_leituras")
    assert dados.execute("SELECT evento_id, numero, nfoto, url FROM inventario_fotos").fetchall() == [(eid, 1001, 1, "https://x/inventario/INV1_BEM_1001_1.webp")]
    db.criar_esquema(dados)                                                          # de novo: nada muda
    assert dados.execute("SELECT count(*) FROM inventario_fotos").fetchone()[0] == 1
    dados.execute("DELETE FROM inventario_leituras WHERE numero = 1001")            # cascata
    assert dados.execute("SELECT count(*) FROM inventario_fotos").fetchone()[0] == 0
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_inventario.py::test_migracao_foto_url_para_inventario_fotos -q`
Expected: FAIL (`no such table: inventario_fotos`).

- [ ] **Step 3: Implementar**

Em `db.ESQUEMA`, remover a linha `  foto_url    TEXT,` do `CREATE TABLE IF NOT EXISTS inventario_leituras` (a
linha anterior `observacao  TEXT,` passa a ser `observacao  TEXT,` seguida de `UNIQUE (evento_id, numero)`; cuidado
com a vírgula). Depois do bloco de `inventario_bens_encerrados`, acrescentar:

```sql
CREATE TABLE IF NOT EXISTS inventario_fotos (
  evento_id INTEGER NOT NULL,
  numero    INTEGER NOT NULL,
  nfoto     INTEGER NOT NULL,
  url       TEXT NOT NULL,
  criado_em TEXT NOT NULL,
  PRIMARY KEY (evento_id, numero, nfoto),
  FOREIGN KEY (evento_id, numero) REFERENCES inventario_leituras(evento_id, numero) ON DELETE CASCADE
);
```

Em `criar_esquema`, antes do `conn.commit()`:

```python
    # Fase 3 (2026-09-17): a foto única da leitura virou a tabela inventario_fotos (várias por bem).
    if "foto_url" in _colunas(conn, "inventario_leituras"):
        conn.execute("""INSERT OR IGNORE INTO inventario_fotos (evento_id, numero, nfoto, url, criado_em)
                        SELECT evento_id, numero, 1, foto_url, lido_em FROM inventario_leituras
                        WHERE foto_url IS NOT NULL AND foto_url <> ''""")
        conn.execute("ALTER TABLE inventario_leituras DROP COLUMN foto_url")
```

Conferir a versão do SQLite: `.venv/bin/python -c "import sqlite3; print(sqlite3.sqlite_version)"` deve ser ≥ 3.35.0
(o projeto já usa `DROP COLUMN` para `responsaveis.tratamento`).

- [ ] **Step 4: Rodar o teste novo**

Run: `.venv/bin/pytest tests/test_inventario.py::test_migracao_foto_url_para_inventario_fotos -q`
Expected: PASS. (A suíte inteira ainda falha: `foto_url` sumiu e `inventario.py` ainda a usa — segue para a Task 3.)

- [ ] **Step 5: Não commitar ainda** — o commit é conjunto ao fim da Task 4.

---

### Task 3: funções de dados das fotos (`inventario.py`)

**Files:**
- Modify: `inventario.py` (imports; `abrir_evento`; `_LEITURA`; `atualizar_leitura`; `bens_da_sala`; `desfazer_leituras`; `_CAMPOS_REL`; novas funções após `atualizar_leitura`)
- Test: `tests/test_inventario.py`

**Interfaces:**
- Consumes: `fotos.pasta`, `fotos.chave_bem` (Task 1); tabela `inventario_fotos` (Task 2).
- Produces:
  - `pasta_do_evento(conn, evento_id: int) -> str`
  - `fotos_do_bem_no_evento(conn, evento_id: int, numero: int) -> list[dict]` — `[{"nfoto", "url", "criado_em"}]` por nfoto
  - `adicionar_foto(conn, evento_id: int, numero: int, enviar) -> list[dict]` — `enviar(chave: str) -> str`
  - `apagar_foto(conn, evento_id: int, numero: int, nfoto: int) -> str | None`
  - `fotos_do_bem(conn, numero: int) -> list[dict]` — `[{"evento_id", "evento", "aberto_em", "encerrado_em", "lido_em", "fotos": [{"nfoto", "url"}]}]`
  - `desfazer_leituras` mantém a assinatura `(conn, evento_id, numeros) -> tuple[list, int]`, urls de todas as fotos
  - linhas de `bens_da_sala`/`relatorio` ganham `n_fotos: int`; `foto_url` = primeira foto ou `None`; itens de `bens_da_sala()["bens"]` e `["trazidos"]` ganham `fotos: list[dict]`
  - `atualizar_leitura` deixa de aceitar `foto_url`

- [ ] **Step 1: Ajustar os testes existentes que usam `foto_url` e escrever os novos**

Em `tests/test_inventario.py`, acrescentar após `semear_inventario`:

```python
def foto_falsa(conn, eid, numero, url=None):
    """adicionar_foto com envio falso; devolve a lista de fotos do bem no evento."""
    return inventario.adicionar_foto(conn, eid, numero, lambda chave: url or f"https://x/{chave}")
```

Trocar, em `test_xlsx_cabecalho_filtros_e_fotos`, a linha
`inventario.atualizar_leitura(dados, eid, 1001, foto_url="https://x/1001.webp")` por
`foto_falsa(dados, eid, 1001, "https://x/1001.webp")`.

Trocar, em `test_ler_lote_e_desfazer_leituras`, a linha
`inventario.atualizar_leitura(dados, eid, 1001, foto_url="http://x/1001.webp")` por
`foto_falsa(dados, eid, 1001, "http://x/1001.webp")` e a asserção seguinte por
`assert inventario.desfazer_leituras(dados, eid, [1001, 2001, 1004]) == (["http://x/1001.webp"], 2)` (igual; só confirmando que continua).

Acrescentar ao fim do arquivo:

```python
def test_adicionar_e_apagar_fotos_do_bem(dados):
    eid = semear_inventario(dados)
    assert inventario.pasta_do_evento(dados, eid) == "inventario2026"
    with pytest.raises(db.ErroDeNegocio):                                            # sem leitura
        foto_falsa(dados, eid, 1001)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    chaves = []
    f1 = inventario.adicionar_foto(dados, eid, 1001, lambda c: chaves.append(c) or "https://x/" + c)
    f2 = inventario.adicionar_foto(dados, eid, 1001, lambda c: chaves.append(c) or "https://x/" + c)
    assert chaves == ["inventario2026/1-1001.webp", "inventario2026/2-1001.webp"]
    assert [f["nfoto"] for f in f1] == [1] and [(f["nfoto"], f["url"]) for f in f2] == [(1, "https://x/inventario2026/1-1001.webp"), (2, "https://x/inventario2026/2-1001.webp")]
    assert inventario.fotos_do_bem_no_evento(dados, eid, 1001) == f2 and f2[0]["criado_em"]
    # envio falhou: nada gravado
    with pytest.raises(RuntimeError):
        inventario.adicionar_foto(dados, eid, 1001, lambda c: (_ for _ in ()).throw(RuntimeError("bucket")))
    assert len(inventario.fotos_do_bem_no_evento(dados, eid, 1001)) == 2
    # apagar a 2 e tirar outra → 3 (número nunca reaproveitado)
    assert inventario.apagar_foto(dados, eid, 1001, 2) == "https://x/inventario2026/2-1001.webp"
    assert inventario.apagar_foto(dados, eid, 1001, 2) is None
    f3 = foto_falsa(dados, eid, 1001)
    assert [f["nfoto"] for f in f3] == [1, 3] and chaves[-1] == "inventario2026/2-1001.webp"   # chaves só tem os 2 primeiros envios
    # bens_da_sala/relatorio: primeira foto + contagem; lista completa na sala
    b = {x["numero"]: x for x in inventario.bens_da_sala(dados, eid, "01 - SALA CCI")["bens"]}
    assert b[1001]["foto_url"] == "https://x/inventario2026/1-1001.webp" and b[1001]["n_fotos"] == 2 and [f["nfoto"] for f in b[1001]["fotos"]] == [1, 3]
    assert b[1002]["foto_url"] is None and b[1002]["n_fotos"] == 0 and b[1002]["fotos"] == []
    r = {x["numero"]: x for x in inventario.relatorio(dados, eid)}
    assert r[1001]["n_fotos"] == 2 and r[1001]["foto_url"].endswith("/1-1001.webp") and inventario.contar_fotos(r.values()) == 1
    assert [x["numero"] for x in inventario.relatorio(dados, eid, foto="com")] == [1001]
    # trazido de outra sala também traz a lista
    inventario.ler(dados, eid, "01 - SALA CCI", 2001, "Fulano")
    foto_falsa(dados, eid, 2001)
    t = inventario.bens_da_sala(dados, eid, "01 - SALA CCI")["trazidos"]
    assert [f["nfoto"] for f in t[0]["fotos"]] == [1]
    # desfazer leva todas as urls; cascata limpa a tabela
    urls, n = inventario.desfazer_leituras(dados, eid, [1001])
    assert sorted(urls) == ["https://x/inventario2026/1-1001.webp", "https://x/inventario2026/3-1001.webp"] and n == 1
    assert inventario.fotos_do_bem_no_evento(dados, eid, 1001) == []
    with pytest.raises(db.ErroDeNegocio):
        inventario.atualizar_leitura(dados, eid, 2001, foto_url="x")                 # campo saiu
    inventario.encerrar_evento(dados, eid)
    with pytest.raises(db.ErroDeNegocio):
        foto_falsa(dados, eid, 2001)
    with pytest.raises(db.ErroDeNegocio):
        inventario.apagar_foto(dados, eid, 2001, 1)


def test_fotos_do_bem_agrupadas_por_evento(dados):
    eid1 = semear_inventario(dados)
    inventario.ler(dados, eid1, "01 - SALA CCI", 1001, "Fulano")
    foto_falsa(dados, eid1, 1001); foto_falsa(dados, eid1, 1001)
    inventario.encerrar_evento(dados, eid1)
    eid2 = inventario.abrir_evento(dados, "Inventário 2027", None, ["Fulano"])
    inventario.ler(dados, eid2, "01 - SALA CCI", 1001, "Fulano")
    foto_falsa(dados, eid2, 1001)
    inventario.ler(dados, eid2, "01 - SALA CCI", 1002, "Fulano")                       # lido sem foto: não aparece
    g = inventario.fotos_do_bem(dados, 1001)
    assert [x["evento"] for x in g] == ["Inventário 2027", "Inventário 2026"]           # mais recente primeiro
    assert [[f["nfoto"] for f in x["fotos"]] for x in g] == [[1], [1, 2]]
    assert g[1]["encerrado_em"] and g[0]["encerrado_em"] is None and g[0]["lido_em"] and g[0]["evento_id"] == eid2
    assert g[0]["fotos"][0]["url"] == "https://x/inventario2027/1-1001.webp"
    assert inventario.fotos_do_bem(dados, 1002) == [] and inventario.fotos_do_bem(dados, 99999) == []


def test_abrir_evento_recusa_pasta_de_fotos_repetida(dados):
    eid = semear_inventario(dados)
    inventario.encerrar_evento(dados, eid)
    with pytest.raises(db.ErroDeNegocio) as e:
        inventario.abrir_evento(dados, "INVENTÁRIO 2026", None, ["Fulano"])            # mesma pasta: inventario2026
    assert "inventario2026" in str(e.value)
    assert inventario.abrir_evento(dados, "Inventário 2026 B", None, ["Fulano"])
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_inventario.py -q -k "fotos or lote or xlsx_cabecalho or pasta"`
Expected: FAIL (`AttributeError: ... adicionar_foto` e `no such column: r.foto_url`).

- [ ] **Step 3: Implementar**

No topo de `inventario.py`, acrescentar `import fotos` após `import db`.

`abrir_evento`: logo após `nome = _obrigatorio(nome, "Nome do evento")`:

```python
    nova = fotos.pasta(nome, 0)
    for outro in eventos(conn):
        if fotos.pasta(outro["nome"], outro["id"]) == nova:
            raise ErroDeNegocio(f"Já existe um evento com esse nome (pasta de fotos '{nova}'); escolha outro nome.")
```

`_LEITURA` e `_CAMPOS_REL`: trocar `r.foto_url` pelo par de subconsultas. Definir a constante `_FOTO_SQL` imediatamente acima de `_LEITURA` (seção `# --- leituras`) e usá-la nos dois:

```python
_FOTO_SQL = """(SELECT f.url FROM inventario_fotos f WHERE f.evento_id = r.evento_id AND f.numero = r.numero ORDER BY f.nfoto LIMIT 1) AS foto_url,
        (SELECT COUNT(*) FROM inventario_fotos f WHERE f.evento_id = r.evento_id AND f.numero = r.numero) AS n_fotos"""
_LEITURA = f"r.localizacao AS lido_em_sala, r.lido_em, r.integrante, r.conservacao, r.quem_usa, r.observacao, {_FOTO_SQL}"
```

e em `_CAMPOS_REL` trocar `r.observacao, r.foto_url,` por `r.observacao, {_FOTO_SQL},` (a string vira f-string).

`atualizar_leitura`: docstring "Campos: conservacao, quem_usa, observacao ..." e `permitidos = {"conservacao", "quem_usa", "observacao"}`.

`bens_da_sala`: antes do `return`, acrescentar:

```python
    for b in bens + trazidos:
        b["fotos"] = fotos_do_bem_no_evento(conn, evento_id, b["numero"]) if b["lido_em"] else []
```

`desfazer_leituras`: trocar a consulta das urls por

```python
    urls = [r[0] for r in conn.execute(
        f"SELECT url FROM inventario_fotos WHERE evento_id = ? AND numero IN ({marcas})", (evento_id, *numeros))]
```

(o DELETE das leituras apaga as fotos em cascata; `PRAGMA foreign_keys = ON` está em `db.conectar`).

Novas funções, logo após `atualizar_leitura`:

```python
# ---------------------------------------------------------------- fotos dos bens
def pasta_do_evento(conn, evento_id: int) -> str:
    e = _um(conn, "SELECT id, nome FROM inventario_eventos WHERE id = ?", evento_id)
    if not e:
        raise ErroDeNegocio("Evento de inventário não encontrado.")
    return fotos.pasta(e["nome"], e["id"])


def fotos_do_bem_no_evento(conn, evento_id: int, numero: int) -> list[dict]:
    return _todos(conn, "SELECT nfoto, url, criado_em FROM inventario_fotos WHERE evento_id = ? AND numero = ? ORDER BY nfoto",
                  evento_id, numero)


def adicionar_foto(conn, evento_id: int, numero: int, enviar) -> list[dict]:
    """Mais uma foto do bem neste evento. `enviar(chave) -> url` grava no bucket (fotos.enviar com os bytes já
    comprimidos, ou um callable falso nos testes) e roda ANTES do INSERT: se falhar, nada é gravado.
    nfoto = maior + 1, nunca reaproveitado. Devolve a lista atualizada."""
    _evento_aberto_ou_erro(conn, evento_id)
    if not _um(conn, "SELECT id FROM inventario_leituras WHERE evento_id = ? AND numero = ?", evento_id, numero):
        raise ErroDeNegocio("Leia o bem antes de fotografar.")
    nfoto = conn.execute("SELECT COALESCE(MAX(nfoto), 0) + 1 FROM inventario_fotos WHERE evento_id = ? AND numero = ?",
                         (evento_id, numero)).fetchone()[0]
    url = enviar(fotos.chave_bem(pasta_do_evento(conn, evento_id), nfoto, numero))
    conn.execute("INSERT INTO inventario_fotos (evento_id, numero, nfoto, url, criado_em) VALUES (?,?,?,?,?)",
                 (evento_id, numero, nfoto, url, _agora()))
    conn.commit()
    return fotos_do_bem_no_evento(conn, evento_id, numero)


def apagar_foto(conn, evento_id: int, numero: int, nfoto: int) -> str | None:
    """Apaga a linha e devolve a url (para a rota apagar no bucket); None se não existia."""
    _evento_aberto_ou_erro(conn, evento_id)
    f = _um(conn, "SELECT url FROM inventario_fotos WHERE evento_id = ? AND numero = ? AND nfoto = ?", evento_id, numero, nfoto)
    if not f:
        return None
    conn.execute("DELETE FROM inventario_fotos WHERE evento_id = ? AND numero = ? AND nfoto = ?", (evento_id, numero, nfoto))
    conn.commit()
    return f["url"]


def fotos_do_bem(conn, numero: int) -> list[dict]:
    """Para o cadastro do bem: um bloco por evento em que o bem tem foto, do mais recente para o mais antigo."""
    grupos = _todos(conn, """
        SELECT e.id AS evento_id, e.nome AS evento, e.aberto_em, e.encerrado_em, r.lido_em
        FROM inventario_eventos e JOIN inventario_leituras r ON r.evento_id = e.id AND r.numero = ?
        WHERE EXISTS (SELECT 1 FROM inventario_fotos f WHERE f.evento_id = e.id AND f.numero = r.numero)
        ORDER BY e.aberto_em DESC, e.id DESC""", numero)
    for g in grupos:
        g["fotos"] = [{"nfoto": f["nfoto"], "url": f["url"]} for f in fotos_do_bem_no_evento(conn, g["evento_id"], numero)]
    return grupos
```

- [ ] **Step 4: Rodar**

Run: `.venv/bin/pytest tests/test_inventario.py tests/test_fotos.py -q`
Expected: os testes de fotos, lote, xlsx e pasta passam. Ainda falham `test_planilha_*` e
`test_aba_inv_bens_encerrados_*` (a planilha ainda exporta `foto_url` de `inv_leituras`) — Task 4.

- [ ] **Step 5: Não commitar ainda** — commit conjunto ao fim da Task 4.

---

### Task 4: planilha de cadastros com aba `inv_fotos` (`inventario.py`, `db.py`)

**Files:**
- Modify: `inventario.py` (`ABAS`, `validar_abas`, `substituir_tabelas`); `db.py` (`_ler_aba_cadastro`, `importar_cadastros`)
- Test: `tests/test_inventario.py`

**Interfaces:**
- Produces: `ABAS["inv_leituras"]` sem `foto_url`; `ABAS["inv_fotos"] = ["evento_id", "numero", "nfoto", "url", "criado_em"]`
  (exportada por último); `db._ler_aba_cadastro(wb, tabela, problemas, colunas=None, opcional=False, opcionais=())`;
  `validar_abas` lê `r.get("foto_url")` de `inv_leituras` e preenche `linhas["inv_fotos"]` quando a aba está vazia.

- [ ] **Step 1: Ajustar testes existentes e escrever os novos**

Em `test_planilha_de_cadastros_exporta_e_importa_abas_de_inventario`:
- `assert wb.sheetnames == [..., "inv_sobras", "inv_bens_encerrados", "inv_fotos"]`
- a linha `wb["inv_leituras"].append([eid, 2001, "02 - SALA B", "2026-01-20 10:00:00", "Antigo", "Regular", "", "migrado", ""])`
  perde o último `""` (8 colunas agora).

Em `test_planilha_de_inventario_validacoes`: as duas linhas `wb["inv_leituras"].append([...])` perdem o último `""`.

Em `test_planilha_sem_abas_de_inventario_nao_toca_nas_tabelas` nada muda (remove todas as `ABAS`).

Acrescentar ao fim do arquivo:

```python
def test_aba_inv_fotos_exporta_importa_e_valida(dados, tmp_path):
    from openpyxl import load_workbook
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    foto_falsa(dados, eid, 1001); foto_falsa(dados, eid, 1001)
    inventario.apagar_foto(dados, eid, 1001, 1)                                        # fica só a 2
    wb = load_workbook(db.exportar_cadastros(dados, tmp_path / "c.xlsx"))
    assert wb.sheetnames[-1] == "inv_fotos"
    assert list(wb["inv_leituras"].iter_rows(values_only=True))[0] == tuple(inventario.ABAS["inv_leituras"]) and "foto_url" not in inventario.ABAS["inv_leituras"]
    linhas = list(wb["inv_fotos"].iter_rows(values_only=True))
    assert linhas[0] == ("evento_id", "numero", "nfoto", "url", "criado_em") and linhas[1][:4] == (eid, 1001, 2, "https://x/inventario2026/2-1001.webp") and len(linhas) == 2
    with open(tmp_path / "c.xlsx", "rb") as f:
        r = db.importar_cadastros(dados, f)
    assert r["inv_fotos"] == 1 and [x["nfoto"] for x in inventario.fotos_do_bem_no_evento(dados, eid, 1001)] == [2]
    assert [x["nfoto"] for x in foto_falsa(dados, eid, 1001)] == [2, 3]                # contador continua do maior
    # aba ausente + inv_leituras com foto_url (planilha anterior à Fase 3): vira foto 1
    wb.remove(wb["inv_fotos"])
    ws = wb["inv_leituras"]
    ws.cell(row=1, column=9, value="foto_url")
    ws.cell(row=2, column=9, value="https://x/inventario/INV1_BEM_1001_1.webp")
    wb.save(tmp_path / "antiga.xlsx")
    with open(tmp_path / "antiga.xlsx", "rb") as f:
        r = db.importar_cadastros(dados, f)
    assert r["inv_fotos"] == 1 and inventario.fotos_do_bem_no_evento(dados, eid, 1001) [0]["url"] == "https://x/inventario/INV1_BEM_1001_1.webp"
    # aba ausente e sem foto_url: tabela fica vazia (a planilha é a fonte de verdade)
    ws.cell(row=2, column=9, value=None)
    wb.save(tmp_path / "vazia.xlsx")
    with open(tmp_path / "vazia.xlsx", "rb") as f:
        r = db.importar_cadastros(dados, f)
    assert r["inv_fotos"] == 0 and inventario.fotos_do_bem_no_evento(dados, eid, 1001) == []
    # validações: leitura inexistente, nfoto inválido, repetido, url vazia, pasta repetida
    wb = load_workbook(tmp_path / "c.xlsx")
    wb["inv_fotos"].append([eid, 1002, 1, "https://x/a.webp", "2026-01-01 00:00:00"])      # 1002 não foi lido
    wb["inv_fotos"].append([eid, 1001, "x", "https://x/b.webp", "2026-01-01 00:00:00"])
    wb["inv_fotos"].append([eid, 1001, 2, "https://x/c.webp", "2026-01-01 00:00:00"])        # repete (eid, 1001, 2)
    wb["inv_fotos"].append([eid, 1001, 5, "", "2026-01-01 00:00:00"])
    wb["inv_eventos"].append([9, "INVENTARIO 2026", None, "2026-02-01 00:00:00", "2026-02-02 00:00:00"])
    wb.save(tmp_path / "ruim.xlsx")
    with open(tmp_path / "ruim.xlsx", "rb") as f, pytest.raises(db.ImportacaoInvalida) as ex:
        db.importar_cadastros(dados, f)
    msg = str(ex.value)
    assert "não tem leitura" in msg and "nfoto inválido" in msg and "repetida" in msg and "url vazia" in msg and "pasta de fotos repetida" in msg
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_inventario.py -q -k planilha or inv_fotos or inv_bens`
Expected: FAIL.

- [ ] **Step 3: Implementar**

`inventario.ABAS`:

```python
ABAS = {
    "inv_eventos": ["id", "nome", "descricao", "aberto_em", "encerrado_em"],
    "inv_integrantes": ["evento_id", "nome"],
    "inv_salas": ["evento_id", "localizacao"],
    "inv_leituras": ["evento_id", "numero", "localizacao", "lido_em", "integrante", "conservacao", "quem_usa", "observacao"],
    "inv_sobras": ["evento_id", "localizacao", "descricao", "complemento", "observacao", "foto_url", "integrante", "criado_em"],
    "inv_bens_encerrados": ["evento_id", "numero", "situacao", "descricao", "complemento", "classificacao", "localizacao"],
    "inv_fotos": ["evento_id", "numero", "nfoto", "url", "criado_em"],
}
ABAS_OPCIONAIS = ("inv_bens_encerrados", "inv_fotos")   # podem faltar mesmo quando as 5 originais vêm
```

`validar_abas`:
1. Depois do laço de `inv_eventos` e do teste `if abertos > 1`, checar pasta repetida:

```python
    pastas: dict = {}
    for eid, nome, *_ in linhas["inv_eventos"]:
        p = fotos.pasta(nome, eid)
        if p in pastas:
            problemas.append(f"inv_eventos: pasta de fotos repetida '{p}' (eventos {pastas[p]} e {eid}); mude um dos nomes")
        pastas.setdefault(p, eid)
```

2. No laço de `inv_leituras`: a tupla final perde `_texto(r["foto_url"]) or None`; guardar as leituras aceitas e a
   foto antiga:

```python
    leituras_ok: set = set()
    fotos_antigas: list = []
    ...
        vistos.add((eid, num))
        leituras_ok.add((eid, num))
        linhas["inv_leituras"].append((eid, num, loc, lido, integ, cons, _texto(r["quem_usa"]) or None, _texto(r["observacao"]) or None))
        if _texto(r.get("foto_url")):
            fotos_antigas.append((eid, num, 1, _texto(r["foto_url"]), lido))
```

3. Antes do `return linhas, problemas`, o bloco de `inv_fotos`:

```python
    vistos = set()
    for r in brutos["inv_fotos"]:
        rot = f"inv_fotos linha {r['_linha']}"
        eid = evento_ok(r, rot)
        if eid is None:
            continue
        num = db._numero(r["numero"])
        if num is None or num != int(num):
            problemas.append(f"{rot}: número inválido ({_texto(r['numero']) or '(vazio)'})"); continue
        num = int(num)
        if (eid, num) not in leituras_ok:
            problemas.append(f"{rot}: bem {num} não tem leitura no evento {eid} (aba inv_leituras)"); continue
        nfoto = db._numero(r["nfoto"])
        if nfoto is None or nfoto != int(nfoto) or nfoto < 1:
            problemas.append(f"{rot}: nfoto inválido ({_texto(r['nfoto']) or '(vazio)'})"); continue
        nfoto = int(nfoto)
        url = _texto(r["url"])
        if not url:
            problemas.append(f"{rot}: url vazia"); continue
        if (eid, num, nfoto) in vistos:
            problemas.append(f"{rot}: foto {nfoto} do bem {num} repetida no evento {eid}"); continue
        criado = _data_iso(r["criado_em"], "data criado_em", rot, problemas, True)
        if criado is None:
            continue
        vistos.add((eid, num, nfoto))
        linhas["inv_fotos"].append((eid, num, nfoto, url, criado))
    if not brutos["inv_fotos"]:
        linhas["inv_fotos"] = fotos_antigas          # planilha anterior à Fase 3: foto_url de inv_leituras vira foto 1
```

`substituir_tabelas`: o `INSERT` de `inventario_leituras` perde `foto_url` (8 colunas, 8 `?`); acrescentar ao fim:

```python
    conn.executemany("INSERT INTO inventario_fotos (evento_id, numero, nfoto, url, criado_em) VALUES (?,?,?,?,?)", linhas["inv_fotos"])
```

(o `DELETE` em ordem reversa de `ABAS` já apaga `inventario_fotos` primeiro.)

`db._ler_aba_cadastro`: assinatura `(wb, tabela, problemas, colunas=None, opcional=False, opcionais=())` e:

```python
    faltando = [c for c in colunas if c not in cabecalho and c not in opcionais]
    ...
    idx = {c: (cabecalho.index(c) if c in cabecalho else None) for c in colunas}
    ...
        linhas.append({"_linha": n, **{c: (r[i] if i is not None and i < len(r) else None) for c, i in idx.items()}})
```

`db.importar_cadastros`:

```python
        inv_brutos = {aba: _ler_aba_cadastro(wb, aba, problemas, colunas=cols + (["foto_url"] if aba == "inv_leituras" else []),
                                             opcional=True, opcionais=("foto_url",)) for aba, cols in inventario.ABAS.items()}
    ...
    faltam = [aba for aba, v in inv_brutos.items() if v is None and aba not in inventario.ABAS_OPCIONAIS]
```

O restante (`inv_linhas["inv_bens_encerrados"] = None` quando ausente) fica. `resultado.update(...)` já conta `inv_fotos`.

- [ ] **Step 4: Rodar a suíte inteira**

Run: `.venv/bin/pytest -q`
Expected: só `tests/test_app.py` falha (rotas e templates ainda usam `fotos.nome_bem`, `foto_url` de
`atualizar_leitura` e `foto_excluir` sem `nfoto`). Tudo em `test_inventario.py`, `test_fotos.py`, `test_db.py`,
`test_migrar_inventario.py` passa.

- [ ] **Step 5: Commit (Tasks 2 + 3 + 4)**

```bash
git add db.py inventario.py tests/test_inventario.py
git commit -m "Inventário: tabela inventario_fotos (várias fotos por bem), migração de foto_url, aba inv_fotos e pasta única por evento

test_app fica vermelho até as rotas mudarem (Task 5).

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 5: rotas de foto (`app_inventario.py`)

**Files:**
- Modify: `app_inventario.py` (`foto_leitura`, `foto_excluir`, `sobra`)
- Test: `tests/test_app.py`

**Interfaces:**
- Consumes: `inventario.adicionar_foto`, `inventario.apagar_foto`, `inventario.pasta_do_evento`, `fotos.chave_sobra`, `fotos.enviar(chave, dados)`.
- Produces: `POST /inventario/<id>/leitura/<numero>/foto` → `{"fotos": [{"nfoto", "url", "criado_em"}]}`;
  `POST /inventario/<id>/leitura/<numero>/foto/<int:nfoto>/excluir` (form `volta`) → redirect à sala, flash "Foto removida.";
  endpoint `inventario.foto_excluir` passa a exigir `nfoto` no `url_for`.

- [ ] **Step 1: Atualizar os testes**

Em `tests/test_app.py`, `test_inventario_fotos_e_sobras`: trocar o trecho de "fotos ativas" até a exclusão da sobra por:

```python
    # fotos ativas: cliente falso
    for v in fotos.VARIAVEIS:
        monkeypatch.setenv(v, "x")
    monkeypatch.setenv("R2_PUBLIC_URL", "https://f.exemplo.org")
    enviados = []
    monkeypatch.setattr(fotos, "enviar", lambda chave, dados: enviados.append(chave) or f"https://f.exemplo.org/{chave}")
    apagados = []
    monkeypatch.setattr(fotos, "apagar", lambda url: apagados.append(url))
    r = cliente.post(f"/inventario/{eid}/leitura/1001/foto", data={"foto": (io.BytesIO(imagem), "a.png")}, content_type="multipart/form-data")
    assert r.status_code == 200 and [f["nfoto"] for f in r.get_json()["fotos"]] == [1] and enviados == ["inv/1-1001.webp"]
    r = cliente.post(f"/inventario/{eid}/leitura/1001/foto", data={"foto": (io.BytesIO(imagem), "a.png")}, content_type="multipart/form-data")
    assert [f["url"] for f in r.get_json()["fotos"]] == ["https://f.exemplo.org/inv/1-1001.webp", "https://f.exemplo.org/inv/2-1001.webp"]
    r = cliente.post(f"/inventario/{eid}/leitura/1002/foto", data={"foto": (io.BytesIO(imagem), "a.png")}, content_type="multipart/form-data")
    assert r.status_code == 409 and "Leia o bem" in r.get_json()["erro"]                  # 1002 não foi lido
    r = cliente.post(f"/inventario/{eid}/leitura/1001/foto/1/excluir", data={"volta": "01 - SALA CCI"}, follow_redirects=True)
    assert apagados == ["https://f.exemplo.org/inv/1-1001.webp"] and b"Foto removida" in r.data
    assert cliente.post(f"/inventario/{eid}/leitura/1001/foto/1/excluir", follow_redirects=True).status_code == 404
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/sobra", data={"descricao": "CADEIRA VELHA", "observacao": "x"}, follow_redirects=True)
    assert "precisa de foto".encode() in r.data
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/sobra", data={"descricao": "CADEIRA VELHA", "observacao": "x", "foto": (io.BytesIO(imagem), "b.jpg")}, content_type="multipart/form-data", follow_redirects=True)
    assert b"CADEIRA VELHA" in r.data and enviados[-1].startswith("inv/sobra-") and enviados[-1].endswith(".webp")
    # falha no envio: sobra não fica registrada; foto de bem não fica registrada
    monkeypatch.setattr(fotos, "enviar", lambda chave, dados: (_ for _ in ()).throw(RuntimeError("bucket fora")))
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/sobra", data={"descricao": "MESA VELHA", "observacao": "x", "foto": (io.BytesIO(imagem), "c.jpg")}, content_type="multipart/form-data", follow_redirects=True)
    assert b"MESA VELHA" not in r.data and "não registrada".encode() in r.data
    r = cliente.post(f"/inventario/{eid}/leitura/1001/foto", data={"foto": (io.BytesIO(imagem), "a.png")}, content_type="multipart/form-data")
    assert r.status_code == 409 and "Falha ao enviar" in r.get_json()["erro"]
    import db, inventario
    assert [f["nfoto"] for f in inventario.fotos_do_bem_no_evento(db.conectar(), eid, 1001)] == [2]
    sobras = inventario.bens_da_sala(db.conectar(), eid, "01 - SALA CCI")["sobras"]
    assert [s["descricao"] for s in sobras] == ["CADEIRA VELHA", "VENTILADOR"]
    r = cliente.post(f"/inventario/{eid}/sobra/{sobras[0]['id']}/excluir", follow_redirects=True)
    assert b"CADEIRA VELHA" not in r.data and len(apagados) == 2
    # evento encerrado: não mexe em foto no bucket nem aceita novo envio
    monkeypatch.setattr(fotos, "enviar", lambda chave, dados: enviados.append(chave) or f"https://f.exemplo.org/{chave}")
    cliente.post(f"/inventario/{eid}/encerrar", data={"confirmar": "1"})
    n_apagados, n_enviados = len(apagados), len(enviados)
    r = cliente.post(f"/inventario/{eid}/leitura/1001/foto/2/excluir", follow_redirects=True)
    assert len(apagados) == n_apagados and b"Evento encerrado" in r.data
    r = cliente.post(f"/inventario/{eid}/leitura/1001/foto", data={"foto": (io.BytesIO(imagem), "d.png")}, content_type="multipart/form-data")
    assert r.status_code == 409 and len(enviados) == n_enviados
```

(`_abrir` cria o evento com nome "Inv" → pasta `inv`.)

Em `test_inventario_relatorio_filtros_ordem_modal_e_xlsx_com_fotos` e `test_inventario_lote_marcar_e_desmarcar`,
trocar `inventario.atualizar_leitura(db.conectar(), eid, 1001, foto_url="https://x/1001.webp")` por
`inventario.adicionar_foto(db.conectar(), eid, 1001, lambda chave: "https://x/1001.webp")`.

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_app.py -q -k "fotos_e_sobras or relatorio_filtros or lote"`
Expected: FAIL.

- [ ] **Step 3: Implementar**

Em `app_inventario.py`, substituir `foto_leitura` e `foto_excluir`:

```python
@inventario_bp.route("/<int:id>/leitura/<int:numero>/foto", methods=["POST"])
def foto_leitura(id, numero):
    """Mais uma foto do bem neste evento. Devolve a lista completa (nfoto, url) para a tela redesenhar a célula."""
    conn = _conn()
    try:
        inventario._evento_aberto_ou_erro(conn, id)
        dados = _foto_processada()

        def enviar(chave):
            try:
                return fotos.enviar(chave, dados)
            except Exception:
                raise db.ErroDeNegocio("Falha ao enviar a foto.")

        lista = inventario.adicionar_foto(conn, id, numero, enviar)
    except db.ErroDeNegocio as e:
        return _json_erro(e)
    return jsonify({"fotos": lista})


@inventario_bp.route("/<int:id>/leitura/<int:numero>/foto/<int:nfoto>/excluir", methods=["POST"])
def foto_excluir(id, numero, nfoto):
    conn = _conn()
    leitura = conn.execute("SELECT localizacao FROM inventario_leituras WHERE evento_id = ? AND numero = ?", (id, numero)).fetchone()
    if not leitura:
        abort(404)
    url = inventario.apagar_foto(conn, id, numero, nfoto)
    if url is None:
        abort(404)
    fotos.apagar(url)
    flash("Foto removida.", "success")
    return redirect(url_for("inventario.sala_tela", id=id, localizacao=request.form.get("volta") or leitura["localizacao"]))
```

Em `sobra`, trocar `fotos.nome_sobra(id, sid)` por `fotos.chave_sobra(inventario.pasta_do_evento(conn, id), sid)`.

Atenção: `inventario_sala.html` ainda chama `url_for('inventario.foto_excluir', id=e.id, numero=...)` sem `nfoto`
e quebra a renderização da sala; a Task 6 corrige o template. Para este commit fechar verde, a Task 6 é feita
antes de rodar a suíte inteira (ver Step 4).

- [ ] **Step 4: Rodar só as rotas de dados**

Run: `.venv/bin/pytest tests/test_app.py -q -k "fotos_e_sobras"`
Expected: as asserções até a primeira renderização da sala passam; `follow_redirects=True` para a sala falha com
`BuildError` do `url_for` — seguir imediatamente para a Task 6 e só então rodar a suíte inteira.

- [ ] **Step 5: Não commitar ainda** — commit conjunto ao fim da Task 6.

---

### Task 6: tela da sala com várias miniaturas (`inventario_sala.html`)

**Files:**
- Modify: `templates/inventario_sala.html` (célula `.foto` da tabela "Bens da sala"; JS `URL_FOTO_EXCLUIR`, `preencherFoto`, resposta do upload, texto do `confirm`)
- Test: `tests/test_app.py`

**Interfaces:**
- Consumes: `bens[i].fotos` (lista) de `bens_da_sala`; rota `foto_excluir(id, numero, nfoto)`; resposta `{"fotos": [...]}`.

- [ ] **Step 1: Teste**

Acrescentar a `tests/test_app.py`:

```python
def test_sala_mostra_varias_fotos_e_camera(cliente, monkeypatch):
    import fotos
    eid = _abrir(cliente)
    for v in fotos.VARIAVEIS:
        monkeypatch.setenv(v, "x")
    cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    import db, inventario
    conn = db.conectar()
    inventario.adicionar_foto(conn, eid, 1001, lambda c: "https://x/1.webp")
    inventario.adicionar_foto(conn, eid, 1001, lambda c: "https://x/2.webp")
    r = cliente.get(f"/inventario/{eid}/sala/01 - SALA CCI")
    linha = r.data.split(b'data-numero="1001"')[1].split(b"</tr>")[0]
    assert linha.count(b'class="dsgov-miniatura"') == 2 and b'src="https://x/2.webp"' in linha
    assert f"/inventario/{eid}/leitura/1001/foto/1/excluir".encode() in linha and f"/foto/2/excluir".encode() in linha
    assert b'class="foto-input"' in linha and b"foto-input\" hidden disabled" not in linha         # câmera continua, habilitada
    linha2 = r.data.split(b'data-numero="1002"')[1].split(b"</tr>")[0]
    assert b"dsgov-miniatura" not in linha2 and b'class="foto-input" hidden disabled' in linha2      # não lido: câmera desabilitada
    assert b"/foto/0/excluir" in r.data                                                              # molde da URL para o JS
    assert b"as fotos s" in r.data                                                                   # confirm do Desmarcar fala em fotos
    cliente.post(f"/inventario/{eid}/encerrar", data={"confirmar": "1"})
    r = cliente.get(f"/inventario/{eid}/sala/01 - SALA CCI")
    linha = r.data.split(b'data-numero="1001"')[1].split(b"</tr>")[0]
    assert linha.count(b'class="dsgov-miniatura"') == 2 and b"/excluir" not in linha and b"foto-input" not in linha   # (o JS da página ainda cita foto-input; por isso a checagem é só na linha)
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_app.py::test_sala_mostra_varias_fotos_e_camera -q`
Expected: FAIL (`BuildError` de `url_for('inventario.foto_excluir')` sem `nfoto`).

- [ ] **Step 3: Implementar**

Célula Foto (substituir o `<td class="foto">…</td>` inteiro da tabela "Bens da sala"):

```html
    <td class="foto">{% for f in b.fotos %}<span class="d-inline-block mr-1"><a href="{{ f.url }}" target="_blank" rel="noopener"><img src="{{ f.url }}" alt="Foto {{ f.nfoto }} do bem {{ b.numero }}" class="dsgov-miniatura"/></a>{% if not fechado %}<form method="post" action="{{ url_for('inventario.foto_excluir', id=e.id, numero=b.numero, nfoto=f.nfoto) }}" class="d-inline"><input type="hidden" name="volta" value="{{ localizacao }}"/><button class="br-button circle small" type="submit" aria-label="Excluir foto {{ f.nfoto }}"><i class="fas fa-trash" aria-hidden="true"></i></button></form>{% endif %}</span>{% endfor %}
      {% if fotos_ativas and not fechado %}<label class="br-button circle small{% if not b.lido_em %} disabled{% endif %}" aria-label="Nova foto do bem {{ b.numero }}"><i class="fas fa-camera" aria-hidden="true"></i><input type="file" accept="image/*" capture="environment" class="foto-input" hidden{% if not b.lido_em %} disabled{% endif %}/></label>{% endif %}</td>
```

JS:

```js
  var URL_FOTO_EXCLUIR = {{ url_for('inventario.foto_excluir', id=e.id, numero=0, nfoto=0)|tojson }};   // troca os 0 por número e nfoto
```

`confirm` do Desmarcar: `"... Eles voltam a pendentes e as fotos são apagadas."`.

`preencherFoto` vira:

```js
  function preencherFoto(td, fotosLista, numero) {
    td.textContent = "";
    fotosLista.forEach(function (f) {
      var wrap = document.createElement("span"); wrap.className = "d-inline-block mr-1";
      var a = document.createElement("a"); a.href = f.url; a.target = "_blank"; a.rel = "noopener";
      var img = document.createElement("img"); img.src = f.url; img.alt = "Foto " + f.nfoto; img.className = "dsgov-miniatura";
      a.appendChild(img); wrap.appendChild(a);
      var form = document.createElement("form"); form.method = "post"; form.className = "d-inline";
      form.action = URL_FOTO_EXCLUIR.replace(/\/0\/foto\/0\/excluir$/, "/" + numero + "/foto/" + f.nfoto + "/excluir");
      var input = document.createElement("input"); input.type = "hidden"; input.name = "volta"; input.value = SALA;
      var btn = document.createElement("button"); btn.className = "br-button circle small"; btn.type = "submit"; btn.setAttribute("aria-label", "Excluir foto " + f.nfoto);
      btn.innerHTML = '<i class="fas fa-trash" aria-hidden="true"></i>';
      form.appendChild(input); form.appendChild(btn); wrap.appendChild(form);
      td.appendChild(wrap);
    });
    var label = document.createElement("label"); label.className = "br-button circle small"; label.setAttribute("aria-label", "Nova foto do bem " + numero);
    label.innerHTML = '<i class="fas fa-camera" aria-hidden="true"></i>';
    var inp = document.createElement("input"); inp.type = "file"; inp.accept = "image/*"; inp.setAttribute("capture", "environment"); inp.className = "foto-input"; inp.hidden = true;
    label.appendChild(inp); td.appendChild(label);
  }
```

e no upload: `preencherFoto(tr.querySelector(".foto"), res.j.fotos, tr.dataset.numero);`.

- [ ] **Step 4: Rodar a suíte inteira**

Run: `.venv/bin/pytest -q`
Expected: tudo PASS (inclusive `test_inventario_fotos_e_sobras` da Task 5).

- [ ] **Step 5: Commit (Tasks 5 + 6)**

```bash
git add app_inventario.py templates/inventario_sala.html tests/test_app.py
git commit -m "Sala: várias fotos por bem (câmera sempre disponível, lixeira por foto); rotas de foto devolvem a lista e excluem por nfoto

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 7: relatório mostra "+N" (`inventario_relatorio.html`)

**Files:**
- Modify: `templates/inventario_relatorio.html` (célula Foto)
- Test: `tests/test_app.py`

**Interfaces:**
- Consumes: `x.n_fotos` das linhas de `relatorio` (Task 3).

- [ ] **Step 1: Teste**

Em `tests/test_app.py::test_inventario_relatorio_filtros_ordem_modal_e_xlsx_com_fotos`, logo após a linha que chama
`inventario.adicionar_foto(...)`, acrescentar `inventario.adicionar_foto(db.conectar(), eid, 1001, lambda chave: "https://x/1001b.webp")`
e, após a asserção com `data-foto="https://x/1001.webp"`, acrescentar:

```python
    assert b'<span class="br-tag small">+1</span>' in r.data and b'data-foto="https://x/1001b.webp"' not in r.data
```

O xlsx continua com `=_xlfn.IMAGE("https://x/1001.webp")` (primeira foto) — as asserções existentes já cobrem.

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_app.py::test_inventario_relatorio_filtros_ordem_modal_e_xlsx_com_fotos -q`
Expected: FAIL na asserção do `+1`.

- [ ] **Step 3: Implementar**

Na célula Foto do relatório, após o `</button>` da miniatura e antes do `{% else %}`:

```html
{% if x.n_fotos > 1 %} <span class="br-tag small">+{{ x.n_fotos - 1 }}</span>{% endif %}
```

- [ ] **Step 4: Rodar**

Run: `.venv/bin/pytest tests/test_app.py -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add templates/inventario_relatorio.html tests/test_app.py
git commit -m "Relatório do inventário: primeira foto com +N quando o bem tem mais fotos

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 8: fotos no cadastro do bem (`app.py`, `bem.html`)

**Files:**
- Modify: `app.py` (rota `bem`), `templates/bem.html`
- Test: `tests/test_app.py`

**Interfaces:**
- Consumes: `inventario.fotos_do_bem(conn, numero)` (Task 3).

- [ ] **Step 1: Teste**

```python
def test_bem_mostra_fotos_por_evento(cliente):
    assert b"Fotos do invent" not in cliente.get("/bem?numero=1001").data
    eid = _abrir(cliente)
    cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    import db, inventario
    conn = db.conectar()
    inventario.adicionar_foto(conn, eid, 1001, lambda c: "https://x/a.webp")
    inventario.adicionar_foto(conn, eid, 1001, lambda c: "https://x/b.webp")
    cliente.post(f"/inventario/{eid}/encerrar", data={"confirmar": "1"})
    cliente.post("/inventario/abrir", data={"nome": "Inv 2", "integrantes": "Fulano", "escopo": "todas"})
    eid2 = inventario.evento_aberto(conn)["id"]
    cliente.post(f"/inventario/{eid2}/integrante", data={"integrante": "Fulano"})
    cliente.post(f"/inventario/{eid2}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    inventario.adicionar_foto(conn, eid2, 1001, lambda c: "https://x/c.webp")
    r = cliente.get("/bem?numero=1001")
    assert b"Fotos do invent" in r.data and r.data.count(b'class="dsgov-miniatura"') == 3
    assert r.data.index(b"Inv 2") < r.data.index(b">Inv<") and r.data.index(b"https://x/c.webp") < r.data.index(b"https://x/a.webp") < r.data.index(b"https://x/b.webp")
    assert b"encerrado em" in r.data and b"lido em" in r.data
    assert b"Fotos do invent" not in cliente.get("/bem?numero=1002").data
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_app.py::test_bem_mostra_fotos_por_evento -q`
Expected: FAIL.

- [ ] **Step 3: Implementar**

`app.py`, rota `bem`: acrescentar `fotos=inventario.fotos_do_bem(obter_conn(), int(numero))` ao `render_template`.

`templates/bem.html`, antes de `{% endblock %}`:

```html
{% if fotos %}
<h2 class="text-up-01 mt-4 mb-2">Fotos do inventário</h2>
<div class="br-card"><div class="card-content">
  {% for g in fotos %}
  <div class="{% if not loop.last %}mb-3{% endif %}">
    <div class="mb-1"><strong>{{ g.evento }}</strong> <span class="text-gray-70">· aberto em {{ g.aberto_em[8:10] }}/{{ g.aberto_em[5:7] }}/{{ g.aberto_em[:4] }}{% if g.encerrado_em %} · encerrado em {{ g.encerrado_em[8:10] }}/{{ g.encerrado_em[5:7] }}/{{ g.encerrado_em[:4] }}{% endif %} · lido em {{ g.lido_em[8:10] }}/{{ g.lido_em[5:7] }}/{{ g.lido_em[:4] }} {{ g.lido_em[11:16] }}</span></div>
    <div class="d-flex flex-wrap">{% for f in g.fotos %}<a class="mr-1 mb-1" href="{{ f.url }}" target="_blank" rel="noopener"><img src="{{ f.url }}" alt="Foto {{ f.nfoto }} do bem {{ bem.numero }} no evento {{ g.evento }}" class="dsgov-miniatura"/></a>{% endfor %}</div>
  </div>
  {% endfor %}
</div></div>
{% endif %}
```

Observação para o teste: o nome do evento aparece como `<strong>Inv</strong>` → a asserção usa `b">Inv<"`.

- [ ] **Step 4: Rodar**

Run: `.venv/bin/pytest tests/test_app.py -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add app.py templates/bem.html tests/test_app.py
git commit -m "Cadastro do bem: card Fotos do inventário, agrupadas por evento do mais recente para o mais antigo

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 9: menu "Inventário" como grupo (`app.py`, `base.html`)

**Files:**
- Modify: `app.py` (`contexto_dsgov`), `templates/base.html` (laço do menu)
- Test: `tests/test_app.py`

**Interfaces:**
- Produces: `MENU` = lista de `(rotulo, icone, url, filhos)`; `filhos` = lista de `(rotulo, url)`.

- [ ] **Step 1: Teste**

```python
def test_menu_inventario_e_grupo_com_telas_do_evento_aberto(cliente):
    r = cliente.get("/").data
    menu = r.split(b'id="main-navigation"')[1].split(b"menu-footer")[0]
    assert b"menu-folder" in menu and b">Eventos<" in menu and b"/painel" not in menu
    assert b'href="/inventario"' in menu and menu.count(b"menu-folder") == 1
    eid = _abrir(cliente, integrante=None)
    menu = cliente.get("/").data.split(b'id="main-navigation"')[1].split(b"menu-footer")[0]
    assert f'href="/inventario/{eid}"'.encode() in menu and f'href="/inventario/{eid}/painel"'.encode() in menu and f'href="/inventario/{eid}/relatorio"'.encode() in menu
    assert b">Inv<" in menu and b">Painel<" in menu and b">Relat" in menu
    cliente.post(f"/inventario/{eid}/encerrar", data={"confirmar": "1"})
    assert b"/painel" not in cliente.get("/").data.split(b'id="main-navigation"')[1].split(b"menu-footer")[0]
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_app.py::test_menu_inventario_e_grupo_com_telas_do_evento_aberto -q`
Expected: FAIL (`menu-folder` ausente).

- [ ] **Step 3: Implementar**

`app.py`, `contexto_dsgov`:

```python
@app.context_processor
def contexto_dsgov():
    t = textos.obter(obter_conn())
    dsgov = dict(DSGOV_FIXO, ORGAO=t["orgao_nome"], SUBTITULO=t["unidade_sigla"])
    inv = [("Eventos", url_for("inventario.eventos_tela"))]
    if (e := inventario.evento_aberto(obter_conn())):
        inv += [(e["nome"], url_for("inventario.evento_tela", id=e["id"])),
                ("Painel", url_for("inventario.painel_tela", id=e["id"])),
                ("Relatório", url_for("inventario.relatorio_tela", id=e["id"]))]
    return {"DSGOV": dsgov, "MENU": [
        ("Início", "fa-home", url_for("home"), []),
        ("Termo por centro de custo", "fa-building", url_for("centro_custos"), []),
        ("Termo individual", "fa-user-check", url_for("termos_individuais"), []),
        ("Termo de devolução", "fa-box-open", url_for("termo_devolucao"), []),
        ("Termos emitidos", "fa-history", url_for("termos_emitidos_tela"), []),
        ("Recorte", "fa-filter", url_for("recorte"), []),
        ("Inventário", "fa-clipboard-check", url_for("inventario.eventos_tela"), inv),
        ("Cadastros", "fa-address-book", url_for("cadastros", aba="responsaveis"), []),
        ("Textos", "fa-file-signature", url_for("textos_tela"), []),
        ("Atualizar base", "fa-upload", url_for("upload"), []),
    ]}
```

`templates/base.html`, laço do menu:

```html
                {% for rotulo, icone, url, filhos in MENU %}
                {% if filhos %}
                <div class="menu-folder">
                  <div class="menu-item"><span class="icon"><i class="fas {{ icone }}" aria-hidden="true"></i></span><span class="content">{{ rotulo }}</span></div>
                  <ul role="group">{% for r, u in filhos %}<li><a class="menu-item" href="{{ u }}" role="treeitem"><span class="content">{{ r }}</span></a></li>{% endfor %}</ul>
                </div>
                {% else %}
                <a class="menu-item" href="{{ url }}" role="treeitem"><span class="icon"><i class="fas {{ icone }}" aria-hidden="true"></i></span><span class="content">{{ rotulo }}</span></a>
                {% endif %}
                {% endfor %}
```

O título do grupo é `div.menu-item` (não `<a>`): o `core.min.js` do DSGov trata `.menu-folder:not(.drop-menu)`
como grupo sempre expandido. Se ao testar no navegador o grupo aparecer recolhido, trocar o `div` por
`<a class="menu-item" href="{{ url }}" role="treeitem">` (vira pasta com seta, um toque para abrir) — e
registrar isso no commit.

- [ ] **Step 4: Rodar**

Run: `.venv/bin/pytest -q`
Expected: PASS (todos; `grep -rn "MENU" templates/ app.py` não mostra outro consumidor).

- [ ] **Step 5: Commit**

```bash
git add app.py templates/base.html tests/test_app.py
git commit -m "Menu: Inventário vira grupo com Eventos e, com evento aberto, o evento, Painel e Relatório

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 10: README e revisão final

**Files:**
- Modify: `README.md` (itens 1 e 6 da seção de uso; linha da seção "Proposta comercial" fica)
- Test: suíte inteira

- [ ] **Step 1: README**

No item 1, trocar a frase "Fotos vão para o bucket R2 configurado em `secrets/.env` (variáveis `R2_*`); sem ele,
fotos ficam desativadas." por:

```
   Fotos vão para o bucket R2 configurado em `secrets/.env` (variáveis `R2_*`); sem ele, fotos ficam
   desativadas. Cada bem aceita várias fotos por evento (a sala mostra todas; relatório e `.xlsx` só a
   primeira, com "+N"); a chave no bucket é `<pasta>/<n>-<número do bem>.webp`, com a pasta sendo o nome
   do evento em minúsculas sem acento (ex.: `inventario2026/1-12334.webp`), por isso dois eventos não podem
   ter nomes que gerem a mesma pasta. O cadastro do bem (`/bem`) lista as fotos agrupadas por evento.
```

No item 6, trocar "mais 6 abas `inv_*`" por "mais 7 abas `inv_*`" e acrescentar ao fim do item:
"A 7ª aba, `inv_fotos`, tem as fotos dos bens (evento, número, nfoto, url); também é opcional — ausente, uma
coluna `foto_url` em `inv_leituras` (planilhas anteriores) vira a foto 1 de cada bem."

Acrescentar ao fim do item 1 (ou como frase própria): "No menu, *Inventário* é um grupo com *Eventos* e, quando há
evento aberto, o próprio evento, *Painel* e *Relatório*."

- [ ] **Step 2: Conferências finais**

```bash
.venv/bin/pytest -q
grep -rn "nome_bem\|nome_sobra\|PREFIXO\b\|atualizar_leitura(.*foto_url" --include=*.py --include=*.html . | grep -v "\.venv\|docs/"
grep -rn "foto_url" --include=*.py --include=*.html . | grep -v "\.venv\|docs/"
```

Expected: suíte verde (≈ 204 + 9 novos). O primeiro `grep` não devolve nada. No segundo, cada ocorrência que
restou é de um destes tipos: sobras (`inventario_sobras.foto_url`, `inv_sobras`, `definir_foto_sobra`, template da
sobra), a subconsulta `AS foto_url` em `_FOTO_SQL`, `_tem_foto`/`_celula_foto`/`exportar_xlsx` (primeira foto),
templates que mostram a primeira foto (`x.foto_url`, `data-foto`), e a compatibilidade da planilha antiga
(`opcionais`, `r.get("foto_url")`, `fotos_antigas`, `if aba == "inv_leituras"`). Qualquer outra é resto a remover.

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "README: várias fotos por bem, chave simples no R2, aba inv_fotos e grupo do menu

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

- [ ] **Step 4: Entrega** (feita pelo orquestrador, não pelo subagente): revisão final do diff da branch, merge
em `main`, `git push`, `docker compose up -d --build` na VPS, conferir `docker compose logs --tail 20` (a
migração roda no primeiro start; nada a fazer no bucket). Depois o usuário valida no celular: tirar duas fotos de
um bem, apagar uma, abrir `/bem?numero=...`, abrir o menu.
