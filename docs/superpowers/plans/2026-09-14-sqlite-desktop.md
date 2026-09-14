# SQLite + programa de desktop — Plano de implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Trocar as planilhas Excel por SQLite mantido pelo app, adotar o visual DSGov, emitir os termos em HTML com "Copiar para o SEI", e empacotar como programa Windows (pywebview + PyInstaller).

**Architecture:** Flask continua sendo o app; `db.py` concentra esquema, importação e consultas (funções puras que recebem `conn`); `termos_html.py` monta o corpo HTML dos termos no padrão do `/gelic`; os três geradores `.docx` passam a receber listas de dicts. `config.py` decide onde os dados vivem (`PASTA_DADOS`). `main.py` sobe o Flask numa thread e abre a janela.

**Tech Stack:** Python 3.14, Flask 3.1, sqlite3 (stdlib), openpyxl 3.1, python-docx 1.2, pywebview 6.2, PyInstaller 6.22, pytest 9. DSGov 3.7.0 vendorizado (CSS/JS/fontes) de `~/.claude/skills/dsgov/assets/`.

**Spec:** `docs/superpowers/specs/2026-09-14-sqlite-desktop-design.md`

## Global Constraints

- Simplicidade: adaptar, não reescrever. Nenhuma dependência além de `flask openpyxl python-docx pywebview pyinstaller pytest`.
- `pandas` e `gunicorn` saem. `Procfile` sai.
- Nenhum `<style>`, nenhum hex, nenhum `<select>` nativo nos templates DSGov. Um único `br-button primary` por tela. Um único `h1` por página.
- **Exceção:** `termos_html.py` usa `style=` inline em `table`/`th`/`td` (o que a área de transferência leva pro SEI). Só ali.
- Largura da tabela do documento: **80 %** (individual e devolução), **100 %** (centro de custo).
- Regra de negócio: bem em `atribuicoes` **não entra** no termo do centro de custo.
- Todos os caminhos passam por `config.py`; a variável de ambiente `TERMOS_DADOS` sobrepõe `PASTA_DADOS` (é como os testes isolam o banco).
- Código, comentários, mensagens e commits em português, como o resto do repositório.
- Rodar sempre com `.venv/bin/python` / `.venv/bin/pytest` (a venv já existe em `/opt/web/termos-responsabilidade/.venv`).
- Nesta VPS não há WebKitGTK: `main.py` cai no modo navegador. A janela pywebview só é validada no Windows do usuário.

---

## Estrutura de arquivos

| Arquivo | Responsabilidade |
|---|---|
| `config.py` (novo) | `pasta_dados()`, `caminho_db()`, `caminho_timbrado()`, `pasta_saida()`, `pasta_recursos()`, `preparar_pastas()` |
| `db.py` (novo) | esquema, `conectar`, `importar_bens`, consultas, operações de cadastro, exceções de negócio |
| `termos_html.py` (novo) | `documento()`, `corpo_ccusto()`, `corpo_individual()`, `corpo_devolucao()` |
| `Script_Termo_Individual.py` | `criar_termo_responsabilidade(nome, bens, destino)` — sem pandas |
| `Termo_de_Responsabilidade.py` | `gerar_termo_centro(ccustos, responsavel, bens, destino)` + `gerar_planilha_centro(bens, destino)` — sem pandas |
| `termo_devolucao.py` | `gerar_termo_devolucao(nome, bens, destino)` — sem pandas |
| `app.py` | rotas Flask; conexão por request em `g` |
| `importar_planilhas.py` (novo) | migração inicial das duas planilhas |
| `main.py` (novo) | thread Flask + janela pywebview (fallback navegador) |
| `build.bat` (novo) | comando PyInstaller |
| `templates/base.html`, `templates/_macros.html`, demais templates | DSGov |
| `templates/termo_base.html` | folha de estilo de documento (do gelic) |
| `static/dsgov/` | vendor DSGov |
| `tests/conftest.py`, `tests/test_*.py` | pytest |

---

### Task 1: `config.py`, dependências, `.gitignore`, esqueleto de testes

**Files:**
- Create: `config.py`, `tests/__init__.py`, `tests/conftest.py`, `tests/test_config.py`
- Modify: `requirements.txt`, `.gitignore`

**Interfaces:**
- Produces: `config.pasta_dados() -> Path`, `config.caminho_db() -> Path`, `config.caminho_timbrado() -> Path`, `config.pasta_saida() -> Path`, `config.pasta_recursos() -> Path`, `config.preparar_pastas() -> None`. Env `TERMOS_DADOS` sobrepõe a pasta.

- [ ] **Step 1: Escrever o teste que falha**

`tests/__init__.py` vazio. `tests/test_config.py`:

```python
from pathlib import Path


def test_env_sobrepoe_pasta_dados(tmp_path, monkeypatch):
    monkeypatch.setenv("TERMOS_DADOS", str(tmp_path))
    import config
    assert config.pasta_dados() == tmp_path
    assert config.caminho_db() == tmp_path / "termos.db"
    assert config.caminho_timbrado() == tmp_path / "timbrado.docx"
    assert config.pasta_saida() == tmp_path / "saida"


def test_preparar_pastas_cria_saida_e_copia_timbrado(tmp_path, monkeypatch):
    monkeypatch.setenv("TERMOS_DADOS", str(tmp_path))
    import config
    config.preparar_pastas()
    assert (tmp_path / "saida").is_dir()
    assert (tmp_path / "timbrado.docx").stat().st_size > 1000


def test_sem_env_usa_pasta_dados_do_projeto(monkeypatch):
    monkeypatch.delenv("TERMOS_DADOS", raising=False)
    import config
    assert config.pasta_dados() == Path(config.__file__).parent / "dados"
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_config.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'config'`

- [ ] **Step 3: Implementar `config.py`**

```python
"""Onde os dados vivem. Tudo que lê ou grava arquivo passa por aqui.

Windows empacotado (PyInstaller): pasta `dados/` ao lado do .exe.
Desenvolvimento: `./dados/`. A variável TERMOS_DADOS sobrepõe (usada nos testes).
"""
import os
import shutil
import sys
from pathlib import Path


def pasta_recursos() -> Path:
    """Arquivos embutidos no programa (templates, static, timbrado.docx original)."""
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


def pasta_dados() -> Path:
    if os.environ.get("TERMOS_DADOS"):
        return Path(os.environ["TERMOS_DADOS"])
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / "dados"
    return Path(__file__).resolve().parent / "dados"


def caminho_db() -> Path:
    return pasta_dados() / "termos.db"


def caminho_timbrado() -> Path:
    return pasta_dados() / "timbrado.docx"


def pasta_saida() -> Path:
    return pasta_dados() / "saida"


def preparar_pastas() -> None:
    """Cria dados/saida e copia o timbrado na primeira execução."""
    pasta_saida().mkdir(parents=True, exist_ok=True)
    if not caminho_timbrado().exists():
        shutil.copy(pasta_recursos() / "timbrado.docx", caminho_timbrado())
```

- [ ] **Step 4: `requirements.txt` e `.gitignore`**

`requirements.txt` (o arquivo atual está em UTF-16; regravar em UTF-8):

```
flask
openpyxl
python-docx
pywebview
pyinstaller
pytest
```

Acrescentar ao final de `.gitignore`:

```
# Dados do programa (banco, termos gerados) e build
dados/
build/
dist/
*.spec
```

- [ ] **Step 5: `tests/conftest.py` com fixture de dados isolados**

```python
"""Fixtures compartilhadas: banco SQLite temporário com dados de exemplo."""
import pytest


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


def semear(conn):
    """Cenário mínimo: 1 centro (CCI), 1 sala mapeada, 4 bens, 1 pessoa com 1 bem atribuído."""
    conn.execute("INSERT INTO responsaveis VALUES ('CCI','Prezada','JAQUELINE PORTELA','j@cfc.org.br','46','coordenadora')")
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
    conn.execute("INSERT INTO pessoas VALUES ('ANA SILVA')")
    conn.execute("INSERT INTO atribuicoes VALUES ('ANA SILVA', 1002)")
    conn.commit()
```

(`db.conectar` e `db.criar_esquema` chegam na Task 2; por enquanto só `test_config.py` roda.)

- [ ] **Step 6: Rodar e ver passar**

Run: `.venv/bin/pytest tests/test_config.py -v`
Expected: 3 passed

- [ ] **Step 7: Commit**

```bash
git add config.py requirements.txt .gitignore tests/
git commit -m "config: pasta de dados única, requirements sem pandas/gunicorn, esqueleto de testes"
```

---

### Task 2: `db.py` — esquema e conexão

**Files:**
- Create: `db.py`, `tests/test_db.py`

**Interfaces:**
- Produces: `db.conectar(caminho: Path | None = None) -> sqlite3.Connection` (row_factory=Row, foreign_keys ON); `db.criar_esquema(conn) -> None` (idempotente); `db.inicializar() -> None` (prepara pastas + esquema); `db.ErroDeNegocio(Exception)`.

- [ ] **Step 1: Teste que falha**

`tests/test_db.py`:

```python
import sqlite3

import pytest

import db
from tests.conftest import semear


def test_esquema_cria_cinco_tabelas(dados):
    nomes = {r["name"] for r in dados.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    assert {"bens", "responsaveis", "localizacoes", "pessoas", "atribuicoes"} <= nomes


def test_esquema_e_idempotente(dados):
    db.criar_esquema(dados)  # segunda vez não pode falhar


def test_foreign_keys_ligadas(dados):
    semear(dados)
    with pytest.raises(sqlite3.IntegrityError):
        dados.execute("INSERT INTO localizacoes VALUES ('02 - X', 'NAO_EXISTE')")
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_db.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'db'`

- [ ] **Step 3: Implementar**

`db.py`:

```python
"""Banco SQLite: esquema, importação de bens, consultas e cadastros.

Todas as funções recebem a conexão como primeiro argumento; quem abre e fecha é o chamador
(o Flask, por request; os testes, por fixture). Nenhuma função aqui usa Flask.
"""
import sqlite3
from pathlib import Path

import config

ESQUEMA = """
CREATE TABLE IF NOT EXISTS bens (
  numero        INTEGER PRIMARY KEY,
  situacao      TEXT NOT NULL,
  descricao     TEXT NOT NULL,
  complemento   TEXT,
  classificacao TEXT,
  localizacao   TEXT,
  data_entrada  TEXT,
  valor_compra  REAL,
  valor_atual   REAL
);
CREATE TABLE IF NOT EXISTS responsaveis (
  ccustos     TEXT PRIMARY KEY,
  tratamento  TEXT,
  responsavel TEXT NOT NULL,
  email       TEXT,
  matricula   TEXT,
  funcao      TEXT
);
CREATE TABLE IF NOT EXISTS localizacoes (
  localizacao TEXT PRIMARY KEY,
  ccustos     TEXT NOT NULL REFERENCES responsaveis(ccustos) ON UPDATE CASCADE
);
CREATE TABLE IF NOT EXISTS pessoas (
  nome TEXT PRIMARY KEY
);
CREATE TABLE IF NOT EXISTS atribuicoes (
  nome   TEXT    NOT NULL REFERENCES pessoas(nome) ON UPDATE CASCADE ON DELETE CASCADE,
  numero INTEGER NOT NULL REFERENCES bens(numero) DEFERRABLE INITIALLY DEFERRED,
  PRIMARY KEY (nome, numero)
);
"""


class ErroDeNegocio(Exception):
    """Erro que vira mensagem para o usuário (flash), não traceback."""


def conectar(caminho: Path | None = None) -> sqlite3.Connection:
    conn = sqlite3.connect(str(caminho or config.caminho_db()))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def criar_esquema(conn: sqlite3.Connection) -> None:
    conn.executescript(ESQUEMA)


def inicializar() -> None:
    """Primeira execução do programa: pastas + esquema vazio."""
    config.preparar_pastas()
    conn = conectar()
    try:
        criar_esquema(conn)
    finally:
        conn.close()
```

- [ ] **Step 4: Rodar e ver passar**

Run: `.venv/bin/pytest tests/test_db.py -v`
Expected: 3 passed

- [ ] **Step 5: Commit**

```bash
git add db.py tests/test_db.py
git commit -m "db: esquema SQLite com cinco tabelas e conexão com FK ligadas"
```

---

### Task 3: `db.importar_bens` — upload do export do sistema

**Files:**
- Modify: `db.py`, `tests/test_db.py`

**Interfaces:**
- Produces: `db.importar_bens(conn, arquivo: Path | file-like) -> dict` com chaves `total`, `ativos`, `sem_centro` (lista de localizações ATIVAS sem mapa). Levanta `db.ImportacaoInvalida(ErroDeNegocio)` com mensagem; nesse caso `bens` não muda.
- Produces: `db.localizacoes_sem_centro(conn) -> list[str]`.

- [ ] **Step 1: Testes que falham**

Acrescentar a `tests/test_db.py`:

```python
from openpyxl import Workbook

CABECALHO = ["Número Bem", "Situação", "Descrição", "Complemento", "Classificação Contábil",
             "Localização", "Data Entrada", "Valor Compra", "Valor Atual"]


def xlsx(tmp_path, linhas, cabecalho=CABECALHO, aba="base"):
    wb = Workbook()
    ws = wb.active
    ws.title = aba
    ws.append(cabecalho)
    for l in linhas:
        ws.append(l)
    caminho = tmp_path / "export.xlsx"
    wb.save(caminho)
    return caminho


def test_importar_substitui_bens_e_conta(dados, tmp_path):
    semear(dados)
    arq = xlsx(tmp_path, [
        [1002, "ATIVO", "NOTEBOOK", "DELL NOVO", "EQUIP", "01 - SALA CCI", "06/12/2012", 3000, 1400],
        [2001, "ATIVO", "MONITOR", "LG", "EQUIP", "01 - SALA CCI", "01/01/2020", 900, 800],
        [2002, "DOADO", "CADEIRA", "", "MÓVEIS", "CFC", "01/01/2000", 10, 1],
    ])
    resumo = db.importar_bens(dados, arq)
    assert resumo["total"] == 3 and resumo["ativos"] == 2
    assert [r["numero"] for r in dados.execute("SELECT numero FROM bens ORDER BY numero")] == [1002, 2001, 2002]
    assert dados.execute("SELECT complemento FROM bens WHERE numero=1002").fetchone()[0] == "DELL NOVO"
    # outras tabelas intactas
    assert dados.execute("SELECT count(*) FROM atribuicoes").fetchone()[0] == 1


def test_importar_cabecalho_errado_nao_altera_nada(dados, tmp_path):
    semear(dados)
    arq = xlsx(tmp_path, [[1, "ATIVO"]], cabecalho=["Patrimônio", "Situação"])
    with pytest.raises(db.ImportacaoInvalida):
        db.importar_bens(dados, arq)
    assert dados.execute("SELECT count(*) FROM bens").fetchone()[0] == 4


def test_importar_que_some_com_bem_atribuido_e_revertida(dados, tmp_path):
    semear(dados)  # ANA tem o 1002
    arq = xlsx(tmp_path, [[1001, "ATIVO", "CADEIRA", "", "MÓVEIS", "01 - SALA CCI", "x", 1, 1]])
    with pytest.raises(db.ImportacaoInvalida) as e:
        db.importar_bens(dados, arq)
    assert "1002" in str(e.value)
    assert dados.execute("SELECT count(*) FROM bens").fetchone()[0] == 4


def test_importar_usa_aba_pelo_cabecalho_e_converte_data(dados, tmp_path):
    from datetime import datetime
    wb = Workbook()
    wb.active.title = "outra"
    wb.active.append(["Nada", "aqui"])
    ws = wb.create_sheet("base")
    ws.append(CABECALHO)
    ws.append([3001, "ATIVO", "MESA", None, "MÓVEIS", "02 - X", datetime(2020, 3, 9), 100, 90])
    arq = tmp_path / "e.xlsx"
    wb.save(arq)
    db.importar_bens(dados, arq)
    r = dados.execute("SELECT data_entrada, complemento FROM bens WHERE numero=3001").fetchone()
    assert r["data_entrada"] == "09/03/2020" and r["complemento"] == ""


def test_localizacoes_sem_centro(dados):
    semear(dados)
    assert db.localizacoes_sem_centro(dados) == ["99 - SEM MAPA"]
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_db.py -v -k importar`
Expected: FAIL — `AttributeError: module 'db' has no attribute 'importar_bens'`

- [ ] **Step 3: Implementar**

Acrescentar a `db.py` (depois de `inicializar`):

```python
from datetime import date, datetime

from openpyxl import load_workbook

COLUNAS_EXPORT = {
    "Número Bem": "numero",
    "Situação": "situacao",
    "Descrição": "descricao",
    "Complemento": "complemento",
    "Classificação Contábil": "classificacao",
    "Localização": "localizacao",
    "Data Entrada": "data_entrada",
    "Valor Compra": "valor_compra",
    "Valor Atual": "valor_atual",
}


class ImportacaoInvalida(ErroDeNegocio):
    pass


def _texto(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    return " ".join(str(v).split())


def _numero(v):
    try:
        return float(v) if v not in (None, "") else None
    except (TypeError, ValueError):
        return None


def _aba_do_export(wb):
    for ws in wb.worksheets:
        cabecalho = [_texto(c) for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())]
        if "Número Bem" in cabecalho:
            return ws, cabecalho
    raise ImportacaoInvalida("Nenhuma aba com a coluna 'Número Bem'. Envie o export do sistema de patrimônio.")


def importar_bens(conn: sqlite3.Connection, arquivo) -> dict:
    """Substitui a tabela `bens` pelo conteúdo do export. Tudo ou nada.

    Devolve {"total", "ativos", "sem_centro"}. Levanta ImportacaoInvalida (e não altera nada)
    se faltar coluna ou se algum bem atribuído a pessoa deixar de existir.
    """
    wb = load_workbook(arquivo, read_only=True, data_only=True)
    ws, cabecalho = _aba_do_export(wb)
    faltando = [c for c in COLUNAS_EXPORT if c not in cabecalho]
    if faltando:
        raise ImportacaoInvalida("Colunas ausentes no export: " + ", ".join(faltando))
    indice = {campo: cabecalho.index(col) for col, campo in COLUNAS_EXPORT.items()}

    linhas = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        num = _numero(r[indice["numero"]])
        if num is None:
            continue
        linhas.append((
            int(num), _texto(r[indice["situacao"]]), _texto(r[indice["descricao"]]),
            _texto(r[indice["complemento"]]), _texto(r[indice["classificacao"]]),
            _texto(r[indice["localizacao"]]), _texto(r[indice["data_entrada"]]),
            _numero(r[indice["valor_compra"]]), _numero(r[indice["valor_atual"]]),
        ))
    wb.close()

    try:
        conn.execute("DELETE FROM bens")
        conn.executemany("INSERT INTO bens VALUES (?,?,?,?,?,?,?,?,?)", linhas)
        orfaos = [str(r[0]) for r in conn.execute(
            "SELECT numero FROM atribuicoes WHERE numero NOT IN (SELECT numero FROM bens) ORDER BY numero")]
        if orfaos:
            raise ImportacaoInvalida(
                "O export não traz bens que estão atribuídos a pessoas: " + ", ".join(orfaos)
                + ". Remova a atribuição na aba Pessoas ou use um export completo.")
        conn.commit()
    except Exception:
        conn.rollback()
        raise

    total = conn.execute("SELECT count(*) FROM bens").fetchone()[0]
    ativos = conn.execute("SELECT count(*) FROM bens WHERE situacao='ATIVO'").fetchone()[0]
    return {"total": total, "ativos": ativos, "sem_centro": localizacoes_sem_centro(conn)}


def localizacoes_sem_centro(conn: sqlite3.Connection) -> list[str]:
    """Localizações de bens ATIVOS que não têm centro de custo mapeado."""
    return [r[0] for r in conn.execute(
        "SELECT DISTINCT localizacao FROM bens WHERE situacao='ATIVO' AND localizacao <> '' "
        "AND localizacao NOT IN (SELECT localizacao FROM localizacoes) ORDER BY localizacao")]
```

Mover os `import` novos (`datetime`, `openpyxl`) para o topo do arquivo.

- [ ] **Step 4: Rodar e ver passar**

Run: `.venv/bin/pytest tests/test_db.py -v`
Expected: 8 passed

- [ ] **Step 5: Commit**

```bash
git add db.py tests/test_db.py
git commit -m "db: importação do export de bens (tudo ou nada) e localizações sem centro"
```

---

### Task 4: `db.py` — consultas que substituem as planilhas

**Files:**
- Modify: `db.py`, `tests/test_db.py`

**Interfaces:**
- Produces (todas `(conn, ...) -> list[dict] | dict | None`):
  `centros(conn)` → linhas de `responsaveis` ordenadas por `ccustos`;
  `responsavel(conn, ccustos)` → dict ou None;
  `bens_do_centro(conn, ccustos)` → bens ATIVO, mapeados, não atribuídos, por número;
  `pessoas(conn)` → `list[str]`;
  `bens_da_pessoa(conn, nome)` → bens atribuídos, por número;
  `buscar_bem(conn, numero)` → dict ou None;
  `pessoa_do_bem(conn, numero)` → nome ou None;
  `ficha_do_bem(conn, numero)` → dict `{**bem, "ccustos", "responsavel", "pessoa"}` ou None;
  `localizacoes_mapeadas(conn)` → `[{"localizacao","ccustos"}]`.
  Linhas são `dict` (não `sqlite3.Row`), para os geradores e o Jinja.

- [ ] **Step 1: Testes que falham**

Acrescentar a `tests/test_db.py`:

```python
def test_bens_do_centro_exclui_baixados_atribuidos_e_sem_mapa(dados):
    semear(dados)
    assert [b["numero"] for b in db.bens_do_centro(dados, "CCI")] == [1001]


def test_bens_da_pessoa(dados):
    semear(dados)
    bens = db.bens_da_pessoa(dados, "ANA SILVA")
    assert [b["numero"] for b in bens] == [1002] and bens[0]["descricao"] == "NOTEBOOK"


def test_centros_responsavel_pessoas(dados):
    semear(dados)
    assert [c["ccustos"] for c in db.centros(dados)] == ["CCI"]
    assert db.responsavel(dados, "CCI")["responsavel"] == "JAQUELINE PORTELA"
    assert db.responsavel(dados, "XX") is None
    assert db.pessoas(dados) == ["ANA SILVA"]


def test_ficha_do_bem_com_setor_e_pessoa(dados):
    semear(dados)
    f = db.ficha_do_bem(dados, 1002)
    assert f["ccustos"] == "CCI" and f["responsavel"] == "JAQUELINE PORTELA" and f["pessoa"] == "ANA SILVA"
    f = db.ficha_do_bem(dados, 1004)
    assert f["ccustos"] is None and f["pessoa"] is None
    assert db.ficha_do_bem(dados, 9999) is None
    assert db.buscar_bem(dados, 1001)["descricao"] == "CADEIRA"
    assert db.localizacoes_mapeadas(dados) == [{"localizacao": "01 - SALA CCI", "ccustos": "CCI"}]
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_db.py -v -k "centro or pessoa or ficha"`
Expected: FAIL — `AttributeError: ... 'bens_do_centro'`

- [ ] **Step 3: Implementar**

Acrescentar a `db.py`:

```python
def _todos(conn, sql, *args) -> list[dict]:
    return [dict(r) for r in conn.execute(sql, args)]


def _um(conn, sql, *args) -> dict | None:
    r = conn.execute(sql, args).fetchone()
    return dict(r) if r else None


def centros(conn) -> list[dict]:
    return _todos(conn, "SELECT * FROM responsaveis ORDER BY ccustos")


def responsavel(conn, ccustos: str) -> dict | None:
    return _um(conn, "SELECT * FROM responsaveis WHERE ccustos = ?", ccustos)


def bens_do_centro(conn, ccustos: str) -> list[dict]:
    """Bens ATIVOS nas localizações do centro, excluindo os atribuídos a pessoas (setor OU pessoa)."""
    return _todos(conn, """
        SELECT b.* FROM bens b JOIN localizacoes l ON l.localizacao = b.localizacao
        WHERE l.ccustos = ? AND b.situacao = 'ATIVO'
          AND b.numero NOT IN (SELECT numero FROM atribuicoes)
        ORDER BY b.numero""", ccustos)


def pessoas(conn) -> list[str]:
    return [r[0] for r in conn.execute("SELECT nome FROM pessoas ORDER BY nome")]


def bens_da_pessoa(conn, nome: str) -> list[dict]:
    return _todos(conn, """
        SELECT b.* FROM atribuicoes a JOIN bens b ON b.numero = a.numero
        WHERE a.nome = ? ORDER BY b.numero""", nome)


def buscar_bem(conn, numero: int) -> dict | None:
    return _um(conn, "SELECT * FROM bens WHERE numero = ?", numero)


def pessoa_do_bem(conn, numero: int) -> str | None:
    r = conn.execute("SELECT nome FROM atribuicoes WHERE numero = ?", (numero,)).fetchone()
    return r[0] if r else None


def ficha_do_bem(conn, numero: int) -> dict | None:
    """Bem + centro de custo/responsável do setor (pela localização) + pessoa (pela atribuição)."""
    bem = buscar_bem(conn, numero)
    if not bem:
        return None
    setor = _um(conn, """
        SELECT r.ccustos, r.responsavel FROM localizacoes l JOIN responsaveis r ON r.ccustos = l.ccustos
        WHERE l.localizacao = ?""", bem["localizacao"]) or {"ccustos": None, "responsavel": None}
    return {**bem, **setor, "pessoa": pessoa_do_bem(conn, numero)}


def localizacoes_mapeadas(conn) -> list[dict]:
    return _todos(conn, "SELECT localizacao, ccustos FROM localizacoes ORDER BY localizacao")
```

- [ ] **Step 4: Rodar e ver passar**

Run: `.venv/bin/pytest tests/test_db.py -v`
Expected: 12 passed

- [ ] **Step 5: Commit**

```bash
git add db.py tests/test_db.py
git commit -m "db: consultas dos termos, ficha do bem e regra setor-ou-pessoa"
```

---

### Task 5: `db.py` — operações de cadastro

**Files:**
- Modify: `db.py`, `tests/test_db.py`

**Interfaces:**
- Produces: `incluir_responsavel(conn, dados: dict)`, `excluir_responsavel(conn, ccustos)`, `renomear_centro(conn, antigo, novo)`, `incluir_localizacao(conn, localizacao, ccustos)`, `excluir_localizacao(conn, localizacao)`, `incluir_pessoa(conn, nome)`, `excluir_pessoa(conn, nome)`, `atribuir(conn, nome, numero, confirmar=False)`, `desatribuir(conn, nome, numero)`. Todas fazem `commit`. Exceções (subclasses de `ErroDeNegocio`): `CentroEmUso`, `BemNaoEncontrado`, `JaAtribuido` (atributo `.pessoa`).

- [ ] **Step 1: Testes que falham**

```python
def test_renomear_centro_cascateia(dados):
    semear(dados)
    db.renomear_centro(dados, "CCI", "GESERV")
    assert db.localizacoes_mapeadas(dados)[0]["ccustos"] == "GESERV"
    assert db.responsavel(dados, "CCI") is None and db.responsavel(dados, "GESERV")


def test_excluir_centro_em_uso_falha(dados):
    semear(dados)
    with pytest.raises(db.CentroEmUso):
        db.excluir_responsavel(dados, "CCI")
    db.excluir_localizacao(dados, "01 - SALA CCI")
    db.excluir_responsavel(dados, "CCI")
    assert db.centros(dados) == []


def test_incluir_responsavel_e_localizacao(dados):
    semear(dados)
    db.incluir_responsavel(dados, {"ccustos": " decom ", "tratamento": "Prezado", "responsavel": "THIAGO",
                                   "email": "t@cfc", "matricula": "481", "funcao": "gerente"})
    assert db.responsavel(dados, "DECOM")["responsavel"] == "THIAGO"
    db.incluir_localizacao(dados, "99 - SEM MAPA", "DECOM")
    assert db.localizacoes_sem_centro(dados) == []
    with pytest.raises(db.ErroDeNegocio):
        db.incluir_responsavel(dados, {"ccustos": "", "responsavel": "X"})


def test_atribuir_bem_livre_e_transferencia(dados):
    semear(dados)
    db.incluir_pessoa(dados, "  BRUNO LIMA ")
    db.atribuir(dados, "BRUNO LIMA", 1001)
    assert [b["numero"] for b in db.bens_da_pessoa(dados, "BRUNO LIMA")] == [1001]
    with pytest.raises(db.BemNaoEncontrado):
        db.atribuir(dados, "BRUNO LIMA", 9999)
    with pytest.raises(db.JaAtribuido) as e:
        db.atribuir(dados, "BRUNO LIMA", 1002)
    assert e.value.pessoa == "ANA SILVA"
    db.atribuir(dados, "BRUNO LIMA", 1002, confirmar=True)
    assert db.bens_da_pessoa(dados, "ANA SILVA") == []
    assert db.pessoa_do_bem(dados, 1002) == "BRUNO LIMA"
    db.atribuir(dados, "BRUNO LIMA", 1002)  # já é dele: não é erro


def test_desatribuir_e_excluir_pessoa(dados):
    semear(dados)
    db.desatribuir(dados, "ANA SILVA", 1002)
    assert [b["numero"] for b in db.bens_do_centro(dados, "CCI")] == [1001, 1002]
    db.atribuir(dados, "ANA SILVA", 1002)
    db.excluir_pessoa(dados, "ANA SILVA")
    assert db.pessoas(dados) == [] and db.pessoa_do_bem(dados, 1002) is None
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_db.py -v -k "renomear or excluir or incluir or atribuir"`
Expected: FAIL — `AttributeError`

- [ ] **Step 3: Implementar**

```python
class CentroEmUso(ErroDeNegocio):
    pass


class BemNaoEncontrado(ErroDeNegocio):
    pass


class JaAtribuido(ErroDeNegocio):
    def __init__(self, numero, pessoa):
        super().__init__(f"O bem {numero} está com {pessoa}.")
        self.pessoa = pessoa


def _obrigatorio(valor, rotulo) -> str:
    v = " ".join(str(valor or "").split())
    if not v:
        raise ErroDeNegocio(f"{rotulo} é obrigatório.")
    return v


def incluir_responsavel(conn, dados: dict) -> None:
    sigla = _obrigatorio(dados.get("ccustos"), "Centro de custo").upper()
    nome = _obrigatorio(dados.get("responsavel"), "Responsável")
    if responsavel(conn, sigla):
        raise ErroDeNegocio(f"O centro de custo {sigla} já existe.")
    conn.execute("INSERT INTO responsaveis VALUES (?,?,?,?,?,?)", (
        sigla, _texto(dados.get("tratamento")), nome, _texto(dados.get("email")),
        _texto(dados.get("matricula")), _texto(dados.get("funcao"))))
    conn.commit()


def excluir_responsavel(conn, ccustos: str) -> None:
    n = conn.execute("SELECT count(*) FROM localizacoes WHERE ccustos = ?", (ccustos,)).fetchone()[0]
    if n:
        raise CentroEmUso(f"{ccustos} tem {n} localização(ões) mapeada(s). Remapeie-as antes de excluir.")
    conn.execute("DELETE FROM responsaveis WHERE ccustos = ?", (ccustos,))
    conn.commit()


def renomear_centro(conn, antigo: str, novo: str) -> None:
    novo = _obrigatorio(novo, "Nova sigla").upper()
    if responsavel(conn, novo):
        raise ErroDeNegocio(f"Já existe um centro de custo {novo}.")
    if not responsavel(conn, antigo):
        raise ErroDeNegocio(f"Centro de custo {antigo} não encontrado.")
    conn.execute("UPDATE responsaveis SET ccustos = ? WHERE ccustos = ?", (novo, antigo))  # cascateia
    conn.commit()


def incluir_localizacao(conn, localizacao: str, ccustos: str) -> None:
    loc = _obrigatorio(localizacao, "Localização")
    if not responsavel(conn, ccustos):
        raise ErroDeNegocio(f"Centro de custo {ccustos} não cadastrado.")
    conn.execute("INSERT OR REPLACE INTO localizacoes VALUES (?, ?)", (loc, ccustos))
    conn.commit()


def excluir_localizacao(conn, localizacao: str) -> None:
    conn.execute("DELETE FROM localizacoes WHERE localizacao = ?", (localizacao,))
    conn.commit()


def incluir_pessoa(conn, nome: str) -> str:
    nome = _obrigatorio(nome, "Nome").upper()
    conn.execute("INSERT OR IGNORE INTO pessoas VALUES (?)", (nome,))
    conn.commit()
    return nome


def excluir_pessoa(conn, nome: str) -> None:
    conn.execute("DELETE FROM pessoas WHERE nome = ?", (nome,))  # atribuições vão junto (cascade)
    conn.commit()


def atribuir(conn, nome: str, numero: int, confirmar: bool = False) -> None:
    """Coloca o bem sob responsabilidade da pessoa. Se já está com outra, exige confirmar=True."""
    if not buscar_bem(conn, numero):
        raise BemNaoEncontrado(f"Bem {numero} não encontrado na base.")
    atual = pessoa_do_bem(conn, numero)
    if atual == nome:
        return
    if atual and not confirmar:
        raise JaAtribuido(numero, atual)
    conn.execute("DELETE FROM atribuicoes WHERE numero = ?", (numero,))
    conn.execute("INSERT INTO atribuicoes VALUES (?, ?)", (nome, numero))
    conn.commit()


def desatribuir(conn, nome: str, numero: int) -> None:
    conn.execute("DELETE FROM atribuicoes WHERE nome = ? AND numero = ?", (nome, numero))
    conn.commit()
```

- [ ] **Step 4: Rodar e ver passar**

Run: `.venv/bin/pytest tests/test_db.py -v`
Expected: 17 passed

- [ ] **Step 5: Commit**

```bash
git add db.py tests/test_db.py
git commit -m "db: cadastros de responsáveis, localizações, pessoas e atribuições"
```

---

### Task 6: `importar_planilhas.py` — migração inicial

**Files:**
- Create: `importar_planilhas.py`, `tests/test_importar_planilhas.py`

**Interfaces:**
- Consumes: `db.conectar`, `db.criar_esquema`, `db.importar_bens`, `db._texto`.
- Produces: `importar_planilhas.migrar(conn, acervo: Path, geral: Path) -> dict` (contagens + avisos); CLI `python importar_planilhas.py [acervo.xlsx] [geral.xlsx]`.

- [ ] **Step 1: Teste que falha**

`tests/test_importar_planilhas.py`:

```python
from openpyxl import Workbook

import db
import importar_planilhas as ip
from tests.test_db import CABECALHO


def planilhas(tmp_path):
    acervo = Workbook()
    ws = acervo.active
    ws.title = "acervo"
    ws.append(["numero", "situacao"])  # aba ignorada
    r = acervo.create_sheet("responsavel")
    r.append(["ccustos", "tratamento", "responsavel", "email", "matricula", "funcao"])
    r.append(["CCI", "Prezada", "JAQUELINE", "j@cfc", 46, "coordenadora"])
    c = acervo.create_sheet("ccustos")
    c.append(["localizacao", "ccustos"])
    c.append(["01 - SALA CCI", "CCI"])
    c.append(["02 - GAB", "TERMOS INDIVIDUAIS"])
    c.append(["03 - DEPOSITO", "SEPAT"])  # sem responsável cadastrado
    pa = tmp_path / "acervo.xlsx"
    acervo.save(pa)

    geral = Workbook()
    d = geral.active
    d.title = "dados"
    d.append(["Nome", "Patrimônio", "Situação", "Descrição", "Complemento", "Valor Atual"])
    d.append(["ANA SILVA", 1002, None, None, None, None])
    d.append(["CARLOS ", 1001, None, None, None, None])
    b = geral.create_sheet("base")
    b.append(CABECALHO)
    b.append([1001, "ATIVO", "CADEIRA", "", "MÓVEIS", "01 - SALA CCI", "31/12/1996", 75.94, 64.54])
    b.append([1002, "ATIVO", "NOTEBOOK", "DELL", "EQUIP", "02 - GAB", "06/12/2012", 3000, 1500])
    n = geral.create_sheet("nomes")
    n.append(["Nº", "responsavel"])
    n.append([1, "ANA SILVA"])
    n.append([2, "DANIEL"])
    pg = tmp_path / "geral.xlsx"
    geral.save(pg)
    return pa, pg


def test_migrar_popula_cinco_tabelas(dados, tmp_path):
    pa, pg = planilhas(tmp_path)
    resumo = ip.migrar(dados, pa, pg)
    assert resumo["bens"] == 2 and resumo["responsaveis"] == 2 and resumo["localizacoes"] == 2
    assert db.pessoas(dados) == ["ANA SILVA", "CARLOS", "DANIEL"]
    assert db.pessoa_do_bem(dados, 1001) == "CARLOS"
    assert db.responsavel(dados, "SEPAT")["responsavel"] == "(preencher)"
    assert [l["localizacao"] for l in db.localizacoes_mapeadas(dados)] == ["01 - SALA CCI", "03 - DEPOSITO"]
    assert "SEPAT" in resumo["avisos"][0]
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_importar_planilhas.py -v`
Expected: FAIL — `ModuleNotFoundError: importar_planilhas`

- [ ] **Step 3: Implementar**

```python
"""Migração inicial: popula termos.db a partir de acervo.xlsx e geral.xlsx. Roda UMA vez.

    .venv/bin/python importar_planilhas.py [acervo.xlsx] [geral.xlsx]
"""
import sys
from pathlib import Path

from openpyxl import load_workbook

import db


def _linhas(caminho: Path, aba: str) -> list[dict]:
    wb = load_workbook(caminho, read_only=True, data_only=True)
    ws = wb[aba]
    it = ws.iter_rows(values_only=True)
    cab = [db._texto(c) for c in next(it)]
    saida = [dict(zip(cab, r)) for r in it if any(v is not None for v in r)]
    wb.close()
    return saida


def migrar(conn, acervo: Path, geral: Path) -> dict:
    avisos = []
    db.criar_esquema(conn)
    for t in ("atribuicoes", "pessoas", "localizacoes", "responsaveis"):
        conn.execute(f"DELETE FROM {t}")
    conn.commit()

    resumo_bens = db.importar_bens(conn, geral)  # acha a aba 'base' pelo cabeçalho

    for r in _linhas(acervo, "responsavel"):
        sigla = db._texto(r["ccustos"]).upper()
        if not sigla:
            continue
        conn.execute("INSERT OR REPLACE INTO responsaveis VALUES (?,?,?,?,?,?)", (
            sigla, db._texto(r.get("tratamento")), db._texto(r.get("responsavel")) or "(preencher)",
            db._texto(r.get("email")), db._texto(r.get("matricula")), db._texto(r.get("funcao"))))

    for r in _linhas(acervo, "ccustos"):
        loc, sigla = db._texto(r["localizacao"]), db._texto(r["ccustos"]).upper()
        if not loc or not sigla or sigla == "TERMOS INDIVIDUAIS":
            continue
        if not db.responsavel(conn, sigla):
            conn.execute("INSERT INTO responsaveis (ccustos, responsavel) VALUES (?, '(preencher)')", (sigla,))
            avisos.append(f"Centro de custo {sigla} sem responsável: criado como '(preencher)'.")
        conn.execute("INSERT OR REPLACE INTO localizacoes VALUES (?, ?)", (loc, sigla))

    nomes = {db._texto(r["responsavel"]).upper() for r in _linhas(geral, "nomes")}
    dados = _linhas(geral, "dados")
    nomes |= {db._texto(r["Nome"]).upper() for r in dados}
    conn.executemany("INSERT OR IGNORE INTO pessoas VALUES (?)", [(n,) for n in sorted(nomes) if n])

    for r in dados:
        nome, num = db._texto(r["Nome"]).upper(), db._numero(r["Patrimônio"])
        if not nome or num is None:
            continue
        if not db.buscar_bem(conn, int(num)):
            avisos.append(f"Bem {int(num)} de {nome} não existe na base; atribuição ignorada.")
            continue
        conn.execute("INSERT OR IGNORE INTO atribuicoes VALUES (?, ?)", (nome, int(num)))
    conn.commit()

    contar = lambda t: conn.execute(f"SELECT count(*) FROM {t}").fetchone()[0]
    return {"bens": resumo_bens["total"], "responsaveis": contar("responsaveis"),
            "localizacoes": contar("localizacoes"), "pessoas": contar("pessoas"),
            "atribuicoes": contar("atribuicoes"), "sem_centro": resumo_bens["sem_centro"], "avisos": avisos}


if __name__ == "__main__":
    acervo = Path(sys.argv[1] if len(sys.argv) > 1 else "acervo.xlsx")
    geral = Path(sys.argv[2] if len(sys.argv) > 2 else "geral.xlsx")
    db.inicializar()
    conn = db.conectar()
    r = migrar(conn, acervo, geral)
    conn.close()
    for k, v in r.items():
        if k not in ("avisos", "sem_centro"):
            print(f"{k}: {v}")
    print("localizações ativas sem centro de custo:", ", ".join(r["sem_centro"]) or "nenhuma")
    for a in r["avisos"]:
        print("AVISO", a)
```

- [ ] **Step 4: Rodar e ver passar**

Run: `.venv/bin/pytest tests/test_importar_planilhas.py -v`
Expected: 1 passed

- [ ] **Step 5: Rodar a migração real e conferir**

Run: `.venv/bin/python importar_planilhas.py`
Expected (aproximado): `bens: 7003`, `responsaveis: 31` ou 32, `localizacoes: ~100`, `pessoas: ~198`, `atribuicoes: 114`, 8 localizações sem centro listadas, aviso sobre `SEPAT` se não tiver responsável. Conferir com:

`.venv/bin/python -c "import db; c=db.conectar(); print(c.execute('select count(*) from bens where situacao=\"ATIVO\"').fetchone()[0])"` → 3096.

`dados/` está no `.gitignore`; o banco não é versionado.

- [ ] **Step 6: Commit**

```bash
git add importar_planilhas.py tests/test_importar_planilhas.py
git commit -m "Migração inicial das planilhas para o SQLite"
```

---

### Task 7: Geradores `.docx` sem pandas, com listas de dicts

**Files:**
- Modify: `Script_Termo_Individual.py`, `Termo_de_Responsabilidade.py`, `termo_devolucao.py`
- Create: `tests/test_docx.py`

**Interfaces:**
- Produces:
  `Script_Termo_Individual.criar_termo_responsabilidade(nome: str, bens: list[dict], destino: Path) -> Path`
  `Termo_de_Responsabilidade.gerar_termo_centro(ccustos: str, responsavel: dict, bens: list[dict], destino: Path) -> Path`
  `Termo_de_Responsabilidade.gerar_planilha_centro(bens: list[dict], destino: Path) -> Path`
  `termo_devolucao.gerar_termo_devolucao(nome: str, bens: list[dict], destino: Path) -> Path | None`
  `bens` têm as chaves de `db.bens` (`numero, descricao, complemento, localizacao, valor_atual`). O modelo vem de `config.caminho_timbrado()`.

- [ ] **Step 1: Testes que falham**

`tests/test_docx.py`:

```python
from docx import Document
from openpyxl import load_workbook

from tests.conftest import semear


def bens():
    return [
        {"numero": 1001, "descricao": "CADEIRA", "complemento": "GIRATÓRIA", "localizacao": "01 - SALA", "valor_atual": 64.54},
        {"numero": 1002, "descricao": "NOTEBOOK", "complemento": None, "localizacao": "01 - SALA", "valor_atual": 1500.0},
    ]


def linhas(caminho):
    return Document(caminho).tables[0].rows


def test_termo_individual(dados, tmp_path):
    from Script_Termo_Individual import criar_termo_responsabilidade
    destino = criar_termo_responsabilidade("ANA SILVA", bens(), tmp_path / "t.docx")
    tab = linhas(destino)
    assert len(tab) == 4  # cabeçalho + 2 + total
    assert tab[3].cells[3].text == "R$ 1.564,54"
    assert tab[2].cells[2].text == ""  # complemento None vira vazio


def test_termo_centro_e_planilha(dados, tmp_path):
    from Termo_de_Responsabilidade import gerar_planilha_centro, gerar_termo_centro
    resp = {"ccustos": "CCI", "responsavel": "JAQUELINE", "matricula": "46", "funcao": "coordenadora"}
    destino = gerar_termo_centro("CCI", resp, bens(), tmp_path / "c.docx")
    tab = linhas(destino)
    assert len(tab) == 4 and tab[3].cells[4].text == "R$ 1.564,54"
    assert "JAQUELINE" in Document(destino).paragraphs[-1].text
    xlsx = gerar_planilha_centro(bens(), tmp_path / "c.xlsx")
    ws = load_workbook(xlsx).active
    assert ws.max_row == 3 and ws["A1"].value == "numero"


def test_termo_devolucao(dados, tmp_path):
    from termo_devolucao import gerar_termo_devolucao
    destino = gerar_termo_devolucao("ANA SILVA", bens(), tmp_path / "d.docx")
    assert len(linhas(destino)) == 4
    assert gerar_termo_devolucao("ANA SILVA", [], tmp_path / "vazio.docx") is None
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_docx.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'pandas'` (os módulos ainda importam pandas)

- [ ] **Step 3: `Script_Termo_Individual.py`**

Trocar o topo e a assinatura; o miolo da tabela e o texto não mudam:

```python
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Pt

import config
```

(remover `import pandas as pd`). Assinatura e abertura:

```python
def criar_termo_responsabilidade(nome, bens, destino):
    """bens: lista de dicts com numero, descricao, complemento, valor_atual. Grava em destino."""
    doc = Document(str(config.caminho_timbrado()))
```

Laço das linhas (substitui o `for _, bem in bens.iterrows()`):

```python
    for bem in bens:
        row_cells = tabela.add_row().cells
        row_cells[0].text = str(bem["numero"])
        row_cells[1].text = bem["descricao"] or ""
        row_cells[2].text = bem["complemento"] or ""
        row_cells[3].text = formatar_moeda(bem["valor_atual"] or 0)
        for cell in row_cells:
            centralizar_celula(cell)

    valor_total = sum(b["valor_atual"] or 0 for b in bens)
```

Final:

```python
    doc.save(str(destino))
    return Path(destino)
```

Apagar o bloco `if __name__ == "__main__":`.

- [ ] **Step 4: `Termo_de_Responsabilidade.py`**

Reescrever mantendo `texto_padrao`, a tabela e a assinatura exatamente como estão:

```python
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook

import config


def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")


TEXTO_PADRAO = """
    Pelo presente termo, eu, {responsavel}, ... (copiar o texto_padrao atual, sem alterar)
    """


def gerar_termo_centro(ccustos, responsavel, bens, destino):
    """Um termo para um centro de custo. responsavel: dict de db.responsaveis; bens: dicts de db.bens."""
    documento = Document(str(config.caminho_timbrado()))
    cabecalho = documento.add_paragraph()
    cabecalho_run = cabecalho.add_run(f'Termo de Responsabilidade - {ccustos}')
    cabecalho_run.font.bold = True
    cabecalho.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cabecalho_run.font.size = Pt(16)

    soma_valores = sum(b["valor_atual"] or 0 for b in bens)

    for paragraph in TEXTO_PADRAO.strip().split('\n'):
        paragrafo = documento.add_paragraph()
        paragrafo.paragraph_format.first_line_indent = Inches(0.59)
        paragrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragrafo.add_run(paragraph.format(
            responsavel=responsavel['responsavel'], matricula=responsavel['matricula'],
            funcao=responsavel['funcao'], ccustos=ccustos))

    # ... tabela: igual à atual, com o laço abaixo no lugar de iterrows ...
    for bem in sorted(bens, key=lambda b: b["numero"]):
        row_cells = tabela.add_row().cells
        row_cells[0].text = str(bem['numero'])
        row_cells[1].text = bem['descricao'] or ""
        row_cells[2].text = bem['complemento'] or ""
        row_cells[3].text = bem['localizacao'] or ""
        # (formatação do valor e das fontes: igual à atual)

    # ... linha TOTAL e assinatura: iguais às atuais ...
    documento.save(str(destino))
    return Path(destino)


def gerar_planilha_centro(bens, destino):
    """Planilha com os bens do termo, para quem quiser analisar os dados."""
    wb = Workbook()
    ws = wb.active
    ws.title = "bens"
    ws.append(["numero", "descricao", "complemento", "localizacao", "valor_atual"])
    for b in sorted(bens, key=lambda b: b["numero"]):
        ws.append([b["numero"], b["descricao"], b["complemento"], b["localizacao"], b["valor_atual"]])
    wb.save(str(destino))
    return Path(destino)
```

A função antiga `gerar_termos(filtrar_ccusto=None)` é removida.

- [ ] **Step 5: `termo_devolucao.py`**

Remover `import pandas as pd`; acrescentar `from pathlib import Path` e `import config`. Assinatura:

```python
def gerar_termo_devolucao(nome, bens, destino):
    """bens: dicts de db.bens. Devolve None se a lista estiver vazia."""
    if not bens:
        return None
    doc = Document(str(config.caminho_timbrado()))
```

Laço e total:

```python
    for bem in bens:
        row_cells = tabela.add_row().cells
        row_cells[0].text = str(bem['numero'])
        row_cells[1].text = bem['descricao'] or ""
        row_cells[2].text = bem['complemento'] or ""
        row_cells[3].text = formatar_moeda(bem['valor_atual'] or 0)
        for cell in row_cells:
            centralizar_celula(cell)

    total = sum(b["valor_atual"] or 0 for b in bens)
```

Final: `doc.save(str(destino)); return Path(destino)`.

- [ ] **Step 6: Rodar e ver passar**

Run: `.venv/bin/pytest tests/test_docx.py -v`
Expected: 3 passed. Também: `grep -l pandas *.py` → nenhum resultado.

- [ ] **Step 7: Commit**

```bash
git add Script_Termo_Individual.py Termo_de_Responsabilidade.py termo_devolucao.py tests/test_docx.py
git commit -m "Geradores .docx recebem listas de dicts; pandas removido"
```

---

### Task 8: `termos_html.py` — documento HTML no padrão gelic

**Files:**
- Create: `termos_html.py`, `templates/termo_base.html`, `tests/test_termos_html.py`

**Interfaces:**
- Produces: `termos_html.corpo_ccusto(ccustos, responsavel: dict, bens) -> str`, `corpo_individual(nome, bens) -> str`, `corpo_devolucao(nome, bens, hoje: date | None = None) -> str`, `documento(titulo, corpo) -> str` (HTML completo), `formatar_moeda(v) -> str`.

- [ ] **Step 1: Testes que falham**

`tests/test_termos_html.py`:

```python
from datetime import date

import termos_html as th
from tests.test_docx import bens


def test_individual_tabela_80_por_cento_e_total():
    html = th.corpo_individual("ANA SILVA", bens())
    assert "width:80%" in html and "TERMO DE RESPONSABILIDADE" in html
    assert "<b>ANA SILVA</b>" in html and "R$ 1.564,54" in html
    assert html.count("<tr") == 4  # cabeçalho + 2 + total


def test_ccusto_tabela_100_por_cento_e_texto_do_responsavel():
    resp = {"ccustos": "CCI", "responsavel": "JAQUELINE", "matricula": "46", "funcao": "coordenadora"}
    html = th.corpo_ccusto("CCI", resp, bens())
    assert "width:100%" in html and "Termo de Responsabilidade - CCI" in html
    assert "matrícula n.º 46" in html and "Localização" in html


def test_devolucao_com_data_e_assinaturas():
    html = th.corpo_devolucao("ANA SILVA", bens(), hoje=date(2026, 9, 14))
    assert "width:80%" in html and "TERMO DE DEVOLUÇÃO" in html
    assert "Brasília (DF), 14 de setembro de 2026" in html
    assert "Supervisor de Patrimônio" in html


def test_documento_envelopa_com_titulo_e_escapa():
    html = th.documento("Termo", th.corpo_individual("A <B>", []))
    assert html.startswith("<!DOCTYPE html>") and "<title>Termo</title>" in html
    assert "A &lt;B&gt;" in html
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_termos_html.py -v`
Expected: FAIL — `ModuleNotFoundError: termos_html`

- [ ] **Step 3: `templates/termo_base.html`** (folha de estilo copiada de `~/.claude/skills/gelic/assets/base.html`, sem as classes específicas de despacho)

```html
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>{{titulo}}</title>
<style>
  @page { size: A4; margin: 2cm 1.8cm 2cm 2cm; }
  body { font-family: "Times New Roman"; font-size: 11.5pt; line-height: 1.3; color: #000; }
  h1 { font-size: 14pt; text-align: center; margin: 10pt 0 12pt; }
  p { text-align: justify; margin: 0 0 7pt; text-indent: 1.25cm; }
  p.semrecuo { text-indent: 0; }
  p.centro { text-align: center; text-indent: 0; }
  p.direita { text-align: right; text-indent: 0; }
  .assinatura { margin-top: 20pt; text-indent: 0; text-align: center; }
</style>
</head>
<body>
{{corpo}}
</body>
</html>
```

- [ ] **Step 4: `termos_html.py`**

```python
"""Corpo HTML dos três termos, no padrão do /gelic: funções pequenas que montam o HTML com html.escape.

As tabelas levam style= inline de propósito: é o que a área de transferência carrega para o SEI.
É o único lugar do projeto com inline style.
"""
import html
from datetime import date
from pathlib import Path

import config

TABELA = "border-collapse:collapse;width:{largura};margin:8pt auto;font-size:10.5pt"
TH = "border:1px solid #000;padding:3pt 5pt;text-align:center;background:#e6e6e6;font-weight:bold"
TD = "border:1px solid #000;padding:3pt 5pt;text-align:{alinhamento}"

MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
         'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']


def formatar_moeda(valor) -> str:
    return f"R$ {valor or 0:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")


def esc(s) -> str:
    return html.escape("" if s is None else str(s))


def tabela(colunas: list[str], linhas: list[list], largura: str, total: float, colunas_total: int) -> str:
    """colunas_total: quantas colunas o rótulo TOTAL ocupa (mescladas); o valor vai na última."""
    cab = "".join(f'<th style="{TH}">{esc(c)}</th>' for c in colunas)
    corpo = ""
    for linha in linhas:
        celulas = ""
        for i, v in enumerate(linha):
            alinhamento = "right" if i == len(linha) - 1 else "center" if i == 0 else "left"
            celulas += f'<td style="{TD.format(alinhamento=alinhamento)}">{esc(v)}</td>'
        corpo += f"<tr>{celulas}</tr>"
    rodape = (f'<tr><td colspan="{colunas_total}" style="{TD.format(alinhamento="center")};font-weight:bold">TOTAL</td>'
              f'<td style="{TD.format(alinhamento="right")};font-weight:bold">{formatar_moeda(total)}</td></tr>')
    return (f'<table style="{TABELA.format(largura=largura)}"><thead><tr>{cab}</tr></thead>'
            f'<tbody>{corpo}{rodape}</tbody></table>')


def _total(bens) -> float:
    return sum(b["valor_atual"] or 0 for b in bens)


COMPROMISSOS_INDIVIDUAL = [
    "1) zelar pela guarda, uso adequado e conservação do(s) bem(ns), utilizando-o(s) exclusivamente para fins profissionais do CFC;",
    "2) informar imediatamente ao Setor de Patrimônio qualquer dano, inutilização, perda ou roubo, apresentando boletim de ocorrência quando necessário;",
    "3) ressarcir o CFC por danos ou perdas decorrentes de negligência do responsável, após decisão da Câmara de Assuntos Administrativos (CAD) e homologação pelo Plenário do CFC, em conformidade com o Manual de Gestão Patrimonial do CFC;",
    "4) devolver o(s) equipamento(s) e acessórios ao término do vínculo, mediante solicitação ou em caso de substituição, em condições compatíveis com o uso; e",
    "5) fornecer informações sobre o(s) bem(ns) sempre que solicitado, especialmente durante o inventário patrimonial.",
]


def corpo_individual(nome: str, bens: list[dict]) -> str:
    linhas = [[b["numero"], b["descricao"], b["complemento"], formatar_moeda(b["valor_atual"])] for b in bens]
    return (
        "<h1>TERMO DE RESPONSABILIDADE</h1>"
        f"<p class=\"semrecuo\">Pelo presente termo, eu, <b>{esc(nome)}</b>, declaro que o(s) equipamento(s) abaixo "
        "discriminado(s) se encontra(m) sob a minha guarda e responsabilidade.</p>"
        + tabela(["Patrimônio", "Descrição", "Complemento", "Valor Atual"], linhas, "80%", _total(bens), 3)
        + "<p class=\"semrecuo\">Comprometo-me a:</p>"
        + "".join(f"<p class=\"semrecuo\">{esc(c)}</p>" for c in COMPROMISSOS_INDIVIDUAL)
        + "<p class=\"semrecuo\">Declaro estar ciente das responsabilidades mencionadas acima e assumo total "
        "responsabilidade pelos bens listados.</p>"
        f"<p class=\"assinatura\"><b>{esc(nome)}</b><br>Assinado eletronicamente via SEI</p>"
    )


PARAGRAFOS_CCUSTO = [
    "Pelo presente termo, eu, {responsavel}, matrícula n.º {matricula}, {funcao} do(a) {ccustos} do CFC, declaro que os bens patrimoniais abaixo discriminados se encontram na localização sob a minha guarda e responsabilidade.",
    "Assumo TOTAL responsabilidade pelos referidos bens, comprometendo-me a informar o Setor de Patrimônio quanto a qualquer alteração e/ou irregularidade, bem como zelar pela guarda e bom uso do patrimônio público.",
    "Em caso de extravio ou dano a bem sob a minha responsabilidade, comprometo-me a ressarcir o CFC dos prejuízos causados.",
    "Observações:",
    "Em caso de perda ou roubo do bem, o responsável deverá registrar boletim de ocorrência policial e apresentar ao Setor de Patrimônio;",
    "Ao final do mandato, função ou designação, o responsável deverá devolver o bem, se for o caso.",
    "No caso de movimentação e transferência de bens entre as unidades administrativas, o Setor de Patrimônio utilizará o Termo de Transferência disponível no SEI, que será apensado a processo específico até a emissão de um novo termo atualizado.",
]


def corpo_ccusto(ccustos: str, responsavel: dict, bens: list[dict]) -> str:
    campos = {k: esc(responsavel.get(k)) for k in ("responsavel", "matricula", "funcao")}
    campos["ccustos"] = esc(ccustos)
    linhas = [[b["numero"], b["descricao"], b["complemento"], b["localizacao"], formatar_moeda(b["valor_atual"])]
              for b in sorted(bens, key=lambda b: b["numero"])]
    return (
        f"<h1>Termo de Responsabilidade - {esc(ccustos)}</h1>"
        + "".join(f"<p>{p.format(**campos)}</p>" for p in PARAGRAFOS_CCUSTO)
        + tabela(["Número Bem", "Descrição", "Complemento", "Localização", "Valor Atual"], linhas, "100%", _total(bens), 4)
        + f"<p class=\"assinatura\">{campos['responsavel']}<br>{campos['funcao']} do(a) {campos['ccustos']} do CFC</p>"
    )


def corpo_devolucao(nome: str, bens: list[dict], hoje: date | None = None) -> str:
    hoje = hoje or date.today()
    linhas = [[b["numero"], b["descricao"], b["complemento"], formatar_moeda(b["valor_atual"])] for b in bens]
    return (
        "<h1>TERMO DE DEVOLUÇÃO</h1>"
        f"<p class=\"semrecuo\">Pelo presente termo, eu, <b>{esc(nome)}</b>, declaro que devolvi ao Setor de Patrimônio "
        "o(s) bem(ns) abaixo discriminado(s), que se encontrava(m) sob minha guarda e responsabilidade:</p>"
        + tabela(["Patrimônio", "Descrição", "Complemento", "Valor Atual"], linhas, "80%", _total(bens), 3)
        + f"<p class=\"direita\">Brasília (DF), {hoje.day} de {MESES[hoje.month - 1]} de {hoje.year}</p>"
        f"<p class=\"assinatura\"><b>{esc(nome)}</b><br>Assinado eletronicamente via SEI</p>"
        "<p class=\"semrecuo\">Declaro que recebi o(s) bem(ns) acima especificado(s):</p>"
        "<p class=\"assinatura\"><b>ANTÔNIO RODRIGUES DE SOUSA JÚNIOR</b><br>Supervisor de Patrimônio<br>"
        "Assinado eletronicamente via SEI</p>"
    )


def documento(titulo: str, corpo: str) -> str:
    base = (config.pasta_recursos() / "templates" / "termo_base.html").read_text(encoding="utf8")
    return base.replace("{{titulo}}", esc(titulo)).replace("{{corpo}}", corpo)
```

- [ ] **Step 5: Rodar e ver passar**

Run: `.venv/bin/pytest tests/test_termos_html.py -v`
Expected: 4 passed

- [ ] **Step 6: Commit**

```bash
git add termos_html.py templates/termo_base.html tests/test_termos_html.py
git commit -m "termos_html: corpo dos três termos em HTML no padrão gelic (80%/100%)"
```

---

### Task 9: DSGov vendorizado + `base.html` + macros

**Files:**
- Create: `static/dsgov/` (vendor), `templates/base.html`, `templates/_macros.html`
- Delete: `static/style.css`

**Interfaces:**
- Produces: `base.html` com blocos `{% block titulo %}`, `{% block conteudo %}`, `{% block scripts %}`; espera no contexto `trilha` (lista de `(rotulo, url|None)`) e usa `get_flashed_messages`. Macros: `select(nome, rotulo, opcoes, selecionado=None, obrigatorio=True)` (br-select), `cabecalho_tabela(titulo, id)` (br-table com busca), `mensagens()`.
- Consumes (Task 10): `app` registra `context_processor` com `DSGOV` e `MENU`.

- [ ] **Step 1: Copiar os assets**

```bash
S=~/.claude/skills/dsgov/assets
mkdir -p static/dsgov/css static/dsgov/js static/dsgov/vendor/govbr-ds static/dsgov/vendor/fontawesome
cp $S/vendor/govbr-ds/3.7.0/{core-tokens.css,core.min.css,core.min.js,LICENSE} static/dsgov/vendor/govbr-ds/
cp -r $S/vendor/fontawesome/5.15.4/{css,webfonts,LICENSE.txt} static/dsgov/vendor/fontawesome/
cp -r $S/vendor/fonts static/dsgov/vendor/fonts
cp $S/projeto/core/static/dsgov/css/{dsgov.css,fontes.css} static/dsgov/css/
cp $S/projeto/core/static/dsgov/js/dsgov.js static/dsgov/js/
git rm -q static/style.css
du -sh static/dsgov   # ~2 MB
```

Não editar nenhum desses arquivos. (`dsgov.js` referencia eventos do HTMX que nunca disparam aqui — inofensivo.)

- [ ] **Step 2: `templates/base.html`** (tradução Jinja do `base.html` + partials da skill, sem login, busca e HTMX)

```html
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>{% block titulo %}{{ DSGOV.SISTEMA }}{% endblock %} — {{ DSGOV.SISTEMA }}</title>
  <link rel="stylesheet" href="{{ url_for('static', filename='dsgov/css/fontes.css') }}"/>
  <link rel="stylesheet" href="{{ url_for('static', filename='dsgov/vendor/govbr-ds/core.min.css') }}"/>
  <link rel="stylesheet" href="{{ url_for('static', filename='dsgov/vendor/fontawesome/css/all.min.css') }}"/>
  <link rel="stylesheet" href="{{ url_for('static', filename='dsgov/css/dsgov.css') }}"/>
</head>
<body>
<div class="template-base">
  <nav class="br-skiplink" role="menubar">
    <a class="br-item" href="#main-content" role="menuitem" accesskey="1">Ir para o conteúdo <span aria-hidden="true">(1/3)</span></a>
    <a class="br-item" href="#header-navigation" role="menuitem" accesskey="2">Ir para o menu <span aria-hidden="true">(2/3)</span></a>
    <a class="br-item" href="#footer" role="menuitem" accesskey="3">Ir para o rodapé <span aria-hidden="true">(3/3)</span></a>
  </nav>
  <header class="br-header mb-4" data-no-search id="header" data-sticky="data-sticky">
    <div class="container-lg">
      <div class="header-top">
        <div class="header-logo">
          <img src="{{ url_for('static', filename='logo.png') }}" alt="{{ DSGOV.ORGAO }}"/>
          <span class="br-divider vertical"></span>
          <div class="header-sign">{{ DSGOV.ORGAO }}</div>
        </div>
        <div class="header-actions"></div>
      </div>
      <div class="header-bottom">
        <div class="header-menu">
          <div class="header-menu-trigger" id="header-navigation">
            <button class="br-button small circle" type="button" aria-label="Menu" data-toggle="menu" data-target="#main-navigation" id="navigation"><i class="fas fa-bars" aria-hidden="true"></i></button>
          </div>
          <div class="header-info">
            <div class="header-title">{{ DSGOV.SISTEMA }}</div>
            <div class="header-subtitle">{{ DSGOV.SUBTITULO }}</div>
          </div>
        </div>
      </div>
    </div>
  </header>
  <main class="d-flex flex-fill mb-5" id="main">
    <div class="container-lg d-flex">
      <div class="row">
        <div class="br-menu" id="main-navigation">
          <div class="menu-container">
            <div class="menu-panel">
              <div class="menu-header">
                <div class="menu-title"><span>{{ DSGOV.SISTEMA }}</span></div>
                <div class="menu-close">
                  <button class="br-button circle" type="button" aria-label="Fechar o menu" data-dismiss="menu"><i class="fas fa-times" aria-hidden="true"></i></button>
                </div>
              </div>
              <nav class="menu-body" role="tree">
                {% for rotulo, icone, url in MENU %}
                <a class="menu-item" href="{{ url }}" role="treeitem"><span class="icon"><i class="fas {{ icone }}" aria-hidden="true"></i></span><span class="content">{{ rotulo }}</span></a>
                {% endfor %}
              </nav>
              <div class="menu-footer"><div class="menu-info"><div class="text-center text-down-01">{{ DSGOV.ORGAO }}</div></div></div>
            </div>
            <div class="menu-scrim" data-dismiss="menu" tabindex="0"></div>
          </div>
        </div>
        <div class="col mb-5">
          {% if trilha %}
          <nav class="br-breadcrumb" aria-label="Trilha de navegação">
            <ol class="crumb-list" role="list">
              <li class="crumb home"><a class="br-button circle" href="{{ url_for('home') }}"><span class="sr-only">Página inicial</span><i class="fas fa-home" aria-hidden="true"></i></a></li>
              {% for rotulo, url in trilha %}
                {% if url %}<li class="crumb"><i class="icon fas fa-chevron-right" aria-hidden="true"></i><a href="{{ url }}">{{ rotulo }}</a></li>
                {% else %}<li class="crumb" data-active="active"><i class="icon fas fa-chevron-right" aria-hidden="true"></i><span tabindex="0" aria-current="page">{{ rotulo }}</span></li>{% endif %}
              {% endfor %}
            </ol>
          </nav>
          {% endif %}
          <div class="main-content pl-sm-3 mt-4" id="main-content">
            {% from "_macros.html" import mensagens %}{{ mensagens() }}
            {% block conteudo %}{% endblock %}
          </div>
        </div>
      </div>
    </div>
  </main>
  <footer class="br-footer pt-3" id="footer">
    <div class="container-lg">
      <div class="logo"><span class="text-up-02 text-weight-bold">CFC</span></div>
      <span class="br-divider my-3"></span>
      <div class="info"><div class="text-down-01 text-medium pb-3">{{ DSGOV.ORGAO }} · {{ DSGOV.SISTEMA }} · {{ DSGOV.SUBTITULO }}</div></div>
    </div>
  </footer>
  <div class="br-cookiebar default d-none" tabindex="-1"></div>
</div>
<script src="{{ url_for('static', filename='dsgov/vendor/govbr-ds/core.min.js') }}"></script>
<script src="{{ url_for('static', filename='dsgov/js/dsgov.js') }}"></script>
{% block scripts %}{% endblock %}
</body>
</html>
```

- [ ] **Step 3: `templates/_macros.html`**

```html
{# Componentes DSGov repetidos. HTML canônico de references/componentes/{message,select,table}.md #}

{% macro mensagens() %}
<div id="mensagens">
{% for categoria, texto in get_flashed_messages(with_categories=true) %}
  {% set tipo, icone, titulo = {'error': ('danger', 'fa-times-circle', 'Erro.'), 'warning': ('warning', 'fa-exclamation-triangle', 'Atenção.'), 'success': ('success', 'fa-check-circle', 'Sucesso.')}.get(categoria, ('info', 'fa-info-circle', 'Informação.')) %}
  <div class="br-message {{ tipo }}">
    <div class="icon"><i class="fas {{ icone }} fa-lg" aria-hidden="true"></i></div>
    <div class="content" role="alert"><span class="message-title">{{ titulo }}</span><span class="message-body"> {{ texto }}</span></div>
    <div class="close"><button class="br-button circle small" type="button" aria-label="Fechar a mensagem"><i class="fas fa-times" aria-hidden="true"></i></button></div>
  </div>
{% endfor %}
</div>
{% endmacro %}

{% macro select(nome, rotulo, opcoes, selecionado=None, obrigatorio=True) %}
{# br-select: input de exibição (não postado) + lista de radios que carregam name/value #}
<div class="br-select">
  <div class="br-input">
    <label for="{{ nome }}_exibicao">{{ rotulo }}{% if obrigatorio %} <span class="text-danger" aria-hidden="true">*</span>{% endif %}</label>
    <input id="{{ nome }}_exibicao" type="text" placeholder="Selecione" autocomplete="off"/>
    <button class="br-button" type="button" aria-label="Exibir lista" tabindex="-1" data-trigger="data-trigger"><i class="fas fa-angle-down" aria-hidden="true"></i></button>
  </div>
  <div class="br-list" tabindex="0">
    {% for opcao in opcoes %}
    <div class="br-item" tabindex="-1">
      <div class="br-radio">
        <input id="{{ nome }}_{{ loop.index }}" name="{{ nome }}" type="radio" value="{{ opcao }}"{% if opcao == selecionado %} checked="checked"{% endif %}{% if obrigatorio %} required{% endif %}/>
        <label for="{{ nome }}_{{ loop.index }}">{{ opcao }}</label>
      </div>
    </div>
    {% endfor %}
  </div>
</div>
{% endmacro %}

{% macro cabecalho_tabela(titulo, id) %}
{# abre um br-table com busca; quem chama fecha com </table></div> #}
<div class="br-table" data-search="data-search">
  <div class="table-header">
    <div class="top-bar">
      <div class="table-title">{{ titulo }}</div>
      <div class="search-trigger">
        <button class="br-button circle" type="button" id="busca-{{ id }}" data-toggle="search" aria-label="Abrir busca" aria-controls="campo-busca-{{ id }}"><i class="fas fa-search" aria-hidden="true"></i></button>
      </div>
    </div>
    <div class="search-bar">
      <div class="br-input">
        <label for="campo-busca-{{ id }}">Buscar na tabela</label>
        <input id="campo-busca-{{ id }}" type="search" placeholder="Buscar na tabela" aria-labelledby="busca-{{ id }}" aria-label="Buscar na tabela"/>
        <button class="br-button" type="button" aria-label="Buscar"><i class="fas fa-search" aria-hidden="true"></i></button>
      </div>
      <button class="br-button circle" type="button" data-dismiss="search" aria-label="Fechar busca"><i class="fas fa-times" aria-hidden="true"></i></button>
    </div>
  </div>
  <table>
    <caption class="sr-only">{{ titulo }}</caption>
{% endmacro %}
```

- [ ] **Step 4: Verificar que renderiza** (sem app ainda — só sintaxe Jinja)

Run: `.venv/bin/python -c "from jinja2 import Environment, FileSystemLoader; e=Environment(loader=FileSystemLoader('templates')); e.get_template('base.html'); e.get_template('_macros.html'); print('ok')"`
Expected: `ok`

- [ ] **Step 5: Commit**

```bash
git add -A static templates/base.html templates/_macros.html
git commit -m "DSGov 3.7.0 vendorizado: base.html, macros e assets offline"
```

---

### Task 10: `app.py` — núcleo: início, ficha do bem, termos por centro e individual, upload

**Files:**
- Rewrite: `app.py`
- Create: `templates/index.html`, `templates/bem.html`, `templates/centro_custos.html`, `templates/termos_individuais.html`, `templates/termo.html`, `templates/upload.html` (os quatro últimos substituem os atuais), `tests/test_app.py`
- Delete: nada ainda (devolução e cadastros vêm nas Tasks 11 e 12)

**Interfaces:**
- Consumes: tudo de `db`, `termos_html`, os três geradores, `config`.
- Produces rotas: `GET /` (`home`), `GET /bem?numero=` (`bem`), `GET /centro-custos` (`centro_custos`), `POST /gerar` → redirect `/termo/ccusto/<sigla>`, `GET /termos-individuais`, `POST /gerar-individual` → redirect `/termo/individual/<nome>`, `GET /termo/<tipo>/<chave>` (`termo`), `GET /termo/<tipo>/<chave>/documento` (`termo_documento`), `GET /termo/<tipo>/<chave>/docx` (`termo_docx`), `GET /termo/ccusto/<chave>/planilha` (`termo_planilha`), `GET|POST /upload`. Helper `obter_conn()` (conexão em `g`). Função `_bens_do_termo(conn, tipo, chave) -> (titulo, corpo_html, bens, extra)` usada pelas rotas de termo (o ramo `devolucao` já está nela; a rota que alimenta a sessão vem na Task 12).

- [ ] **Step 1: Testes que falham**

`tests/test_app.py`:

```python
import io

import pytest
from openpyxl import Workbook

from tests.conftest import semear
from tests.test_db import CABECALHO


@pytest.fixture
def cliente(dados):
    semear(dados)
    from app import app
    app.config["TESTING"] = True
    with app.test_client() as c:
        yield c


def test_home_e_busca_de_bem(cliente):
    assert cliente.get("/").status_code == 200
    r = cliente.get("/bem?numero=1002")
    assert b"NOTEBOOK" in r.data and b"ANA SILVA" in r.data and b"JAQUELINE PORTELA" in r.data
    r = cliente.get("/bem?numero=9999", follow_redirects=True)
    assert "não encontrado".encode() in r.data


def test_termo_ccusto_documento_docx_planilha(cliente):
    r = cliente.post("/gerar", data={"ccusto": "CCI"})
    assert r.status_code == 302 and r.headers["Location"].endswith("/termo/ccusto/CCI")
    assert cliente.get("/termo/ccusto/CCI").status_code == 200
    doc = cliente.get("/termo/ccusto/CCI/documento").data.decode()
    assert "width:100%" in doc and "1001" in doc and "1002" not in doc  # 1002 está com ANA
    assert cliente.get("/termo/ccusto/CCI/docx").headers["Content-Disposition"].endswith('Termo_de_Responsabilidade_CCI.docx')
    assert cliente.get("/termo/ccusto/CCI/planilha").headers["Content-Disposition"].endswith('planilha_CCI.xlsx')


def test_termo_individual(cliente):
    r = cliente.post("/gerar-individual", data={"nome": "ANA SILVA"})
    assert r.status_code == 302
    doc = cliente.get("/termo/individual/ANA SILVA/documento").data.decode()
    assert "width:80%" in doc and "NOTEBOOK" in doc
    assert cliente.get("/termo/individual/ANA SILVA/docx").status_code == 200
    assert cliente.get("/termo/individual/NINGUEM").status_code == 404


def test_upload_importa_e_lista_sem_centro(cliente):
    wb = Workbook()
    ws = wb.active
    ws.append(CABECALHO)
    ws.append([1002, "ATIVO", "NOTEBOOK", "DELL", "EQ", "01 - SALA CCI", "x", 1, 1])
    ws.append([5000, "ATIVO", "TV", "", "EQ", "77 - NOVA SALA", "x", 1, 1])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    r = cliente.post("/upload", data={"arquivo": (buf, "export.xlsx")}, content_type="multipart/form-data",
                     follow_redirects=True)
    assert "2 bens importados (2 ativos)".encode() in r.data and b"77 - NOVA SALA" in r.data


def test_upload_invalido_mostra_erro(cliente):
    r = cliente.post("/upload", data={"arquivo": (io.BytesIO(b"nada"), "x.txt")}, content_type="multipart/form-data",
                     follow_redirects=True)
    assert b".xlsx" in r.data
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_app.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'pandas'` (o `app.py` atual)

- [ ] **Step 3: Reescrever `app.py`**

```python
"""Termos de Responsabilidade — CFC. Rotas Flask; dados em db.py; documentos em termos_html.py e nos geradores."""
from pathlib import Path

from flask import Flask, abort, flash, g, redirect, render_template, request, send_file, session, url_for

import config
import db
import termos_html
from Script_Termo_Individual import criar_termo_responsabilidade
from Termo_de_Responsabilidade import gerar_planilha_centro, gerar_termo_centro
from termo_devolucao import gerar_termo_devolucao

app = Flask(__name__, template_folder=str(config.pasta_recursos() / "templates"),
            static_folder=str(config.pasta_recursos() / "static"))
app.secret_key = "termos-cfc-local"  # sessão só guarda seleção de bens; programa roda em 127.0.0.1

DSGOV = {"ORGAO": "Conselho Federal de Contabilidade", "SISTEMA": "Termos de Responsabilidade",
         "SUBTITULO": "Setor de Patrimônio"}


@app.context_processor
def contexto_dsgov():
    return {"DSGOV": DSGOV, "MENU": [
        ("Início", "fa-home", url_for("home")),
        ("Termo por centro de custo", "fa-building", url_for("centro_custos")),
        ("Termo individual", "fa-user-check", url_for("termos_individuais")),
        ("Termo de devolução", "fa-box-open", url_for("termo_devolucao")),
        ("Cadastros", "fa-address-book", url_for("cadastros", aba="responsaveis")),
        ("Atualizar base", "fa-upload", url_for("upload")),
    ]}


def obter_conn():
    if "conn" not in g:
        g.conn = db.conectar()
    return g.conn


@app.teardown_appcontext
def fechar_conn(_exc):
    conn = g.pop("conn", None)
    if conn is not None:
        conn.close()


@app.errorhandler(db.ErroDeNegocio)
def erro_de_negocio(e):
    flash(str(e), "error")
    return redirect(request.referrer or url_for("home"))


def _nome_arquivo(s: str) -> str:
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in s)


# ---------------------------------------------------------------- início e ficha do bem
@app.route("/")
def home():
    return render_template("index.html", trilha=[])


@app.route("/bem")
def bem():
    numero = request.args.get("numero", "").strip()
    ficha = db.ficha_do_bem(obter_conn(), int(numero)) if numero.isdigit() else None
    if not ficha:
        flash(f"Bem {numero or '(vazio)'} não encontrado.", "error")
        return redirect(url_for("home"))
    return render_template("bem.html", bem=ficha, trilha=[(f"Bem {numero}", None)])


# ---------------------------------------------------------------- termos
def _bens_do_termo(conn, tipo, chave):
    """Devolve (titulo, corpo_html, bens, extra) do termo pedido; 404 se não existir."""
    if tipo == "ccusto":
        resp = db.responsavel(conn, chave) or abort(404)
        bens = db.bens_do_centro(conn, chave)
        return f"Termo de Responsabilidade - {chave}", termos_html.corpo_ccusto(chave, resp, bens), bens, resp
    if tipo == "individual":
        if chave not in db.pessoas(conn):
            abort(404)
        bens = db.bens_da_pessoa(conn, chave)
        return f"Termo de Responsabilidade - {chave}", termos_html.corpo_individual(chave, bens), bens, None
    if tipo == "devolucao":
        numeros = session.get("bens_selecionados", [])
        bens = [b for b in (db.buscar_bem(conn, int(n)) for n in numeros) if b]
        return f"Termo de Devolução - {chave}", termos_html.corpo_devolucao(chave, bens), bens, None
    abort(404)


@app.route("/centro-custos")
def centro_custos():
    return render_template("centro_custos.html", centros=db.centros(obter_conn()),
                           trilha=[("Termo por centro de custo", None)])


@app.route("/gerar", methods=["POST"])
def gerar():
    return redirect(url_for("termo", tipo="ccusto", chave=request.form["ccusto"]))


@app.route("/termos-individuais")
def termos_individuais():
    return render_template("termos_individuais.html", nomes=db.pessoas(obter_conn()),
                           trilha=[("Termo individual", None)])


@app.route("/gerar-individual", methods=["POST"])
def gerar_individual():
    return redirect(url_for("termo", tipo="individual", chave=request.form["nome"]))


@app.route("/termo/<tipo>/<chave>")
def termo(tipo, chave):
    titulo, _, bens, _ = _bens_do_termo(obter_conn(), tipo, chave)
    return render_template("termo.html", tipo=tipo, chave=chave, titulo=titulo, quantidade=len(bens),
                           trilha=[(titulo, None)])


@app.route("/termo/<tipo>/<chave>/documento")
def termo_documento(tipo, chave):
    titulo, corpo, _, _ = _bens_do_termo(obter_conn(), tipo, chave)
    html = termos_html.documento(titulo, corpo)
    (config.pasta_saida() / f"{_nome_arquivo(titulo)}.html").write_text(html, encoding="utf8")
    return html


@app.route("/termo/<tipo>/<chave>/docx")
def termo_docx(tipo, chave):
    conn = obter_conn()
    _, _, bens, extra = _bens_do_termo(conn, tipo, chave)
    saida = config.pasta_saida()
    if tipo == "ccusto":
        destino = gerar_termo_centro(chave, extra, bens, saida / f"Termo_de_Responsabilidade_{_nome_arquivo(chave)}.docx")
    elif tipo == "individual":
        destino = criar_termo_responsabilidade(chave, bens, saida / f"Termo_{_nome_arquivo(chave)}.docx")
    else:
        destino = gerar_termo_devolucao(chave, bens, saida / f"Termo_Devolucao_{_nome_arquivo(chave)}.docx")
        if destino is None:
            flash("Nenhum bem selecionado.", "error")
            return redirect(url_for("termo_devolucao"))
    return send_file(destino, as_attachment=True, download_name=destino.name)


@app.route("/termo/ccusto/<chave>/planilha")
def termo_planilha(chave):
    _, _, bens, _ = _bens_do_termo(obter_conn(), "ccusto", chave)
    destino = gerar_planilha_centro(bens, config.pasta_saida() / f"planilha_{_nome_arquivo(chave)}.xlsx")
    return send_file(destino, as_attachment=True, download_name=destino.name)


# ---------------------------------------------------------------- atualizar base
@app.route("/upload", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        arquivo = request.files.get("arquivo")
        if not arquivo or not arquivo.filename.lower().endswith(".xlsx"):
            flash("Envie o export do sistema em .xlsx.", "error")
            return redirect(url_for("upload"))
        resumo = db.importar_bens(obter_conn(), arquivo.stream)
        flash(f"{resumo['total']} bens importados ({resumo['ativos']} ativos).", "success")
        return redirect(url_for("upload"))
    return render_template("upload.html", sem_centro=db.localizacoes_sem_centro(obter_conn()),
                           trilha=[("Atualizar base", None)])


if __name__ == "__main__":
    db.inicializar()
    app.run(host="127.0.0.1", port=5000, debug=True)
```

As rotas `termo_devolucao` e `cadastros` referenciadas no menu chegam nas Tasks 11 e 12. **Para esta task compilar**, acrescentar provisoriamente:

```python
@app.route("/termo_devolucao")
def termo_devolucao():
    return redirect(url_for("home"))


@app.route("/cadastros/<aba>")
def cadastros(aba):
    return redirect(url_for("home"))
```

- [ ] **Step 4: Templates**

`templates/index.html`:

```html
{% extends "base.html" %}
{% block titulo %}Início{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Início</h1></div>
<div class="br-card mb-4">
  <div class="card-header"><div class="text-weight-semi-bold text-up-01">Consultar bem</div></div>
  <div class="card-content">
    <form method="get" action="{{ url_for('bem') }}" class="d-flex align-items-end">
      <div class="br-input mr-3">
        <label for="numero">Número do patrimônio</label>
        <input id="numero" name="numero" type="text" inputmode="numeric" placeholder="Ex.: 14359" required/>
      </div>
      <button class="br-button primary" type="submit"><i class="fas fa-search mr-1" aria-hidden="true"></i>Consultar</button>
    </form>
  </div>
</div>
<div class="row">
  {% for rotulo, icone, url in MENU[1:] %}
  <div class="col-sm-6 col-md-4 mb-3">
    <div class="br-card h-100"><div class="card-content">
      <a class="br-button secondary block" href="{{ url }}"><i class="fas {{ icone }} mr-1" aria-hidden="true"></i>{{ rotulo }}</a>
    </div></div>
  </div>
  {% endfor %}
</div>
{% endblock %}
```

`templates/bem.html`:

```html
{% extends "base.html" %}
{% block titulo %}Bem {{ bem.numero }}{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4">
  <h1 class="mb-0">Bem {{ bem.numero }}</h1>
  <div class="ml-auto"><a class="br-button" href="{{ url_for('home') }}">Voltar</a></div>
</div>
<div class="br-card"><div class="card-content">
  <dl class="dsgov-detalhe row">
    <div class="col-md-6"><dt>Descrição</dt><dd>{{ bem.descricao }}</dd></div>
    <div class="col-md-6"><dt>Complemento</dt><dd>{{ bem.complemento or '—' }}</dd></div>
    <div class="col-md-6"><dt>Situação</dt><dd>{{ bem.situacao }}</dd></div>
    <div class="col-md-6"><dt>Localização</dt><dd>{{ bem.localizacao or '—' }}</dd></div>
    <div class="col-md-6"><dt>Valor atual</dt><dd>R$ {{ '%.2f'|format(bem.valor_atual or 0) }}</dd></div>
    <div class="col-md-6"><dt>Centro de custo</dt>
      <dd>{% if bem.ccustos %}<a href="{{ url_for('termo', tipo='ccusto', chave=bem.ccustos) }}">{{ bem.ccustos }}</a> — {{ bem.responsavel }}{% else %}sem mapeamento{% endif %}</dd></div>
    <div class="col-md-6"><dt>Responsável individual</dt>
      <dd>{% if bem.pessoa %}<a href="{{ url_for('termo', tipo='individual', chave=bem.pessoa) }}">{{ bem.pessoa }}</a>{% else %}— (responde o setor){% endif %}</dd></div>
  </dl>
</div></div>
{% endblock %}
```

`templates/centro_custos.html`:

```html
{% extends "base.html" %}
{% from "_macros.html" import select %}
{% block titulo %}Termo por centro de custo{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Termo por centro de custo</h1></div>
<form method="post" action="{{ url_for('gerar') }}" class="col-md-8">
  <div class="mb-3">{{ select('ccusto', 'Centro de custo', centros|map(attribute='ccustos')|list) }}</div>
  <button class="br-button primary" type="submit"><i class="fas fa-file-alt mr-1" aria-hidden="true"></i>Gerar termo</button>
</form>
{% endblock %}
```

`templates/termos_individuais.html`: idêntico, com `h1` "Termo individual", `action="{{ url_for('gerar_individual') }}"`, `select('nome', 'Pessoa', nomes)`.

`templates/termo.html`:

```html
{% extends "base.html" %}
{% block titulo %}{{ titulo }}{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4">
  <h1 class="mb-0">{{ titulo }}</h1>
  <div class="ml-auto">
    <button class="br-button primary" type="button" id="copiar"><i class="fas fa-copy mr-1" aria-hidden="true"></i>Copiar para o SEI</button>
    <a class="br-button secondary ml-2" href="{{ url_for('termo_docx', tipo=tipo, chave=chave) }}"><i class="fas fa-download mr-1" aria-hidden="true"></i>Baixar .docx</a>
    {% if tipo == 'ccusto' %}<a class="br-button ml-2" href="{{ url_for('termo_planilha', chave=chave) }}">Baixar planilha</a>{% endif %}
  </div>
</div>
<div id="aviso-copiado" class="br-message success" hidden>
  <div class="icon"><i class="fas fa-check-circle fa-lg" aria-hidden="true"></i></div>
  <div class="content" role="alert"><span class="message-title">Copiado.</span><span class="message-body"> Cole no editor do SEI (Ctrl+V).</span></div>
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
  var aviso = document.getElementById("aviso-copiado");
  aviso.hidden = false;
  setTimeout(function () { aviso.hidden = true; }, 3000);
});
</script>
{% endblock %}
```

`templates/upload.html`:

```html
{% extends "base.html" %}
{% block titulo %}Atualizar base{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Atualizar base de bens</h1></div>
<form method="post" enctype="multipart/form-data" class="col-md-8 mb-4">
  <p class="text-gray-70">Envie o export do sistema de patrimônio (.xlsx, aba com a coluna <em>Número Bem</em>). A tabela de bens é substituída inteira; responsáveis, localizações, pessoas e atribuições não mudam.</p>
  <div class="br-upload mb-3">
    <label class="upload-label" for="arquivo"><span>Arquivo .xlsx</span></label>
    <input class="upload-input" id="arquivo" name="arquivo" type="file" accept=".xlsx" required/>
    <div class="upload-list"></div>
  </div>
  <button class="br-button primary" type="submit"><i class="fas fa-upload mr-1" aria-hidden="true"></i>Importar</button>
</form>
{% if sem_centro %}
<div class="br-message warning">
  <div class="icon"><i class="fas fa-exclamation-triangle fa-lg" aria-hidden="true"></i></div>
  <div class="content" role="alert"><span class="message-title">Localizações ativas sem centro de custo:</span>
    <span class="message-body"> {{ sem_centro|join(' · ') }}. <a href="{{ url_for('cadastros', aba='localizacoes') }}">Mapear</a>.</span></div>
</div>
{% endif %}
{% endblock %}
```

- [ ] **Step 5: Rodar e ver passar**

Run: `.venv/bin/pytest tests/test_app.py -v`
Expected: 5 passed

- [ ] **Step 6: Verificação manual no navegador**

Run: `.venv/bin/python app.py` e abrir `http://127.0.0.1:5000` (o `dados/termos.db` da Task 6 já existe). Conferir: header e menu lateral abrem (botão ☰), `br-select` abre a lista, termo por centro renderiza no iframe com tabela 100 %, "Copiar para o SEI" mostra "Copiado" e colar num editor rico (ex.: e-mail HTML) traz a tabela. Parar com Ctrl+C.

- [ ] **Step 7: Commit**

```bash
git add app.py templates tests/test_app.py
git commit -m "app: rotas com SQLite, ficha do bem, página do termo com Copiar para o SEI, upload importa"
```

---

### Task 11: Tela Cadastros (três abas)

**Files:**
- Modify: `app.py` (substituir o stub `cadastros`), `tests/test_app.py`
- Create: `templates/cadastros.html`

**Interfaces:**
- Produces rotas: `GET /cadastros/<aba>` (aba ∈ `responsaveis|localizacoes|pessoas`; `pessoas` aceita `?nome=`), `POST /cadastros/<aba>/incluir`, `POST /cadastros/<aba>/excluir`, `POST /cadastros/responsaveis/renomear`, `POST /cadastros/pessoas/atribuir` (campos `nome`, `numero`, opcional `confirmar=1`), `POST /cadastros/pessoas/desatribuir`.

- [ ] **Step 1: Testes que falham**

Acrescentar a `tests/test_app.py`:

```python
def test_cadastro_responsaveis_incluir_renomear_excluir(cliente):
    r = cliente.post("/cadastros/responsaveis/incluir", data={"ccustos": "decom", "tratamento": "Prezado",
                     "responsavel": "THIAGO", "email": "", "matricula": "481", "funcao": "gerente"}, follow_redirects=True)
    assert b"DECOM" in r.data
    r = cliente.post("/cadastros/responsaveis/renomear", data={"antigo": "CCI", "novo": "GESERV"}, follow_redirects=True)
    assert b"GESERV" in r.data and b">CCI<" not in r.data
    r = cliente.post("/cadastros/responsaveis/excluir", data={"ccustos": "GESERV"}, follow_redirects=True)
    assert "Remapeie".encode() in r.data


def test_cadastro_localizacoes(cliente):
    r = cliente.get("/cadastros/localizacoes")
    assert b"99 - SEM MAPA" in r.data
    r = cliente.post("/cadastros/localizacoes/incluir", data={"localizacao": "99 - SEM MAPA", "ccustos": "CCI"}, follow_redirects=True)
    assert r.data.count(b"99 - SEM MAPA") >= 1 and b"pendente" not in r.data.lower()
    cliente.post("/cadastros/localizacoes/excluir", data={"localizacao": "99 - SEM MAPA"})
    assert b"99 - SEM MAPA" in cliente.get("/cadastros/localizacoes").data


def test_cadastro_pessoas_atribuir_com_confirmacao(cliente):
    cliente.post("/cadastros/pessoas/incluir", data={"nome": "bruno lima"})
    r = cliente.get("/cadastros/pessoas?nome=BRUNO LIMA")
    assert b"BRUNO LIMA" in r.data
    r = cliente.post("/cadastros/pessoas/atribuir", data={"nome": "BRUNO LIMA", "numero": "1002"}, follow_redirects=True)
    assert b"ANA SILVA" in r.data and b"confirmar" in r.data  # pede confirmação
    r = cliente.post("/cadastros/pessoas/atribuir", data={"nome": "BRUNO LIMA", "numero": "1002", "confirmar": "1"},
                     follow_redirects=True)
    assert b"NOTEBOOK" in r.data
    r = cliente.post("/cadastros/pessoas/desatribuir", data={"nome": "BRUNO LIMA", "numero": "1002"}, follow_redirects=True)
    assert b"NOTEBOOK" not in r.data
    r = cliente.post("/cadastros/pessoas/atribuir", data={"nome": "BRUNO LIMA", "numero": "9999"}, follow_redirects=True)
    assert "não encontrado".encode() in r.data
    r = cliente.post("/cadastros/pessoas/excluir", data={"nome": "BRUNO LIMA"}, follow_redirects=True)
    assert b"confirmar" in r.data
    cliente.post("/cadastros/pessoas/excluir", data={"nome": "BRUNO LIMA", "confirmar": "1"})
    assert b"BRUNO LIMA" not in cliente.get("/cadastros/pessoas").data
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_app.py -v -k cadastro`
Expected: FAIL (404/405 nas rotas POST)

- [ ] **Step 3: Rotas em `app.py`** (substituem o stub `cadastros`)

```python
# ---------------------------------------------------------------- cadastros
ABAS = ("responsaveis", "localizacoes", "pessoas")


@app.route("/cadastros/<aba>")
def cadastros(aba):
    if aba not in ABAS:
        abort(404)
    conn = obter_conn()
    nome = request.args.get("nome") or None
    return render_template(
        "cadastros.html", aba=aba, trilha=[("Cadastros", None)],
        centros=db.centros(conn), mapeadas=db.localizacoes_mapeadas(conn), pendentes=db.localizacoes_sem_centro(conn),
        pessoas=db.pessoas(conn), nome=nome, bens_pessoa=db.bens_da_pessoa(conn, nome) if nome else [],
        confirmar=request.args.get("confirmar"))


def _volta(aba, **args):
    return redirect(url_for("cadastros", aba=aba, **args))


@app.route("/cadastros/responsaveis/incluir", methods=["POST"])
def responsaveis_incluir():
    db.incluir_responsavel(obter_conn(), request.form)
    flash("Responsável incluído.", "success")
    return _volta("responsaveis")


@app.route("/cadastros/responsaveis/excluir", methods=["POST"])
def responsaveis_excluir():
    db.excluir_responsavel(obter_conn(), request.form["ccustos"])
    flash("Centro de custo excluído.", "success")
    return _volta("responsaveis")


@app.route("/cadastros/responsaveis/renomear", methods=["POST"])
def responsaveis_renomear():
    db.renomear_centro(obter_conn(), request.form["antigo"], request.form["novo"])
    flash(f"{request.form['antigo']} renomeado para {request.form['novo'].upper()}; localizações atualizadas.", "success")
    return _volta("responsaveis")


@app.route("/cadastros/localizacoes/incluir", methods=["POST"])
def localizacoes_incluir():
    db.incluir_localizacao(obter_conn(), request.form["localizacao"], request.form["ccustos"])
    flash("Localização mapeada.", "success")
    return _volta("localizacoes")


@app.route("/cadastros/localizacoes/excluir", methods=["POST"])
def localizacoes_excluir():
    db.excluir_localizacao(obter_conn(), request.form["localizacao"])
    flash("Mapeamento removido.", "success")
    return _volta("localizacoes")


@app.route("/cadastros/pessoas/incluir", methods=["POST"])
def pessoas_incluir():
    nome = db.incluir_pessoa(obter_conn(), request.form["nome"])
    return _volta("pessoas", nome=nome)


@app.route("/cadastros/pessoas/excluir", methods=["POST"])
def pessoas_excluir():
    nome = request.form["nome"]
    if not request.form.get("confirmar"):
        flash(f"Excluir {nome} remove também os bens atribuídos a ela. Clique em confirmar para prosseguir.", "warning")
        return _volta("pessoas", nome=nome, confirmar="excluir")
    db.excluir_pessoa(obter_conn(), nome)
    flash(f"{nome} excluída.", "success")
    return _volta("pessoas")


@app.route("/cadastros/pessoas/atribuir", methods=["POST"])
def pessoas_atribuir():
    nome, numero = request.form["nome"], request.form.get("numero", "").strip()
    if not numero.isdigit():
        raise db.ErroDeNegocio("Digite o número do bem.")
    try:
        db.atribuir(obter_conn(), nome, int(numero), confirmar=bool(request.form.get("confirmar")))
    except db.JaAtribuido as e:
        flash(f"O bem {numero} está com {e.pessoa}. Clique em confirmar para transferir a {nome}.", "warning")
        return _volta("pessoas", nome=nome, confirmar=numero)
    flash(f"Bem {numero} atribuído a {nome}.", "success")
    return _volta("pessoas", nome=nome)


@app.route("/cadastros/pessoas/desatribuir", methods=["POST"])
def pessoas_desatribuir():
    db.desatribuir(obter_conn(), request.form["nome"], int(request.form["numero"]))
    flash("Atribuição removida; o bem volta a responder pelo setor.", "success")
    return _volta("pessoas", nome=request.form["nome"])
```

O `errorhandler(db.ErroDeNegocio)` da Task 10 já cobre `CentroEmUso`, `BemNaoEncontrado` etc. (redireciona ao `referrer` com flash).

- [ ] **Step 4: `templates/cadastros.html`**

```html
{% extends "base.html" %}
{% from "_macros.html" import select, cabecalho_tabela %}
{% block titulo %}Cadastros{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Cadastros</h1></div>
<div class="br-tab">
  <nav class="tab-nav" aria-label="Cadastros">
    <ul role="tablist">
      {% for id, rotulo in [('responsaveis', 'Responsáveis'), ('localizacoes', 'Localizações'), ('pessoas', 'Pessoas')] %}
      <li class="tab-item{% if aba == id %} active{% endif %}" title="{{ rotulo }}" role="presentation">
        <button type="button" role="tab" id="tab-{{ id }}" data-panel="painel-{{ id }}" aria-controls="painel-{{ id }}" aria-selected="{{ 'true' if aba == id else 'false' }}"
                onclick="history.replaceState(null, '', '{{ url_for('cadastros', aba=id) }}')"><span class="name">{{ rotulo }}</span></button>
      </li>
      {% endfor %}
    </ul>
  </nav>
  <div class="tab-content">

    <div class="tab-panel{% if aba == 'responsaveis' %} active{% endif %}" id="painel-responsaveis" role="tabpanel" aria-labelledby="tab-responsaveis">
      <form method="post" action="{{ url_for('responsaveis_incluir') }}" class="row mb-4">
        {% for campo, rotulo in [('ccustos', 'Sigla do centro de custo'), ('responsavel', 'Responsável'), ('tratamento', 'Tratamento (Prezado/Prezada)'), ('funcao', 'Função'), ('matricula', 'Matrícula'), ('email', 'E-mail')] %}
        <div class="col-md-4 mb-3"><div class="br-input"><label for="r-{{ campo }}">{{ rotulo }}</label><input id="r-{{ campo }}" name="{{ campo }}" type="text"{% if campo in ('ccustos', 'responsavel') %} required{% endif %}/></div></div>
        {% endfor %}
        <div class="col-12"><button class="br-button primary" type="submit"><i class="fas fa-plus mr-1" aria-hidden="true"></i>Incluir</button></div>
      </form>
      {{ cabecalho_tabela('Responsáveis por centro de custo', 'resp') }}
        <thead><tr><th scope="col">Centro</th><th scope="col">Responsável</th><th scope="col">Função</th><th scope="col">Matrícula</th><th scope="col">E-mail</th><th scope="col" class="dsgov-acoes">Ações</th></tr></thead>
        <tbody>
        {% for c in centros %}
        <tr>
          <td>{{ c.ccustos }}</td><td>{{ c.responsavel }}</td><td>{{ c.funcao }}</td><td>{{ c.matricula }}</td><td>{{ c.email }}</td>
          <td class="dsgov-acoes">
            <form method="post" action="{{ url_for('responsaveis_renomear') }}" class="d-inline-flex align-items-center">
              <input type="hidden" name="antigo" value="{{ c.ccustos }}"/>
              <div class="br-input small mr-1"><label class="sr-only" for="novo-{{ loop.index }}">Nova sigla</label><input id="novo-{{ loop.index }}" name="novo" type="text" placeholder="Nova sigla" required/></div>
              <button class="br-button circle small" type="submit" aria-label="Renomear {{ c.ccustos }}"><i class="fas fa-pen" aria-hidden="true"></i></button>
            </form>
            <form method="post" action="{{ url_for('responsaveis_excluir') }}" class="d-inline">
              <input type="hidden" name="ccustos" value="{{ c.ccustos }}"/>
              <button class="br-button circle small" type="submit" aria-label="Excluir {{ c.ccustos }}"><i class="fas fa-trash" aria-hidden="true"></i></button>
            </form>
          </td>
        </tr>
        {% endfor %}
        </tbody>
      </table></div>
    </div>

    <div class="tab-panel{% if aba == 'localizacoes' %} active{% endif %}" id="painel-localizacoes" role="tabpanel" aria-labelledby="tab-localizacoes">
      {% if pendentes %}
      <div class="br-message warning"><div class="icon"><i class="fas fa-exclamation-triangle fa-lg" aria-hidden="true"></i></div>
        <div class="content" role="alert"><span class="message-title">Pendentes:</span><span class="message-body"> {{ pendentes|length }} localização(ões) com bens ativos sem centro de custo.</span></div></div>
      {% endif %}
      <form method="post" action="{{ url_for('localizacoes_incluir') }}" class="row mb-4">
        <div class="col-md-6 mb-3">{{ select('localizacao', 'Localização (ativas sem centro)', pendentes) }}</div>
        <div class="col-md-4 mb-3">{{ select('ccustos', 'Centro de custo', centros|map(attribute='ccustos')|list) }}</div>
        <div class="col-md-2 mb-3 d-flex align-items-end"><button class="br-button primary" type="submit">Mapear</button></div>
      </form>
      {{ cabecalho_tabela('Localizações mapeadas', 'loc') }}
        <thead><tr><th scope="col">Localização</th><th scope="col">Centro de custo</th><th scope="col" class="dsgov-acoes">Ações</th></tr></thead>
        <tbody>
        {% for l in mapeadas %}
        <tr><td>{{ l.localizacao }}</td><td>{{ l.ccustos }}</td>
          <td class="dsgov-acoes"><form method="post" action="{{ url_for('localizacoes_excluir') }}" class="d-inline">
            <input type="hidden" name="localizacao" value="{{ l.localizacao }}"/>
            <button class="br-button circle small" type="submit" aria-label="Remover mapeamento de {{ l.localizacao }}"><i class="fas fa-trash" aria-hidden="true"></i></button></form></td></tr>
        {% endfor %}
        </tbody>
      </table></div>
    </div>

    <div class="tab-panel{% if aba == 'pessoas' %} active{% endif %}" id="painel-pessoas" role="tabpanel" aria-labelledby="tab-pessoas">
      <div class="row mb-4">
        <form method="get" action="{{ url_for('cadastros', aba='pessoas') }}" class="col-md-6 mb-3 d-flex align-items-end">
          <div class="flex-fill mr-2">{{ select('nome', 'Pessoa', pessoas, selecionado=nome) }}</div>
          <button class="br-button secondary" type="submit">Ver bens</button>
        </form>
        <form method="post" action="{{ url_for('pessoas_incluir') }}" class="col-md-6 mb-3 d-flex align-items-end">
          <div class="br-input flex-fill mr-2"><label for="nova-pessoa">Cadastrar nova pessoa</label><input id="nova-pessoa" name="nome" type="text" required/></div>
          <button class="br-button secondary" type="submit">Incluir</button>
        </form>
      </div>
      {% if nome %}
      <div class="d-flex align-items-center mb-3">
        <h2 class="mb-0 text-up-01">{{ nome }}</h2>
        <form method="post" action="{{ url_for('pessoas_excluir') }}" class="ml-auto">
          <input type="hidden" name="nome" value="{{ nome }}"/>
          {% if confirmar == 'excluir' %}<input type="hidden" name="confirmar" value="1"/>
          <button class="br-button secondary" type="submit"><i class="fas fa-check mr-1" aria-hidden="true"></i>Confirmar exclusão de {{ nome }}</button>
          {% else %}<button class="br-button" type="submit"><i class="fas fa-trash mr-1" aria-hidden="true"></i>Excluir pessoa</button>{% endif %}
        </form>
      </div>
      <form method="post" action="{{ url_for('pessoas_atribuir') }}" class="d-flex align-items-end mb-4">
        <input type="hidden" name="nome" value="{{ nome }}"/>
        <div class="br-input mr-2"><label for="numero-bem">Nº do patrimônio</label><input id="numero-bem" name="numero" type="text" inputmode="numeric" value="{{ confirmar if confirmar and confirmar != 'excluir' else '' }}" required/></div>
        {% if confirmar and confirmar != 'excluir' %}<input type="hidden" name="confirmar" value="1"/>
        <button class="br-button primary" type="submit"><i class="fas fa-check mr-1" aria-hidden="true"></i>Confirmar transferência</button>
        {% else %}<button class="br-button primary" type="submit"><i class="fas fa-plus mr-1" aria-hidden="true"></i>Atribuir</button>{% endif %}
      </form>
      {{ cabecalho_tabela('Bens de ' ~ nome, 'bens') }}
        <thead><tr><th scope="col">Patrimônio</th><th scope="col">Descrição</th><th scope="col">Complemento</th><th scope="col" class="dsgov-numero">Valor atual</th><th scope="col" class="dsgov-acoes">Ações</th></tr></thead>
        <tbody>
        {% for b in bens_pessoa %}
        <tr><td>{{ b.numero }}</td><td>{{ b.descricao }}</td><td>{{ b.complemento }}</td><td class="dsgov-numero">R$ {{ '%.2f'|format(b.valor_atual or 0) }}</td>
          <td class="dsgov-acoes"><form method="post" action="{{ url_for('pessoas_desatribuir') }}" class="d-inline">
            <input type="hidden" name="nome" value="{{ nome }}"/><input type="hidden" name="numero" value="{{ b.numero }}"/>
            <button class="br-button circle small" type="submit" aria-label="Remover {{ b.numero }} de {{ nome }}"><i class="fas fa-trash" aria-hidden="true"></i></button></form></td></tr>
        {% endfor %}
        </tbody>
      </table></div>
      <p class="mt-3"><a class="br-button" href="{{ url_for('termo', tipo='individual', chave=nome) }}"><i class="fas fa-file-alt mr-1" aria-hidden="true"></i>Ver termo individual</a></p>
      {% endif %}
    </div>

  </div>
</div>
{% endblock %}
```

Observações para quem implementa: a página tem um único `br-button primary` **visível por aba** (Incluir / Mapear / Atribuir) — os outros painéis estão ocultos pelo `br-tab`; o `h2` dentro da aba Pessoas não fere a regra do `h1` único.

- [ ] **Step 5: Rodar e ver passar**

Run: `.venv/bin/pytest tests/test_app.py -v`
Expected: 8 passed

- [ ] **Step 6: Verificação manual**

`.venv/bin/python app.py` → `/cadastros/responsaveis`: abas trocam, busca da `br-table` filtra, renomear `COLOG`→`GESERV` reflete na aba Localizações.

- [ ] **Step 7: Commit**

```bash
git add app.py templates/cadastros.html tests/test_app.py
git commit -m "Cadastros: responsáveis (com renomear), localizações e pessoas/atribuições"
```

---

### Task 12: Termo de devolução no novo fluxo

**Files:**
- Modify: `app.py` (substituir o stub `termo_devolucao`), `tests/test_app.py`
- Rewrite: `templates/termo_devolucao.html`

**Interfaces:**
- Produces: `GET|POST /termo_devolucao`. POST com `numero_bem` adiciona à `session["bens_selecionados"]`; com `remover=<n>` remove; com `gerar=1` redireciona a `/termo/devolucao/<nome>` (a lista fica na sessão para a página do termo; `limpar=1` esvazia).

- [ ] **Step 1: Testes que falham**

```python
from urllib.parse import unquote


def test_termo_devolucao_fluxo(cliente):
    r = cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "numero_bem": "1001"}, follow_redirects=True)
    assert b"CADEIRA" in r.data
    r = cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "numero_bem": "9999"}, follow_redirects=True)
    assert "não encontrado".encode() in r.data
    r = cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "gerar": "1"})
    assert unquote(r.headers["Location"]).endswith("/termo/devolucao/ANA SILVA")
    doc = cliente.get("/termo/devolucao/ANA SILVA/documento").data.decode()
    assert "TERMO DE DEVOLUÇÃO" in doc and "CADEIRA" in doc and "width:80%" in doc
    assert cliente.get("/termo/devolucao/ANA SILVA/docx").status_code == 200
    r = cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "remover": "1001"}, follow_redirects=True)
    assert b"CADEIRA" not in r.data
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_app.py -v -k devolucao`
Expected: FAIL (o stub redireciona para `/`)

- [ ] **Step 3: Rota** (substitui o stub)

```python
# ---------------------------------------------------------------- termo de devolução
@app.route("/termo_devolucao", methods=["GET", "POST"])
def termo_devolucao():
    conn = obter_conn()
    nome = request.form.get("nome") or session.get("nome_devolucao")
    selecionados = session.setdefault("bens_selecionados", [])
    if request.method == "POST":
        if nome:
            session["nome_devolucao"] = nome
        if request.form.get("limpar"):
            session["bens_selecionados"] = []
        elif request.form.get("gerar"):
            if not nome or not selecionados:
                flash("Escolha a pessoa e adicione ao menos um bem.", "error")
            else:
                return redirect(url_for("termo", tipo="devolucao", chave=nome))
        elif request.form.get("remover"):
            session["bens_selecionados"] = [n for n in selecionados if n != request.form["remover"]]
        else:
            numero = request.form.get("numero_bem", "").strip()
            if not numero.isdigit() or not db.buscar_bem(conn, int(numero)):
                flash(f"Bem {numero or '(vazio)'} não encontrado. Verifique o número digitado.", "error")
            elif numero not in selecionados:
                session["bens_selecionados"] = selecionados + [numero]
        session.modified = True
        return redirect(url_for("termo_devolucao"))
    bens = [b for b in (db.buscar_bem(conn, int(n)) for n in selecionados) if b]
    return render_template("termo_devolucao.html", nomes=db.pessoas(conn), nome=nome, bens=bens,
                           total=sum(b["valor_atual"] or 0 for b in bens), trilha=[("Termo de devolução", None)])
```

- [ ] **Step 4: `templates/termo_devolucao.html`**

```html
{% extends "base.html" %}
{% from "_macros.html" import select %}
{% block titulo %}Termo de devolução{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Termo de devolução</h1></div>
<form method="post" class="row mb-4">
  <div class="col-md-6 mb-3">{{ select('nome', 'Pessoa que devolve', nomes, selecionado=nome) }}</div>
  <div class="col-md-4 mb-3"><div class="br-input"><label for="numero_bem">Nº do patrimônio</label><input id="numero_bem" name="numero_bem" type="text" inputmode="numeric" placeholder="Digite e adicione"/></div></div>
  <div class="col-md-2 mb-3 d-flex align-items-end"><button class="br-button secondary" type="submit"><i class="fas fa-plus mr-1" aria-hidden="true"></i>Adicionar</button></div>
</form>
{% if bens %}
<div class="br-table mb-3">
  <div class="table-header"><div class="top-bar"><div class="table-title">Bens a devolver</div></div></div>
  <table>
    <caption class="sr-only">Bens a devolver</caption>
    <thead><tr><th scope="col">Patrimônio</th><th scope="col">Descrição</th><th scope="col">Complemento</th><th scope="col" class="dsgov-numero">Valor atual</th><th scope="col" class="dsgov-acoes">Ações</th></tr></thead>
    <tbody>
    {% for b in bens %}
    <tr><td>{{ b.numero }}</td><td>{{ b.descricao }}</td><td>{{ b.complemento }}</td><td class="dsgov-numero">R$ {{ '%.2f'|format(b.valor_atual or 0) }}</td>
      <td class="dsgov-acoes"><form method="post" class="d-inline"><input type="hidden" name="nome" value="{{ nome or '' }}"/><input type="hidden" name="remover" value="{{ b.numero }}"/>
        <button class="br-button circle small" type="submit" aria-label="Remover {{ b.numero }}"><i class="fas fa-trash" aria-hidden="true"></i></button></form></td></tr>
    {% endfor %}
    <tr><td colspan="3" class="text-weight-bold">TOTAL</td><td class="dsgov-numero text-weight-bold">R$ {{ '%.2f'|format(total) }}</td><td></td></tr>
    </tbody>
  </table>
</div>
<form method="post" class="d-flex">
  <input type="hidden" name="nome" value="{{ nome or '' }}"/>
  <button class="br-button primary" type="submit" name="gerar" value="1"><i class="fas fa-file-alt mr-1" aria-hidden="true"></i>Gerar termo</button>
  <button class="br-button ml-2" type="submit" name="limpar" value="1">Limpar lista</button>
</form>
{% endif %}
{% endblock %}
```

- [ ] **Step 5: Rodar e ver passar**

Run: `.venv/bin/pytest -v`
Expected: todos passam (config 3, db 17, importar 1, docx 3, html 4, app 9).

- [ ] **Step 6: Commit**

```bash
git add app.py templates/termo_devolucao.html tests/test_app.py
git commit -m "Termo de devolução: seleção em sessão, página do termo com Copiar para o SEI"
```

---

### Task 13: `main.py`, `build.bat`, limpeza e README

**Files:**
- Create: `main.py`, `build.bat`
- Delete: `Procfile`, `acervo.xlsx`, `geral.xlsx`
- Modify: `README.md`

**Interfaces:**
- Consumes: `app.app`, `db.inicializar`.

- [ ] **Step 1: `main.py`**

```python
"""Programa de desktop: sobe o Flask numa thread e abre a janela. Fechar a janela encerra o programa.

Sem WebView2/WebKit disponível, abre o navegador padrão e fica servindo até Ctrl+C.
Porta 5000 ocupada = o programa já está aberto: abre o navegador nele e sai.
"""
import socket
import threading
import webbrowser

import db
from app import app

PORTA = 5000
URL = f"http://127.0.0.1:{PORTA}"


def porta_livre() -> bool:
    with socket.socket() as s:
        return s.connect_ex(("127.0.0.1", PORTA)) != 0


def servidor():
    app.run(host="127.0.0.1", port=PORTA, debug=False, use_reloader=False)


def main():
    if not porta_livre():
        webbrowser.open(URL)
        return
    db.inicializar()
    threading.Thread(target=servidor, daemon=True).start()
    try:
        import webview
        webview.create_window("Termos de Responsabilidade – CFC", URL, width=1100, height=750)
        webview.start()
    except Exception:
        webbrowser.open(URL)
        threading.Event().wait()  # mantém o servidor vivo até Ctrl+C


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Testar o fallback aqui**

Run: `timeout 8 .venv/bin/python main.py; echo "saiu com $?"` — em outra aba, `curl -s -o /dev/null -w "%{http_code}\n" http://127.0.0.1:5000/` durante os 8 s.
Expected: `200`; o processo é morto pelo `timeout` (código 124). Nesta VPS o `webview.start()` falha (sem WebKitGTK) e cai no navegador — é o comportamento esperado.

- [ ] **Step 3: `build.bat`** (roda no Windows do usuário)

```bat
@echo off
REM Gera dist\TermosCFC\TermosCFC.exe. Rode dentro da venv: python -m venv .venv && .venv\Scripts\activate && pip install -r requirements.txt
pyinstaller --noconfirm --clean --onedir --windowed --name TermosCFC ^
  --add-data "templates;templates" --add-data "static;static" --add-data "timbrado.docx;." ^
  main.py
if errorlevel 1 exit /b 1
if not exist dist\TermosCFC\dados mkdir dist\TermosCFC\dados
echo.
echo Pronto: dist\TermosCFC\TermosCFC.exe  (copie termos.db para dist\TermosCFC\dados\ se quiser levar os dados)
```

- [ ] **Step 4: Limpeza**

```bash
git rm -q Procfile acervo.xlsx geral.xlsx
grep -rn "pandas\|gunicorn\|render" --include=*.py --include=*.txt --include=*.md . | grep -v .venv | grep -v docs/
```

Expected: nenhum resultado além de `render_template` no `app.py`.

- [ ] **Step 5: README.md** (substituir inteiro)

```markdown
# Termos de Responsabilidade — CFC

Programa local (Windows) do Setor de Patrimônio para emitir Termos de Responsabilidade (por centro de
custo e individuais) e Termos de Devolução, com botão **Copiar para o SEI** e download em `.docx`.
Os dados ficam num SQLite (`dados/termos.db`) mantido pelo próprio programa.

## Uso

1. Abra `TermosCFC.exe`. A janela abre em `http://127.0.0.1:5000`.
2. **Atualizar base**: envie o export do sistema de patrimônio (`.xlsx`). Só a tabela de bens muda.
3. **Cadastros**: responsáveis por centro de custo (com *renomear*), localização → centro de custo,
   pessoas e bens atribuídos. Bem atribuído a pessoa não entra no termo do setor.
4. **Termos**: escolha o centro/pessoa → página do termo → *Copiar para o SEI* ou *Baixar .docx*.

Backup = copiar a pasta `dados/`.

## Desenvolvimento

    python -m venv .venv && .venv/bin/pip install -r requirements.txt
    .venv/bin/pytest
    .venv/bin/python app.py        # http://127.0.0.1:5000 (debug)
    .venv/bin/python main.py       # como o programa: janela (ou navegador, se não houver WebView)

Migração inicial a partir das planilhas antigas: `python importar_planilhas.py acervo.xlsx geral.xlsx`.

## Gerar o executável (Windows)

    python -m venv .venv && .venv\Scripts\activate && pip install -r requirements.txt
    build.bat

Sai em `dist\TermosCFC\`. Distribua a pasta inteira (zip). Requer o WebView2 Runtime (já vem no
Windows 10/11 atualizados); sem ele o programa abre no navegador padrão.

## Arquivos

| Arquivo | Função |
|---|---|
| `app.py` | rotas Flask |
| `db.py` | esquema, importação, consultas, cadastros |
| `termos_html.py` | corpo HTML dos termos (padrão gelic; tabelas 80 % / 100 %) |
| `Script_Termo_Individual.py`, `Termo_de_Responsabilidade.py`, `termo_devolucao.py` | geradores `.docx` |
| `config.py` | pasta de dados (`TERMOS_DADOS` sobrepõe) |
| `main.py`, `build.bat` | programa de desktop e build |
| `templates/`, `static/dsgov/` | telas DSGov 3.7.0 (offline) |
```

- [ ] **Step 6: Rodar tudo**

Run: `.venv/bin/pytest -q`
Expected: todos passam.

- [ ] **Step 7: Commit**

```bash
git add main.py build.bat README.md
git commit -m "Programa de desktop (pywebview + PyInstaller); remove Render e planilhas"
```

- [ ] **Step 8: Entrega ao usuário (fora do repositório)**

No Windows: clonar, `build.bat`, copiar `dados/termos.db` gerado na Task 6 (`scp` da VPS) para `dist\TermosCFC\dados\`, abrir o `.exe`. Validar: janela abre, menu funciona, "Copiar para o SEI" cola tabela no editor do SEI.

---

## Self-review

**Spec coverage**
- §4.1 esquema → Task 2. §4.2 importação → Task 3 + rota Task 10. §4.3 consultas → Task 4. §4.4 migração → Task 6. §4.5 pastas → Task 1.
- §5.1 renomear/de-para → Task 5 + Task 11. §5.2 setor OU pessoa → Task 4 (`bens_do_centro`) + Task 5 (`atribuir` com confirmação) + Task 11. §5.3 ficha do bem → Task 4 + Task 10.
- §6 DSGov → Task 9; §6.1 cadastros → Task 11. §7 termo HTML + copiar + docx + planilha → Tasks 8, 10, 12. §8 desktop/build → Task 13. §10 testes → em cada task. §3 remoções → Tasks 1, 7, 9, 13.
- Planilha `planilha_<ccustos>.xlsx` (adendo da spec) → Task 7 + rota `termo_planilha` Task 10.

**Placeholders**: nenhum "TBD/TODO". Task 7 diz "igual à atual" para trechos de formatação do docx que **não mudam** — o implementador mantém o código existente do arquivo; o que muda está escrito.

**Consistência de nomes**: `db.conectar/criar_esquema/inicializar/importar_bens/localizacoes_sem_centro/centros/responsavel/bens_do_centro/pessoas/bens_da_pessoa/buscar_bem/pessoa_do_bem/ficha_do_bem/localizacoes_mapeadas/incluir_responsavel/excluir_responsavel/renomear_centro/incluir_localizacao/excluir_localizacao/incluir_pessoa/excluir_pessoa/atribuir/desatribuir`, exceções `ErroDeNegocio/ImportacaoInvalida/CentroEmUso/BemNaoEncontrado/JaAtribuido` — usados com esses nomes nas Tasks 4–12. Geradores: `criar_termo_responsabilidade(nome, bens, destino)`, `gerar_termo_centro(ccustos, responsavel, bens, destino)`, `gerar_planilha_centro(bens, destino)`, `gerar_termo_devolucao(nome, bens, destino)` — Tasks 7 e 10. `termos_html.corpo_ccusto/corpo_individual/corpo_devolucao/documento` — Tasks 8 e 10. Rotas nomeadas no menu (`home, centro_custos, termos_individuais, termo_devolucao, cadastros, upload`) existem a partir da Task 10 (stubs) e são substituídas nas Tasks 11–12.
