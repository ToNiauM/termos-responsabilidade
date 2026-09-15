# Módulo de inventário — plano de implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Evento de inventário com salas, leitura por código de barras (leitor/câmera/digitação), situação localizado/divergente/pendente, sobras com foto no R2, relatório e `.xlsx`, abas de inventário na planilha de cadastros.

**Architecture:** Módulo separado: `inventario.py` (dados, `conn` primeiro, sem Flask), `fotos.py` (validação, WebP, R2 via boto3), `app_inventario.py` (blueprint Flask com prefixo `/inventario`), templates `inventario_*.html` (DSGov) e um JS de leitura embutido na tela da sala com `html5-qrcode` vendorizado. `db.py` só ganha o esquema, `localizacoes_ativas` e a extensão da planilha de cadastros.

**Tech Stack:** Python 3.12, Flask (Blueprint), SQLite, openpyxl, Pillow, boto3, html5-qrcode 2.3.8 (vendorizado), DSGov 3.7.0, pytest.

**Spec:** `docs/superpowers/specs/2026-09-15-inventario-design.md` (regras do sistema antigo em `docs/superpowers/notes/2026-09-15-inventario-existente.md`).

## Global Constraints

- Só acréscimos ao esquema (`CREATE TABLE IF NOT EXISTS`); `bens` nunca é alterada pelo inventário.
- Funções de dados recebem `conn` primeiro e não importam Flask; erros para o usuário são `db.ErroDeNegocio`; datas via `db._agora()` (ISO `YYYY-MM-DD HH:MM:SS`).
- SQL sempre com parâmetros ligados; nenhum valor do usuário em f-string.
- DSGov: `br-button`, `br-input`, `br-select` (macro `select`), `br-table` (macro `cabecalho_tabela`), `br-message`, `br-tag`; textos em português; sem `style` inline.
- JS só em `{% block scripts %}` da própria tela; `html5-qrcode` só na tela da sala; sem CDN.
- Testes sem rede: `fotos._cliente` sempre substituído por um cliente falso; variáveis `R2_*` via `monkeypatch.setenv`.
- Suíte inteira antes de cada commit: `.venv/bin/pytest -q` (122 passam na base). Um commit por tarefa, em português, rodapé:
  ```
  Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>
  Claude-Session: https://claude.ai/code/session_01Xz5P5QQJyvfjPDkCPVSG18
  ```
- Fixtures: `dados` (conexão com esquema), `semear(conn)` (centro CCI; sala "01 - SALA CCI" com bens 1001 CADEIRA, 1002 NOTEBOOK (atribuído a ANA SILVA), 1003 MESA BAIXADO; 1004 ARMÁRIO em "99 - SEM MAPA"), `cliente` (test client semeado, `tests/test_app.py`), `xlsx(tmp_path, linhas)` (`tests/test_db.py`).

---

## Mapa de arquivos

| arquivo | papel |
|---|---|
| `db.py` | esquema (5 tabelas), `localizacoes_ativas`, `_ler_aba_cadastro(..., colunas=None, opcional=False)`, integração das abas `inv_*` em exportar/importar |
| `inventario.py` (novo) | eventos, salas, leituras, sobras, relatório, resumo, xlsx, abas da planilha |
| `fotos.py` (novo) | `configurado`, `validar`, `comprimir`, `enviar`, `apagar`, nomes |
| `app_inventario.py` (novo) | blueprint `/inventario` |
| `app.py` | `register_blueprint`, item de menu, card do painel |
| `templates/inventario_eventos.html`, `inventario_evento.html`, `inventario_sala.html`, `inventario_relatorio.html` (novos) | telas |
| `templates/index.html` | card "Inventário em andamento" |
| `static/dsgov/vendor/html5-qrcode/html5-qrcode.min.js` + `LICENSE` (novos) | leitor por câmera |
| `requirements.txt`, `Dockerfile`, `compose.yml`, `.gitignore` | `boto3`, `Pillow`; `secrets/.env` opcional |
| `tests/test_inventario.py`, `tests/test_fotos.py` (novos), `tests/test_app.py`, `tests/test_db.py` | testes |

---

### Task 1: Esquema, `localizacoes_ativas`, eventos e salas

**Files:**
- Modify: `db.py` (`ESQUEMA`; nova função após `localizacoes_sem_centro`)
- Create: `inventario.py`
- Test: `tests/test_inventario.py`

**Interfaces:**
- Produces: `db.localizacoes_ativas(conn) -> list[str]`; `inventario.CONSERVACAO`; `inventario.evento_aberto(conn)`, `eventos(conn)`, `evento(conn, id)` (com `integrantes` e `resumo`), `abrir_evento(conn, nome, descricao, integrantes, salas=None) -> int`, `encerrar_evento(conn, id)`, `salas(conn, evento_id) -> list[dict]` (`localizacao, ccustos, total, localizados, divergentes, pendentes`), `resumo(conn, evento_id)` (`salas, salas_iniciadas, bens, lidos, divergentes, pendentes, sobras, pct_bens`).

- [ ] **Step 1: Testes que falham**

`tests/test_inventario.py`:

```python
"""Módulo de inventário: eventos, salas, leituras, sobras, relatório, xlsx."""
import pytest

import db
import inventario
from tests.conftest import semear


def semear_inventario(conn):
    """semear() + 2ª sala com 2 bens e 1 sala vazia de escopo; devolve o id do evento aberto."""
    semear(conn)
    conn.execute("INSERT INTO bens VALUES (2001,'ATIVO','MONITOR','LG','EQUIPAMENTOS','02 - SALA B','01/01/2020',900,800)")
    conn.execute("INSERT INTO bens VALUES (2002,'ATIVO','IMPRESSORA','HP','EQUIPAMENTOS','02 - SALA B','01/01/2020',1200,1000)")
    conn.commit()
    return inventario.abrir_evento(conn, "Inventário 2026", "Portaria 1/2026", ["Fulano", "Beltrana"])


def test_localizacoes_ativas(dados):
    semear(dados)
    assert db.localizacoes_ativas(dados) == ["01 - SALA CCI", "99 - SEM MAPA"]   # 1003 é BAIXADO, não muda nada


def test_abrir_evento_todas_as_salas_e_integrantes(dados):
    eid = semear_inventario(dados)
    e = inventario.evento(dados, eid)
    assert e["nome"] == "Inventário 2026" and e["encerrado_em"] is None and e["integrantes"] == ["Beltrana", "Fulano"]
    assert [s["localizacao"] for s in inventario.salas(dados, eid)] == ["01 - SALA CCI", "02 - SALA B", "99 - SEM MAPA"]
    assert inventario.evento_aberto(dados)["id"] == eid
    assert e["resumo"] == {"salas": 3, "salas_iniciadas": 0, "bens": 5, "lidos": 0, "divergentes": 0,
                           "pendentes": 5, "sobras": 0, "pct_bens": 0.0}


def test_abrir_evento_amostragem_e_validacoes(dados):
    semear(dados)
    with pytest.raises(db.ErroDeNegocio):
        inventario.abrir_evento(dados, "", "", ["A"])
    with pytest.raises(db.ErroDeNegocio):
        inventario.abrir_evento(dados, "X", "", [" ", ""])
    with pytest.raises(db.ErroDeNegocio):
        inventario.abrir_evento(dados, "X", "", ["A"], salas=["NÃO EXISTE"])
    eid = inventario.abrir_evento(dados, "Amostra", None, ["A", "A ", "b"], salas=["99 - SEM MAPA"])
    assert [s["localizacao"] for s in inventario.salas(dados, eid)] == ["99 - SEM MAPA"]
    assert inventario.evento(dados, eid)["integrantes"] == ["A", "b"]
    with pytest.raises(db.ErroDeNegocio):
        inventario.abrir_evento(dados, "Outro", "", ["A"])          # já há aberto
    inventario.encerrar_evento(dados, eid)
    assert inventario.evento_aberto(dados) is None
    e2 = inventario.abrir_evento(dados, "Outro", "", ["A"])
    assert [x["id"] for x in inventario.eventos(dados)] == [e2, eid]   # aberto primeiro


def test_salas_contadores_e_resumo(dados):
    eid = semear_inventario(dados)
    s = {x["localizacao"]: x for x in inventario.salas(dados, eid)}
    assert s["01 - SALA CCI"]["total"] == 2 and s["01 - SALA CCI"]["ccustos"] == "CCI" and s["02 - SALA B"]["ccustos"] is None
    assert (s["01 - SALA CCI"]["localizados"], s["01 - SALA CCI"]["pendentes"], s["01 - SALA CCI"]["divergentes"]) == (0, 2, 0)
    assert "concluida_em" not in s["01 - SALA CCI"]
    dados.execute("INSERT INTO inventario_leituras (evento_id, numero, localizacao, lido_em, integrante) VALUES (?,?,?,?,?)",
                  (eid, 2001, "02 - SALA B", "2026-09-15 10:00:00", "Fulano"))
    dados.commit()
    r = inventario.resumo(dados, eid)
    assert (r["salas_iniciadas"], r["lidos"], r["pendentes"], r["pct_bens"]) == (1, 1, 4, 20.0)
    inventario.encerrar_evento(dados, eid)
    assert inventario.evento(dados, eid)["encerrado_em"] is not None
    inventario.encerrar_evento(dados, eid)                                        # idempotente
    with pytest.raises(db.ErroDeNegocio):
        inventario.encerrar_evento(dados, 999)
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest -q tests/test_inventario.py`
Expected: `ModuleNotFoundError: No module named 'inventario'`.

- [ ] **Step 3: Esquema e `localizacoes_ativas`**

Em `db.ESQUEMA`, após `importacoes_mudancas`:

```sql
CREATE TABLE IF NOT EXISTS inventario_eventos (
  id           INTEGER PRIMARY KEY,
  nome         TEXT NOT NULL,
  descricao    TEXT,
  aberto_em    TEXT NOT NULL,
  encerrado_em TEXT
);
CREATE TABLE IF NOT EXISTS inventario_integrantes (
  evento_id INTEGER NOT NULL REFERENCES inventario_eventos(id) ON DELETE CASCADE,
  nome      TEXT NOT NULL,
  PRIMARY KEY (evento_id, nome)
);
CREATE TABLE IF NOT EXISTS inventario_salas (
  evento_id    INTEGER NOT NULL REFERENCES inventario_eventos(id) ON DELETE CASCADE,
  localizacao  TEXT NOT NULL,
  PRIMARY KEY (evento_id, localizacao)
);
CREATE TABLE IF NOT EXISTS inventario_leituras (
  id          INTEGER PRIMARY KEY,
  evento_id   INTEGER NOT NULL REFERENCES inventario_eventos(id) ON DELETE CASCADE,
  numero      INTEGER NOT NULL,
  localizacao TEXT NOT NULL,
  lido_em     TEXT NOT NULL,
  integrante  TEXT NOT NULL,
  conservacao TEXT CHECK (conservacao IN ('Bom','Regular','Ruim','Inservível')),
  quem_usa    TEXT,
  observacao  TEXT,
  foto_url    TEXT,
  UNIQUE (evento_id, numero)
);
CREATE TABLE IF NOT EXISTS inventario_sobras (
  id          INTEGER PRIMARY KEY,
  evento_id   INTEGER NOT NULL REFERENCES inventario_eventos(id) ON DELETE CASCADE,
  localizacao TEXT NOT NULL,
  descricao   TEXT NOT NULL,
  complemento TEXT,
  observacao  TEXT NOT NULL,
  foto_url    TEXT NOT NULL,
  integrante  TEXT NOT NULL,
  criado_em   TEXT NOT NULL
);
```

Após `localizacoes_sem_centro` em `db.py`:

```python
def localizacoes_ativas(conn) -> list[str]:
    """Localizações distintas com bens ATIVO (as "salas" que um inventário pode conferir)."""
    return [r[0] for r in conn.execute(
        "SELECT DISTINCT localizacao FROM bens WHERE situacao = 'ATIVO' AND localizacao <> '' ORDER BY localizacao")]
```

- [ ] **Step 4: `inventario.py` (parte 1)**

```python
"""Módulo de inventário: eventos (campanhas), salas, leituras por código de barras, sobras, relatório e
planilhas. Só dados: toda função recebe `conn` primeiro e não importa Flask (mesmo padrão de db.py).

Regras (spec 2026-09-15): `bens` é espelho do SPW e nunca muda aqui; "local sistema" = bens.localizacao,
"local inventário" = sala onde o bem foi lido; divergente = os dois diferem (calculado, nunca gravado)."""
import db
from db import ErroDeNegocio, _agora, _obrigatorio, _texto, _todos, _um

CONSERVACAO = ("Bom", "Regular", "Ruim", "Inservível")
ROTULO_SITUACAO = {"localizado": "Localizado", "divergente": "Divergente", "pendente": "Não localizado"}


class BemNaoEncontrado(ErroDeNegocio):
    """Número lido não existe em `bens`: a tela oferece registrar como sobra."""

    def __init__(self, numero):
        super().__init__(f"Bem {numero} não está na base.")
        self.numero = numero


# ---------------------------------------------------------------- eventos
def evento_aberto(conn) -> dict | None:
    return _um(conn, "SELECT * FROM inventario_eventos WHERE encerrado_em IS NULL")


def eventos(conn) -> list[dict]:
    return _todos(conn, "SELECT * FROM inventario_eventos ORDER BY (encerrado_em IS NULL) DESC, aberto_em DESC, id DESC")


def evento(conn, id: int) -> dict | None:
    e = _um(conn, "SELECT * FROM inventario_eventos WHERE id = ?", id)
    if e:
        e["integrantes"] = [r[0] for r in conn.execute(
            "SELECT nome FROM inventario_integrantes WHERE evento_id = ? ORDER BY nome", (id,))]
        e["resumo"] = resumo(conn, id)
    return e


def _evento_aberto_ou_erro(conn, id: int) -> dict:
    e = _um(conn, "SELECT * FROM inventario_eventos WHERE id = ?", id)
    if not e:
        raise ErroDeNegocio("Evento de inventário não encontrado.")
    if e["encerrado_em"]:
        raise ErroDeNegocio("Evento encerrado: não aceita alterações.")
    return e


def abrir_evento(conn, nome: str, descricao, integrantes: list, salas: list | None = None) -> int:
    """Um evento aberto por vez. salas=None → todas as localizações com bens ATIVO; lista → amostragem."""
    nome = _obrigatorio(nome, "Nome do evento")
    if evento_aberto(conn):
        raise ErroDeNegocio("Já existe um evento de inventário aberto; encerre-o antes de abrir outro.")
    nomes = sorted({" ".join(_texto(n).split()) for n in integrantes if _texto(n).strip()})
    if not nomes:
        raise ErroDeNegocio("Informe ao menos um integrante da comissão.")
    ativas = db.localizacoes_ativas(conn)
    escolhidas = ativas if salas is None else [s for s in ativas if s in set(salas)]
    if not escolhidas:
        raise ErroDeNegocio("Nenhuma sala com bens ativos no escopo do evento.")
    cur = conn.execute("INSERT INTO inventario_eventos (nome, descricao, aberto_em) VALUES (?,?,?)",
                       (nome, _texto(descricao) or None, _agora()))
    eid = cur.lastrowid
    conn.executemany("INSERT INTO inventario_integrantes VALUES (?,?)", [(eid, n) for n in nomes])
    conn.executemany("INSERT INTO inventario_salas (evento_id, localizacao) VALUES (?,?)", [(eid, s) for s in escolhidas])
    conn.commit()
    return eid


def encerrar_evento(conn, id: int) -> None:
    e = _um(conn, "SELECT * FROM inventario_eventos WHERE id = ?", id)
    if not e:
        raise ErroDeNegocio("Evento de inventário não encontrado.")
    if not e["encerrado_em"]:
        conn.execute("UPDATE inventario_eventos SET encerrado_em = ? WHERE id = ?", (_agora(), id))
        conn.commit()


# ---------------------------------------------------------------- salas
def salas(conn, evento_id: int) -> list[dict]:
    """Por sala do escopo: bens ativos (total), localizados aqui, divergentes lidos aqui, pendentes."""
    linhas = _todos(conn, """
        SELECT s.localizacao, l.ccustos,
          (SELECT count(*) FROM bens b WHERE b.localizacao = s.localizacao AND b.situacao = 'ATIVO') AS total,
          (SELECT count(*) FROM inventario_leituras r JOIN bens b ON b.numero = r.numero
             WHERE r.evento_id = s.evento_id AND r.localizacao = s.localizacao
               AND b.localizacao = s.localizacao AND b.situacao = 'ATIVO') AS localizados,
          (SELECT count(*) FROM inventario_leituras r JOIN bens b ON b.numero = r.numero
             WHERE r.evento_id = s.evento_id AND r.localizacao = s.localizacao AND b.localizacao <> s.localizacao) AS divergentes
        FROM inventario_salas s LEFT JOIN localizacoes l ON l.localizacao = s.localizacao
        WHERE s.evento_id = ? ORDER BY s.localizacao""", evento_id)
    for s in linhas:
        s["pendentes"] = s["total"] - s["localizados"]
    return linhas


def _sala_ou_erro(conn, evento_id: int, localizacao: str) -> dict:
    s = _um(conn, "SELECT * FROM inventario_salas WHERE evento_id = ? AND localizacao = ?", evento_id, localizacao)
    if not s:
        raise ErroDeNegocio("Sala fora do escopo deste evento.")
    return s


def resumo(conn, evento_id: int) -> dict:
    """Progresso do evento pela quantidade de bens localizados (não há "concluir sala")."""
    ss = salas(conn, evento_id)
    bens = sum(s["total"] for s in ss)
    lidos = sum(s["localizados"] for s in ss)
    sobras = conn.execute("SELECT count(*) FROM inventario_sobras WHERE evento_id = ?", (evento_id,)).fetchone()[0]
    return {"salas": len(ss), "salas_iniciadas": sum(1 for s in ss if s["localizados"] or s["divergentes"]),
            "bens": bens, "lidos": lidos, "divergentes": sum(s["divergentes"] for s in ss),
            "pendentes": sum(s["pendentes"] for s in ss), "sobras": sobras,
            "pct_bens": round(100 * lidos / bens, 1) if bens else 0.0}
```

- [ ] **Step 5: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam (126).

- [ ] **Step 6: Commit**

```bash
git add db.py inventario.py tests/test_inventario.py
git commit -m "Inventário: esquema, eventos (abrir/encerrar, um aberto por vez), salas e resumo por bens"
```

---

### Task 2: Leitura, bens da sala, atualização e sobras

**Files:**
- Modify: `inventario.py`
- Test: `tests/test_inventario.py`

**Interfaces:**
- Produces: `inventario.ler(conn, evento_id, localizacao, numero, integrante) -> dict` (`situacao, bem, cadastrado_em, ativo, reler, leitura_anterior, lido_em, integrante`; levanta `BemNaoEncontrado`); `bens_da_sala(conn, evento_id, localizacao) -> {"bens", "trazidos", "sobras"}`; `atualizar_leitura(conn, evento_id, numero, **campos)`; `registrar_sobra(conn, evento_id, localizacao, descricao, complemento, observacao, foto_url, integrante, exigir_foto=True) -> int`; `definir_foto_sobra(conn, sobra_id, foto_url)`; `excluir_sobra(conn, evento_id, sobra_id) -> dict`.

- [ ] **Step 1: Testes que falham**

```python
def test_ler_localizado_divergente_reler_e_erros(dados):
    eid = semear_inventario(dados)
    r = inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    assert r["situacao"] == "localizado" and r["reler"] is False and r["ativo"] and r["bem"]["descricao"] == "CADEIRA"
    r = inventario.ler(dados, eid, "01 - SALA CCI", 2001, "Fulano")             # MONITOR é da SALA B
    assert r["situacao"] == "divergente" and r["cadastrado_em"] == "02 - SALA B"
    r = inventario.ler(dados, eid, "02 - SALA B", 2001, "Beltrana")             # reler: atualiza a mesma linha
    assert r["situacao"] == "localizado" and r["reler"] and r["leitura_anterior"]["localizacao"] == "01 - SALA CCI"
    assert dados.execute("SELECT count(*) FROM inventario_leituras WHERE evento_id = ?", (eid,)).fetchone()[0] == 2
    r = inventario.ler(dados, eid, "01 - SALA CCI", 1003, "Fulano")             # BAIXADO: registra, avisa
    assert r["ativo"] is False and r["situacao"] == "localizado"
    with pytest.raises(inventario.BemNaoEncontrado) as e:
        inventario.ler(dados, eid, "01 - SALA CCI", 99999, "Fulano")
    assert e.value.numero == 99999
    with pytest.raises(db.ErroDeNegocio):
        inventario.ler(dados, eid, "SALA QUE NÃO EXISTE", 1001, "Fulano")
    with pytest.raises(db.ErroDeNegocio):
        inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Ninguém")
    s = {x["localizacao"]: x for x in inventario.salas(dados, eid)}
    assert (s["01 - SALA CCI"]["localizados"], s["01 - SALA CCI"]["pendentes"]) == (1, 1)
    assert (s["02 - SALA B"]["localizados"], s["02 - SALA B"]["divergentes"]) == (1, 0)
    inventario.encerrar_evento(dados, eid)
    with pytest.raises(db.ErroDeNegocio):
        inventario.ler(dados, eid, "01 - SALA CCI", 1002, "Fulano")


def test_bens_da_sala_e_atualizar_leitura(dados):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.ler(dados, eid, "01 - SALA CCI", 2002, "Fulano")                 # trazido da SALA B
    inventario.ler(dados, eid, "02 - SALA B", 1002, "Fulano")                   # bem da CCI lido na SALA B
    inventario.ler(dados, eid, "01 - SALA CCI", 1003, "Fulano")                 # BAIXADO da própria sala
    d = inventario.bens_da_sala(dados, eid, "01 - SALA CCI")
    por = {b["numero"]: b for b in d["bens"]}
    assert set(por) == {1001, 1002}                                              # só ativos da sala
    assert por[1001]["situacao_inv"] == "localizado" and por[1002]["situacao_inv"] == "divergente" and por[1002]["lido_em_sala"] == "02 - SALA B"
    assert [t["numero"] for t in d["trazidos"]] == [1003, 2002] and d["sobras"] == []
    inventario.atualizar_leitura(dados, eid, 1001, conservacao="Ruim", quem_usa="Ciclana", observacao="pé quebrado")
    b = {x["numero"]: x for x in inventario.bens_da_sala(dados, eid, "01 - SALA CCI")["bens"]}[1001]
    assert (b["conservacao"], b["quem_usa"], b["observacao"]) == ("Ruim", "Ciclana", "pé quebrado")
    inventario.atualizar_leitura(dados, eid, 1001, conservacao="")                # limpa
    assert {x["numero"]: x for x in inventario.bens_da_sala(dados, eid, "01 - SALA CCI")["bens"]}[1001]["conservacao"] is None
    with pytest.raises(db.ErroDeNegocio):
        inventario.atualizar_leitura(dados, eid, 1001, conservacao="Ótimo")
    with pytest.raises(db.ErroDeNegocio):
        inventario.atualizar_leitura(dados, eid, 2001, quem_usa="x")             # sem leitura


def test_sobras(dados):
    eid = semear_inventario(dados)
    with pytest.raises(db.ErroDeNegocio):
        inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "", "", "achado", "http://x/1.webp", "Fulano")
    with pytest.raises(db.ErroDeNegocio):
        inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "VENTILADOR", "", "", "http://x/1.webp", "Fulano")
    with pytest.raises(db.ErroDeNegocio):
        inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "VENTILADOR", "", "achado", "", "Fulano")
    sid = inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "VENTILADOR", "", "achado", "", "Fulano", exigir_foto=False)
    inventario.definir_foto_sobra(dados, sid, "http://x/1.webp")
    s = inventario.bens_da_sala(dados, eid, "01 - SALA CCI")["sobras"]
    assert len(s) == 1 and s[0]["foto_url"] == "http://x/1.webp" and inventario.resumo(dados, eid)["sobras"] == 1
    with pytest.raises(db.ErroDeNegocio):
        inventario.excluir_sobra(dados, eid, 999)
    assert inventario.excluir_sobra(dados, eid, sid)["descricao"] == "VENTILADOR"
    assert inventario.resumo(dados, eid)["sobras"] == 0
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest -q tests/test_inventario.py -k "ler_ or bens_da_sala or sobras"`
Expected: 3 falhas por atributo inexistente.

- [ ] **Step 3: Implementar** (após `resumo` em `inventario.py`)

```python
# ---------------------------------------------------------------- leituras
_LEITURA = "r.localizacao AS lido_em_sala, r.lido_em, r.integrante, r.conservacao, r.quem_usa, r.observacao, r.foto_url"


def ler(conn, evento_id: int, localizacao: str, numero: int, integrante: str) -> dict:
    """Registra (ou atualiza) a leitura do bem nesta sala. Qualquer bem cadastrado é aceito em qualquer sala;
    a divergência é só sinalizada (regra do sistema antigo)."""
    _evento_aberto_ou_erro(conn, evento_id)
    _sala_ou_erro(conn, evento_id, localizacao)
    if not conn.execute("SELECT 1 FROM inventario_integrantes WHERE evento_id = ? AND nome = ?", (evento_id, integrante)).fetchone():
        raise ErroDeNegocio("Escolha o integrante da comissão antes de ler.")
    bem = db.buscar_bem(conn, numero)
    if not bem:
        raise BemNaoEncontrado(numero)
    anterior = _um(conn, "SELECT * FROM inventario_leituras WHERE evento_id = ? AND numero = ?", evento_id, numero)
    agora = _agora()
    if anterior:
        conn.execute("UPDATE inventario_leituras SET localizacao = ?, lido_em = ?, integrante = ? WHERE id = ?",
                     (localizacao, agora, integrante, anterior["id"]))
    else:
        conn.execute("INSERT INTO inventario_leituras (evento_id, numero, localizacao, lido_em, integrante) VALUES (?,?,?,?,?)",
                     (evento_id, numero, localizacao, agora, integrante))
    conn.commit()
    return {"situacao": "localizado" if bem["localizacao"] == localizacao else "divergente", "bem": bem,
            "cadastrado_em": bem["localizacao"], "ativo": bem["situacao"] == "ATIVO", "reler": bool(anterior),
            "leitura_anterior": anterior, "lido_em": agora, "integrante": integrante}


def bens_da_sala(conn, evento_id: int, localizacao: str) -> dict:
    """bens: ativos cadastrados na sala (com a leitura do evento, se houver) e situacao_inv;
    trazidos: leituras feitas nesta sala de bens de outra sala ou não ativos; sobras: desta sala."""
    bens = _todos(conn, f"""
        SELECT b.*, {_LEITURA} FROM bens b
        LEFT JOIN inventario_leituras r ON r.numero = b.numero AND r.evento_id = ?
        WHERE b.localizacao = ? AND b.situacao = 'ATIVO' ORDER BY b.numero""", evento_id, localizacao)
    for b in bens:
        b["situacao_inv"] = "pendente" if not b["lido_em"] else ("localizado" if b["lido_em_sala"] == localizacao else "divergente")
    trazidos = _todos(conn, f"""
        SELECT b.*, {_LEITURA} FROM inventario_leituras r JOIN bens b ON b.numero = r.numero
        WHERE r.evento_id = ? AND r.localizacao = ? AND (b.localizacao <> ? OR b.situacao <> 'ATIVO')
        ORDER BY r.lido_em DESC""", evento_id, localizacao, localizacao)
    sobras = _todos(conn, "SELECT * FROM inventario_sobras WHERE evento_id = ? AND localizacao = ? ORDER BY id DESC",
                    evento_id, localizacao)
    return {"bens": bens, "trazidos": trazidos, "sobras": sobras}


def atualizar_leitura(conn, evento_id: int, numero: int, **campos) -> None:
    """Campos: conservacao, quem_usa, observacao, foto_url (só os presentes são gravados; '' vira NULL)."""
    _evento_aberto_ou_erro(conn, evento_id)
    permitidos = {"conservacao", "quem_usa", "observacao", "foto_url"}
    extra = set(campos) - permitidos
    if extra:
        raise ErroDeNegocio(f"Campo desconhecido: {', '.join(sorted(extra))}.")
    if "conservacao" in campos and campos["conservacao"] and campos["conservacao"] not in CONSERVACAO:
        raise ErroDeNegocio("Estado de conservação inválido.")
    if not _um(conn, "SELECT id FROM inventario_leituras WHERE evento_id = ? AND numero = ?", evento_id, numero):
        raise ErroDeNegocio("Leia o bem antes de preencher os dados.")
    for campo, valor in campos.items():
        conn.execute(f"UPDATE inventario_leituras SET {campo} = ? WHERE evento_id = ? AND numero = ?",
                     (_texto(valor) or None, evento_id, numero))
    conn.commit()


# ---------------------------------------------------------------- sobras
def registrar_sobra(conn, evento_id, localizacao, descricao, complemento, observacao, foto_url, integrante, exigir_foto=True) -> int:
    """Bem sem cadastro encontrado na sala. Foto obrigatória quando as fotos estão ativas (exigir_foto)."""
    _evento_aberto_ou_erro(conn, evento_id)
    _sala_ou_erro(conn, evento_id, localizacao)
    descricao = _obrigatorio(descricao, "Descrição")
    observacao = _obrigatorio(observacao, "Observação")
    integrante = _obrigatorio(integrante, "Integrante")
    if exigir_foto and not _texto(foto_url):
        raise ErroDeNegocio("A sobra precisa de foto.")
    cur = conn.execute("""INSERT INTO inventario_sobras (evento_id, localizacao, descricao, complemento, observacao, foto_url, integrante, criado_em)
                          VALUES (?,?,?,?,?,?,?,?)""",
                       (evento_id, localizacao, descricao, _texto(complemento) or None, observacao, _texto(foto_url), integrante, _agora()))
    conn.commit()
    return cur.lastrowid


def definir_foto_sobra(conn, sobra_id: int, foto_url: str) -> None:
    conn.execute("UPDATE inventario_sobras SET foto_url = ? WHERE id = ?", (_texto(foto_url), sobra_id))
    conn.commit()


def excluir_sobra(conn, evento_id: int, sobra_id: int) -> dict:
    """Só sobras podem ser apagadas (leituras de bens cadastrados, nunca). Devolve a sobra para apagar a foto."""
    _evento_aberto_ou_erro(conn, evento_id)
    s = _um(conn, "SELECT * FROM inventario_sobras WHERE evento_id = ? AND id = ?", evento_id, sobra_id)
    if not s:
        raise ErroDeNegocio("Sobra não encontrada.")
    conn.execute("DELETE FROM inventario_sobras WHERE id = ?", (sobra_id,))
    conn.commit()
    return s
```

`db.buscar_bem(conn, numero)` já existe.

- [ ] **Step 4: Rodar e ver passar** — `.venv/bin/pytest -q` → 129.

- [ ] **Step 5: Commit**

```bash
git add inventario.py tests/test_inventario.py
git commit -m "Inventário: leitura (localizado/divergente/reler), bens da sala, campos por bem e sobras"
```

---

### Task 3: Relatório, resumo por situação e `.xlsx`

**Files:**
- Modify: `inventario.py`
- Test: `tests/test_inventario.py`

**Interfaces:**
- Produces: `inventario.relatorio(conn, evento_id, localizacao=None, situacao=None) -> list[dict]` (`numero, descricao, complemento, classificacao, local_sistema, local_inventario, situacao_inv, lido_em, integrante, conservacao, quem_usa, observacao, foto_url`), `exportar_xlsx(conn, evento_id, destino, localizacao=None) -> destino`, `COLUNAS_XLSX`.

- [ ] **Step 1: Testes que falham**

```python
def test_relatorio_e_xlsx(dados, tmp_path):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.ler(dados, eid, "01 - SALA CCI", 2001, "Fulano")                     # divergente (é da SALA B)
    inventario.atualizar_leitura(dados, eid, 1001, conservacao="Bom", quem_usa="Ciclana")
    inventario.registrar_sobra(dados, eid, "02 - SALA B", "VENTILADOR", "ARNO", "sem plaqueta", "http://x/s.webp", "Beltrana")
    r = inventario.relatorio(dados, eid)
    assert [(x["numero"], x["situacao_inv"]) for x in r] == [(1001, "localizado"), (1002, "pendente"), (2001, "divergente"), (2002, "pendente"), (1004, "pendente")]
    assert r[2]["local_sistema"] == "02 - SALA B" and r[2]["local_inventario"] == "01 - SALA CCI"
    assert [x["numero"] for x in inventario.relatorio(dados, eid, localizacao="01 - SALA CCI")] == [1001, 1002, 2001]   # inclui o trazido
    assert [x["numero"] for x in inventario.relatorio(dados, eid, situacao="pendente")] == [1002, 2002, 1004]
    from openpyxl import load_workbook
    wb = load_workbook(inventario.exportar_xlsx(dados, eid, tmp_path / "inv.xlsx"))
    assert wb.sheetnames == ["Bens", "Sobras"]
    linhas = list(wb["Bens"].iter_rows(values_only=True))
    assert linhas[0][0].startswith("Inventário 2026") and linhas[1] == tuple(inventario.COLUNAS_XLSX)
    assert linhas[2][:3] == (1001, "CADEIRA", "GIRATÓRIA") and linhas[2][6] == "Localizado" and linhas[2][7] == "Bom"
    sobras = list(wb["Sobras"].iter_rows(values_only=True))
    assert sobras[1][0] == "Sala" and sobras[2][:2] == ("02 - SALA B", "VENTILADOR") and sobras[2][6] == "http://x/s.webp"
    so_cci = load_workbook(inventario.exportar_xlsx(dados, eid, tmp_path / "cci.xlsx", localizacao="01 - SALA CCI"))["Bens"]
    assert so_cci.max_row == 2 + 3
```

- [ ] **Step 2: Rodar e ver falhar** — `-k relatorio_e_xlsx` → `AttributeError`.

- [ ] **Step 3: Implementar** (após `excluir_sobra`)

```python
# ---------------------------------------------------------------- relatório e planilha do evento
COLUNAS_XLSX = ["Patrimônio", "Descrição", "Complemento", "Classificação", "Local sistema", "Local inventário",
                "Situação", "Conservação", "Quem usa", "Observação", "Integrante", "Data/hora", "Foto"]
COLUNAS_SOBRAS = ["Sala", "Descrição", "Complemento", "Observação", "Integrante", "Data/hora", "Foto"]
_CAMPOS_REL = """b.numero, b.descricao, b.complemento, b.classificacao, b.localizacao AS local_sistema,
        r.localizacao AS local_inventario, r.lido_em, r.integrante, r.conservacao, r.quem_usa, r.observacao, r.foto_url"""


def relatorio(conn, evento_id: int, localizacao: str | None = None, situacao: str | None = None) -> list[dict]:
    """Uma linha por bem ativo das salas do escopo (ou da sala pedida), mais os lidos nela vindos de fora
    do escopo. situacao filtra por localizado | divergente | pendente."""
    params = [evento_id, evento_id]
    filtro_sala = ""
    if localizacao:
        filtro_sala = " AND (s.localizacao = ? OR r.localizacao = ?)"
        params += [localizacao, localizacao]
    linhas = _todos(conn, f"""
        SELECT {_CAMPOS_REL},
          CASE WHEN r.id IS NULL THEN 'pendente' WHEN r.localizacao = b.localizacao THEN 'localizado' ELSE 'divergente' END AS situacao_inv
        FROM inventario_salas s JOIN bens b ON b.localizacao = s.localizacao AND b.situacao = 'ATIVO'
        LEFT JOIN inventario_leituras r ON r.evento_id = s.evento_id AND r.numero = b.numero
        WHERE s.evento_id = ?{filtro_sala}
        UNION ALL
        SELECT {_CAMPOS_REL}, 'divergente' AS situacao_inv
        FROM inventario_leituras r JOIN bens b ON b.numero = r.numero
        WHERE r.evento_id = ? AND b.localizacao NOT IN (SELECT localizacao FROM inventario_salas WHERE evento_id = r.evento_id)
          {"AND r.localizacao = ?" if localizacao else ""}
        ORDER BY local_sistema, numero""", *([evento_id] + ([localizacao, localizacao] if localizacao else []) + [evento_id] + ([localizacao] if localizacao else [])))
    if situacao:
        linhas = [x for x in linhas if x["situacao_inv"] == situacao]
    return linhas


def _data_br(iso):
    return f"{iso[8:10]}/{iso[5:7]}/{iso[:4]} {iso[11:16]}" if iso else ""


def exportar_xlsx(conn, evento_id: int, destino, localizacao: str | None = None):
    from openpyxl import Workbook
    e = evento(conn, evento_id)
    wb = Workbook()
    ws = wb.active
    ws.title = "Bens"
    ws.append([f"{e['nome']} — gerado em {_data_br(_agora())} — {('sala ' + localizacao) if localizacao else 'todas as salas'}"])
    ws.append(COLUNAS_XLSX)
    for x in relatorio(conn, evento_id, localizacao):
        ws.append([x["numero"], x["descricao"], x["complemento"], x["classificacao"], x["local_sistema"], x["local_inventario"],
                   ROTULO_SITUACAO[x["situacao_inv"]], x["conservacao"], x["quem_usa"], x["observacao"], x["integrante"],
                   _data_br(x["lido_em"]), x["foto_url"]])
    ws2 = wb.create_sheet("Sobras")
    ws2.append([f"{e['nome']} — sobras (bens sem cadastro)"])
    ws2.append(COLUNAS_SOBRAS)
    sql = "SELECT * FROM inventario_sobras WHERE evento_id = ?" + (" AND localizacao = ?" if localizacao else "") + " ORDER BY localizacao, id"
    for s in _todos(conn, sql, *([evento_id, localizacao] if localizacao else [evento_id])):
        ws2.append([s["localizacao"], s["descricao"], s["complemento"], s["observacao"], s["integrante"], _data_br(s["criado_em"]), s["foto_url"]])
    wb.save(destino)
    return destino
```

Atenção à montagem dos parâmetros em `relatorio`: a lista final deve ser, na ordem dos `?` do SQL: `evento_id`, (`localizacao`, `localizacao` se filtrar), `evento_id`, (`localizacao` se filtrar). O primeiro `params` do rascunho acima é descartado; use só a expressão passada a `_todos`. O trazido de fora do escopo aparece no relatório da sala onde foi lido; um bem de sala do escopo lido em outra sala do escopo aparece uma vez, na sala cadastral, como divergente com `local_inventario` = onde foi lido (é o que o teste da SALA CCI espera: 1001, 1002 e o 2001 trazido).

- [ ] **Step 4: Rodar e ver passar** — `.venv/bin/pytest -q` → 130.

- [ ] **Step 5: Commit**

```bash
git add inventario.py tests/test_inventario.py
git commit -m "Inventário: relatório por sala/situação e planilha do evento (abas Bens e Sobras)"
```

---

### Task 4: `fotos.py`, dependências e infra

**Files:**
- Create: `fotos.py`, `tests/test_fotos.py`
- Modify: `requirements.txt`, `Dockerfile`, `compose.yml`, `.gitignore`

**Interfaces:**
- Produces: `fotos.configurado()`, `fotos.validar(arquivo) -> bytes`, `fotos.comprimir(dados) -> bytes`, `fotos.enviar(nome, dados) -> str`, `fotos.apagar(url)`, `fotos.nome_bem(evento_id, numero)`, `fotos.nome_sobra(evento_id, sobra_id)`, `fotos.VARIAVEIS`.

- [ ] **Step 1: Instalar dependências**

```bash
printf 'flask\nopenpyxl\npython-docx\npywebview\npyinstaller\npytest\nPillow\nboto3\n' > requirements.txt
.venv/bin/pip install -q Pillow boto3
sed -i 's/RUN pip install --no-cache-dir flask openpyxl python-docx waitress/RUN pip install --no-cache-dir flask openpyxl python-docx waitress Pillow boto3/' Dockerfile
```

`compose.yml`, dentro de `web:` (após `restart`):

```yaml
    env_file:
      # Credenciais do bucket de fotos (R2_ACCESS_KEY_ID, R2_SECRET_ACCESS_KEY, R2_ENDPOINT_URL, R2_BUCKET_NAME, R2_PUBLIC_URL).
      # Sem o arquivo, as fotos do inventário ficam desativadas e o resto funciona.
      - path: secrets/.env
        required: false
```

`.gitignore`: acrescentar `secrets/`.

- [ ] **Step 2: Testes que falham**

`tests/test_fotos.py`:

```python
"""fotos.py: validação, compressão WebP e envio ao R2 (cliente falso; nada de rede)."""
import io

import pytest
from PIL import Image

import db
import fotos


class Arquivo:
    def __init__(self, nome, dados):
        self.filename, self._dados = nome, dados

    def read(self):
        return self._dados


def png(largura=100, altura=50):
    buf = io.BytesIO()
    Image.new("RGB", (largura, altura), (200, 10, 10)).save(buf, "PNG")
    return buf.getvalue()


def test_validar(monkeypatch):
    assert fotos.validar(Arquivo("a.PNG", png())) == png()
    with pytest.raises(db.ErroDeNegocio):
        fotos.validar(Arquivo("a.gif", png()))
    with pytest.raises(db.ErroDeNegocio):
        fotos.validar(Arquivo("a.jpg", b"isto nao e imagem"))
    monkeypatch.setattr(fotos, "TAMANHO_MAX", 10)
    with pytest.raises(db.ErroDeNegocio):
        fotos.validar(Arquivo("a.png", png()))


def test_comprimir_webp_e_redimensiona():
    saida = fotos.comprimir(png(4000, 1000))
    img = Image.open(io.BytesIO(saida))
    assert img.format == "WEBP" and img.size == (1920, 480)
    assert Image.open(io.BytesIO(fotos.comprimir(png()))).size == (100, 50)


class ClienteFalso:
    def __init__(self):
        self.enviados, self.apagados = [], []

    def put_object(self, Bucket, Key, Body, ContentType):
        self.enviados.append((Bucket, Key, len(Body), ContentType))

    def delete_object(self, Bucket, Key):
        self.apagados.append((Bucket, Key))


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
    url = fotos.enviar("INV1_BEM_1001_20260915120000.webp", b"webp")
    assert url == "https://fotos.exemplo.org/inventario/INV1_BEM_1001_20260915120000.webp"
    assert falso.enviados == [("fotos", "inventario/INV1_BEM_1001_20260915120000.webp", 4, "image/webp")]
    fotos.apagar(url)
    assert falso.apagados == [("fotos", "inventario/INV1_BEM_1001_20260915120000.webp")]
    fotos.apagar("https://outro/sem-prefixo.webp")                       # ignora
    assert len(falso.apagados) == 1
    monkeypatch.delenv("R2_PUBLIC_URL")
    assert fotos.enviar("a.webp", b"1") == "https://acc.r2.cloudflarestorage.com/fotos/inventario/a.webp"
    assert fotos.nome_bem(3, 1001).startswith("INV3_BEM_1001_") and fotos.nome_sobra(3, 7).startswith("INV3_SOBRA_7_")
    assert fotos.nome_bem(3, 1001).endswith(".webp")
```

- [ ] **Step 3: Rodar e ver falhar** — `.venv/bin/pytest -q tests/test_fotos.py` → `ModuleNotFoundError: fotos`.

- [ ] **Step 4: `fotos.py`**

```python
"""Fotos do inventário no Cloudflare R2 (API S3): validação, compressão para WebP e envio.
Transcrito do sistema antigo (sga/cfc: app/services/storage.py e images.py), sem Flask.

Credenciais nas variáveis de ambiente R2_* (compose.yml lê secrets/.env). Sem elas, `configurado()` é
falso e as telas desativam os botões de foto; o resto do inventário funciona (programa Windows offline)."""
import io
import os
from datetime import datetime

from db import ErroDeNegocio

VARIAVEIS = ("R2_ACCESS_KEY_ID", "R2_SECRET_ACCESS_KEY", "R2_ENDPOINT_URL", "R2_BUCKET_NAME")
EXTENSOES = {".jpg", ".jpeg", ".png", ".webp"}
TAMANHO_MAX = 5 * 1024 * 1024
LARGURA_MAX, ALTURA_MAX, QUALIDADE = 1920, 1080, 85
PREFIXO = "inventario/"


def configurado() -> bool:
    return all(os.environ.get(v) for v in VARIAVEIS)


def validar(arquivo) -> bytes:
    """arquivo: objeto com .filename e .read() (FileStorage do Flask). Devolve os bytes se for imagem válida."""
    nome = (getattr(arquivo, "filename", "") or "").lower()
    if os.path.splitext(nome)[1] not in EXTENSOES:
        raise ErroDeNegocio("Foto: envie um arquivo .jpg, .png ou .webp.")
    dados = arquivo.read()
    if len(dados) > TAMANHO_MAX:
        raise ErroDeNegocio("Foto maior que 5 MB.")
    from PIL import Image
    try:
        Image.open(io.BytesIO(dados)).verify()
    except Exception:
        raise ErroDeNegocio("O arquivo não é uma imagem válida.")
    return dados


def comprimir(dados: bytes) -> bytes:
    """WebP qualidade 85, no máximo 1920×1080 mantendo a proporção, orientação EXIF aplicada."""
    from PIL import Image, ImageOps
    img = ImageOps.exif_transpose(Image.open(io.BytesIO(dados)))
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGB")
    largura, altura = img.size
    if largura > LARGURA_MAX or altura > ALTURA_MAX:
        razao = min(LARGURA_MAX / largura, ALTURA_MAX / altura)
        img = img.resize((int(largura * razao), int(altura * razao)), Image.Resampling.LANCZOS)
    saida = io.BytesIO()
    img.save(saida, "WEBP", quality=QUALIDADE, optimize=True)
    return saida.getvalue()


def _cliente():
    import boto3
    from botocore.config import Config
    return boto3.client("s3", aws_access_key_id=os.environ["R2_ACCESS_KEY_ID"],
                        aws_secret_access_key=os.environ["R2_SECRET_ACCESS_KEY"], region_name="auto",
                        endpoint_url=os.environ["R2_ENDPOINT_URL"],
                        config=Config(signature_version="s3v4", s3={"addressing_style": "path"}))


def _url(chave: str) -> str:
    base = os.environ.get("R2_PUBLIC_URL")
    if base:
        return f"{base.rstrip('/')}/{chave}"
    return f"{os.environ['R2_ENDPOINT_URL'].rstrip('/')}/{os.environ['R2_BUCKET_NAME']}/{chave}"


def enviar(nome: str, dados: bytes) -> str:
    """Grava `inventario/<nome>` no bucket e devolve a URL pública."""
    if not configurado():
        raise ErroDeNegocio("Fotos desativadas: bucket não configurado.")
    chave = PREFIXO + nome
    _cliente().put_object(Bucket=os.environ["R2_BUCKET_NAME"], Key=chave, Body=dados, ContentType="image/webp")
    return _url(chave)


def apagar(url: str | None) -> None:
    """Apaga o objeto pela chave contida na URL; erro só é ignorado (a URL some do banco de qualquer jeito)."""
    if not url or not configurado():
        return
    pos = url.find(PREFIXO)
    if pos < 0:
        return
    try:
        _cliente().delete_object(Bucket=os.environ["R2_BUCKET_NAME"], Key=url[pos:])
    except Exception:
        pass


def _carimbo() -> str:
    return datetime.now().strftime("%Y%m%d%H%M%S")


def nome_bem(evento_id: int, numero: int) -> str:
    return f"INV{evento_id}_BEM_{numero}_{_carimbo()}.webp"


def nome_sobra(evento_id: int, sobra_id: int) -> str:
    return f"INV{evento_id}_SOBRA_{sobra_id}_{_carimbo()}.webp"
```

- [ ] **Step 5: Rodar e ver passar** — `.venv/bin/pytest -q` → 133.

- [ ] **Step 6: Commit**

```bash
git add fotos.py tests/test_fotos.py requirements.txt Dockerfile compose.yml .gitignore
git commit -m "Inventário: fotos no R2 (validação, WebP, envio) e dependências Pillow/boto3"
```

---

### Task 5: Abas `inv_*` na planilha de cadastros

**Files:**
- Modify: `inventario.py` (`ABAS`, `exportar_abas`, `validar_abas`, `substituir_tabelas`), `db.py` (`_ler_aba_cadastro` com `colunas`/`opcional`; `exportar_cadastros`; `importar_cadastros`)
- Test: `tests/test_inventario.py`

**Interfaces:**
- Produces: `inventario.ABAS` (dict aba → colunas), `inventario.exportar_abas(conn, wb)`, `inventario.validar_abas(conn, brutos) -> (linhas_por_aba, problemas)`, `inventario.substituir_tabelas(conn, linhas_por_aba)` (sem commit; chamado dentro da transação de `importar_cadastros`).

- [ ] **Step 1: Testes que falham**

```python
def test_planilha_de_cadastros_exporta_e_importa_abas_de_inventario(dados, tmp_path):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.registrar_sobra(dados, eid, "02 - SALA B", "VENTILADOR", None, "achado", "http://x/s.webp", "Fulano")
    from openpyxl import load_workbook
    caminho = db.exportar_cadastros(dados, tmp_path / "c.xlsx")
    wb = load_workbook(caminho)
    assert wb.sheetnames == ["responsaveis", "localizacoes", "pessoas", "atribuicoes", "inv_eventos", "inv_integrantes", "inv_salas", "inv_leituras", "inv_sobras"]
    assert list(wb["inv_leituras"].iter_rows(values_only=True))[1][:3] == (eid, 1001, "01 - SALA CCI")
    # editar: encerra o evento, acrescenta uma leitura migrada de outro sistema, e reimporta
    ws = wb["inv_eventos"]
    ws.cell(row=2, column=5, value="2026-01-31")                                       # encerrado_em só com data
    wb["inv_leituras"].append([eid, 2001, "02 - SALA B", "2026-01-20 10:00:00", "Antigo", "Regular", "", "migrado", ""])
    wb.save(tmp_path / "c2.xlsx")
    with open(tmp_path / "c2.xlsx", "rb") as f:
        r = db.importar_cadastros(dados, f)
    assert r["inv_leituras"] == 2 and r["inv_eventos"] == 1
    assert inventario.evento(dados, eid)["encerrado_em"] == "2026-01-31 00:00:00"
    assert {x["numero"]: x["situacao_inv"] for x in inventario.relatorio(dados, eid)}[2001] == "localizado"
    assert inventario.resumo(dados, eid)["sobras"] == 1


def test_planilha_sem_abas_de_inventario_nao_toca_nas_tabelas(dados, tmp_path):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    from openpyxl import load_workbook
    wb = load_workbook(db.exportar_cadastros(dados, tmp_path / "c.xlsx"))
    for aba in list(inventario.ABAS):
        wb.remove(wb[aba])
    wb.save(tmp_path / "so4.xlsx")
    with open(tmp_path / "so4.xlsx", "rb") as f:
        r = db.importar_cadastros(dados, f)
    assert "inv_leituras" not in r and inventario.resumo(dados, eid)["lidos"] == 1


def test_planilha_de_inventario_validacoes(dados, tmp_path):
    eid = semear_inventario(dados)
    from openpyxl import load_workbook
    wb = load_workbook(db.exportar_cadastros(dados, tmp_path / "c.xlsx"))
    wb["inv_leituras"].append([eid, 99999, "01 - SALA CCI", "2026-01-20 10:00:00", "Fulano", "", "", "", ""])   # bem inexistente
    wb["inv_leituras"].append([eid, 1001, "01 - SALA CCI", "x", "Fulano", "Ótimo", "", "", ""])                  # data e conservação
    wb["inv_salas"].append([77, "01 - SALA CCI"])                                                                # evento inexistente
    wb["inv_eventos"].append([2, "Outro aberto", None, "2026-02-01 00:00:00", None])                             # 2 abertos
    wb.save(tmp_path / "ruim.xlsx")
    with open(tmp_path / "ruim.xlsx", "rb") as f, pytest.raises(db.ImportacaoInvalida) as e:
        db.importar_cadastros(dados, f)
    msg = str(e.value)
    assert "99999" in msg and "conservação" in msg and "77" in msg and "aberto" in msg and "data" in msg
    assert inventario.resumo(dados, eid)["salas"] == 3                                                            # nada mudou
```

- [ ] **Step 2: Rodar e ver falhar** — `-k planilha` → falhas (`AttributeError: ABAS`, abas ausentes).

- [ ] **Step 3: `db.py`**

`_ler_aba_cadastro(wb, tabela, problemas, colunas=None, opcional=False)`: `colunas = colunas or CADASTROS[tabela]`; se `tabela not in wb.sheetnames`: `if opcional: return None` senão o `problemas.append(...)` atual. Resto igual.

`exportar_cadastros`: após o `for` das 4 abas e antes de `wb.save`: `import inventario; inventario.exportar_abas(conn, wb)`.

`importar_cadastros`: após ler `brutos` das 4 abas (e antes do primeiro `if problemas`):

```python
    import inventario
    inv_brutos = {aba: _ler_aba_cadastro(wb, aba, problemas, colunas=cols, opcional=True) for aba, cols in inventario.ABAS.items()}
    tem_inventario = any(v is not None for v in inv_brutos.values())
```

Depois das validações das 4 abas (antes do `if problemas:` final que levanta `ImportacaoInvalida`):

```python
    inv_linhas = {}
    if tem_inventario:
        inv_linhas, inv_problemas = inventario.validar_abas(conn, {a: (v or []) for a, v in inv_brutos.items()})
        problemas += inv_problemas
```

Dentro do `try` da transação, após os 4 `executemany` e antes do `commit`: `if tem_inventario: inventario.substituir_tabelas(conn, inv_linhas)`. No `return`, se `tem_inventario`, acrescente `{aba: len(l) for aba, l in inv_linhas.items()}`.

- [ ] **Step 4: `inventario.py`** (ao final)

```python
# ---------------------------------------------------------------- planilha de cadastros (migração de inventários)
ABAS = {
    "inv_eventos": ["id", "nome", "descricao", "aberto_em", "encerrado_em"],
    "inv_integrantes": ["evento_id", "nome"],
    "inv_salas": ["evento_id", "localizacao"],
    "inv_leituras": ["evento_id", "numero", "localizacao", "lido_em", "integrante", "conservacao", "quem_usa", "observacao", "foto_url"],
    "inv_sobras": ["evento_id", "localizacao", "descricao", "complemento", "observacao", "foto_url", "integrante", "criado_em"],
}
_TABELA = {aba: "inventario_" + aba[4:] for aba in ABAS}


def exportar_abas(conn, wb) -> None:
    for aba, colunas in ABAS.items():
        ws = wb.create_sheet(aba)
        ws.append(colunas)
        for linha in conn.execute(f"SELECT {', '.join(colunas)} FROM {_TABELA[aba]} ORDER BY {colunas[0]}, {colunas[1]}"):
            ws.append(list(linha))


def _data_iso(valor, rotulo, linha, problemas, obrigatoria):
    """Aceita datetime do Excel, 'YYYY-MM-DD HH:MM:SS' ou 'YYYY-MM-DD' (vira 00:00:00). Vazio → None."""
    from datetime import datetime
    if valor is None or _texto(valor) == "":
        if obrigatoria:
            problemas.append(f"{linha}: {rotulo} vazia")
        return None
    if isinstance(valor, datetime):
        return valor.strftime("%Y-%m-%d %H:%M:%S")
    t = _texto(valor)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(t, fmt).strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            pass
    problemas.append(f"{linha}: {rotulo} inválida ({t}); use AAAA-MM-DD HH:MM:SS")
    return None


def validar_abas(conn, brutos: dict) -> tuple[dict, list]:
    """brutos: {aba: [linhas dict com _linha]} (aba ausente = []). Devolve ({aba: [tuplas p/ INSERT]}, problemas)."""
    problemas: list[str] = []
    linhas: dict = {aba: [] for aba in ABAS}
    ids, abertos = set(), 0
    for r in brutos["inv_eventos"]:
        rot = f"inv_eventos linha {r['_linha']}"
        try:
            eid = int(r["id"])
        except (TypeError, ValueError):
            problemas.append(f"{rot}: id inválido"); continue
        if eid in ids:
            problemas.append(f"{rot}: id {eid} repetido"); continue
        nome = _texto(r["nome"])
        if not nome:
            problemas.append(f"{rot}: nome vazio"); continue
        aberto = _data_iso(r["aberto_em"], "aberto_em", rot, problemas, True)
        encerrado = _data_iso(r["encerrado_em"], "encerrado_em", rot, problemas, False)
        if encerrado is None and _texto(r["encerrado_em"]) == "":
            abertos += 1
        ids.add(eid)
        linhas["inv_eventos"].append((eid, nome, _texto(r["descricao"]) or None, aberto, encerrado))
    if abertos > 1:
        problemas.append("inv_eventos: mais de um evento aberto (sem encerrado_em)")

    def evento_ok(r, rot):
        try:
            eid = int(r["evento_id"])
        except (TypeError, ValueError):
            problemas.append(f"{rot}: evento_id inválido"); return None
        if eid not in ids:
            problemas.append(f"{rot}: evento {eid} não está na aba inv_eventos"); return None
        return eid

    vistos = set()
    for r in brutos["inv_integrantes"]:
        rot = f"inv_integrantes linha {r['_linha']}"
        eid, nome = evento_ok(r, rot), " ".join(_texto(r["nome"]).split())
        if eid is None or not nome or (eid, nome) in vistos:
            continue
        vistos.add((eid, nome))
        linhas["inv_integrantes"].append((eid, nome))
    vistos = set()
    for r in brutos["inv_salas"]:
        rot = f"inv_salas linha {r['_linha']}"
        eid, loc = evento_ok(r, rot), _texto(r["localizacao"])
        if eid is None or not loc or (eid, loc) in vistos:
            continue
        vistos.add((eid, loc))
        linhas["inv_salas"].append((eid, loc))
    vistos = set()
    for r in brutos["inv_leituras"]:
        rot = f"inv_leituras linha {r['_linha']}"
        eid = evento_ok(r, rot)
        num = db._numero(r["numero"])
        if eid is None:
            continue
        if num is None or num != int(num) or not db.buscar_bem(conn, int(num)):
            problemas.append(f"{rot}: bem {_texto(r['numero']) or '(vazio)'} não existe na base"); continue
        num = int(num)
        if (eid, num) in vistos:
            problemas.append(f"{rot}: bem {num} repetido no evento {eid}"); continue
        cons = _texto(r["conservacao"]) or None
        if cons and cons not in CONSERVACAO:
            problemas.append(f"{rot}: conservação inválida ({cons})"); continue
        loc, integ = _texto(r["localizacao"]), _texto(r["integrante"])
        if not loc or not integ:
            problemas.append(f"{rot}: localização e integrante são obrigatórios"); continue
        lido = _data_iso(r["lido_em"], "data lido_em", rot, problemas, True)
        if lido is None:
            continue
        vistos.add((eid, num))
        linhas["inv_leituras"].append((eid, num, loc, lido, integ, cons, _texto(r["quem_usa"]) or None,
                                       _texto(r["observacao"]) or None, _texto(r["foto_url"]) or None))
    for r in brutos["inv_sobras"]:
        rot = f"inv_sobras linha {r['_linha']}"
        eid = evento_ok(r, rot)
        if eid is None:
            continue
        loc, desc, obs, integ = (_texto(r[c]) for c in ("localizacao", "descricao", "observacao", "integrante"))
        if not (loc and desc and obs and integ):
            problemas.append(f"{rot}: localização, descrição, observação e integrante são obrigatórios"); continue
        criado = _data_iso(r["criado_em"], "data criado_em", rot, problemas, True)
        if criado is None:
            continue
        linhas["inv_sobras"].append((eid, loc, desc, _texto(r["complemento"]) or None, obs, _texto(r["foto_url"]), integ, criado))
    return linhas, problemas


def substituir_tabelas(conn, linhas: dict) -> None:
    """Dentro da transação de db.importar_cadastros: apaga e regrava as 5 tabelas (ids de evento preservados)."""
    for aba in reversed(list(ABAS)):
        conn.execute(f"DELETE FROM {_TABELA[aba]}")
    conn.executemany("INSERT INTO inventario_eventos (id, nome, descricao, aberto_em, encerrado_em) VALUES (?,?,?,?,?)", linhas["inv_eventos"])
    conn.executemany("INSERT INTO inventario_integrantes VALUES (?,?)", linhas["inv_integrantes"])
    conn.executemany("INSERT INTO inventario_salas VALUES (?,?)", linhas["inv_salas"])
    conn.executemany("INSERT INTO inventario_leituras (evento_id, numero, localizacao, lido_em, integrante, conservacao, quem_usa, observacao, foto_url) VALUES (?,?,?,?,?,?,?,?,?)", linhas["inv_leituras"])
    conn.executemany("INSERT INTO inventario_sobras (evento_id, localizacao, descricao, complemento, observacao, foto_url, integrante, criado_em) VALUES (?,?,?,?,?,?,?,?)", linhas["inv_sobras"])
```

`db._numero` já existe (converte célula em número ou `None`). A mensagem da validação de conservação precisa conter a palavra "conservação", e a de data a palavra "data" (o teste procura). Em `inv_eventos` o teste do "2 abertos" espera a palavra "aberto".

- [ ] **Step 5: Rodar e ver passar** — `.venv/bin/pytest -q` → 136. Verifique que `test_exportar_importar_cadastros` antigo (4 abas) ainda passa: a exportação agora tem 9 abas e a importação continua aceitando 4.

- [ ] **Step 6: Commit**

```bash
git add db.py inventario.py tests/test_inventario.py
git commit -m "Planilha de cadastros: abas inv_* (exportar e importar inventários, migração de sistemas antigos)"
```

---

### Task 6: Blueprint — eventos e salas do evento, menu

**Files:**
- Create: `app_inventario.py`, `templates/inventario_eventos.html`, `templates/inventario_evento.html`
- Modify: `app.py` (`register_blueprint`, MENU)
- Test: `tests/test_app.py`

**Interfaces:**
- Produces: blueprint `inventario_bp` com rotas `inventario.eventos_tela` (`/inventario`), `inventario.abrir` (POST), `inventario.evento_tela` (`/inventario/<id>`), `inventario.encerrar` (POST), `inventario.integrante` (POST); helper `app_inventario._conn()`.

- [ ] **Step 1: Teste que falha**

```python
def test_inventario_eventos_abrir_e_encerrar(cliente):
    r = cliente.get("/inventario")
    assert r.status_code == 200 and b"Abrir evento" in r.data and b"Nenhum evento aberto" in r.data
    r = cliente.post("/inventario/abrir", data={"nome": "Inventário 2026", "descricao": "Portaria 1", "integrantes": "Fulano\nBeltrana", "escopo": "todas"}, follow_redirects=True)
    assert "Inventário 2026".encode() in r.data and b"01 - SALA CCI" in r.data and b"99 - SEM MAPA" in r.data
    assert b"Inventário" in cliente.get("/").data                                  # menu
    r = cliente.post("/inventario/abrir", data={"nome": "Outro", "integrantes": "X", "escopo": "todas"}, follow_redirects=True)
    assert "já existe".encode() in r.data.lower() or "Já existe".encode() in r.data
    import db, inventario
    eid = inventario.evento_aberto(db.conectar())["id"]
    r = cliente.post(f"/inventario/{eid}/integrante", data={"integrante": "Fulano", "volta": f"/inventario/{eid}"}, follow_redirects=True)
    assert b"Fulano" in r.data
    r = cliente.post(f"/inventario/{eid}/encerrar", data={}, follow_redirects=True)
    assert b"Confirmar encerramento" in r.data
    r = cliente.post(f"/inventario/{eid}/encerrar", data={"confirmar": "1"}, follow_redirects=True)
    assert b"encerrado" in r.data
    assert cliente.get("/inventario/999").status_code == 404


def test_inventario_abrir_com_amostragem(cliente):
    r = cliente.post("/inventario/abrir", data={"nome": "Amostra", "integrantes": "A", "escopo": "escolher", "salas": ["99 - SEM MAPA"]}, follow_redirects=True)
    assert b"99 - SEM MAPA" in r.data and b"01 - SALA CCI" not in r.data.split(b"<tbody>")[1]
```

- [ ] **Step 2: Rodar e ver falhar** — `-k inventario_` → 404.

- [ ] **Step 3: `app_inventario.py`**

```python
"""Rotas do módulo de inventário (blueprint /inventario). Dados em inventario.py; fotos em fotos.py.
A conexão por request e o errorhandler de ErroDeNegocio são os de app.py (g.conn e handler global)."""
import io

from flask import Blueprint, abort, flash, g, jsonify, redirect, render_template, request, send_file, session, url_for

import db
import fotos
import inventario

inventario_bp = Blueprint("inventario", __name__, url_prefix="/inventario")


def _conn():
    if "conn" not in g:
        g.conn = db.conectar()
    return g.conn


def _evento_ou_404(conn, id):
    return inventario.evento(conn, id) or abort(404)


def _trilha(e=None, *resto):
    t = [("Inventário", url_for("inventario.eventos_tela"))]
    if e:
        t.append((e["nome"], url_for("inventario.evento_tela", id=e["id"])))
    t += list(resto)
    t[-1] = (t[-1][0], None)
    return t


@inventario_bp.route("")
def eventos_tela():
    conn = _conn()
    aberto = inventario.evento_aberto(conn)
    return render_template("inventario_eventos.html", aberto=inventario.evento(conn, aberto["id"]) if aberto else None,
                           eventos=[e for e in inventario.eventos(conn) if e["encerrado_em"]],
                           salas_ativas=db.localizacoes_ativas(conn), trilha=_trilha())


@inventario_bp.route("/abrir", methods=["POST"])
def abrir():
    f = request.form
    salas = None if f.get("escopo", "todas") == "todas" else f.getlist("salas")
    eid = inventario.abrir_evento(_conn(), f.get("nome", ""), f.get("descricao", ""), f.get("integrantes", "").splitlines(), salas)
    flash("Evento aberto. Escolha o integrante e comece pelas salas.", "success")
    return redirect(url_for("inventario.evento_tela", id=eid))


@inventario_bp.route("/<int:id>")
def evento_tela(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    return render_template("inventario_evento.html", e=e, salas=inventario.salas(conn, id),
                           integrante=session.get("integrante"), confirmar=request.args.get("confirmar"), trilha=_trilha(e))


@inventario_bp.route("/<int:id>/encerrar", methods=["POST"])
def encerrar(id):
    conn = _conn()
    _evento_ou_404(conn, id)
    if not request.form.get("confirmar"):
        return redirect(url_for("inventario.evento_tela", id=id, confirmar="encerrar"))
    inventario.encerrar_evento(conn, id)
    flash("Evento encerrado. As leituras ficam congeladas; relatório e planilha continuam disponíveis.", "success")
    return redirect(url_for("inventario.evento_tela", id=id))


@inventario_bp.route("/<int:id>/integrante", methods=["POST"])
def integrante(id):
    e = _evento_ou_404(_conn(), id)
    nome = request.form.get("integrante", "")
    if nome not in e["integrantes"]:
        raise db.ErroDeNegocio("Integrante não está na comissão deste evento.")
    session["integrante"] = nome
    volta = request.form.get("volta") or url_for("inventario.evento_tela", id=id)
    return redirect(volta if volta.startswith("/") else url_for("inventario.evento_tela", id=id))
```

Em `app.py`, após as importações: `from app_inventario import inventario_bp` e, logo após `app.secret_key = ...`: `app.register_blueprint(inventario_bp)`. No `MENU`, após `("Recorte", ...)`: `("Inventário", "fa-clipboard-check", url_for("inventario.eventos_tela")),`.

- [ ] **Step 4: Templates**

`templates/inventario_eventos.html`:

```jinja
{% extends "base.html" %}
{% from "_macros.html" import cabecalho_tabela %}
{% block titulo %}Inventário{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Inventário</h1></div>
{% if aberto %}
<div class="br-card mb-4">
  <div class="card-header"><div class="text-weight-semi-bold text-up-02">{{ aberto.nome }}</div>
    <div class="text-down-01 text-gray-70">Aberto em {{ aberto.aberto_em[8:10] }}/{{ aberto.aberto_em[5:7] }}/{{ aberto.aberto_em[:4] }}{% if aberto.descricao %} · {{ aberto.descricao }}{% endif %} · Comissão: {{ aberto.integrantes|join(', ') }}</div></div>
  <div class="card-content">
    {% set r = aberto.resumo %}
    <div class="row">
      {% for valor, legenda in [(r.salas_iniciadas ~ ' / ' ~ r.salas, 'salas iniciadas'), (r.lidos ~ ' / ' ~ r.bens, 'bens localizados'), (r.divergentes, 'divergentes'), (r.pendentes, 'pendentes'), (r.sobras, 'sobras')] %}
      <div class="col-6 col-md mb-3"><div class="text-up-03 text-weight-bold">{{ valor }}</div><div class="text-down-01 text-gray-70">{{ legenda }}</div></div>
      {% endfor %}
    </div>
    <div class="br-message info mb-3"><div class="icon"><i class="fas fa-info-circle fa-lg" aria-hidden="true"></i></div>
      <div class="content"><span class="message-body">{{ r.pct_bens }}% dos bens localizados.</span></div></div>
    <a class="br-button primary" href="{{ url_for('inventario.evento_tela', id=aberto.id) }}"><i class="fas fa-door-open mr-1" aria-hidden="true"></i>Salas e leitura</a>
    <a class="br-button secondary ml-2" href="{{ url_for('inventario.relatorio_tela', id=aberto.id) }}"><i class="fas fa-table mr-1" aria-hidden="true"></i>Relatório</a>
  </div>
</div>
{% else %}
<p class="text-gray-70">Nenhum evento aberto. Abra um evento para começar a conferência.</p>
<form method="post" action="{{ url_for('inventario.abrir') }}" class="br-card mb-4"><div class="card-content">
  <div class="text-weight-semi-bold text-up-01 mb-3">Abrir evento</div>
  <div class="row">
    <div class="col-md-6 mb-3"><div class="br-input"><label for="nome">Nome</label><input id="nome" name="nome" type="text" placeholder="Ex.: Inventário 2026" required/></div></div>
    <div class="col-md-6 mb-3"><div class="br-input"><label for="descricao">Descrição (portaria, observações)</label><input id="descricao" name="descricao" type="text"/></div></div>
    <div class="col-md-6 mb-3"><div class="br-textarea"><label for="integrantes">Integrantes da comissão (um por linha)</label><textarea id="integrantes" name="integrantes" rows="4" required></textarea></div></div>
    <div class="col-md-6 mb-3">
      <div class="text-weight-semi-bold mb-2">Escopo</div>
      <div class="br-radio mb-2"><input id="escopo-todas" name="escopo" type="radio" value="todas" checked/><label for="escopo-todas">Todas as salas com bens ativos ({{ salas_ativas|length }})</label></div>
      <div class="br-radio mb-2"><input id="escopo-escolher" name="escopo" type="radio" value="escolher"/><label for="escopo-escolher">Escolher salas (amostragem)</label></div>
      <div id="lista-salas" hidden>
        {% for s in salas_ativas %}<div class="br-checkbox"><input id="sala-{{ loop.index }}" name="salas" type="checkbox" value="{{ s }}"/><label for="sala-{{ loop.index }}">{{ s }}</label></div>{% endfor %}
      </div>
    </div>
  </div>
  <button class="br-button primary" type="submit"><i class="fas fa-play mr-1" aria-hidden="true"></i>Abrir evento</button>
</div></form>
{% endif %}
{% if eventos %}
{{ cabecalho_tabela('Eventos encerrados', 'eventos') }}
  <thead><tr><th scope="col">Evento</th><th scope="col">Aberto em</th><th scope="col">Encerrado em</th><th scope="col" class="dsgov-acoes">Relatório</th></tr></thead>
  <tbody>{% for e in eventos %}
  <tr><td><a href="{{ url_for('inventario.evento_tela', id=e.id) }}">{{ e.nome }}</a></td><td>{{ e.aberto_em[8:10] }}/{{ e.aberto_em[5:7] }}/{{ e.aberto_em[:4] }}</td><td>{{ e.encerrado_em[8:10] }}/{{ e.encerrado_em[5:7] }}/{{ e.encerrado_em[:4] }}</td>
    <td class="dsgov-acoes"><a class="br-button circle small" href="{{ url_for('inventario.relatorio_tela', id=e.id) }}" aria-label="Relatório de {{ e.nome }}"><i class="fas fa-table" aria-hidden="true"></i></a></td></tr>
  {% endfor %}</tbody>
</table></div>
{% endif %}
{% endblock %}
{% block scripts %}
<script>
document.querySelectorAll('input[name="escopo"]').forEach(function (r) {
  r.addEventListener("change", function () { document.getElementById("lista-salas").hidden = this.value !== "escolher"; });
});
</script>
{% endblock %}
```

`templates/inventario_evento.html`:

```jinja
{% extends "base.html" %}
{% from "_macros.html" import select, cabecalho_tabela %}
{% block titulo %}{{ e.nome }}{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-3">
  <h1 class="mb-0">{{ e.nome }}</h1>
  <div class="ml-auto">
    <a class="br-button secondary" href="{{ url_for('inventario.relatorio_tela', id=e.id) }}"><i class="fas fa-table mr-1" aria-hidden="true"></i>Relatório</a>
    {% if not e.encerrado_em %}
    <form method="post" action="{{ url_for('inventario.encerrar', id=e.id) }}" class="d-inline ml-2">
      {% if confirmar == 'encerrar' %}<input type="hidden" name="confirmar" value="1"/><button class="br-button primary" type="submit"><i class="fas fa-check mr-1" aria-hidden="true"></i>Confirmar encerramento</button>
      {% else %}<button class="br-button" type="submit"><i class="fas fa-lock mr-1" aria-hidden="true"></i>Encerrar evento</button>{% endif %}
    </form>
    {% endif %}
  </div>
</div>
{% if e.encerrado_em %}
<div class="br-message warning mb-3"><div class="icon"><i class="fas fa-lock fa-lg" aria-hidden="true"></i></div>
  <div class="content"><span class="message-title">Evento encerrado</span><span class="message-body"> em {{ e.encerrado_em[8:10] }}/{{ e.encerrado_em[5:7] }}/{{ e.encerrado_em[:4] }}. Somente consulta.</span></div></div>
{% else %}
<form method="post" action="{{ url_for('inventario.integrante', id=e.id) }}" class="row mb-3 align-items-end">
  <input type="hidden" name="volta" value="{{ request.path }}"/>
  <div class="col-md-4 mb-2">{{ select('integrante', 'Quem está lendo (integrante da comissão)', e.integrantes, selecionado=integrante) }}</div>
  <div class="col-md-2 mb-2"><button class="br-button secondary" type="submit">Escolher</button></div>
  {% if integrante %}<div class="col-md-6 mb-2 text-gray-70">Leituras serão registradas por <strong>{{ integrante }}</strong>.</div>{% endif %}
</form>
{% endif %}
{% set r = e.resumo %}
<p class="text-gray-70">{{ r.salas_iniciadas }} de {{ r.salas }} salas iniciadas · {{ r.lidos }} de {{ r.bens }} bens localizados ({{ r.pct_bens }}%) · {{ r.divergentes }} divergentes · {{ r.sobras }} sobras</p>
{{ cabecalho_tabela('Salas do evento', 'salas') }}
  <thead><tr><th scope="col">Sala</th><th scope="col">Centro</th><th scope="col" class="dsgov-numero">Bens</th><th scope="col" class="dsgov-numero">Localizados</th><th scope="col" class="dsgov-numero">Divergentes</th><th scope="col" class="dsgov-numero">Pendentes</th><th scope="col">Situação</th><th scope="col" class="dsgov-acoes">Ler</th></tr></thead>
  <tbody>{% for s in salas %}
  <tr><td>{{ s.localizacao }}</td><td>{{ s.ccustos or '—' }}</td><td class="dsgov-numero">{{ s.total }}</td><td class="dsgov-numero">{{ s.localizados }}</td><td class="dsgov-numero">{{ s.divergentes }}</td><td class="dsgov-numero">{{ s.pendentes }}</td>
    <td>{% if s.total and s.pendentes == 0 %}<span class="br-tag bg-success text-pure-0"><span>completa</span></span>{% elif s.localizados or s.divergentes %}<span class="br-tag bg-warning"><span>em andamento</span></span>{% else %}<span class="br-tag bg-gray-20"><span>não iniciada</span></span>{% endif %}</td>
    <td class="dsgov-acoes"><a class="br-button circle small primary" href="{{ url_for('inventario.sala_tela', id=e.id, localizacao=s.localizacao) }}" aria-label="Ler {{ s.localizacao }}"><i class="fas fa-barcode" aria-hidden="true"></i></a></td></tr>
  {% endfor %}</tbody>
</table></div>
{% endblock %}
```

Os `url_for` de `inventario.relatorio_tela` e `inventario.sala_tela` só existem nas Tasks 7 e 8: nesta task, crie as duas rotas **provisórias** em `app_inventario.py` (substituídas depois):

```python
@inventario_bp.route("/<int:id>/sala/<path:localizacao>")
def sala_tela(id, localizacao):
    return redirect(url_for("inventario.evento_tela", id=id))      # completada na Task 7


@inventario_bp.route("/<int:id>/relatorio")
def relatorio_tela(id):
    return redirect(url_for("inventario.evento_tela", id=id))      # completada na Task 8
```

- [ ] **Step 5: Rodar e ver passar** — `.venv/bin/pytest -q` → 138.

- [ ] **Step 6: Commit**

```bash
git add app_inventario.py app.py templates/inventario_eventos.html templates/inventario_evento.html tests/test_app.py
git commit -m "Inventário: blueprint, telas de eventos e salas do evento, integrante por sessão, menu"
```

---

### Task 7: Tela de leitura da sala (JSON, câmera, campos, fotos, sobras)

**Files:**
- Modify: `app_inventario.py` (substituir `sala_tela`; rotas novas), `templates/_macros.html` (nada), `static/dsgov/css/dsgov.css`
- Create: `templates/inventario_sala.html`, `static/dsgov/vendor/html5-qrcode/html5-qrcode.min.js`, `static/dsgov/vendor/html5-qrcode/LICENSE`
- Test: `tests/test_app.py`

**Interfaces:**
- Consumes: Tasks 2 e 4.
- Produces: `GET /inventario/<id>/sala/<localizacao>`; `POST .../ler` (JSON); `POST /inventario/<id>/leitura/<numero>` (JSON); `POST .../leitura/<numero>/foto` (multipart → JSON `{"foto_url"}`); `POST .../leitura/<numero>/foto/excluir`; `POST .../sala/<localizacao>/sobra` (multipart); `POST /inventario/<id>/sobra/<sobra_id>/excluir`.

- [ ] **Step 1: Vendorizar o leitor**

```bash
mkdir -p static/dsgov/vendor/html5-qrcode
curl -sL -o static/dsgov/vendor/html5-qrcode/html5-qrcode.min.js https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js
curl -sL -o static/dsgov/vendor/html5-qrcode/LICENSE https://raw.githubusercontent.com/mebjas/html5-qrcode/master/LICENSE
head -c 200 static/dsgov/vendor/html5-qrcode/html5-qrcode.min.js   # deve começar com código JS, não HTML de erro
```

- [ ] **Step 2: Testes que falham**

```python
def _abrir(cliente, integrante="Fulano"):
    cliente.post("/inventario/abrir", data={"nome": "Inv", "integrantes": "Fulano\nBeltrana", "escopo": "todas"})
    import db, inventario
    eid = inventario.evento_aberto(db.conectar())["id"]
    if integrante:
        cliente.post(f"/inventario/{eid}/integrante", data={"integrante": integrante})
    return eid


def test_inventario_sala_leitura_json(cliente):
    eid = _abrir(cliente, integrante=None)
    r = cliente.get(f"/inventario/{eid}/sala/01 - SALA CCI")
    assert r.status_code == 200 and b'id="leitura"' in r.data and b"html5-qrcode" in r.data and b"Escolha o integrante" in r.data
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    assert r.status_code == 409 and "integrante" in r.get_json()["erro"].lower()
    cliente.post(f"/inventario/{eid}/integrante", data={"integrante": "Fulano"})
    j = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "001001"}).get_json()
    assert j["situacao"] == "localizado" and j["numero"] == 1001 and j["descricao"] == "CADEIRA" and j["reler"] is False
    j = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1004"}).get_json()
    assert j["situacao"] == "divergente" and j["cadastrado_em"] == "99 - SEM MAPA"
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "99999"})
    assert r.status_code == 404 and r.get_json()["numero"] == 99999
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "abc"})
    assert r.status_code == 404
    j = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"}).get_json()
    assert j["reler"] and j["leitura_anterior"]["integrante"] == "Fulano"
    r = cliente.get(f"/inventario/{eid}/sala/01 - SALA CCI")
    assert b"Localizado" in r.data and b"Divergente" in r.data and b"1004" in r.data
    r = cliente.post(f"/inventario/{eid}/leitura/1001", json={"conservacao": "Ruim", "quem_usa": "Ciclana"})
    assert r.status_code == 200 and r.get_json()["ok"]
    assert cliente.post(f"/inventario/{eid}/leitura/1001", json={"conservacao": "Péssimo"}).status_code == 409
    assert cliente.post(f"/inventario/{eid}/leitura/1002", json={"observacao": "x"}).status_code == 409   # não lido


def test_inventario_fotos_e_sobras(cliente, monkeypatch, tmp_path):
    import io
    from PIL import Image
    import fotos
    eid = _abrir(cliente)
    buf = io.BytesIO(); Image.new("RGB", (30, 20), (1, 2, 3)).save(buf, "PNG"); imagem = buf.getvalue()
    # fotos desativadas: sobra sem foto é aceita; foto de bem recusada
    for v in fotos.VARIAVEIS:
        monkeypatch.delenv(v, raising=False)
    r = cliente.get(f"/inventario/{eid}/sala/01 - SALA CCI")
    assert b"Fotos desativadas" in r.data
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/sobra", data={"descricao": "VENTILADOR", "observacao": "sem plaqueta"}, follow_redirects=True)
    assert b"VENTILADOR" in r.data
    cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    r = cliente.post(f"/inventario/{eid}/leitura/1001/foto", data={"foto": (io.BytesIO(imagem), "a.png")}, content_type="multipart/form-data")
    assert r.status_code == 409 and "desativadas" in r.get_json()["erro"]
    # fotos ativas: cliente falso
    for v in fotos.VARIAVEIS:
        monkeypatch.setenv(v, "x")
    monkeypatch.setenv("R2_PUBLIC_URL", "https://f.exemplo.org")
    enviados = []
    monkeypatch.setattr(fotos, "enviar", lambda nome, dados: enviados.append(nome) or f"https://f.exemplo.org/inventario/{nome}")
    apagados = []
    monkeypatch.setattr(fotos, "apagar", lambda url: apagados.append(url))
    r = cliente.post(f"/inventario/{eid}/leitura/1001/foto", data={"foto": (io.BytesIO(imagem), "a.png")}, content_type="multipart/form-data")
    assert r.status_code == 200 and r.get_json()["foto_url"].startswith("https://f.exemplo.org/inventario/INV") and enviados[-1].startswith(f"INV{eid}_BEM_1001_")
    r = cliente.post(f"/inventario/{eid}/leitura/1001/foto/excluir", follow_redirects=True)
    assert apagados and b"Foto removida" in r.data
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/sobra", data={"descricao": "CADEIRA VELHA", "observacao": "x"}, follow_redirects=True)
    assert "precisa de foto".encode() in r.data
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/sobra", data={"descricao": "CADEIRA VELHA", "observacao": "x", "foto": (io.BytesIO(imagem), "b.jpg")}, content_type="multipart/form-data", follow_redirects=True)
    assert b"CADEIRA VELHA" in r.data and enviados[-1].startswith(f"INV{eid}_SOBRA_")
    # falha no envio: sobra não fica registrada
    monkeypatch.setattr(fotos, "enviar", lambda nome, dados: (_ for _ in ()).throw(RuntimeError("bucket fora")))
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/sobra", data={"descricao": "MESA VELHA", "observacao": "x", "foto": (io.BytesIO(imagem), "c.jpg")}, content_type="multipart/form-data", follow_redirects=True)
    assert b"MESA VELHA" not in r.data and "não registrada".encode() in r.data
    import db, inventario
    sobras = inventario.bens_da_sala(db.conectar(), eid, "01 - SALA CCI")["sobras"]
    assert [s["descricao"] for s in sobras] == ["CADEIRA VELHA", "VENTILADOR"]
    r = cliente.post(f"/inventario/{eid}/sobra/{sobras[0]['id']}/excluir", follow_redirects=True)
    assert b"CADEIRA VELHA" not in r.data and len(apagados) == 2
```

- [ ] **Step 3: Rodar e ver falhar** — `-k "sala_leitura or fotos_e_sobras"`.

- [ ] **Step 4: Rotas** (substituir a `sala_tela` provisória)

```python
def _json_erro(e, status=409):
    return jsonify({"erro": str(e)}), status


def _numero_lido(texto) -> int | None:
    t = "".join(ch for ch in str(texto or "") if ch.isdigit()).lstrip("0")
    return int(t) if t else None


def _bem_json(r):
    b = r["bem"]
    return {"situacao": r["situacao"], "numero": b["numero"], "descricao": b["descricao"], "complemento": b["complemento"],
            "situacao_bem": b["situacao"], "ativo": r["ativo"], "cadastrado_em": r["cadastrado_em"], "reler": r["reler"],
            "leitura_anterior": r["leitura_anterior"], "lido_em": r["lido_em"], "integrante": r["integrante"]}


@inventario_bp.route("/<int:id>/sala/<path:localizacao>")
def sala_tela(id, localizacao):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    if not any(s["localizacao"] == localizacao for s in inventario.salas(conn, id)):
        abort(404)
    d = inventario.bens_da_sala(conn, id, localizacao)
    sala = next(s for s in inventario.salas(conn, id) if s["localizacao"] == localizacao)
    return render_template("inventario_sala.html", e=e, sala=sala, localizacao=localizacao, integrante=session.get("integrante"),
                           conservacao=inventario.CONSERVACAO, fotos_ativas=fotos.configurado(), **d,
                           trilha=_trilha(e, (localizacao, None)))


@inventario_bp.route("/<int:id>/sala/<path:localizacao>/ler", methods=["POST"])
def ler(id, localizacao):
    conn = _conn()
    numero = _numero_lido((request.get_json(silent=True) or {}).get("numero"))
    if numero is None:
        return jsonify({"erro": "Número inválido.", "numero": None}), 404
    try:
        r = inventario.ler(conn, id, localizacao, numero, session.get("integrante") or "")
    except inventario.BemNaoEncontrado as e:
        return jsonify({"erro": str(e), "numero": e.numero}), 404
    except db.ErroDeNegocio as e:
        return _json_erro(e)
    return jsonify(_bem_json(r))


@inventario_bp.route("/<int:id>/leitura/<int:numero>", methods=["POST"])
def atualizar_leitura(id, numero):
    dados = request.get_json(silent=True) or {}
    try:
        inventario.atualizar_leitura(_conn(), id, numero, **{k: v for k, v in dados.items() if k in ("conservacao", "quem_usa", "observacao")})
    except db.ErroDeNegocio as e:
        return _json_erro(e)
    return jsonify({"ok": True})


def _foto_processada():
    """Valida e comprime a foto enviada em request.files['foto']; ErroDeNegocio se faltar ou fotos desativadas."""
    if not fotos.configurado():
        raise db.ErroDeNegocio("Fotos desativadas: bucket não configurado.")
    arquivo = request.files.get("foto")
    if not arquivo or not arquivo.filename:
        raise db.ErroDeNegocio("Envie a foto.")
    return fotos.comprimir(fotos.validar(arquivo))


@inventario_bp.route("/<int:id>/leitura/<int:numero>/foto", methods=["POST"])
def foto_leitura(id, numero):
    conn = _conn()
    try:
        dados = _foto_processada()
        url = fotos.enviar(fotos.nome_bem(id, numero), dados)
        inventario.atualizar_leitura(conn, id, numero, foto_url=url)
    except db.ErroDeNegocio as e:
        return _json_erro(e)
    return jsonify({"foto_url": url})


@inventario_bp.route("/<int:id>/leitura/<int:numero>/foto/excluir", methods=["POST"])
def foto_excluir(id, numero):
    conn = _conn()
    atual = conn.execute("SELECT foto_url, localizacao FROM inventario_leituras WHERE evento_id = ? AND numero = ?", (id, numero)).fetchone()
    if not atual:
        abort(404)
    fotos.apagar(atual["foto_url"])
    inventario.atualizar_leitura(conn, id, numero, foto_url="")
    flash("Foto removida.", "success")
    return redirect(url_for("inventario.sala_tela", id=id, localizacao=request.form.get("volta") or atual["localizacao"]))


@inventario_bp.route("/<int:id>/sala/<path:localizacao>/sobra", methods=["POST"])
def sobra(id, localizacao):
    conn = _conn()
    f = request.form
    integrante = session.get("integrante") or ""
    exigir = fotos.configurado()
    dados = None
    if exigir:
        if not (request.files.get("foto") and request.files["foto"].filename):
            raise db.ErroDeNegocio("A sobra precisa de foto.")
        dados = fotos.comprimir(fotos.validar(request.files["foto"]))
    sid = inventario.registrar_sobra(conn, id, localizacao, f.get("descricao", ""), f.get("complemento", ""),
                                     f.get("observacao", ""), "", integrante, exigir_foto=False)
    if exigir:
        try:
            url = fotos.enviar(fotos.nome_sobra(id, sid), dados)
        except Exception:
            inventario.excluir_sobra(conn, id, sid)
            raise db.ErroDeNegocio("Falha ao enviar a foto; sobra não registrada. Tente de novo.")
        inventario.definir_foto_sobra(conn, sid, url)
    flash("Sobra registrada.", "success")
    return redirect(url_for("inventario.sala_tela", id=id, localizacao=localizacao))


@inventario_bp.route("/<int:id>/sobra/<int:sobra_id>/excluir", methods=["POST"])
def sobra_excluir(id, sobra_id):
    s = inventario.excluir_sobra(_conn(), id, sobra_id)
    fotos.apagar(s["foto_url"])
    flash("Sobra excluída.", "success")
    return redirect(url_for("inventario.sala_tela", id=id, localizacao=s["localizacao"]))


```

Observação: a validação de foto na rota `sobra` acontece **antes** de criar a linha (o sistema antigo criava e depois apagava; aqui só o envio pode falhar depois da criação, e aí a linha é apagada). O teste "não registrada" cobre isso.

- [ ] **Step 5: Template `templates/inventario_sala.html`**

```jinja
{% extends "base.html" %}
{% from "_macros.html" import cabecalho_tabela %}
{% block titulo %}{{ localizacao }}{% endblock %}
{% block conteudo %}
{% set fechado = e.encerrado_em is not none %}
<div class="d-flex align-items-center mb-2">
  <div><h1 class="mb-0">{{ localizacao }}</h1><div class="text-gray-70">{{ e.nome }}{% if sala.ccustos %} · centro {{ sala.ccustos }}{% endif %} ·
    {% if integrante %}lendo como <strong>{{ integrante }}</strong>{% else %}<strong>Escolha o integrante</strong>{% endif %}
    (<a href="{{ url_for('inventario.evento_tela', id=e.id) }}">trocar</a>)</div></div>
  <div class="ml-auto text-right" id="contadores" data-total="{{ sala.total }}">
    <span class="br-tag bg-success text-pure-0"><span id="n-localizados">{{ sala.localizados }}</span></span>
    <span class="br-tag bg-warning"><span id="n-divergentes">{{ sala.divergentes }}</span></span>
    <span class="br-tag bg-gray-20"><span id="n-pendentes">{{ sala.pendentes }}</span></span>
    <div class="text-down-01 text-gray-70">localizados · divergentes · pendentes de {{ sala.total }}</div>
  </div>
</div>
{% if fechado %}
<div class="br-message warning mb-3"><div class="icon"><i class="fas fa-lock fa-lg" aria-hidden="true"></i></div><div class="content"><span class="message-body">Evento encerrado: somente consulta.</span></div></div>
{% endif %}
{% if not fotos_ativas %}
<div class="br-message info mb-3"><div class="icon"><i class="fas fa-camera fa-lg" aria-hidden="true"></i></div><div class="content"><span class="message-body">Fotos desativadas: bucket não configurado. Leituras e sobras funcionam sem foto.</span></div></div>
{% endif %}

<div class="br-card mb-3"><div class="card-content">
  <div class="d-flex align-items-end">
    <div class="br-input flex-fill mr-2">
      <label for="leitura">Ler plaqueta (leitor de código de barras ou câmera)</label>
      <input id="leitura" type="text" inputmode="none" autocomplete="off" enterkeyhint="done" placeholder="Aproxime o leitor…"{% if fechado or not integrante %} disabled{% endif %} autofocus/>
    </div>
    <button class="br-button secondary mr-2" type="button" id="btn-camera"{% if fechado or not integrante %} disabled{% endif %}><i class="fas fa-camera mr-1" aria-hidden="true"></i>Câmera</button>
    <button class="br-button" type="button" id="btn-digitar"{% if fechado or not integrante %} disabled{% endif %}><i class="fas fa-keyboard mr-1" aria-hidden="true"></i>Digitar</button>
  </div>
  <div id="aviso" class="br-message mt-3" hidden><div class="icon"><i class="fas fa-info-circle fa-lg" aria-hidden="true"></i></div><div class="content" role="alert"><span class="message-body" id="aviso-texto"></span>
    <button class="br-button small secondary ml-2" type="button" id="btn-sobra" hidden>Registrar sobra</button></div></div>
  <div id="camera" class="mt-3" hidden><div id="leitor-camera" class="dsgov-leitor"></div><button class="br-button small mt-2" type="button" id="btn-fechar-camera">Fechar câmera</button></div>
</div></div>

{{ cabecalho_tabela('Bens da sala', 'bens') }}
  <thead><tr><th scope="col">Nº</th><th scope="col">Descrição</th><th scope="col">Situação</th><th scope="col">Conservação</th><th scope="col">Quem usa</th><th scope="col">Observação</th><th scope="col">Foto</th></tr></thead>
  <tbody id="tabela-bens">
  {% for b in bens|sort(attribute='situacao_inv') %}
  <tr data-numero="{{ b.numero }}" data-situacao="{{ b.situacao_inv }}">
    <td>{{ b.numero }}</td><td>{{ b.descricao }}{% if b.complemento %} <span class="text-gray-70">{{ b.complemento }}</span>{% endif %}</td>
    <td class="situacao">{% if b.situacao_inv == 'localizado' %}<span class="br-tag bg-success text-pure-0"><span>Localizado</span></span>{% elif b.situacao_inv == 'divergente' %}<span class="br-tag bg-warning"><span>Divergente · lido em {{ b.lido_em_sala }}</span></span>{% else %}<span class="br-tag bg-gray-20"><span>Pendente</span></span>{% endif %}</td>
    <td><select class="campo" data-campo="conservacao"{% if fechado or not b.lido_em %} disabled{% endif %}><option value="">—</option>{% for c in conservacao %}<option{% if b.conservacao == c %} selected{% endif %}>{{ c }}</option>{% endfor %}</select></td>
    <td><input class="campo" data-campo="quem_usa" type="text" value="{{ b.quem_usa or '' }}"{% if fechado or not b.lido_em %} disabled{% endif %}/></td>
    <td><input class="campo" data-campo="observacao" type="text" value="{{ b.observacao or '' }}"{% if fechado or not b.lido_em %} disabled{% endif %}/></td>
    <td class="foto">{% if b.foto_url %}<a href="{{ b.foto_url }}" target="_blank" rel="noopener"><img src="{{ b.foto_url }}" alt="Foto do bem {{ b.numero }}" class="dsgov-miniatura"/></a>
      {% if not fechado %}<form method="post" action="{{ url_for('inventario.foto_excluir', id=e.id, numero=b.numero) }}" class="d-inline"><input type="hidden" name="volta" value="{{ localizacao }}"/><button class="br-button circle small" type="submit" aria-label="Excluir foto"><i class="fas fa-trash" aria-hidden="true"></i></button></form>{% endif %}
      {% elif fotos_ativas and not fechado %}<label class="br-button circle small{% if not b.lido_em %} disabled{% endif %}" aria-label="Foto do bem {{ b.numero }}"><i class="fas fa-camera" aria-hidden="true"></i><input type="file" accept="image/*" capture="environment" class="foto-input" hidden{% if not b.lido_em %} disabled{% endif %}/></label>{% endif %}</td>
  </tr>
  {% endfor %}
  </tbody>
</table></div>

<h2 class="text-up-01 mt-4 mb-2">Trazidos de outras salas ou não ativos</h2>
<div class="br-table"><table><thead><tr><th scope="col">Nº</th><th scope="col">Descrição</th><th scope="col">Cadastrado em</th><th scope="col">Situação</th><th scope="col">Lido em</th></tr></thead>
<tbody id="tabela-trazidos">
{% for t in trazidos %}<tr data-numero="{{ t.numero }}"><td>{{ t.numero }}</td><td>{{ t.descricao }}</td><td>{{ t.localizacao }}</td><td>{{ t.situacao }}</td><td>{{ t.lido_em[8:10] }}/{{ t.lido_em[5:7] }} {{ t.lido_em[11:16] }} · {{ t.integrante }}</td></tr>{% endfor %}
</tbody></table></div>

<h2 class="text-up-01 mt-4 mb-2">Sobras (bens sem cadastro)</h2>
{% if sobras %}
<div class="br-table"><table><thead><tr><th scope="col">Descrição</th><th scope="col">Observação</th><th scope="col">Foto</th><th scope="col">Por</th>{% if not fechado %}<th scope="col" class="dsgov-acoes">Ações</th>{% endif %}</tr></thead><tbody>
{% for s in sobras %}<tr><td>{{ s.descricao }}{% if s.complemento %} <span class="text-gray-70">{{ s.complemento }}</span>{% endif %}</td><td>{{ s.observacao }}</td>
  <td>{% if s.foto_url %}<a href="{{ s.foto_url }}" target="_blank" rel="noopener"><img src="{{ s.foto_url }}" alt="Foto da sobra" class="dsgov-miniatura"/></a>{% else %}—{% endif %}</td><td>{{ s.integrante }}</td>
  {% if not fechado %}<td class="dsgov-acoes"><form method="post" action="{{ url_for('inventario.sobra_excluir', id=e.id, sobra_id=s.id) }}"><button class="br-button circle small" type="submit" aria-label="Excluir sobra"><i class="fas fa-trash" aria-hidden="true"></i></button></form></td>{% endif %}</tr>{% endfor %}
</tbody></table></div>
{% else %}<p class="text-gray-70">Nenhuma sobra nesta sala.</p>{% endif %}
{% if not fechado %}
<form method="post" action="{{ url_for('inventario.sobra', id=e.id, localizacao=localizacao) }}" enctype="multipart/form-data" class="br-card mt-3" id="form-sobra" hidden><div class="card-content">
  <div class="text-weight-semi-bold text-up-01 mb-2">Registrar sobra</div>
  <div class="row">
    <div class="col-md-4 mb-2"><div class="br-input"><label for="s-descricao">Descrição</label><input id="s-descricao" name="descricao" type="text" required/></div></div>
    <div class="col-md-4 mb-2"><div class="br-input"><label for="s-complemento">Complemento</label><input id="s-complemento" name="complemento" type="text"/></div></div>
    <div class="col-md-4 mb-2"><div class="br-input"><label for="s-observacao">Observação (obrigatória)</label><input id="s-observacao" name="observacao" type="text" required/></div></div>
    {% if fotos_ativas %}<div class="col-md-6 mb-2"><div class="br-upload"><label class="upload-label" for="s-foto"><span>Foto (obrigatória)</span></label><input class="upload-input" id="s-foto" name="foto" type="file" accept="image/*" capture="environment" required/><div class="upload-list"></div></div></div>{% endif %}
  </div>
  <button class="br-button primary" type="submit">Registrar sobra</button>
  <button class="br-button ml-2" type="button" id="btn-cancelar-sobra">Cancelar</button>
</div></form>
{% endif %}

<div class="mt-4">
  <a class="br-button" href="{{ url_for('inventario.evento_tela', id=e.id) }}"><i class="fas fa-arrow-left mr-1" aria-hidden="true"></i>Salas</a>
</div>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='dsgov/vendor/html5-qrcode/html5-qrcode.min.js') }}"></script>
<script>
(function () {
  var URL_LER = {{ url_for('inventario.ler', id=e.id, localizacao=localizacao)|tojson }};
  var URL_LEITURA = {{ url_for('inventario.atualizar_leitura', id=e.id, numero=0)|tojson }};   // troca o 0 pelo número
  var SALA = {{ localizacao|tojson }};
  var campo = document.getElementById("leitura"), aviso = document.getElementById("aviso"), avisoTexto = document.getElementById("aviso-texto");
  var btnSobra = document.getElementById("btn-sobra"), formSobra = document.getElementById("form-sobra");
  var ultimoLido = "", ultimoQuando = 0, timerAviso = null;
  if (!campo) return;

  function focar() { if (!campo.disabled) { campo.value = ""; campo.focus(); } }
  function mostrar(tipo, texto, sobra) {
    aviso.className = "br-message mt-3 " + tipo; avisoTexto.textContent = texto; aviso.hidden = false;
    btnSobra.hidden = !sobra; clearTimeout(timerAviso); if (!sobra) timerAviso = setTimeout(function () { aviso.hidden = true; }, 4000);
  }
  function contar() {
    var linhas = document.querySelectorAll("#tabela-bens tr"), loc = 0, div = 0;
    linhas.forEach(function (tr) { var s = tr.dataset.situacao; if (s === "localizado") loc++; else if (s === "divergente") div++; });
    var trazidos = document.querySelectorAll("#tabela-trazidos tr").length;
    document.getElementById("n-localizados").textContent = loc;
    document.getElementById("n-divergentes").textContent = trazidos;
    document.getElementById("n-pendentes").textContent = document.getElementById("contadores").dataset.total - loc;
  }
  function tag(classe, texto) { return '<span class="br-tag ' + classe + '"><span>' + texto + '</span></span>'; }
  function aplicar(j) {
    var tr = document.querySelector('#tabela-bens tr[data-numero="' + j.numero + '"]');
    if (tr) {
      tr.dataset.situacao = j.situacao;
      tr.querySelector(".situacao").innerHTML = j.situacao === "localizado" ? tag("bg-success text-pure-0", "Localizado") : tag("bg-warning", "Divergente · lido em " + SALA);
      tr.querySelectorAll(".campo, .foto-input").forEach(function (el) { el.disabled = false; });
      var rotulo = tr.querySelector("label.disabled"); if (rotulo) rotulo.classList.remove("disabled");
      document.getElementById("tabela-bens").prepend(tr);
    } else {
      var t = document.querySelector('#tabela-trazidos tr[data-numero="' + j.numero + '"]');
      if (t) t.remove();
      var novo = document.createElement("tr"); novo.dataset.numero = j.numero;
      novo.innerHTML = "<td>" + j.numero + "</td><td>" + j.descricao + "</td><td>" + j.cadastrado_em + "</td><td>" + j.situacao_bem + "</td><td>agora · " + j.integrante + "</td>";
      document.getElementById("tabela-trazidos").prepend(novo);
    }
    contar();
  }
  function ler(numero) {
    numero = String(numero || "").trim(); if (!numero) return;
    var agora = Date.now(); if (numero === ultimoLido && agora - ultimoQuando < 3000) return;   // leitura dupla do leitor/câmera
    ultimoLido = numero; ultimoQuando = agora;
    fetch(URL_LER, {method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify({numero: numero})})
      .then(function (r) { return r.json().then(function (j) { return {status: r.status, j: j}; }); })
      .then(function (res) {
        var j = res.j;
        if (res.status === 404) { mostrar("danger", "Bem " + numero + " não está na base.", true); return; }
        if (res.status !== 200) { mostrar("danger", j.erro || "Erro ao registrar a leitura."); return; }
        var texto = j.situacao === "localizado" ? "Bem " + j.numero + " localizado." : "Bem " + j.numero + " cadastrado em " + j.cadastrado_em + "; registrado aqui.";
        if (!j.ativo) texto = "Bem " + j.numero + " está " + j.situacao_bem + "; leitura registrada.";
        if (j.reler && j.leitura_anterior) texto += " Já lido em " + j.leitura_anterior.lido_em.slice(8, 10) + "/" + j.leitura_anterior.lido_em.slice(5, 7) + " por " + j.leitura_anterior.integrante + " em " + j.leitura_anterior.localizacao + ". Atualizado.";
        mostrar(!j.ativo ? "info" : (j.situacao === "localizado" ? "success" : "warning"), texto);
        aplicar(j);
      })
      .catch(function () { mostrar("danger", "Sem conexão; tente de novo."); })
      .finally(focar);
  }
  campo.addEventListener("keydown", function (ev) { if (ev.key === "Enter") { ev.preventDefault(); ler(campo.value); } });
  document.getElementById("btn-digitar").addEventListener("click", function () { campo.inputMode = "numeric"; campo.placeholder = "Digite o número e Enter"; campo.focus(); });
  campo.addEventListener("blur", function () { campo.inputMode = "none"; });
  btnSobra.addEventListener("click", function () { if (formSobra) { formSobra.hidden = false; formSobra.querySelector("#s-descricao").focus(); } });
  var btnCancelar = document.getElementById("btn-cancelar-sobra"); if (btnCancelar) btnCancelar.addEventListener("click", function () { formSobra.hidden = true; focar(); });

  // Câmera (html5-qrcode): fica aberta para leituras seguidas até fechar.
  var leitor = null, cameraDiv = document.getElementById("camera");
  document.getElementById("btn-camera").addEventListener("click", function () {
    cameraDiv.hidden = false;
    if (!leitor) leitor = new Html5QrcodeScanner("leitor-camera", {fps: 10, qrbox: {width: 250, height: 120}}, false);
    leitor.render(function (texto) { ler(texto); }, function () {});
  });
  document.getElementById("btn-fechar-camera").addEventListener("click", function () {
    cameraDiv.hidden = true; if (leitor) leitor.clear().catch(function () {}); leitor = null; focar();
  });

  // Campos por bem: salvam ao mudar.
  document.getElementById("tabela-bens").addEventListener("change", function (ev) {
    var el = ev.target, tr = el.closest("tr"); if (!tr) return;
    if (el.classList.contains("campo")) {
      var corpo = {}; corpo[el.dataset.campo] = el.value;
      fetch(URL_LEITURA.replace(/0$/, tr.dataset.numero), {method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(corpo)})
        .then(function (r) { if (!r.ok) return r.json().then(function (j) { mostrar("danger", j.erro); }); });
    } else if (el.classList.contains("foto-input") && el.files[0]) {
      var fd = new FormData(); fd.append("foto", el.files[0]);
      fetch(URL_LEITURA.replace(/0$/, tr.dataset.numero) + "/foto", {method: "POST", body: fd})
        .then(function (r) { return r.json().then(function (j) { return {ok: r.ok, j: j}; }); })
        .then(function (res) {
          if (!res.ok) { mostrar("danger", res.j.erro); return; }
          tr.querySelector(".foto").innerHTML = '<a href="' + res.j.foto_url + '" target="_blank" rel="noopener"><img src="' + res.j.foto_url + '" alt="Foto" class="dsgov-miniatura"/></a>';
          mostrar("success", "Foto enviada.");
        });
    }
  });
  focar();
})();
</script>
{% endblock %}
```

CSS em `static/dsgov/css/dsgov.css` (ao final):

```css
/* Inventário: miniatura da foto e área do leitor por câmera. */
.dsgov-miniatura { width: 56px; height: 56px; object-fit: cover; border-radius: 4px; }
.dsgov-leitor { max-width: 480px; }
```

- [ ] **Step 6: Rodar e ver passar** — `.venv/bin/pytest -q` → 140.

- [ ] **Step 7: Commit**

```bash
git add app_inventario.py templates/inventario_sala.html static/dsgov/vendor/html5-qrcode static/dsgov/css/dsgov.css tests/test_app.py
git commit -m "Inventário: tela de leitura da sala (leitor/câmera/digitação, JSON, campos, fotos, sobras)"
```

---

### Task 8: Relatório, `.xlsx` e card do painel

**Files:**
- Modify: `app_inventario.py` (substituir `relatorio_tela`; rota `xlsx`), `app.py` (`home` passa `inventario_aberto`), `templates/index.html`
- Create: `templates/inventario_relatorio.html`
- Test: `tests/test_app.py`

- [ ] **Step 1: Testes que falham**

```python
def test_inventario_relatorio_xlsx_e_card_do_painel(cliente):
    assert b"Nenhum invent" in cliente.get("/").data
    eid = _abrir(cliente)
    cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    r = cliente.get("/")
    assert b"Invent\xc3\xa1rio em andamento" in r.data and f"/inventario/{eid}".encode() in r.data
    r = cliente.get(f"/inventario/{eid}/relatorio")
    assert r.status_code == 200 and b"1001" in r.data and b"Localizado" in r.data and b"1002" in r.data
    assert b"1002" not in cliente.get(f"/inventario/{eid}/relatorio?situacao=localizado").data.split(b"<tbody>")[1]
    assert b"1004" not in cliente.get(f"/inventario/{eid}/relatorio?localizacao=01 - SALA CCI").data.split(b"<tbody>")[1]
    r = cliente.get(f"/inventario/{eid}/xlsx?localizacao=01 - SALA CCI")
    assert r.status_code == 200 and r.headers["Content-Disposition"].endswith(".xlsx")
```

- [ ] **Step 2: Rodar e ver falhar**.

- [ ] **Step 3: Rotas** (substituir a provisória)

```python
@inventario_bp.route("/<int:id>/relatorio")
def relatorio_tela(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    loc, sit = request.args.get("localizacao") or None, request.args.get("situacao") or None
    return render_template("inventario_relatorio.html", e=e, linhas=inventario.relatorio(conn, id, loc, sit), localizacao=loc, situacao=sit,
                           salas=[s["localizacao"] for s in inventario.salas(conn, id)], rotulos=inventario.ROTULO_SITUACAO,
                           trilha=_trilha(e, ("Relatório", None)))


@inventario_bp.route("/<int:id>/xlsx")
def xlsx(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    loc = request.args.get("localizacao") or None
    arquivo = inventario.exportar_xlsx(conn, id, io.BytesIO(), loc)
    arquivo.seek(0)
    nome = f"inventario_{id}_{''.join(c if c.isalnum() else '_' for c in (loc or 'todas'))}.xlsx"
    return send_file(arquivo, as_attachment=True, download_name=nome)
```

Em `app.py` `home()`: `import inventario` no topo e passe `inventario_aberto=inventario.evento(obter_conn(), a["id"]) if (a := inventario.evento_aberto(obter_conn())) else None`.

- [ ] **Step 4: Templates**

`templates/inventario_relatorio.html`:

```jinja
{% extends "base.html" %}
{% from "_macros.html" import select, cabecalho_tabela %}
{% block titulo %}Relatório — {{ e.nome }}{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-3"><h1 class="mb-0">Relatório · {{ e.nome }}</h1>
  <a class="br-button primary ml-auto" href="{{ url_for('inventario.xlsx', id=e.id, localizacao=localizacao) }}"><i class="fas fa-file-excel mr-1" aria-hidden="true"></i>Exportar .xlsx</a></div>
<form method="get" class="row mb-3">
  <div class="col-md-4 mb-2">{{ select('localizacao', 'Sala', [('', 'Todas')] + salas, selecionado=localizacao or '', obrigatorio=False) }}</div>
  <div class="col-md-3 mb-2">{{ select('situacao', 'Situação', [('', 'Todas')] + rotulos.items()|list, selecionado=situacao or '', obrigatorio=False) }}</div>
  <div class="col-md-2 mb-2 d-flex align-items-end"><button class="br-button secondary" type="submit">Filtrar</button></div>
</form>
{% set r = e.resumo %}
<p class="text-gray-70">{{ r.lidos }} localizados · {{ r.divergentes }} divergentes · {{ r.pendentes }} pendentes · {{ r.sobras }} sobras · {{ linhas|length }} linha(s) no filtro</p>
{{ cabecalho_tabela('Bens', 'rel') }}
  <thead><tr><th scope="col">Nº</th><th scope="col">Descrição</th><th scope="col">Local sistema</th><th scope="col">Local inventário</th><th scope="col">Situação</th><th scope="col">Conservação</th><th scope="col">Quem usa</th><th scope="col">Observação</th><th scope="col">Integrante</th><th scope="col">Data/hora</th><th scope="col">Foto</th></tr></thead>
  <tbody>{% for x in linhas %}
  <tr><td><a href="{{ url_for('bem', numero=x.numero) }}">{{ x.numero }}</a></td><td>{{ x.descricao }}</td><td>{{ x.local_sistema }}</td><td>{{ x.local_inventario or '—' }}</td>
    <td>{% if x.situacao_inv == 'localizado' %}<span class="br-tag bg-success text-pure-0"><span>Localizado</span></span>{% elif x.situacao_inv == 'divergente' %}<span class="br-tag bg-warning"><span>Divergente</span></span>{% else %}<span class="br-tag bg-gray-20"><span>Não localizado</span></span>{% endif %}</td>
    <td>{{ x.conservacao or '' }}</td><td>{{ x.quem_usa or '' }}</td><td>{{ x.observacao or '' }}</td><td>{{ x.integrante or '' }}</td>
    <td>{% if x.lido_em %}{{ x.lido_em[8:10] }}/{{ x.lido_em[5:7] }}/{{ x.lido_em[:4] }} {{ x.lido_em[11:16] }}{% endif %}</td>
    <td>{% if x.foto_url %}<a href="{{ x.foto_url }}" target="_blank" rel="noopener">ver</a>{% endif %}</td></tr>
  {% endfor %}</tbody>
</table></div>
{% endblock %}
```

Em `templates/index.html`, após o card da última importação (dentro da mesma `row`):

```jinja
  <div class="col-sm-6 col-md-4 col-lg mb-3">
    {% if inventario_aberto %}<a class="br-card h-100 dsgov-kpi" href="{{ url_for('inventario.evento_tela', id=inventario_aberto.id) }}"><div class="card-content">
      <div class="text-up-03 text-weight-bold">{{ inventario_aberto.resumo.pct_bens }}%</div>
      <div class="text-down-01 text-gray-70">Inventário em andamento · {{ inventario_aberto.resumo.divergentes }} divergentes</div></div></a>
    {% else %}<a class="br-card h-100 dsgov-kpi" href="{{ url_for('inventario.eventos_tela') }}"><div class="card-content"><div class="text-up-03 text-weight-bold">—</div><div class="text-down-01 text-gray-70">Nenhum inventário aberto</div></div></a>{% endif %}
  </div>
```

- [ ] **Step 5: Rodar e ver passar** — `.venv/bin/pytest -q` → 141.

- [ ] **Step 6: Commit**

```bash
git add app_inventario.py app.py templates/inventario_relatorio.html templates/index.html tests/test_app.py
git commit -m "Inventário: relatório filtrável, planilha do evento e card no painel"
```

---

### Task 9: README, publicação e push

**Files:** `README.md`

- [ ] **Step 1: README** — na seção "Uso", após o parágrafo do painel/recorte:

```markdown
   **Inventário** (menu próprio): abra um evento (nome, portaria, comissão; todas as salas com bens
   ativos ou uma amostra), escolha o integrante e leia as plaquetas por sala com leitor de código de
   barras, câmera do celular ou digitação. Bem lido na sala cadastrada = localizado; em outra sala =
   divergente (o cadastro do SPW não muda); não lido = pendente; sem cadastro = sobra (com foto).
   Bem baixado lido fica registrado (continua baixado). Conservação, quem usa, observação e foto por bem.
   Relatório e `.xlsx` por evento. O evento fica aberto até ser encerrado; encerrar congela tudo.
   Fotos vão para o bucket R2 configurado em `secrets/.env` (variáveis `R2_*`); sem ele, fotos ficam
   desativadas. A planilha de cadastros ganha abas `inv_*` para exportar/importar inventários inteiros
   (migração de outros sistemas).
```

Na tabela Arquivos: `| `inventario.py`, `fotos.py`, `app_inventario.py` | módulo de inventário (dados, fotos no R2, rotas) |`.

- [ ] **Step 2: Publicar e conferir**

```bash
.venv/bin/pytest -q
git add README.md && git commit -m "README: módulo de inventário"
docker compose up -d --build
sleep 5; for p in /inventario / "/recorte"; do echo "$p $(curl -s -o /dev/null -w '%{http_code}' http://127.0.0.1:12012$p)"; done
docker compose logs --tail=30 | grep -ci traceback
```

Expected: 141 testes; rotas 200; 0 tracebacks. O push é feito pelo controlador ao final (merge na `main`).

---

## Self-review

- **Spec → tasks:** §2.1 esquema (T1); §2.2 situação (T2/T3); §2.3 funções (T1–T3); §2.4 fotos (T4); §3 rotas (T6–T8); §4 telas (T6–T8); §5 xlsx (T3/T8); §6 infra (T4/T7); §7 regras migradas (T2/T4/T7); §8 testes (todas); §9 abas `inv_*` (T5).
- **Nomes:** `inventario.ler/bens_da_sala/atualizar_leitura/registrar_sobra/definir_foto_sobra/excluir_sobra/relatorio/exportar_xlsx/resumo/salas/abrir_evento/encerrar_evento/evento/eventos/evento_aberto`, `fotos.configurado/validar/comprimir/enviar/apagar/nome_bem/nome_sobra/VARIAVEIS`, rotas `inventario.eventos_tela/abrir/evento_tela/encerrar/integrante/sala_tela/ler/atualizar_leitura/foto_leitura/foto_excluir/sobra/sobra_excluir/relatorio_tela/xlsx` — consistentes entre tasks.
- **Placeholders:** nenhum; rotas provisórias da T6 substituídas em T7/T8.
