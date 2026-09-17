# Fase 2 do inventário — plano de implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Completar o módulo de inventário com painel do evento (ECharts), relatório filtrável/ordenável com miniaturas, xlsx com cabeçalho e fotos, marcação em lote na sala e snapshot dos bens no encerramento.

**Architecture:** Só acréscimos. Dados em `inventario.py` (funções `conn`-primeiro, sem Flask), rotas em `app_inventario.py` (blueprint `/inventario`), cards do painel num módulo novo `painel_inventario.py` que só monta opções ECharts com `graficos.py`. Uma tabela nova (`inventario_bens_encerrados`) e um seletor `_fonte_bens()` que faz todas as consultas do evento lerem de `bens` (aberto) ou do snapshot (encerrado).

**Tech Stack:** Python 3, Flask, SQLite (`sqlite3`), openpyxl, Jinja2, DSGov (CSS/JS já no repo), ECharts (já no repo), pytest.

**Spec:** `docs/superpowers/specs/2026-09-17-fase2-inventario-design.md` — leia antes de cada tarefa; o plano argumenta a partir dela.

## Global Constraints

- Branch de trabalho: `fase2-inventario` (criar a partir de `main`). Um commit por tarefa, mensagem em português, assunto ≤ 72 caracteres, rodapé `Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>`.
- `bens` é espelho do SPW e **nunca** é alterada pelo módulo de inventário.
- Toda função de dados recebe `conn` primeiro e não importa Flask (padrão de `db.py`/`inventario.py`).
- Rótulos de UI em português; "Baixar"/"Carregar" em vez de Exportar/Importar em rótulos novos (o botão "Exportar .xlsx" já existe e fica).
- Nada de rede em testes: `fotos.apagar`/`fotos.enviar` só via `monkeypatch`.
- Rodar `python -m pytest -q` (189 testes passando na base) antes de cada commit; nenhum teste antigo pode quebrar sem ajuste consciente descrito na tarefa.
- Não adicionar dependências ao `requirements.txt`.
- Helpers disponíveis em `db.py`: `_um(conn, sql, *params) -> dict|None`, `_todos(conn, sql, *params) -> list[dict]`, `_agora() -> "YYYY-MM-DD HH:MM:SS"`, `_texto(v) -> str`, `_obrigatorio(v, rotulo)`, `_numero(v) -> float|None`, `acrescentar_linha(ws, valores)` (grava strings que começam com `=` como texto), `ErroDeNegocio`, `ImportacaoInvalida`.
- Fixtures de teste: `dados` (conexão com esquema vazio, `tests/conftest.py`), `semear(conn)` (bens 1001, 1002 ATIVO em "01 - SALA CCI"; 1003 BAIXADO em "01 - SALA CCI"; 1004 ATIVO em "99 - SEM MAPA"), `semear_inventario(conn)` em `tests/test_inventario.py` (semear + 2001, 2002 ATIVO em "02 - SALA B" + evento aberto "Inventário 2026" com integrantes Fulano e Beltrana; devolve o id), `cliente` (Flask test client já semeado, `tests/test_app.py`), `_abrir(cliente, integrante="Fulano")` em `tests/test_app.py` (abre evento com todas as salas e escolhe o integrante; devolve o id).

---

## Mapa de arquivos

| arquivo | responsabilidade nesta fase |
|---|---|
| `db.py` | `ESQUEMA` ganha `inventario_bens_encerrados`; `importar_cadastros` trata a 6ª aba como opcional |
| `inventario.py` | snapshot e `_fonte_bens`; `andar`, `ler_lote`, `desfazer_leituras`; filtros/ordem/busca do `relatorio`; `descrever_filtros`, `contar_fotos`; xlsx com cabeçalho e fotos; `painel`; aba `inv_bens_encerrados` |
| `painel_inventario.py` (novo) | cards ECharts do painel do evento |
| `app_inventario.py` | rotas `/painel`, `/lote`; `/relatorio` e `/xlsx` com filtros |
| `templates/inventario_painel.html` (novo) | página do painel |
| `templates/inventario_relatorio.html` | filtros, cabeçalhos ordenáveis, miniatura + modal, "Incluir fotos" |
| `templates/inventario_sala.html` | seleção de linhas + formulário de lote |
| `templates/inventario_evento.html` | botão "Painel" |
| `README.md` | aba `inv_bens_encerrados`; painel, lote, fotos no xlsx |
| `tests/test_inventario.py`, `tests/test_app.py`, `tests/test_painel_inventario.py` (novo) | testes |

---

### Task 0: Branch

- [ ] **Step 1: Criar o branch a partir de `main`**

```bash
cd /opt/web/termos-responsabilidade && git checkout -b fase2-inventario main && python -m pytest -q 2>&1 | tail -1
```
Expected: `189 passed`.

---

### Task 1: Snapshot dos bens no encerramento e fonte dos bens por evento

**Files:**
- Modify: `db.py` (`ESQUEMA`, logo após o `CREATE TABLE IF NOT EXISTS inventario_sobras (...)`)
- Modify: `inventario.py` (`encerrar_evento`, `salas`, `bens_da_sala`, `relatorio`; função nova `_fonte_bens`)
- Test: `tests/test_inventario.py`

**Interfaces:**
- Produces: `_fonte_bens(conn, evento_id: int) -> str` (usado nas tarefas 4 e 6 via `salas`/`relatorio`); constante `_COLS_SNAPSHOT`.

- [ ] **Step 1: Escrever os testes que falham**

Acrescentar ao fim de `tests/test_inventario.py`:

```python
def test_encerrar_grava_snapshot_e_congela_o_evento(dados):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.ler(dados, eid, "01 - SALA CCI", 1003, "Fulano")          # BAIXADO lido: entra no snapshot por ter leitura
    inventario.encerrar_evento(dados, eid)
    snap = {r[0] for r in dados.execute("SELECT numero FROM inventario_bens_encerrados WHERE evento_id = ?", (eid,))}
    assert snap == {1001, 1002, 1003, 1004, 2001, 2002}
    inventario.encerrar_evento(dados, eid)                                 # idempotente: não regrava
    assert dados.execute("SELECT count(*) FROM inventario_bens_encerrados").fetchone()[0] == 6
    # o export do SPW do ano seguinte muda `bens`; o evento encerrado não muda
    dados.execute("UPDATE bens SET localizacao = '02 - SALA B', descricao = 'CADEIRA NOVA' WHERE numero = 1001")
    dados.execute("DELETE FROM bens WHERE numero = 2002")
    dados.execute("INSERT INTO bens VALUES (3001,'ATIVO','TV','LG','EQUIPAMENTOS','01 - SALA CCI','01/01/2027',1,1)")
    dados.commit()
    assert [(s["localizacao"], s["total"], s["localizados"]) for s in inventario.salas(dados, eid)] == \
        [("01 - SALA CCI", 2, 1), ("02 - SALA B", 2, 0), ("99 - SEM MAPA", 1, 0)]
    r = {x["numero"]: x for x in inventario.relatorio(dados, eid)}
    assert set(r) == {1001, 1002, 1003, 1004, 2001, 2002}
    assert r[1001]["descricao"] == "CADEIRA" and r[1001]["situacao_inv"] == "localizado" and r[1003]["situacao_bem"] == "BAIXADO"
    assert inventario.resumo(dados, eid)["bens"] == 5
    assert [b["numero"] for b in inventario.bens_da_sala(dados, eid, "01 - SALA CCI")["bens"]] == [1001, 1002]


def test_evento_encerrado_sem_snapshot_le_bens(dados):
    eid = semear_inventario(dados)
    dados.execute("UPDATE inventario_eventos SET encerrado_em = '2026-01-01 00:00:00' WHERE id = ?", (eid,))
    dados.commit()
    assert inventario.resumo(dados, eid)["bens"] == 5 and inventario._fonte_bens(dados, eid) == "bens"
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `python -m pytest tests/test_inventario.py -k "snapshot or sem_snapshot" -q`
Expected: FAIL (`no such table: inventario_bens_encerrados` / `AttributeError: _fonte_bens`).

- [ ] **Step 3: Esquema em `db.py`**

Dentro da string `ESQUEMA`, após o bloco de `inventario_sobras`:

```sql
CREATE TABLE IF NOT EXISTS inventario_bens_encerrados (
  evento_id     INTEGER NOT NULL REFERENCES inventario_eventos(id) ON DELETE CASCADE,
  numero        INTEGER NOT NULL,
  situacao      TEXT,
  descricao     TEXT,
  complemento   TEXT,
  classificacao TEXT,
  localizacao   TEXT,
  PRIMARY KEY (evento_id, numero)
);
```

- [ ] **Step 4: `_fonte_bens` e `encerrar_evento` em `inventario.py`**

Logo após `_evento_aberto_ou_erro`:

```python
_COLS_SNAPSHOT = "numero, situacao, descricao, complemento, classificacao, localizacao"


def _fonte_bens(conn, evento_id: int) -> str:
    """De onde vêm os bens do evento: `bens` (aberto, ou encerrado antes de existir snapshot) ou a subconsulta
    do snapshot gravado no encerramento. evento_id é int e vai inline no SQL (não há injeção)."""
    e = _um(conn, "SELECT encerrado_em FROM inventario_eventos WHERE id = ?", evento_id)
    if not e or not e["encerrado_em"] or not conn.execute(
            "SELECT 1 FROM inventario_bens_encerrados WHERE evento_id = ? LIMIT 1", (evento_id,)).fetchone():
        return "bens"
    return f"(SELECT {_COLS_SNAPSHOT} FROM inventario_bens_encerrados WHERE evento_id = {int(evento_id)})"
```

Substituir `encerrar_evento`:

```python
def encerrar_evento(conn, id: int) -> None:
    """Grava encerrado_em e congela os bens do evento (ativos das salas do escopo + todo bem lido) em
    inventario_bens_encerrados, para o relatório não mudar quando o export do SPW seguinte for carregado."""
    e = _um(conn, "SELECT * FROM inventario_eventos WHERE id = ?", id)
    if not e:
        raise ErroDeNegocio("Evento de inventário não encontrado.")
    if e["encerrado_em"]:
        return
    conn.execute("UPDATE inventario_eventos SET encerrado_em = ? WHERE id = ?", (_agora(), id))
    conn.execute(f"""INSERT OR IGNORE INTO inventario_bens_encerrados (evento_id, {_COLS_SNAPSHOT})
        SELECT ?, {_COLS_SNAPSHOT} FROM bens
        WHERE (situacao = 'ATIVO' AND localizacao IN (SELECT localizacao FROM inventario_salas WHERE evento_id = ?))
           OR numero IN (SELECT numero FROM inventario_leituras WHERE evento_id = ?)""", (id, id, id))
    conn.commit()
```

- [ ] **Step 5: Trocar `bens` pela fonte nas consultas**

Em `salas`: no início, `B = _fonte_bens(conn, evento_id)`; no f-string do SQL, trocar `FROM bens b` por `FROM {B} b` e os dois `JOIN bens b` por `JOIN {B} b`.

Em `bens_da_sala`: idem — `B = _fonte_bens(conn, evento_id)`, `FROM {B} b` no primeiro SELECT e `JOIN {B} b` no de trazidos.

Em `relatorio`: após a checagem do evento, `B = _fonte_bens(conn, evento_id)`; trocar os dois `JOIN bens b` por `JOIN {B} b`.

(`resumo` usa `salas`, então já herda. `ler` continua com `db.buscar_bem`: só roda em evento aberto.)

- [ ] **Step 6: Rodar tudo**

Run: `python -m pytest -q`
Expected: `191 passed`.

- [ ] **Step 7: Commit**

```bash
git add db.py inventario.py tests/test_inventario.py
git commit -m "Inventário: snapshot dos bens no encerramento; consultas leem do snapshot em evento encerrado" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 2: Aba `inv_bens_encerrados` na planilha de cadastros

**Files:**
- Modify: `inventario.py` (`ABAS`, `validar_abas`, `substituir_tabelas`)
- Modify: `db.py` (`importar_cadastros`: lista `faltam` e `inv_linhas`)
- Modify: `README.md` (parágrafo das abas `inv_*`, linhas ~32 e ~52)
- Test: `tests/test_inventario.py`

**Interfaces:**
- Consumes: tabela `inventario_bens_encerrados` (Task 1).
- Produces: `ABAS["inv_bens_encerrados"]`; `substituir_tabelas` aceita `linhas[aba] is None` = "não tocar".

- [ ] **Step 1: Ajustar o teste antigo e escrever o novo**

Em `test_planilha_de_cadastros_exporta_e_importa_abas_de_inventario`, a lista de `sheetnames` ganha `"inv_bens_encerrados"` no fim:

```python
    assert wb.sheetnames == ["responsaveis", "localizacoes", "pessoas", "atribuicoes", "inv_eventos", "inv_integrantes", "inv_salas", "inv_leituras", "inv_sobras", "inv_bens_encerrados"]
```

Acrescentar ao fim do arquivo:

```python
def test_aba_inv_bens_encerrados_exporta_importa_e_valida(dados, tmp_path):
    from openpyxl import load_workbook
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.encerrar_evento(dados, eid)
    wb = load_workbook(db.exportar_cadastros(dados, tmp_path / "c.xlsx"))
    assert wb.sheetnames[-1] == "inv_bens_encerrados"
    linhas = list(wb["inv_bens_encerrados"].iter_rows(values_only=True))
    assert linhas[0] == tuple(inventario.ABAS["inv_bens_encerrados"])
    assert linhas[1] == (eid, 1001, "ATIVO", "CADEIRA", "GIRATÓRIA", "MÓVEIS", "01 - SALA CCI") and len(linhas) == 6
    dados.execute("DELETE FROM bens WHERE numero = 1002")          # o SPW mudou; o snapshot importado preserva 1002
    dados.commit()
    with open(tmp_path / "c.xlsx", "rb") as f:
        r = db.importar_cadastros(dados, f)
    assert r["inv_bens_encerrados"] == 5
    assert {x["numero"] for x in inventario.relatorio(dados, eid)} == {1001, 1002, 1004, 2001, 2002}
    # aba ausente (planilha exportada por versão anterior, com 5 abas inv_*): snapshot não é tocado
    wb.remove(wb["inv_bens_encerrados"])
    wb.save(tmp_path / "cinco.xlsx")
    with open(tmp_path / "cinco.xlsx", "rb") as f:
        r = db.importar_cadastros(dados, f)
    assert "inv_bens_encerrados" not in r and dados.execute("SELECT count(*) FROM inventario_bens_encerrados").fetchone()[0] == 5
    # validações: evento aberto, número inválido, repetido
    wb = load_workbook(tmp_path / "c.xlsx")
    wb["inv_eventos"].cell(row=2, column=5, value=None)
    wb["inv_bens_encerrados"].append([eid, "abc", "", "", "", "", ""])
    wb["inv_bens_encerrados"].append([eid, 1001, "", "", "", "", ""])
    wb.save(tmp_path / "ruim.xlsx")
    with open(tmp_path / "ruim.xlsx", "rb") as f, pytest.raises(db.ImportacaoInvalida) as ex:
        db.importar_cadastros(dados, f)
    msg = str(ex.value)
    assert "não está encerrado" in msg and "número inválido" in msg and "repetido" in msg
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `python -m pytest tests/test_inventario.py -k "abas_de_inventario or inv_bens" -q`
Expected: 2 FAIL (sheetnames sem a aba; `KeyError: 'inv_bens_encerrados'`).

- [ ] **Step 3: `ABAS` e `validar_abas` em `inventario.py`**

Acrescentar a `ABAS` (última entrada):

```python
    "inv_bens_encerrados": ["evento_id", "numero", "situacao", "descricao", "complemento", "classificacao", "localizacao"],
```

Em `validar_abas`, na linha `ids, abertos = set(), 0` acrescentar `encerrados = set()`; dentro do laço de `inv_eventos`, logo após `ids.add(eid)`:

```python
        if encerrado:
            encerrados.add(eid)
```

Antes do `return linhas, problemas`, o bloco da aba nova:

```python
    vistos = set()
    for r in brutos["inv_bens_encerrados"]:
        rot = f"inv_bens_encerrados linha {r['_linha']}"
        eid = evento_ok(r, rot)
        if eid is None:
            continue
        if eid not in encerrados:
            problemas.append(f"{rot}: evento {eid} não está encerrado (o snapshot é só de eventos encerrados)"); continue
        num = db._numero(r["numero"])
        if num is None or num != int(num):
            problemas.append(f"{rot}: número inválido ({_texto(r['numero']) or '(vazio)'})"); continue
        num = int(num)
        if (eid, num) in vistos:
            problemas.append(f"{rot}: bem {num} repetido no evento {eid}"); continue
        vistos.add((eid, num))
        linhas["inv_bens_encerrados"].append((eid, num, *(_texto(r[c]) or None for c in ("situacao", "descricao", "complemento", "classificacao", "localizacao"))))
```

- [ ] **Step 4: `substituir_tabelas` aceita "não tocar"**

`db.conectar()` liga `PRAGMA foreign_keys = ON`, então apagar `inventario_eventos` apaga o snapshot em cascata. Quando a
aba está ausente, o snapshot é guardado em memória antes e regravado depois dos eventos. Substituir a função:

```python
def substituir_tabelas(conn, linhas: dict) -> None:
    """Dentro da transação de db.importar_cadastros: apaga e regrava as tabelas (ids de evento preservados).
    linhas["inv_bens_encerrados"] is None = aba ausente na planilha → o snapshot atual é mantido (guardado antes do
    DELETE em cascata e regravado depois)."""
    snapshot = linhas.get("inv_bens_encerrados")
    if snapshot is None:
        snapshot = conn.execute("SELECT evento_id, numero, situacao, descricao, complemento, classificacao, localizacao FROM inventario_bens_encerrados").fetchall()
    for aba in reversed(list(ABAS)):
        conn.execute(f"DELETE FROM {_TABELA[aba]}")
    conn.executemany("INSERT INTO inventario_eventos (id, nome, descricao, aberto_em, encerrado_em) VALUES (?,?,?,?,?)", linhas["inv_eventos"])
    conn.executemany("INSERT INTO inventario_integrantes VALUES (?,?)", linhas["inv_integrantes"])
    conn.executemany("INSERT INTO inventario_salas VALUES (?,?)", linhas["inv_salas"])
    conn.executemany("INSERT INTO inventario_leituras (evento_id, numero, localizacao, lido_em, integrante, conservacao, quem_usa, observacao, foto_url) VALUES (?,?,?,?,?,?,?,?,?)", linhas["inv_leituras"])
    conn.executemany("INSERT INTO inventario_sobras (evento_id, localizacao, descricao, complemento, observacao, foto_url, integrante, criado_em) VALUES (?,?,?,?,?,?,?,?)", linhas["inv_sobras"])
    conn.executemany("INSERT OR IGNORE INTO inventario_bens_encerrados VALUES (?,?,?,?,?,?,?)", snapshot)
```
(`OR IGNORE` cobre o snapshot guardado de um evento que a planilha nova não traz mais: a FK falharia; com `foreign_keys`
ligado o INSERT de evento inexistente dá `IntegrityError`, não é ignorado — por isso filtrar antes:
`ids = {e[0] for e in linhas["inv_eventos"]}; snapshot = [r for r in snapshot if r[0] in ids]`, logo após montar `snapshot`.)

- [ ] **Step 5: `db.importar_cadastros` — aba opcional**

Em `db.py`, na linha `faltam = [aba for aba, v in inv_brutos.items() if v is None]` trocar por:

```python
    faltam = [aba for aba, v in inv_brutos.items() if v is None and aba != "inv_bens_encerrados"]
```

E onde `inv_linhas, inv_problemas = inventario.validar_abas(...)` é chamado, logo depois:

```python
        if inv_brutos["inv_bens_encerrados"] is None:
            inv_linhas["inv_bens_encerrados"] = None          # aba ausente: tabela mantida
```

E no `resultado.update({aba: len(l) for aba, l in inv_linhas.items()})` trocar por `{aba: len(l) for aba, l in inv_linhas.items() if l is not None}`.

- [ ] **Step 6: README**

No parágrafo que fala das abas `inv_*` (linhas ~52–54), acrescentar uma frase: "A 6ª aba, `inv_bens_encerrados`, é o retrato dos bens de cada evento encerrado (gravado no encerramento); é opcional na importação — ausente, a tabela é mantida."

- [ ] **Step 7: Rodar tudo**

Run: `python -m pytest -q`
Expected: `192 passed`.

- [ ] **Step 8: Commit**

```bash
git add inventario.py db.py README.md tests/test_inventario.py
git commit -m "Planilha de cadastros: aba inv_bens_encerrados (snapshot), opcional na importação" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 3: `andar`, `ler_lote` e `desfazer_leituras`

**Files:**
- Modify: `inventario.py` (após `atualizar_leitura`)
- Test: `tests/test_inventario.py`

**Interfaces:**
- Produces: `ANDAR_SEM = "Sem andar"`; `andar(localizacao: str) -> str`; `ler_lote(conn, evento_id, localizacao, numeros: list[int], integrante: str) -> {"lidos": int, "nao_encontrados": list[int]}`; `desfazer_leituras(conn, evento_id, numeros: list[int]) -> list[str]` (URLs de fotos apagadas).

- [ ] **Step 1: Testes**

```python
def test_andar():
    assert inventario.andar("07 - COAD - SALA DE REUNIÃO") == "07"
    assert inventario.andar("02 - CCOM") == "02"
    assert inventario.andar("TERMOS INDIVIDUAIS") == inventario.ANDAR_SEM
    assert inventario.andar("") == inventario.ANDAR_SEM
    assert inventario.andar(" - X") == inventario.ANDAR_SEM


def test_ler_lote_e_desfazer_leituras(dados):
    eid = semear_inventario(dados)
    r = inventario.ler_lote(dados, eid, "01 - SALA CCI", [1001, 2001, 99999], "Fulano")
    assert r == {"lidos": 2, "nao_encontrados": [99999]}
    assert {b["numero"]: b["situacao_inv"] for b in inventario.bens_da_sala(dados, eid, "01 - SALA CCI")["bens"]} == {1001: "localizado", 1002: "pendente"}
    assert inventario.resumo(dados, eid)["divergentes"] == 1
    with pytest.raises(db.ErroDeNegocio):
        inventario.ler_lote(dados, eid, "01 - SALA CCI", [1002], "Ninguém")
    inventario.atualizar_leitura(dados, eid, 1001, foto_url="http://x/1001.webp")
    assert inventario.desfazer_leituras(dados, eid, [1001, 2001, 1004]) == ["http://x/1001.webp"]   # 1004 sem leitura: ignorado
    assert inventario.resumo(dados, eid)["lidos"] == 0 and inventario.resumo(dados, eid)["divergentes"] == 0
    assert inventario.desfazer_leituras(dados, eid, []) == []
    inventario.encerrar_evento(dados, eid)
    with pytest.raises(db.ErroDeNegocio):
        inventario.desfazer_leituras(dados, eid, [1002])
    with pytest.raises(db.ErroDeNegocio):
        inventario.ler_lote(dados, eid, "01 - SALA CCI", [1002], "Fulano")
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `python -m pytest tests/test_inventario.py -k "andar or lote" -q` → FAIL (`AttributeError`).

- [ ] **Step 3: Implementar**

Após `ROTULO_SITUACAO` no topo: `ANDAR_SEM = "Sem andar"`. Após `atualizar_leitura`:

```python
def andar(localizacao: str) -> str:
    """"07 - COAD - SALA DE REUNIÃO" → "07" (texto antes do primeiro " - "); sem separador → ANDAR_SEM."""
    cabeca, sep, _ = (localizacao or "").partition(" - ")
    return cabeca.strip() if sep and cabeca.strip() else ANDAR_SEM


def ler_lote(conn, evento_id: int, localizacao: str, numeros: list, integrante: str) -> dict:
    """Leitura sem plaqueta de vários bens de uma vez: mesma regra de `ler` para cada número
    (divergente e não ativo são aceitos); bem inexistente é pulado e devolvido em nao_encontrados."""
    lidos, nao_encontrados = 0, []
    for numero in numeros:
        try:
            ler(conn, evento_id, localizacao, int(numero), integrante)
            lidos += 1
        except BemNaoEncontrado:
            nao_encontrados.append(int(numero))
    return {"lidos": lidos, "nao_encontrados": nao_encontrados}


def desfazer_leituras(conn, evento_id: int, numeros: list) -> list:
    """Volta os bens a "não localizado" neste evento (o "alternar status" do sistema antigo): apaga as
    leituras. Devolve as URLs das fotos que existiam, para a rota apagar no bucket."""
    _evento_aberto_ou_erro(conn, evento_id)
    numeros = [int(n) for n in numeros]
    if not numeros:
        return []
    marcas = ",".join("?" * len(numeros))
    urls = [r[0] for r in conn.execute(
        f"SELECT foto_url FROM inventario_leituras WHERE evento_id = ? AND numero IN ({marcas}) AND foto_url IS NOT NULL AND foto_url <> ''",
        (evento_id, *numeros))]
    conn.execute(f"DELETE FROM inventario_leituras WHERE evento_id = ? AND numero IN ({marcas})", (evento_id, *numeros))
    conn.commit()
    return urls
```

- [ ] **Step 4: Rodar tudo** → `194 passed`.

- [ ] **Step 5: Commit**

```bash
git add inventario.py tests/test_inventario.py
git commit -m "Inventário: andar pelo nome da sala, leitura em lote e desfazer leituras" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 4: Relatório com filtros, busca sem acento e ordenação

**Files:**
- Modify: `inventario.py` (`relatorio` e constantes; funções novas `descrever_filtros`, `contar_fotos`)
- Test: `tests/test_inventario.py`

**Interfaces:**
- Produces: `FILTROS_RELATORIO`, `CONSERVACAO_VAZIA = "-"`, `COLUNAS_ORDEM`, `relatorio(conn, evento_id, localizacao=None, situacao=None, integrante=None, conservacao=None, foto=None, busca=None, ordem=None, dir=None)`, `descrever_filtros(f: dict) -> str`, `contar_fotos(linhas) -> int`.

- [ ] **Step 1: Teste**

```python
def test_relatorio_filtros_busca_e_ordem(dados):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.ler(dados, eid, "02 - SALA B", 2001, "Beltrana")
    inventario.atualizar_leitura(dados, eid, 1001, conservacao="Ruim", quem_usa="José", foto_url="https://x/1001.webp")
    inventario.atualizar_leitura(dados, eid, 2001, observacao="tela quebrada")
    num = lambda **f: [x["numero"] for x in inventario.relatorio(dados, eid, **f)]
    assert num() == [1001, 1002, 2001, 2002, 1004]
    assert num(integrante="Fulano") == [1001] and num(integrante="Beltrana") == [2001]
    assert num(conservacao="Ruim") == [1001] and num(conservacao="-") == [1002, 2001, 2002, 1004]
    assert num(foto="com") == [1001] and num(foto="sem") == [1002, 2001, 2002, 1004]
    assert num(busca="jose") == [1001] and num(busca="QUEBRADA tela") == [2001]
    assert num(busca="cadeira giratoria") == [1001] and num(busca="20") == [2001, 2002] and num(busca="   ") == num()
    assert num(ordem="numero", dir="desc") == [2002, 2001, 1004, 1002, 1001]
    assert num(ordem="descricao") == [1004, 1001, 2002, 2001, 1002]            # ARMÁRIO, CADEIRA, IMPRESSORA, MONITOR, NOTEBOOK
    assert num(ordem="integrante") == [2001, 1001, 1002, 2002, 1004]           # vazios sempre por último
    assert num(ordem="integrante", dir="desc") == [1001, 2001, 1002, 2002, 1004]
    assert num(ordem="inexistente") == num()
    assert num(localizacao="02 - SALA B", situacao="pendente") == [2002]
    assert inventario.contar_fotos(inventario.relatorio(dados, eid)) == 1
    assert inventario.descrever_filtros({}) == "Todas as salas"
    assert inventario.descrever_filtros({"localizacao": "01 - SALA CCI", "situacao": "divergente", "integrante": "Fulano",
                                         "conservacao": "-", "foto": "com", "busca": " cadeira ", "ordem": "numero"}) == \
        'Sala 01 - SALA CCI · Situação Divergente · Integrante Fulano · Conservação Não informada · Com foto · Busca "cadeira"'
```

- [ ] **Step 2: Rodar e ver falhar** → `TypeError: relatorio() got an unexpected keyword argument 'integrante'`.

- [ ] **Step 3: Implementar**

Constantes junto de `COLUNAS_XLSX`:

```python
FILTROS_RELATORIO = ("localizacao", "situacao", "integrante", "conservacao", "foto", "busca", "ordem", "dir")
CONSERVACAO_VAZIA = "-"    # valor do filtro/chave para "Não informada"
COLUNAS_ORDEM = ("numero", "descricao", "local_sistema", "local_inventario", "situacao_inv", "conservacao", "quem_usa", "integrante", "lido_em")
_CAMPOS_BUSCA = ("numero", "descricao", "complemento", "quem_usa", "observacao", "local_sistema", "local_inventario")
ROTULO_FOTO = {"com": "Com foto", "sem": "Sem foto"}


def _normalizar(texto) -> str:
    """Sem acento e sem caixa, como cadastro_busca em db.py."""
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFD", str(texto if texto is not None else "").casefold()) if not unicodedata.combining(c))


def _tem_foto(x: dict) -> bool:
    return str(x.get("foto_url") or "").startswith(("http://", "https://"))


def contar_fotos(linhas) -> int:
    return sum(1 for x in linhas if _tem_foto(x))


def descrever_filtros(f: dict) -> str:
    """Frase dos filtros ativos, para a tela e o cabeçalho do xlsx."""
    partes = [f"Sala {f['localizacao']}" if f.get("localizacao") else "Todas as salas"]
    if f.get("situacao") in ROTULO_SITUACAO:
        partes.append(f"Situação {ROTULO_SITUACAO[f['situacao']]}")
    if f.get("integrante"):
        partes.append(f"Integrante {f['integrante']}")
    if f.get("conservacao"):
        partes.append("Conservação " + ("Não informada" if f["conservacao"] == CONSERVACAO_VAZIA else f["conservacao"]))
    if f.get("foto") in ROTULO_FOTO:
        partes.append(ROTULO_FOTO[f["foto"]])
    if (f.get("busca") or "").strip():
        partes.append(f'Busca "{f["busca"].strip()}"')
    return " · ".join(partes)
```

Nova assinatura e cauda de `relatorio` (o SQL do meio fica como está, já com `{B}` da Task 1):

```python
def relatorio(conn, evento_id: int, localizacao=None, situacao=None, integrante=None, conservacao=None, foto=None,
              busca=None, ordem=None, dir=None) -> list[dict]:
    """Uma linha por bem ativo das salas do escopo (ou da sala pedida), mais os lidos nela vindos de fora do
    escopo (ou de bens que deixaram de estar ATIVO). Filtros em Python sobre o resultado (≤ alguns milhares de
    linhas): situacao localizado|divergente|pendente; integrante; conservacao (valor ou "-" = não informada);
    foto com|sem; busca sem acento (todas as palavras, em qualquer campo de _CAMPOS_BUSCA); ordem/dir."""
    ...  # checagem do evento, B, SQL — inalterados
    if situacao:
        linhas = [x for x in linhas if x["situacao_inv"] == situacao]
    if integrante:
        linhas = [x for x in linhas if x["integrante"] == integrante]
    if conservacao:
        linhas = [x for x in linhas if (x["conservacao"] or CONSERVACAO_VAZIA) == conservacao]
    if foto in ROTULO_FOTO:
        linhas = [x for x in linhas if _tem_foto(x) == (foto == "com")]
    palavras = _normalizar(busca).split()
    if palavras:
        linhas = [x for x in linhas if all(any(p in _normalizar(x[c]) for c in _CAMPOS_BUSCA) for p in palavras)]
    if ordem in COLUNAS_ORDEM:
        vazios = [x for x in linhas if x[ordem] in (None, "")]
        cheios = [x for x in linhas if x[ordem] not in (None, "")]
        cheios.sort(key=lambda x: x[ordem] if isinstance(x[ordem], (int, float)) else _normalizar(x[ordem]), reverse=(dir == "desc"))
        linhas = cheios + vazios
    return linhas
```

- [ ] **Step 4: Rodar tudo** → `195 passed` (o teste antigo `test_relatorio_e_xlsx` continua válido: `localizacao=`/`situacao=` nomeados).

- [ ] **Step 5: Commit**

```bash
git add inventario.py tests/test_inventario.py
git commit -m "Relatório do inventário: filtros por integrante, conservação e foto, busca sem acento, ordenação" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 5: xlsx com cabeçalho de 4 linhas, filtros e fotos

**Files:**
- Modify: `inventario.py` (`exportar_xlsx`; função nova `_celula_foto`)
- Test: `tests/test_inventario.py`

**Interfaces:**
- Consumes: `relatorio(**filtros)`, `descrever_filtros` (Task 4).
- Produces: `exportar_xlsx(conn, evento_id, destino, localizacao=None, fotos=False, **filtros)`.

- [ ] **Step 1: Ajustar o teste antigo e escrever o novo**

Em `test_relatorio_e_xlsx`, o cabeçalho desce 3 linhas:

```python
    assert linhas[0][0] == "Inventário 2026" and linhas[4] == tuple(inventario.COLUNAS_XLSX)
    assert linhas[5][:3] == (1001, "CADEIRA", "GIRATÓRIA") and linhas[5][6] == "Localizado" and linhas[5][7] == "Bom"
    ...
    assert so_cci.max_row == 5 + 4
```

(`sobras[1][0] == "Sala"` e `sobras[2]...` ficam: a aba Sobras mantém título + cabeçalho.)

Novo teste:

```python
def test_xlsx_cabecalho_filtros_e_fotos(dados, tmp_path):
    from openpyxl import load_workbook
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.atualizar_leitura(dados, eid, 1001, foto_url="https://x/1001.webp")
    inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "VENTILADOR", None, "sem plaqueta", "https://x/s.webp", "Fulano")
    wb = load_workbook(inventario.exportar_xlsx(dados, eid, tmp_path / "a.xlsx", situacao="localizado", fotos=True))
    ws = wb["Bens"]
    linhas = list(ws.iter_rows(values_only=True))
    assert linhas[0][0] == "Inventário 2026" and linhas[1][0].startswith("Gerado em ")
    assert linhas[2][0] == "Todas as salas · Situação Localizado" and linhas[3][0] == "Total de bens: 1"
    assert linhas[4] == tuple(inventario.COLUNAS_XLSX) and linhas[5][0] == 1001 and len(linhas) == 6
    assert linhas[5][12] == '=_xlfn.IMAGE("https://x/1001.webp")' and ws.row_dimensions[6].height == 60
    assert wb["Sobras"].cell(row=3, column=7).value == '=_xlfn.IMAGE("https://x/s.webp")'
    ws = load_workbook(inventario.exportar_xlsx(dados, eid, tmp_path / "b.xlsx", fotos=True))["Bens"]
    assert ws.cell(row=6, column=13).value.startswith("=_xlfn") and ws.cell(row=7, column=13).value == "-"   # 1002 sem foto
    ws = load_workbook(inventario.exportar_xlsx(dados, eid, tmp_path / "c.xlsx"))["Bens"]
    assert ws.cell(row=6, column=13).value == "https://x/1001.webp" and ws.row_dimensions[6].height is None
```

- [ ] **Step 2: Rodar e ver falhar** → `python -m pytest tests/test_inventario.py -k xlsx -q`: FAIL.

- [ ] **Step 3: Implementar**

```python
def _celula_foto(ws, linha: int, coluna: int, url, fotos: bool) -> None:
    """fotos=True: fórmula IMAGE (o Excel pt-BR mostra =IMAGEM; o nome localizado dá #NOME?) e linha alta;
    sem URL http(s) escreve "-". fotos=False: fica a URL como texto (já gravada por acrescentar_linha)."""
    if not fotos:
        return
    if url and str(url).startswith(("http://", "https://")):
        ws.cell(row=linha, column=coluna).value = f'=_xlfn.IMAGE("{url}")'
        ws.row_dimensions[linha].height = 60
    else:
        ws.cell(row=linha, column=coluna).value = "-"


def exportar_xlsx(conn, evento_id: int, destino, localizacao: str | None = None, fotos: bool = False, **filtros):
    from openpyxl import Workbook
    e = _um(conn, "SELECT nome FROM inventario_eventos WHERE id = ?", evento_id)
    if not e:
        raise ErroDeNegocio("Evento de inventário não encontrado.")
    filtros = {"localizacao": localizacao, **filtros}
    linhas = relatorio(conn, evento_id, **filtros)
    wb = Workbook()
    ws = wb.active
    ws.title = "Bens"
    acrescentar_linha(ws, [e["nome"]])
    acrescentar_linha(ws, [f"Gerado em {_data_br(_agora())}"])
    acrescentar_linha(ws, [descrever_filtros(filtros)])
    acrescentar_linha(ws, [f"Total de bens: {len(linhas)}"])
    ws.append(COLUNAS_XLSX)
    col_foto = COLUNAS_XLSX.index("Foto") + 1
    for x in linhas:
        acrescentar_linha(ws, [x["numero"], x["descricao"], x["complemento"], x["classificacao"], x["local_sistema"], x["local_inventario"],
                                ROTULO_SITUACAO[x["situacao_inv"]], x["conservacao"], x["quem_usa"], x["observacao"], x["integrante"],
                                _data_br(x["lido_em"]), x["foto_url"], x["situacao_bem"]])
        _celula_foto(ws, ws.max_row, col_foto, x["foto_url"], fotos)
    ws2 = wb.create_sheet("Sobras")
    acrescentar_linha(ws2, [f"{e['nome']} — sobras (bens sem cadastro)"])
    ws2.append(COLUNAS_SOBRAS)
    col_foto = COLUNAS_SOBRAS.index("Foto") + 1
    sql = "SELECT * FROM inventario_sobras WHERE evento_id = ?" + (" AND localizacao = ?" if localizacao else "") + " ORDER BY localizacao, id"
    for s in _todos(conn, sql, *([evento_id, localizacao] if localizacao else [evento_id])):
        acrescentar_linha(ws2, [s["localizacao"], s["descricao"], s["complemento"], s["observacao"], s["integrante"], _data_br(s["criado_em"]), s["foto_url"]])
        _celula_foto(ws2, ws2.max_row, col_foto, s["foto_url"], fotos)
    wb.save(destino)
    return destino
```

- [ ] **Step 4: Rodar tudo** → `196 passed` (conferir também `test_titulo_do_xlsx_do_inventario_e_texto_literal`: A1 continua texto).

- [ ] **Step 5: Commit**

```bash
git add inventario.py tests/test_inventario.py
git commit -m "xlsx do inventário: cabeçalho com data, filtros e total; opção de fotos com =IMAGE" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 6: Dados do painel (`inventario.painel`)

**Files:**
- Modify: `inventario.py` (após `resumo`)
- Test: `tests/test_inventario.py`

**Interfaces:**
- Consumes: `salas`, `resumo`, `andar`, `ANDAR_SEM`, `CONSERVACAO`, `CONSERVACAO_VAZIA`.
- Produces: `painel(conn, evento_id, andar_sel=None) -> dict` com as chaves `resumo`, `situacao`, `integrantes`, `conservacao`, `andares`, `salas_do_andar` (formato no teste abaixo).

- [ ] **Step 1: Testes**

```python
def test_painel_dados(dados):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.ler(dados, eid, "01 - SALA CCI", 2001, "Beltrana")      # divergente (é da SALA B)
    inventario.ler(dados, eid, "01 - SALA CCI", 1002, "Fulano")
    inventario.atualizar_leitura(dados, eid, 1001, conservacao="Ruim")
    p = inventario.painel(dados, eid)
    assert p["resumo"]["lidos"] == 2
    assert [(x["chave"], x["rotulo"], x["quantidade"]) for x in p["situacao"]] == \
        [("localizado", "Localizado", 2), ("divergente", "Divergente", 1), ("pendente", "Não localizado", 3)]
    assert [(x["chave"], x["quantidade"]) for x in p["integrantes"]] == [("Fulano", 2), ("Beltrana", 1)]
    assert [(x["chave"], x["rotulo"], x["quantidade"]) for x in p["conservacao"]] == [("Ruim", "Ruim", 1), ("-", "Não informada", 2)]
    assert p["andares"] == [{"andar": "01", "total": 2, "localizados": 2, "pendentes": 0, "divergentes": 1, "salas": 1},
                            {"andar": "02", "total": 2, "localizados": 0, "pendentes": 2, "divergentes": 0, "salas": 1},
                            {"andar": "99", "total": 1, "localizados": 0, "pendentes": 1, "divergentes": 0, "salas": 1}]
    assert p["salas_do_andar"] == []
    assert inventario.painel(dados, eid, "02")["salas_do_andar"] == \
        [{"localizacao": "02 - SALA B", "total": 2, "localizados": 0, "pendentes": 2, "divergentes": 0}]
    with pytest.raises(db.ErroDeNegocio):
        inventario.painel(dados, 999)


def test_painel_sala_sem_andar_vai_por_ultimo(dados):
    semear(dados)
    dados.execute("INSERT INTO bens VALUES (5001,'ATIVO','QUADRO','','MÓVEIS','TERMOS INDIVIDUAIS','01/01/2020',1,1)")
    dados.commit()
    eid = inventario.abrir_evento(dados, "Inv", None, ["Fulano"])
    assert [a["andar"] for a in inventario.painel(dados, eid)["andares"]] == ["01", "99", inventario.ANDAR_SEM]
    assert [x["chave"] for x in inventario.painel(dados, eid)["integrantes"]] == [] and inventario.painel(dados, eid)["conservacao"] == []
```

- [ ] **Step 2: Rodar e ver falhar** → `AttributeError: painel`.

- [ ] **Step 3: Implementar** (após `resumo`)

```python
def painel(conn, evento_id: int, andar_sel: str | None = None) -> dict:
    """Números do painel do evento: situação dos bens, leituras por integrante, conservação informada, progresso
    por andar e, se andar_sel, por sala do andar. Tudo a partir de salas() (que lê da fonte do evento)."""
    if not _um(conn, "SELECT id FROM inventario_eventos WHERE id = ?", evento_id):
        raise ErroDeNegocio("Evento de inventário não encontrado.")
    ss = salas(conn, evento_id)
    r = resumo(conn, evento_id)
    situacao = [{"chave": "localizado", "rotulo": ROTULO_SITUACAO["localizado"], "quantidade": r["lidos"]},
                {"chave": "divergente", "rotulo": ROTULO_SITUACAO["divergente"], "quantidade": r["divergentes"]},
                {"chave": "pendente", "rotulo": ROTULO_SITUACAO["pendente"], "quantidade": r["pendentes"]}]
    integrantes = [{"chave": n, "rotulo": n, "quantidade": q} for n, q in conn.execute(
        "SELECT integrante, count(*) FROM inventario_leituras WHERE evento_id = ? GROUP BY integrante ORDER BY count(*) DESC, integrante", (evento_id,))]
    por_cons = dict(conn.execute("SELECT coalesce(conservacao, ?), count(*) FROM inventario_leituras WHERE evento_id = ? GROUP BY 1",
                                 (CONSERVACAO_VAZIA, evento_id)).fetchall())
    conservacao = [{"chave": c, "rotulo": "Não informada" if c == CONSERVACAO_VAZIA else c, "quantidade": por_cons[c]}
                   for c in (*CONSERVACAO, CONSERVACAO_VAZIA) if por_cons.get(c)]
    andares: dict = {}
    for s in ss:
        a = andares.setdefault(andar(s["localizacao"]), {"andar": andar(s["localizacao"]), "total": 0, "localizados": 0, "pendentes": 0, "divergentes": 0, "salas": 0})
        for k in ("total", "localizados", "pendentes", "divergentes"):
            a[k] += s[k]
        a["salas"] += 1
    lista_andares = [andares[k] for k in sorted(andares, key=lambda k: (k == ANDAR_SEM, k))]
    salas_do_andar = [{k: s[k] for k in ("localizacao", "total", "localizados", "pendentes", "divergentes")}
                      for s in ss if andar_sel and andar(s["localizacao"]) == andar_sel]
    return {"resumo": r, "situacao": situacao, "integrantes": integrantes, "conservacao": conservacao,
            "andares": lista_andares, "salas_do_andar": salas_do_andar}
```

- [ ] **Step 4: Rodar tudo** → `198 passed`.

- [ ] **Step 5: Commit**

```bash
git add inventario.py tests/test_inventario.py
git commit -m "Inventário: dados do painel do evento (situação, integrantes, conservação, andares)" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 7: Cards do painel (`painel_inventario.py`)

**Files:**
- Create: `painel_inventario.py`
- Test: `tests/test_painel_inventario.py` (novo)

**Interfaces:**
- Consumes: `graficos.rosca(fatias, status=None, total=None, rotulos=False, urls=None)` (urls = `{rótulo: url}`; item de dado = `{"name", "value", "url"?}`), `graficos.barras_horizontais(rotulos, valores, nome, escala=False, urls=None)` (urls = lista paralela), `graficos.colunas(categorias, series, empilhado=False, rotulos=False, urls=None, status=None)` (urls = `{nome_série: [url por categoria]}`), `graficos.tabela_dados(colunas, linhas)`; endpoints `inventario.relatorio_tela`, `inventario.painel_tela` (Task 8), `inventario.sala_tela`.
- Produces: `cards(dados: dict, evento_id: int, andar_sel: str | None = None) -> list[dict]` no formato da macro `grafico` (`_macros.html:89`): `{id, titulo, subtitulo, opcoes, resumo, alto, altura, col, tabela}`.

Atenção: o endpoint `inventario.painel_tela` só existe depois da Task 8. Neste teste, registrar uma rota provisória não é necessário: `url_for` para endpoint inexistente levanta `BuildError`. Por isso a Task 7 cria a rota **mínima** `painel_tela` já (Step 3b) e a Task 8 a completa.

- [ ] **Step 1: Teste**

```python
"""Cards ECharts do painel do evento de inventário."""


def _dados():
    return {"resumo": {},
            "situacao": [{"chave": "localizado", "rotulo": "Localizado", "quantidade": 2},
                         {"chave": "divergente", "rotulo": "Divergente", "quantidade": 1},
                         {"chave": "pendente", "rotulo": "Não localizado", "quantidade": 3}],
            "integrantes": [{"chave": f"I{i}", "rotulo": f"I{i}", "quantidade": 10 - i} for i in range(7)],
            "conservacao": [{"chave": "Bom", "rotulo": "Bom", "quantidade": 3}, {"chave": "-", "rotulo": "Não informada", "quantidade": 1}],
            "andares": [{"andar": "02", "total": 4, "localizados": 1, "pendentes": 3, "divergentes": 0, "salas": 2},
                        {"andar": "Sem andar", "total": 1, "localizados": 0, "pendentes": 1, "divergentes": 0, "salas": 1}],
            "salas_do_andar": []}


def test_cards_tipos_urls_e_tabelas(dados):
    from app import app
    import painel_inventario
    with app.test_request_context():
        c = {x["id"]: x for x in painel_inventario.cards(_dados(), 7)}
    assert list(c) == ["g-situacao", "g-integrantes", "g-conservacao", "g-andares"]
    s = c["g-situacao"]["opcoes"]["series"][0]
    assert s["type"] == "pie" and s["data"][0] == {"name": "Localizado", "value": 2, "itemStyle": {"color": "#168821"}, "url": "/inventario/7/relatorio?situacao=localizado"}
    assert c["g-situacao"]["opcoes"]["graphic"][0]["style"]["text"] == "6\nbens"
    i = c["g-integrantes"]["opcoes"]
    assert i["yAxis"]["type"] == "category" and i["series"][0]["data"][0] == {"value": 10, "url": "/inventario/7/relatorio?integrante=I0"}   # 7 itens → barras
    assert c["g-integrantes"]["tabela"]["colunas"] == ["Integrante", "Leituras"] and c["g-integrantes"]["tabela"]["linhas"][0][0]["url"] == "/inventario/7/relatorio?integrante=I0"
    k = c["g-conservacao"]["opcoes"]["series"][0]
    assert k["type"] == "pie" and k["data"][1]["url"] == "/inventario/7/relatorio?conservacao=-"
    a = c["g-andares"]["opcoes"]
    assert a["xAxis"]["data"] == ["02", "Sem andar"] and [x["name"] for x in a["series"]] == ["Localizados", "Pendentes"]
    assert a["series"][0]["stack"] == "total" and a["series"][0]["itemStyle"] == {"color": "#168821"}
    assert a["series"][1]["data"][0] == {"value": 3, "url": "/inventario/7/painel?andar=02"}
    assert c["g-andares"]["tabela"]["colunas"] == ["Andar", "Bens", "Localizados", "Divergentes", "Pendentes"]
    assert c["g-andares"]["tabela"]["linhas"][0] == [{"valor": "02", "url": "/inventario/7/painel?andar=02"}, {"valor": 4}, {"valor": 1}, {"valor": 0}, {"valor": 3}]
    assert c["g-andares"]["col"] is None and c["g-integrantes"]["altura"] is None


def test_card_salas_do_andar(dados):
    from app import app
    import painel_inventario
    d = _dados()
    d["salas_do_andar"] = [{"localizacao": f"02 - SALA {i}", "total": 1, "localizados": 0, "pendentes": 1, "divergentes": 0} for i in range(11)]
    with app.test_request_context():
        c = {x["id"]: x for x in painel_inventario.cards(d, 7, "02")}
        assert "g-salas" not in {x["id"] for x in painel_inventario.cards(d, 7)}       # sem andar escolhido, sem card
    s = c["g-salas"]
    assert s["titulo"] == "Salas do andar 02" and s["col"] == "col-12" and s["opcoes"]["xAxis"]["axisLabel"]["rotate"] == 45
    assert s["opcoes"]["xAxis"]["data"][0] == "SALA 0"
    assert s["opcoes"]["series"][1]["data"][0]["url"] == "/inventario/7/sala/02%20-%20SALA%200"
    assert s["tabela"]["linhas"][0][0] == {"valor": "02 - SALA 0", "url": "/inventario/7/sala/02%20-%20SALA%200"}
    d["salas_do_andar"] = [{"localizacao": "TERMOS INDIVIDUAIS", "total": 1, "localizados": 0, "pendentes": 1, "divergentes": 0}]
    with app.test_request_context():
        s = next(x for x in painel_inventario.cards(d, 7, "Sem andar") if x["id"] == "g-salas")
    assert s["opcoes"]["xAxis"]["data"] == ["TERMOS INDIVIDUAIS"] and s["col"] is None and "rotate" not in str(s["opcoes"]["xAxis"])
```

- [ ] **Step 2: Rodar e ver falhar** → `ModuleNotFoundError: painel_inventario`.

- [ ] **Step 3: Criar `painel_inventario.py`**

```python
"""Cards do painel do evento de inventário, no formato da macro `grafico` (templates/_macros.html) e no padrão do
Recorte: tipo de gráfico pelo nº de itens, tudo clicável, tabela "Ver dados". Só monta opções ECharts com
graficos.py; os números vêm de inventario.painel()."""
from flask import url_for

import graficos
from inventario import ANDAR_SEM

TOP = 20
STATUS_SITUACAO = {"Localizado": "sucesso", "Divergente": "alerta", "Não localizado": "neutro"}
STATUS_PROGRESSO = {"Localizados": "sucesso", "Pendentes": "pendente"}


def _card(id, titulo, opcoes, tabela, resumo_itens, subtitulo=None, altura=None, col=None):
    resumo = f"{titulo}: " + ", ".join(f"{r} {q}" for r, q in resumo_itens[:6])
    return {"id": id, "titulo": titulo, "subtitulo": subtitulo, "opcoes": opcoes, "resumo": resumo,
            "alto": altura == "alto", "altura": altura, "col": col, "tabela": tabela}


def _tabela_contagem(itens, urls, rotulo_col, valor_col):
    return graficos.tabela_dados([rotulo_col, valor_col], [[{"valor": i["rotulo"], "url": u}, i["quantidade"]] for i, u in zip(itens, urls)])


def _tabela_progresso(itens, urls, rotulo_col, chave_rotulo):
    return graficos.tabela_dados([rotulo_col, "Bens", "Localizados", "Divergentes", "Pendentes"],
                                 [[{"valor": i[chave_rotulo], "url": u}, i["total"], i["localizados"], i["divergentes"], i["pendentes"]]
                                  for i, u in zip(itens, urls)])


def _grafico(itens, urls, nome):
    """≤ 5 itens → rosca com o total no centro; senão barras horizontais dos TOP maiores.
    Devolve (opções, subtítulo, altura)."""
    if len(itens) <= 5:
        op = graficos.rosca([(i["rotulo"], i["quantidade"]) for i in itens], total=(sum(i["quantidade"] for i in itens), nome.lower()),
                            urls={i["rotulo"]: u for i, u in zip(itens, urls)})
        return op, None, None
    top, top_urls = itens[:TOP], urls[:TOP]
    op = graficos.barras_horizontais([i["rotulo"] for i in top], [i["quantidade"] for i in top], nome, urls=top_urls)
    sub = f"{TOP} maiores no gráfico; todos na tabela" if len(itens) > TOP else None
    return op, sub, "extra" if len(top) > 15 else ("alto" if len(top) > 8 else None)


def _empilhado(rotulos, itens, urls):
    return graficos.colunas(rotulos, {"Localizados": [i["localizados"] for i in itens], "Pendentes": [i["pendentes"] for i in itens]},
                            empilhado=True, rotulos=True, status=STATUS_PROGRESSO, urls={"Localizados": urls, "Pendentes": urls})


def cards(dados: dict, evento_id: int, andar_sel: str | None = None) -> list[dict]:
    def rel(**f):
        return url_for("inventario.relatorio_tela", id=evento_id, **f)
    lista = []
    it = dados["situacao"]
    urls = [rel(situacao=i["chave"]) for i in it]
    op = graficos.rosca([(i["rotulo"], i["quantidade"]) for i in it], status=STATUS_SITUACAO,
                        total=(sum(i["quantidade"] for i in it), "bens"), urls={i["rotulo"]: u for i, u in zip(it, urls)})
    lista.append(_card("g-situacao", "Bens por situação", op, _tabela_contagem(it, urls, "Situação", "Bens"),
                       [(i["rotulo"], i["quantidade"]) for i in it], subtitulo="clique para abrir o relatório"))
    it = dados["integrantes"]
    if it:
        urls = [rel(integrante=i["chave"]) for i in it]
        op, sub, altura = _grafico(it, urls, "Leituras")
        lista.append(_card("g-integrantes", "Leituras por integrante", op, _tabela_contagem(it, urls, "Integrante", "Leituras"),
                           [(i["rotulo"], i["quantidade"]) for i in it], subtitulo=sub, altura=altura))
    it = dados["conservacao"]
    if it:
        urls = [rel(conservacao=i["chave"]) for i in it]
        op, sub, altura = _grafico(it, urls, "Leituras")
        lista.append(_card("g-conservacao", "Conservação informada", op, _tabela_contagem(it, urls, "Conservação", "Leituras"),
                           [(i["rotulo"], i["quantidade"]) for i in it], subtitulo=sub, altura=altura))
    it = dados["andares"]
    if it:
        urls = [url_for("inventario.painel_tela", id=evento_id, andar=i["andar"]) for i in it]
        lista.append(_card("g-andares", "Progresso por andar", _empilhado([i["andar"] for i in it], it, urls),
                           _tabela_progresso(it, urls, "Andar", "andar"), [(i["andar"], i["total"]) for i in it],
                           subtitulo="localizados e pendentes; clique no andar para ver as salas", col="col-12" if len(it) > 10 else None))
    it = dados["salas_do_andar"]
    if andar_sel and it:
        urls = [url_for("inventario.sala_tela", id=evento_id, localizacao=s["localizacao"]) for s in it]
        rotulos = [s["localizacao"] if andar_sel == ANDAR_SEM else (s["localizacao"].partition(" - ")[2] or s["localizacao"]) for s in it]
        op = _empilhado(rotulos, it, urls)
        if len(it) > 10:
            op["xAxis"]["axisLabel"] = {"interval": 0, "rotate": 45}
        lista.append(_card("g-salas", f"Salas do andar {andar_sel}", op, _tabela_progresso(it, urls, "Sala", "localizacao"),
                           [(s["localizacao"], s["total"]) for s in it],
                           subtitulo="clique na sala para abrir a leitura", col="col-12" if len(it) > 10 else None))
    return lista
```

- [ ] **Step 3b: Rota mínima `painel_tela` em `app_inventario.py`** (a Task 8 completa)

```python
@inventario_bp.route("/<int:id>/painel")
def painel_tela(id):
    abort(501)
```

- [ ] **Step 4: Rodar tudo** → `200 passed`.

- [ ] **Step 5: Commit**

```bash
git add painel_inventario.py app_inventario.py tests/test_painel_inventario.py
git commit -m "Painel do inventário: cards ECharts (situação, integrantes, conservação, andares, salas)" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 8: Página do painel e botão na tela do evento

**Files:**
- Modify: `app_inventario.py` (rota `painel_tela`, import `painel_inventario`)
- Create: `templates/inventario_painel.html`
- Modify: `templates/inventario_evento.html` (botão)
- Test: `tests/test_app.py`

**Interfaces:**
- Consumes: `inventario.painel`, `painel_inventario.cards`, macro `grafico`.

- [ ] **Step 1: Teste** (acrescentar após `test_inventario_relatorio_xlsx_e_card_do_painel`)

```python
def test_inventario_painel(cliente):
    eid = _abrir(cliente)
    cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    r = cliente.get(f"/inventario/{eid}/painel")
    assert r.status_code == 200 and b'id="g-situacao"' in r.data and b'id="g-andares"' in r.data and b'id="g-salas"' not in r.data
    assert b"bens no escopo" in r.data and b"echarts-dsgov.js" in r.data and f"/inventario/{eid}/painel?andar=01".encode() in r.data
    assert b"1 (20.0%)" in r.data                                                       # localizados com %
    r = cliente.get(f"/inventario/{eid}/painel?andar=01")
    assert b'id="g-salas"' in r.data and b"Salas do andar 01" in r.data and b"todos os andares" in r.data
    assert cliente.get("/inventario/999/painel").status_code == 404
    assert f"/inventario/{eid}/painel".encode() in cliente.get(f"/inventario/{eid}").data    # botão Painel
    cliente.post(f"/inventario/{eid}/encerrar", data={"confirmar": "1"})
    assert b"Evento encerrado" in cliente.get(f"/inventario/{eid}/painel").data
```

- [ ] **Step 2: Rodar e ver falhar** → 501.

- [ ] **Step 3: Rota**

Em `app_inventario.py`, `import painel_inventario` junto dos outros imports; substituir a rota provisória:

```python
@inventario_bp.route("/<int:id>/painel")
def painel_tela(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    andar = request.args.get("andar") or None
    dados = inventario.painel(conn, id, andar)
    return render_template("inventario_painel.html", e=e, r=dados["resumo"], andar=andar,
                           cards=painel_inventario.cards(dados, id, andar), trilha=_trilha(e, ("Painel", None)))
```

- [ ] **Step 4: Template `templates/inventario_painel.html`**

```html
{% extends "base.html" %}
{% from "_macros.html" import grafico %}
{% block titulo %}Painel — {{ e.nome }}{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-3">
  <h1 class="mb-0">Painel · {{ e.nome }}</h1>
  <div class="ml-auto">
    <a class="br-button secondary" href="{{ url_for('inventario.relatorio_tela', id=e.id) }}"><i class="fas fa-table mr-1" aria-hidden="true"></i>Relatório</a>
    <a class="br-button ml-2" href="{{ url_for('inventario.evento_tela', id=e.id) }}"><i class="fas fa-door-open mr-1" aria-hidden="true"></i>Salas</a>
  </div>
</div>
{% if e.encerrado_em %}
<div class="br-message warning mb-3"><div class="icon"><i class="fas fa-lock fa-lg" aria-hidden="true"></i></div>
  <div class="content"><span class="message-title">Evento encerrado</span><span class="message-body"> em {{ e.encerrado_em[8:10] }}/{{ e.encerrado_em[5:7] }}/{{ e.encerrado_em[:4] }}. Números congelados no encerramento.</span></div></div>
{% endif %}
<div class="row mb-2">
  {% for valor, rotulo in [(r.bens, 'bens no escopo'), (r.lidos ~ ' (' ~ r.pct_bens ~ '%)', 'localizados'), (r.divergentes, 'divergentes'),
                           (r.pendentes, 'pendentes'), (r.sobras, 'sobras'), (r.salas_iniciadas ~ ' de ' ~ r.salas, 'salas iniciadas')] %}
  <div class="col-sm-6 col-md-2 mb-3"><div class="br-card h-100"><div class="card-content"><div class="text-up-03 text-weight-bold">{{ valor }}</div><div class="text-down-01 text-gray-70">{{ rotulo }}</div></div></div></div>
  {% endfor %}
</div>
{% if andar %}<p class="text-gray-70">Andar <strong>{{ andar }}</strong> · <a href="{{ url_for('inventario.painel_tela', id=e.id) }}">todos os andares</a></p>{% endif %}
<div class="row">{% for g in cards %}{{ grafico(g) }}{% endfor %}</div>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='dsgov/vendor/echarts/echarts.min.js') }}"></script>
<script src="{{ url_for('static', filename='dsgov/js/echarts-dsgov.js') }}"></script>
{% endblock %}
```

- [ ] **Step 5: Botão na tela do evento**

Em `templates/inventario_evento.html`, antes do link "Relatório":

```html
    <a class="br-button secondary" href="{{ url_for('inventario.painel_tela', id=e.id) }}"><i class="fas fa-chart-pie mr-1" aria-hidden="true"></i>Painel</a>
```
(e o link Relatório ganha `ml-2`).

- [ ] **Step 6: Rodar tudo** → `201 passed`. Abrir no navegador local (`python main.py` cai no modo navegador; ou `flask --app app run`) `/inventario/<id>/painel` e conferir que os gráficos renderizam e o clique navega.

- [ ] **Step 7: Commit**

```bash
git add app_inventario.py templates/inventario_painel.html templates/inventario_evento.html tests/test_app.py
git commit -m "Inventário: página Painel do evento com KPIs e gráficos clicáveis" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 9: Tela do relatório: filtros, ordenação, miniatura com modal e xlsx com fotos

**Files:**
- Modify: `app_inventario.py` (`relatorio_tela`, `xlsx`; helper `_filtros_relatorio`)
- Modify: `templates/inventario_relatorio.html` (reescrever)
- Test: `tests/test_app.py`

**Interfaces:**
- Consumes: `inventario.relatorio(**f)`, `descrever_filtros`, `contar_fotos`, `FILTROS_RELATORIO`, `COLUNAS_ORDEM`, `CONSERVACAO`, `ROTULO_SITUACAO`, `exportar_xlsx(..., fotos=, **f)`; macros `select`, `cabecalho_tabela`.

- [ ] **Step 1: Teste**

```python
def test_inventario_relatorio_filtros_ordem_modal_e_xlsx_com_fotos(cliente):
    import io
    from openpyxl import load_workbook
    eid = _abrir(cliente)
    cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    cliente.post(f"/inventario/{eid}/leitura/1001", json={"quem_usa": "José"})
    import db, inventario
    inventario.atualizar_leitura(db.conectar(), eid, 1001, foto_url="https://x/1001.webp")
    r = cliente.get(f"/inventario/{eid}/relatorio?integrante=Fulano&busca=jose&ordem=numero&dir=desc")
    corpo = r.data.split(b"<tbody>")[1]
    assert r.status_code == 200 and b">1001<" in corpo and b">1002<" not in corpo
    assert "Integrante Fulano".encode() in r.data and 'Busca "jose"'.encode() in r.data and b"fa-sort-down" in r.data
    assert b"ordem=numero&amp;dir=asc" in r.data or b"dir=asc&amp;ordem=numero" in r.data      # clique de novo inverte
    assert b'data-foto="https://x/1001.webp"' in r.data and b'id="scrim-foto"' in r.data and b'name="fotos"' in r.data
    assert b'name="ordem" value="numero"' in r.data and b"1 linha(s)" in r.data
    r = cliente.get(f"/inventario/{eid}/relatorio?foto=sem&conservacao=-")
    assert b">1001<" not in r.data.split(b"<tbody>")[1] and b"Sem foto" in r.data and "Não informada".encode() in r.data
    r = cliente.get(f"/inventario/{eid}/xlsx?integrante=Fulano&fotos=1")
    assert r.status_code == 200 and r.headers["Content-Disposition"].endswith(".xlsx")
    ws = load_workbook(io.BytesIO(r.data))["Bens"]
    assert ws.cell(row=3, column=1).value == "Todas as salas · Integrante Fulano" and ws.max_row == 6
    assert ws.cell(row=6, column=13).value == '=_xlfn.IMAGE("https://x/1001.webp")'
    ws = load_workbook(io.BytesIO(cliente.get(f"/inventario/{eid}/xlsx").data))["Bens"]
    assert ws.cell(row=6, column=13).value == "https://x/1001.webp" and ws.max_row == 8      # 5 de cabeçalho + 1001, 1002, 1004
```

- [ ] **Step 2: Rodar e ver falhar** (`test_inventario_relatorio_xlsx_e_card_do_painel` antigo deve continuar passando; o novo falha em `fa-sort-down`).

- [ ] **Step 3: Rotas**

Substituir `relatorio_tela` e `xlsx` em `app_inventario.py`:

```python
def _filtros_relatorio() -> dict:
    return {k: (request.args.get(k) or None) for k in inventario.FILTROS_RELATORIO}


@inventario_bp.route("/<int:id>/relatorio")
def relatorio_tela(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    f = _filtros_relatorio()
    linhas = inventario.relatorio(conn, id, **f)
    return render_template("inventario_relatorio.html", e=e, linhas=linhas, f=f, ativos={k: v for k, v in f.items() if v},
                           descricao=inventario.descrever_filtros(f), n_fotos=inventario.contar_fotos(linhas),
                           salas=[s["localizacao"] for s in inventario.salas(conn, id)], rotulos=inventario.ROTULO_SITUACAO,
                           conservacao=inventario.CONSERVACAO, colunas_ordem=inventario.COLUNAS_ORDEM, trilha=_trilha(e, ("Relatório", None)))


@inventario_bp.route("/<int:id>/xlsx")
def xlsx(id):
    conn = _conn()
    _evento_ou_404(conn, id)
    f = _filtros_relatorio()
    arquivo = inventario.exportar_xlsx(conn, id, io.BytesIO(), fotos=request.args.get("fotos") == "1", **f)
    arquivo.seek(0)
    nome = f"inventario_{id}_{''.join(c if c.isalnum() else '_' for c in (f['localizacao'] or 'todas'))}.xlsx"
    return send_file(arquivo, as_attachment=True, download_name=nome)
```

- [ ] **Step 4: Template `templates/inventario_relatorio.html`** (substituir inteiro)

```html
{% extends "base.html" %}
{% from "_macros.html" import select, cabecalho_tabela %}
{% block titulo %}Relatório — {{ e.nome }}{% endblock %}
{% macro th(col, rotulo) -%}
<th scope="col">{% if col in colunas_ordem %}<a href="{{ url_for('inventario.relatorio_tela', id=e.id, **dict(ativos, ordem=col, dir=('desc' if f.ordem == col and f.dir != 'desc' else 'asc'))) }}">{{ rotulo }}{% if f.ordem == col %} <i class="fas fa-sort-{{ 'down' if f.dir == 'desc' else 'up' }}" aria-hidden="true"></i>{% endif %}</a>{% else %}{{ rotulo }}{% endif %}</th>
{%- endmacro %}
{% block conteudo %}
<div class="d-flex align-items-center mb-3"><h1 class="mb-0">Relatório · {{ e.nome }}</h1>
  <a class="br-button secondary ml-auto" href="{{ url_for('inventario.painel_tela', id=e.id) }}"><i class="fas fa-chart-pie mr-1" aria-hidden="true"></i>Painel</a></div>
<form method="get" class="br-card mb-3"><div class="card-content row">
  <div class="col-md-3 mb-3">{{ select('localizacao', 'Sala', [('', 'Todas')] + salas, selecionado=f.localizacao or '', obrigatorio=False) }}</div>
  <div class="col-md-3 mb-3">{{ select('situacao', 'Situação', [('', 'Todas')] + rotulos.items()|list, selecionado=f.situacao or '', obrigatorio=False) }}</div>
  <div class="col-md-3 mb-3">{{ select('integrante', 'Integrante', [('', 'Todos')] + e.integrantes, selecionado=f.integrante or '', obrigatorio=False) }}</div>
  <div class="col-md-3 mb-3">{{ select('conservacao', 'Conservação', [('', 'Todas'), ('-', 'Não informada')] + conservacao|list, selecionado=f.conservacao or '', obrigatorio=False) }}</div>
  <div class="col-md-3 mb-3">{{ select('foto', 'Foto', [('', 'Todas'), ('com', 'Com foto'), ('sem', 'Sem foto')], selecionado=f.foto or '', obrigatorio=False) }}</div>
  <div class="col-md-4 mb-3"><div class="br-input"><label for="busca">Busca (nº, descrição, quem usa, observação)</label><input id="busca" name="busca" type="search" value="{{ f.busca or '' }}"/></div></div>
  {% if f.ordem %}<input type="hidden" name="ordem" value="{{ f.ordem }}"/><input type="hidden" name="dir" value="{{ f.dir or 'asc' }}"/>{% endif %}
  <div class="col-md-5 mb-3 d-flex align-items-end flex-wrap">
    <button class="br-button primary" type="submit"><i class="fas fa-filter mr-1" aria-hidden="true"></i>Aplicar</button>
    <a class="br-button ml-2" href="{{ url_for('inventario.relatorio_tela', id=e.id) }}">Limpar</a>
    <div class="br-checkbox ml-3 mr-2"><input id="fotos" name="fotos" type="checkbox" value="1"/><label for="fotos">Incluir fotos</label></div>
    <button class="br-button secondary" type="submit" formaction="{{ url_for('inventario.xlsx', id=e.id) }}"><i class="fas fa-file-excel mr-1" aria-hidden="true"></i>Exportar .xlsx</button>
  </div>
</div></form>
{% set r = e.resumo %}
<p class="text-gray-70">{{ descricao }} · {{ linhas|length }} linha(s) · evento: {{ r.lidos }} localizados, {{ r.divergentes }} divergentes, {{ r.pendentes }} pendentes, {{ r.sobras }} sobras</p>
{% if n_fotos >= 50 %}
<div class="br-message warning mb-3"><div class="icon"><i class="fas fa-exclamation-triangle fa-lg" aria-hidden="true"></i></div>
  <div class="content" role="alert"><span class="message-title">Relatório pesado.</span><span class="message-body"> {{ n_fotos }} fotos no filtro: o .xlsx com fotos pode demorar para abrir.</span></div></div>
{% endif %}
{{ cabecalho_tabela('Bens', 'rel') }}
  <thead><tr>{{ th('numero', 'Nº') }}{{ th('descricao', 'Descrição') }}{{ th('local_sistema', 'Local sistema') }}{{ th('local_inventario', 'Local inventário') }}{{ th('situacao_inv', 'Situação') }}{{ th('conservacao', 'Conservação') }}{{ th('quem_usa', 'Quem usa') }}<th scope="col">Observação</th>{{ th('integrante', 'Integrante') }}{{ th('lido_em', 'Data/hora') }}<th scope="col">Foto</th><th scope="col">Situação do bem</th></tr></thead>
  <tbody>{% for x in linhas %}
  <tr><td><a href="{{ url_for('bem', numero=x.numero) }}">{{ x.numero }}</a></td><td>{{ x.descricao }}{% if x.complemento %} <span class="text-gray-70">{{ x.complemento }}</span>{% endif %}</td><td>{{ x.local_sistema }}</td><td>{{ x.local_inventario or '—' }}</td>
    <td>{% if x.situacao_inv == 'localizado' %}<span class="br-tag bg-success text-pure-0"><span>Localizado</span></span>{% elif x.situacao_inv == 'divergente' %}<span class="br-tag bg-warning"><span>Divergente</span></span>{% else %}<span class="br-tag bg-gray-20"><span>Não localizado</span></span>{% endif %}</td>
    <td>{{ x.conservacao or '' }}</td><td>{{ x.quem_usa or '' }}</td><td>{{ x.observacao or '' }}</td><td>{{ x.integrante or '' }}</td>
    <td>{% if x.lido_em %}{{ x.lido_em[8:10] }}/{{ x.lido_em[5:7] }}/{{ x.lido_em[:4] }} {{ x.lido_em[11:16] }}{% endif %}</td>
    <td>{% if x.foto_url %}<button type="button" class="foto-abrir" data-foto="{{ x.foto_url }}" aria-label="Ver foto do bem {{ x.numero }}"><img src="{{ x.foto_url }}" alt="" class="dsgov-miniatura"/></button>{% else %}—{% endif %}</td>
    <td>{{ x.situacao_bem }}</td></tr>
  {% else %}<tr><td colspan="12">Nenhum bem neste filtro.</td></tr>
  {% endfor %}</tbody>
</table></div>
<div class="br-scrim-util foco" id="scrim-foto" hidden>
  <div class="br-modal medium" role="dialog" aria-modal="true" aria-labelledby="modal-foto-titulo">
    <div class="br-modal-header"><div class="br-modal-title" id="modal-foto-titulo">Foto</div>
      <button class="br-button circle small" type="button" id="modal-foto-fechar" aria-label="Fechar"><i class="fas fa-times" aria-hidden="true"></i></button></div>
    <div class="br-modal-body text-center"><img id="modal-foto-img" src="" alt="Foto do bem" style="max-width:100%;max-height:70vh"/></div>
    <div class="br-modal-footer justify-content-end"><a class="br-button secondary" id="modal-foto-link" href="#" target="_blank" rel="noopener">Abrir em nova aba</a></div>
  </div>
</div>
{% endblock %}
{% block scripts %}
<style>.foto-abrir { border: 0; background: none; padding: 0; cursor: pointer; }</style>
<script>
(function () {
  var scrim = document.getElementById("scrim-foto"), img = document.getElementById("modal-foto-img"), link = document.getElementById("modal-foto-link");
  function fechar() { scrim.hidden = true; scrim.classList.remove("active"); img.src = ""; }
  document.querySelectorAll(".foto-abrir").forEach(function (b) {
    b.addEventListener("click", function () { img.src = b.dataset.foto; link.href = b.dataset.foto; scrim.hidden = false; scrim.classList.add("active"); });
  });
  document.getElementById("modal-foto-fechar").addEventListener("click", fechar);
  scrim.addEventListener("click", function (ev) { if (ev.target === scrim) fechar(); });
  document.addEventListener("keydown", function (ev) { if (ev.key === "Escape" && !scrim.hidden) fechar(); });
})();
</script>
{% endblock %}
```

Observação: o `{% macro th %}` fica no topo do template filho, antes de `{% block conteudo %}` (verificado: Jinja aceita macro no topo de template que estende outro, e aceita `**dict(...)` em chamadas).

- [ ] **Step 5: Rodar tudo** → `202 passed`. Conferir no navegador: clicar num cabeçalho ordena e inverte; miniatura abre o modal; "Incluir fotos" + Exportar baixa o xlsx com `=IMAGEM` (abrir no Excel se possível).

- [ ] **Step 6: Commit**

```bash
git add app_inventario.py templates/inventario_relatorio.html tests/test_app.py
git commit -m "Relatório do inventário: filtros, ordenação por coluna, miniatura com modal e xlsx com fotos" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 10: Lote na tela da sala e README

**Files:**
- Modify: `app_inventario.py` (rota `lote`)
- Modify: `templates/inventario_sala.html` (seleção, formulário, JS de confirmação)
- Modify: `README.md` (parágrafo do inventário: painel, relatório, fotos no xlsx, lote, snapshot)
- Test: `tests/test_app.py`

**Interfaces:**
- Consumes: `inventario.ler_lote`, `inventario.desfazer_leituras`, `fotos.apagar(url)`; macros `cabecalho_tabela(titulo, id, selecao=True)`, `th_selecao(id)`, `td_selecao(id, indice, name, value, form=None)`.

- [ ] **Step 1: Teste**

```python
def test_inventario_lote_marcar_e_desmarcar(cliente, monkeypatch):
    import fotos
    eid = _abrir(cliente)
    r = cliente.get(f"/inventario/{eid}/sala/01 - SALA CCI")
    assert b'id="form-lote"' in r.data and b'name="numeros"' in r.data and b'value="marcar"' in r.data and b'value="desmarcar"' in r.data
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/lote", data={"acao": "marcar", "numeros": ["1001", "1002", "99999"]}, follow_redirects=True)
    assert b"2 bem(ns) marcado(s)" in r.data and b"99999" in r.data and r.data.count(b">Localizado<") == 2
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/lote", data={"acao": "marcar"}, follow_redirects=True)
    assert b"Selecione ao menos um bem" in r.data
    import db, inventario
    inventario.atualizar_leitura(db.conectar(), eid, 1001, foto_url="https://x/1001.webp")
    apagadas = []
    monkeypatch.setattr(fotos, "apagar", apagadas.append)
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/lote", data={"acao": "desmarcar", "numeros": ["1001"]}, follow_redirects=True)
    assert b"desfeita" in r.data and apagadas == ["https://x/1001.webp"] and r.data.count(b">Localizado<") == 1
    with cliente.session_transaction() as sess:
        sess.pop("integrante", None)
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/lote", data={"acao": "marcar", "numeros": ["1002"]}, follow_redirects=True)
    assert b"integrante" in r.data.lower()
    cliente.post(f"/inventario/{eid}/integrante", data={"integrante": "Fulano"})
    cliente.post(f"/inventario/{eid}/encerrar", data={"confirmar": "1"})
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/lote", data={"acao": "marcar", "numeros": ["1002"]}, follow_redirects=True)
    assert b"Evento encerrado" in r.data
    assert b'id="form-lote"' not in cliente.get(f"/inventario/{eid}/sala/01 - SALA CCI").data
```

- [ ] **Step 2: Rodar e ver falhar** → 404 em `/lote`.

- [ ] **Step 3: Rota** (após `atualizar_leitura` em `app_inventario.py`)

```python
@inventario_bp.route("/<int:id>/sala/<path:localizacao>/lote", methods=["POST"])
def lote(id, localizacao):
    """Marcar como localizados (leitura sem plaqueta) ou desmarcar (apaga a leitura) os bens selecionados."""
    conn = _conn()
    _evento_ou_404(conn, id)
    volta = redirect(url_for("inventario.sala_tela", id=id, localizacao=localizacao))
    numeros = [int(n) for n in request.form.getlist("numeros") if n.strip().isdigit()]
    if not numeros:
        flash("Selecione ao menos um bem.", "warning")
        return volta
    if request.form.get("acao") == "desmarcar":
        for url in inventario.desfazer_leituras(conn, id, numeros):
            fotos.apagar(url)
        flash("Leitura(s) desfeita(s): os bens voltaram a pendentes.", "success")
        return volta
    r = inventario.ler_lote(conn, id, localizacao, numeros, session.get("integrante") or "")
    msg = f"{r['lidos']} bem(ns) marcado(s) como localizado(s)."
    if r["nao_encontrados"]:
        msg += " Não encontrado(s): " + ", ".join(str(n) for n in r["nao_encontrados"]) + "."
    flash(msg, "success" if r["lidos"] else "warning")
    return volta
```
(`ErroDeNegocio` de evento encerrado/sem integrante sobe para o handler global: flash + redirect.)

- [ ] **Step 4: Template da sala**

Em `templates/inventario_sala.html`:

1. Import: `{% from "_macros.html" import cabecalho_tabela, th_selecao, td_selecao %}`.
2. Antes de `{{ cabecalho_tabela('Bens da sala', 'bens') }}`, o formulário:

```html
{% if not fechado %}
<form method="post" action="{{ url_for('inventario.lote', id=e.id, localizacao=localizacao) }}" id="form-lote" class="mb-2 d-flex align-items-center flex-wrap">
  <button class="br-button secondary small" type="submit" name="acao" value="marcar"{% if not integrante %} disabled{% endif %}><i class="fas fa-check-double mr-1" aria-hidden="true"></i>Marcar selecionados como localizados</button>
  <button class="br-button small ml-2" type="submit" name="acao" value="desmarcar"{% if not integrante %} disabled{% endif %}><i class="fas fa-undo mr-1" aria-hidden="true"></i>Desmarcar</button>
  <span class="text-down-01 text-gray-70 ml-2">plaqueta ilegível ou ausente: marque na lista e use os botões</span>
</form>
{% endif %}
```

3. `{{ cabecalho_tabela('Bens da sala', 'bens', selecao=not fechado) }}`; no `<thead><tr>`, antes de `<th scope="col">Nº</th>`: `{% if not fechado %}{{ th_selecao('bens') }}{% endif %}`; em cada `<tr data-numero=...>`, antes de `<td>{{ b.numero }}</td>`: `{% if not fechado %}{{ td_selecao('bens', loop.index, 'numeros', b.numero, form='form-lote') }}{% endif %}`.

4. No JS, após a linha `if (!campo) return;`:

```js
  var formLote = document.getElementById("form-lote");
  if (formLote) formLote.addEventListener("submit", function (ev) {
    var n = document.querySelectorAll('input[name="numeros"]:checked').length;
    if (!n) { ev.preventDefault(); mostrar("warning", "Selecione ao menos um bem na lista."); return; }
    if (ev.submitter && ev.submitter.value === "desmarcar" && !confirm("Desfazer a leitura de " + n + " bem(ns)? Eles voltam a pendentes e a foto é apagada.")) ev.preventDefault();
  });
```

Conferir que `mostrar(tipo, texto, sobra)` aceita `"warning"` como classe de `br-message` (ver a função em `inventario_sala.html:102-105`); se ela só troca a classe, `warning` já funciona.

- [ ] **Step 5: README**

No parágrafo do inventário (README ~linha 32), acrescentar: "Cada evento tem **Painel** (KPIs e gráficos por situação, integrante, conservação, andar e sala — o andar é o texto antes do primeiro `-` no nome da sala), relatório com filtros/ordenação e `.xlsx` com opção *Incluir fotos* (`=IMAGEM`). Na sala, é possível marcar bens como localizados em lote (plaqueta ilegível) e desmarcar. Ao encerrar, os bens do evento são congelados (`inventario_bens_encerrados`): o relatório de um evento encerrado não muda quando a base do SPW é atualizada."

- [ ] **Step 6: Rodar tudo** → `203 passed`. No navegador: marcar dois bens e usar os botões; o "Desmarcar" pede confirmação.

- [ ] **Step 7: Commit**

```bash
git add app_inventario.py templates/inventario_sala.html README.md tests/test_app.py
git commit -m "Leitura em lote na sala: marcar selecionados como localizados e desmarcar" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

## Depois das tarefas

1. `python -m pytest -q` verde; `git log main..fase2-inventario --oneline` mostra 10 commits.
2. Revisão final do branch (superpowers:requesting-code-review), corrigir o que aparecer.
3. Integração: merge em `main`, `git push`, na VPS `docker compose up -d --build` (o esquema cria a tabela nova no primeiro start). Encerrar o evento 1 só quando o usuário decidir — o snapshot é gravado nessa hora.
4. Pedir ao usuário para validar no celular: painel, relatório (modal da foto), lote, xlsx com fotos no Excel. Só depois ele desliga o `sga-cfc`.
