# Administração de inventários — plano de implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Tela Administração (só admin) com abas Inventários e Usuários; cada inventário não finalizado tem uma chave Abrir/Fechar (ligar em um fecha o outro); Finalizar continua permanente; quem entra na comissão ganha a função Inventário.

**Architecture:** Coluna `suspenso_em` em `inventario_eventos` deriva os três estados (aberto, fechado, finalizado) sem mexer em `aberto_em`/`encerrado_em`. As rotas administrativas existentes (`inventario.abrir/encerrar/comissao/excluir`) ficam com nome e URL, ganham `inventario.abrir_chave` e `inventario.fechar`, e passam a redirecionar para a tela nova `admin.tela` (`/administracao`), que renderiza a aba Inventários; a aba Usuários é a tela `/usuarios` de hoje com a barra de abas. A tela Inventário fica só com a conferência.

**Tech Stack:** Flask + Jinja, SQLite, DSGov 3.7 (só visual), pytest (`.venv/bin/pytest -q`, baseline 3027).

**Spec:** `docs/superpowers/specs/2026-09-18-administracao-inventarios-design.md`

## Global Constraints

- Estados exatamente `aberto`, `fechado`, `finalizado`; regra: `encerrado_em` preenchido = finalizado; senão `suspenso_em` preenchido = fechado; senão aberto. No máximo um aberto.
- Ligar a chave em um evento fecha o aberto na mesma transação, sem erro. Ligar em finalizado → `ErroDeNegocio` com "finalizado". Criar nasce fechado, salvo "abrir agora".
- Leituras, fotos, sobras e edição de leitura em evento fechado → `ErroDeNegocio("Evento fechado: não aceita leituras até ser reaberto.")`. Comissão pode ser editada com o evento fechado; só finalizado recusa.
- Comissão (modo web) aceita qualquer usuário ativo; quem não tem `inventariante` nem `admin` ganha `inventariante` na mesma transação; o flash lista os nomes. Inativo é recusado com "Selecione ao menos um usuário ativo." Modo desktop (sem login, `g.usuario["id"] is None`) continua por nomes, como hoje.
- Endpoints e URLs existentes não mudam de nome (`inventario.abrir` = criar, `inventario.encerrar` = finalizar, `inventario.comissao`, `inventario.excluir`, `usuarios.lista` = `/usuarios`). Novos: `admin.tela` (GET `/administracao`), `inventario.abrir_chave` (POST `/inventario/<id>/abrir`), `inventario.fechar` (POST `/inventario/<id>/fechar`). Todos os administrativos em `permissoes.ADMIN`.
- Item de menu "Administração" (ícone `fa-cogs`) substitui "Usuários"; no modo desktop aparece sem a aba Usuários. Seção da Ajuda mantém o id `usuarios` com título "Administração".
- DSGov só visual: `br-tab`, `br-card`, `br-tag`, `br-message`, macro `cabecalho_tabela`; sem JS novo além do que já existe para o escopo de salas.
- Código, textos e commits em português; prefixos `feat:`/`fix:`/`test:`/`docs:`; commits terminam com `Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>`. Nunca commitar `dados/` ou `secrets/`.
- Ramo `admin-inventarios` a partir de `main`, no próprio checkout. Rodar a suíte inteira antes de cada commit.

---

### Task 1: Estados no banco e em `inventario.py`

**Files:**
- Modify: `db.py` (ESQUEMA, tabela `inventario_eventos`; `criar_esquema`, migrações)
- Modify: `inventario.py:32-60` (consultas de evento), `:105-150` (`abrir_evento`, `encerrar_evento`)
- Test: `tests/test_inventario.py`

**Interfaces:**
- Consumes: `db._um`, `db._todos`, `db._agora`, `db._colunas`, `db.ErroDeNegocio`.
- Produces:
  - `inventario.ABERTO_SQL = "encerrado_em IS NULL AND suspenso_em IS NULL"`, `inventario.FECHADO = "Evento fechado: não aceita leituras até ser reaberto."`
  - `inventario.estado(e: dict) -> str` (`aberto|fechado|finalizado`); `inventario.evento()` passa a incluir `e["estado"]`
  - `inventario.evento_aberto(conn) -> dict | None` (só o aberto), `inventario.evento_corrente(conn) -> dict | None` (aberto, senão o fechado mais recente)
  - `inventario.criar_evento(conn, nome, descricao, integrantes, salas=None, elegiveis=None, confirmar=True, abrir=False) -> int`
  - `inventario.abrir_evento(...)` = `criar_evento(..., abrir=True)` (assinatura de hoje; fecha o que estava aberto em vez de recusar)
  - `inventario.ligar_chave(conn, id) -> dict | None` (devolve o evento que foi fechado, ou None), `inventario.desligar_chave(conn, id) -> None`
  - `inventario._evento_nao_finalizado_ou_erro(conn, id) -> dict` (para comissão); `_evento_aberto_ou_erro` recusa também fechado

- [ ] **Step 1: Testes que falham**

Acrescentar ao fim de `tests/test_inventario.py`:

```python
def test_criar_evento_nasce_fechado_e_a_chave_fecha_o_outro(dados):
    eid = semear_inventario(dados)                                             # abrir_evento: nasce aberto
    assert inventario.evento(dados, eid)["estado"] == "aberto"
    e2 = inventario.criar_evento(dados, "Inventário 2027", "", ["Fulano"])
    assert inventario.evento(dados, e2)["estado"] == "fechado" and inventario.evento_aberto(dados)["id"] == eid
    fechado = inventario.ligar_chave(dados, e2)
    assert fechado["id"] == eid
    assert inventario.evento_aberto(dados)["id"] == e2 and inventario.evento(dados, eid)["estado"] == "fechado"
    assert inventario.ligar_chave(dados, e2) is None                          # já estava aberto: nada muda
    inventario.desligar_chave(dados, e2)
    assert inventario.evento_aberto(dados) is None and inventario.evento_corrente(dados)["id"] == e2   # fechado mais recente
    inventario.desligar_chave(dados, e2)                                       # idempotente
    e3 = inventario.abrir_evento(dados, "Inventário 2028", "", ["Fulano"])    # sem erro mesmo com outros fechados
    assert inventario.evento_aberto(dados)["id"] == e3 and inventario.eventos(dados)[0]["id"] == e3
    assert [e["estado"] for e in map(lambda e: inventario.evento(dados, e["id"]), inventario.eventos(dados))] == ["aberto", "fechado", "fechado"]


def test_evento_fechado_bloqueia_leitura_e_reabrir_mantem(dados):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.desligar_chave(dados, eid)
    with pytest.raises(db.ErroDeNegocio, match="fechado"):
        inventario.ler(dados, eid, "01 - SALA CCI", 1002, "Fulano")
    with pytest.raises(db.ErroDeNegocio, match="fechado"):
        inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "X", "", "", "", "Fulano", exigir_foto=False)
    with pytest.raises(db.ErroDeNegocio, match="fechado"):
        inventario.atualizar_leitura(dados, eid, 1001, conservacao="Bom")
    inventario.editar_comissao(dados, eid, ["Fulano", "Beltrana"])           # comissão muda com o evento fechado
    inventario.ligar_chave(dados, eid)
    assert inventario.resumo(dados, eid)["lidos"] == 1
    inventario.ler(dados, eid, "01 - SALA CCI", 1002, "Fulano")
    assert inventario.resumo(dados, eid)["lidos"] == 2


def test_finalizar_fechado_congela_e_nao_reabre(dados):
    eid = semear_inventario(dados)
    inventario.desligar_chave(dados, eid)
    inventario.encerrar_evento(dados, eid)
    e = inventario.evento(dados, eid)
    assert e["estado"] == "finalizado" and e["suspenso_em"] is None and e["encerrado_em"]
    assert dados.execute("SELECT count(*) FROM inventario_bens_encerrados WHERE evento_id=?", (eid,)).fetchone()[0] == 5
    with pytest.raises(db.ErroDeNegocio, match="finalizado"):
        inventario.ligar_chave(dados, eid)
    with pytest.raises(db.ErroDeNegocio, match="finalizado"):
        inventario.desligar_chave(dados, eid)
    with pytest.raises(db.ErroDeNegocio, match="encerrado"):
        inventario.editar_comissao(dados, eid, ["Fulano"])
    assert inventario.evento_corrente(dados) is None


def test_esquema_acrescenta_suspenso_em_em_banco_antigo(dados):
    semear(dados)
    eid = inventario.abrir_evento(dados, "Antigo", "", ["Fulano"])
    dados.execute("ALTER TABLE inventario_eventos DROP COLUMN suspenso_em")
    dados.commit()
    db.criar_esquema(dados)
    assert "suspenso_em" in db._colunas(dados, "inventario_eventos")
    assert inventario.evento_aberto(dados)["id"] == eid                        # o aberto de antes continua aberto
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_inventario.py -k "chave or fechado or finalizar_fechado or suspenso_em" -v`
Expected: FAIL com `AttributeError: module 'inventario' has no attribute 'criar_evento'` (e afins).

- [ ] **Step 3: Implementar**

Em `db.py`, no `ESQUEMA`, tabela `inventario_eventos`:

```sql
CREATE TABLE IF NOT EXISTS inventario_eventos (
  id           INTEGER PRIMARY KEY,
  nome         TEXT NOT NULL,
  descricao    TEXT,
  aberto_em    TEXT NOT NULL,
  encerrado_em TEXT,
  suspenso_em  TEXT                          -- preenchido = fechado (chave desligada); NULL com encerrado_em NULL = aberto
);
```

Em `criar_esquema`, logo antes da linha `conn.execute("UPDATE inventario_eventos SET encerrado_em = NULL WHERE encerrado_em = ''")`:

```python
    # 2026-09-18: chave aberto/fechado dos inventários (spec administracao-inventarios §3).
    if "suspenso_em" not in _colunas(conn, "inventario_eventos"):
        conn.execute("ALTER TABLE inventario_eventos ADD COLUMN suspenso_em TEXT")
```

Em `inventario.py`, substituir `evento_aberto`, `eventos`, `evento` e `_evento_aberto_ou_erro` por:

```python
ABERTO_SQL = "encerrado_em IS NULL AND suspenso_em IS NULL"
FECHADO = "Evento fechado: não aceita leituras até ser reaberto."


def estado(e) -> str:
    """aberto | fechado | finalizado, derivado de encerrado_em e suspenso_em."""
    if e["encerrado_em"]:
        return "finalizado"
    return "fechado" if e["suspenso_em"] else "aberto"


def evento_aberto(conn) -> dict | None:
    return _um(conn, f"SELECT * FROM inventario_eventos WHERE {ABERTO_SQL}")


def evento_corrente(conn) -> dict | None:
    """O aberto ou, sem aberto, o fechado mais recente: é o que o card do Início e o menu mostram."""
    return _um(conn, """SELECT * FROM inventario_eventos WHERE encerrado_em IS NULL
                        ORDER BY (suspenso_em IS NULL) DESC, aberto_em DESC, id DESC LIMIT 1""")


def eventos(conn) -> list[dict]:
    """Aberto primeiro, depois os fechados, depois os finalizados; dentro de cada grupo o mais recente antes."""
    return _todos(conn, """SELECT * FROM inventario_eventos ORDER BY (encerrado_em IS NULL) DESC,
                           (suspenso_em IS NULL) DESC, aberto_em DESC, id DESC""")


def evento(conn, id: int) -> dict | None:
    e = _um(conn, "SELECT * FROM inventario_eventos WHERE id = ?", id)
    if e:
        e["integrantes"] = [r[0] for r in conn.execute(
            "SELECT nome FROM inventario_integrantes WHERE evento_id = ? ORDER BY nome", (id,))]
        e["resumo"] = resumo(conn, id)
        e["estado"] = estado(e)
    return e


def _evento_nao_finalizado_ou_erro(conn, id: int) -> dict:
    e = _um(conn, "SELECT * FROM inventario_eventos WHERE id = ?", id)
    if not e:
        raise ErroDeNegocio("Evento de inventário não encontrado.")
    if e["encerrado_em"]:
        raise ErroDeNegocio("Evento encerrado: não aceita alterações.")
    return e


def _evento_aberto_ou_erro(conn, id: int) -> dict:
    e = _evento_nao_finalizado_ou_erro(conn, id)
    if e["suspenso_em"]:
        raise ErroDeNegocio(FECHADO)
    return e
```

Em `editar_comissao`, trocar `_evento_aberto_ou_erro(conn, evento_id)` por `_evento_nao_finalizado_ou_erro(conn, evento_id)` e a docstring por "Substitui a comissão do evento aberto ou fechado. Leituras já feitas não mudam: quem sai só deixa de poder ler."

Substituir `abrir_evento` por:

```python
def _fechar_os_outros(conn, eid: int) -> None:
    conn.execute(f"UPDATE inventario_eventos SET suspenso_em = ? WHERE {ABERTO_SQL} AND id <> ?", (_agora(), eid))


def criar_evento(conn, nome: str, descricao, integrantes: list, salas: list | None = None, elegiveis: list | None = None,
                 confirmar: bool = True, abrir: bool = False) -> int:
    """Cria o evento fechado (chave desligada); abrir=True já liga a chave e fecha o que estava aberto.
    salas=None → todas as localizações com bens ATIVO; lista → amostragem.
    confirmar=False deixa a transação aberta para quem chamou (comissoes.criar grava os vínculos junto)."""
    nome = _obrigatorio(nome, "Nome do evento")
    nova = fotos.pasta(nome, 0)
    for outro in eventos(conn):
        if fotos.pasta(outro["nome"], outro["id"]) == nova:
            raise ErroDeNegocio(f"Já existe um evento com esse nome (pasta de fotos '{nova}'); escolha outro nome.")
    nomes = _nomes_da_comissao(integrantes, elegiveis)
    ativas = db.localizacoes_ativas(conn)
    escolhidas = ativas if salas is None else [s for s in ativas if s in set(salas)]
    if not escolhidas:
        raise ErroDeNegocio("Nenhuma sala com bens ativos no escopo do evento.")
    agora = _agora()
    cur = conn.execute("INSERT INTO inventario_eventos (nome, descricao, aberto_em, suspenso_em) VALUES (?,?,?,?)",
                       (nome, _texto(descricao) or None, agora, None if abrir else agora))
    eid = cur.lastrowid
    conn.executemany("INSERT INTO inventario_integrantes VALUES (?,?)", [(eid, n) for n in nomes])
    conn.executemany("INSERT INTO inventario_salas (evento_id, localizacao) VALUES (?,?)", [(eid, s) for s in escolhidas])
    if abrir:
        _fechar_os_outros(conn, eid)
    if confirmar:
        conn.commit()
    return eid


def abrir_evento(conn, nome: str, descricao, integrantes: list, salas: list | None = None, elegiveis: list | None = None,
                 confirmar: bool = True) -> int:
    """Cria já com a chave ligada (o que estava aberto fica fechado)."""
    return criar_evento(conn, nome, descricao, integrantes, salas, elegiveis, confirmar, abrir=True)


def ligar_chave(conn, id: int) -> dict | None:
    """Abre o evento e fecha o que estava aberto. Devolve o evento fechado por causa disso (ou None)."""
    e = _um(conn, "SELECT * FROM inventario_eventos WHERE id = ?", id)
    if not e:
        raise ErroDeNegocio("Evento de inventário não encontrado.")
    if e["encerrado_em"]:
        raise ErroDeNegocio("Evento finalizado não pode ser reaberto.")
    atual = evento_aberto(conn)
    if atual and atual["id"] == id:
        return None
    conn.execute("UPDATE inventario_eventos SET suspenso_em = NULL WHERE id = ?", (id,))
    _fechar_os_outros(conn, id)
    conn.commit()
    return atual


def desligar_chave(conn, id: int) -> None:
    """Fecha o evento (leituras suspensas até reabrir). Fechar um já fechado não faz nada."""
    e = _um(conn, "SELECT * FROM inventario_eventos WHERE id = ?", id)
    if not e:
        raise ErroDeNegocio("Evento de inventário não encontrado.")
    if e["encerrado_em"]:
        raise ErroDeNegocio("Evento finalizado não tem chave.")
    if e["suspenso_em"]:
        return
    conn.execute("UPDATE inventario_eventos SET suspenso_em = ? WHERE id = ?", (_agora(), id))
    conn.commit()
```

Em `encerrar_evento`, trocar o `UPDATE` por:

```python
    conn.execute("UPDATE inventario_eventos SET encerrado_em = ?, suspenso_em = NULL WHERE id = ?", (_agora(), id))
```

e na docstring acrescentar "Vale para evento aberto ou fechado."

- [ ] **Step 4: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: tudo PASS. Se algum teste antigo esperava a recusa "Já existe um evento de inventário aberto", ele não existe mais (conferido: nenhum teste cita essa mensagem).

- [ ] **Step 5: Commit**

```bash
git add db.py inventario.py tests/test_inventario.py
git commit -m "feat: estados aberto/fechado/finalizado dos inventários com chave única"
```

---

### Task 2: Comissão aceita qualquer usuário ativo e concede a função Inventário

**Files:**
- Modify: `usuarios.py:171-175` (`elegiveis_comissao` fica; acrescentar `ativos_para_comissao` e `conceder_funcao`)
- Modify: `comissoes.py:44-83` (`_usuarios_selecionados`, `definir`, `abrir` → `criar`)
- Test: `tests/test_escopo_inventario.py`

**Interfaces:**
- Consumes: `inventario.criar_evento(..., confirmar=False, abrir=...)`, `inventario._evento_nao_finalizado_ou_erro`, `usuarios.listar`, `usuarios.por_id`.
- Produces:
  - `usuarios.ativos_para_comissao(conn) -> list[dict]` (todos os ativos, por nome e login)
  - `usuarios.conceder_funcao(conn, ids, funcao="inventariante") -> list[str]` (nomes de quem ganhou; sem commit)
  - `comissoes.criar(conn, nome, descricao, ids, salas, abrir=False) -> tuple[int, list[str]]` (id do evento, nomes que ganharam a função)
  - `comissoes.abrir(conn, nome, descricao, ids, salas) -> int` (= `criar(..., abrir=True)[0]`, mantém os testes de hoje)
  - `comissoes.definir(conn, eid, ids) -> list[str]` (nomes que ganharam a função; funciona com o evento fechado)

- [ ] **Step 1: Testes que falham**

Acrescentar ao fim de `tests/test_escopo_inventario.py`:

```python
def test_comissao_aceita_qualquer_ativo_e_concede_a_funcao_inventario(dados):
    semear(dados)
    admin = usuarios.criar(dados, "adm", "Admin", SENHA_PADRAO, ["admin"], trocar_senha=False)
    leitor = usuarios.criar(dados, "leitor", "Consulta Teste", SENHA_PADRAO, ["consulta"], trocar_senha=False)
    inativo = usuarios.criar(dados, "ina", "Inativo", SENHA_PADRAO, ["inventariante"], trocar_senha=False)
    usuarios.editar(dados, inativo, "Inativo", ["inventariante"], ativo=False)
    assert [u["login"] for u in usuarios.ativos_para_comissao(dados)] == ["adm", "leitor"]
    eid, concedidos = comissoes.criar(dados, "Inv", "", [admin, leitor], None, abrir=True)
    assert concedidos == ["Consulta Teste"]
    assert usuarios.por_id(dados, leitor)["funcoes"] == ("consulta", "inventariante")
    assert usuarios.por_id(dados, admin)["funcoes"] == ("admin",)                     # admin não precisa da função
    assert _vinculos(dados, eid) == [(admin, "Admin"), (leitor, "Consulta Teste")]
    assert inventario.evento_aberto(dados)["id"] == eid
    with pytest.raises(db.ErroDeNegocio, match="usuário ativo"):
        comissoes.definir(dados, eid, [inativo])
    assert comissoes.definir(dados, eid, [leitor]) == []                              # já tem a função: nada a conceder
    inventario.desligar_chave(dados, eid)
    assert comissoes.definir(dados, eid, [admin]) == [] and _integrantes(dados, eid) == ["Admin"]   # fechado aceita comissão
    inventario.encerrar_evento(dados, eid)
    with pytest.raises(db.ErroDeNegocio, match="encerrado"):
        comissoes.definir(dados, eid, [admin])


def test_criar_sem_abrir_nasce_fechado_e_conceder_e_atomico(dados):
    semear(dados)
    fulano = usuarios.criar(dados, "fulano2", "Fulano Dois", SENHA_PADRAO, ["consulta"], trocar_senha=False)
    eid, concedidos = comissoes.criar(dados, "Preparado", "", [fulano], None)
    assert inventario.evento(dados, eid)["estado"] == "fechado" and concedidos == ["Fulano Dois"]
    with pytest.raises(db.ErroDeNegocio):
        comissoes.criar(dados, "Preparado", "", [fulano], None)                        # nome repetido: nada gravado
    assert len(inventario.eventos(dados)) == 1
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_escopo_inventario.py -k "qualquer_ativo or nasce_fechado" -v`
Expected: FAIL com `AttributeError: module 'usuarios' has no attribute 'ativos_para_comissao'`.

- [ ] **Step 3: Implementar**

Em `usuarios.py`, logo após `elegiveis_comissao`:

```python
def ativos_para_comissao(conn) -> list[dict]:
    """Todos os usuários ativos, por nome: qualquer um pode entrar na comissão (a função Inventário é concedida ao entrar)."""
    return sorted(listar(conn), key=lambda u: (u["nome"], u["login"]))


def conceder_funcao(conn, ids, funcao: str = "inventariante") -> list[str]:
    """Dá a função a quem ainda não a tem (admin já tem tudo). Devolve os nomes de quem ganhou.
    Sem commit: roda dentro da transação de quem chamou (comissoes.criar / comissoes.definir)."""
    if funcao not in FUNCOES:
        raise ValueError(funcao)
    nomes = []
    for uid in ids:
        u = por_id(conn, uid)
        if u and u["ativo"] and funcao not in u["funcoes"] and "admin" not in u["funcoes"]:
            conn.execute("INSERT OR IGNORE INTO usuarios_funcoes VALUES (?,?)", (uid, funcao))
            nomes.append(u["nome"])
    return nomes
```

Em `comissoes.py`, substituir `_usuarios_selecionados`, `definir` e `abrir` por:

```python
def _usuarios_selecionados(conn, ids) -> list[dict]:
    """IDs vindos do formulário → usuários ativos. ID oculto ou inativo é recusado."""
    import usuarios
    try:
        ids = sorted({int(i) for i in ids})
    except (TypeError, ValueError):
        raise db.ErroDeNegocio("Selecione integrantes válidos.")
    ativos = {u["id"]: u for u in usuarios.ativos_para_comissao(conn)}
    if not ids or not set(ids) <= set(ativos):
        raise db.ErroDeNegocio("Selecione ao menos um usuário ativo.")
    return [ativos[i] for i in ids]


def definir(conn, eid, ids) -> list[str]:
    """Substitui a comissão do evento aberto ou fechado: vínculos, nomes exibidos e função Inventário a quem
    não tinha. Leituras já feitas não mudam. Devolve os nomes que ganharam a função."""
    import inventario
    import usuarios
    with conn:
        conn.execute("BEGIN IMMEDIATE")
        inventario._evento_nao_finalizado_ou_erro(conn, eid)
        selecionados = _usuarios_selecionados(conn, ids)
        conn.execute("DELETE FROM inventario_comissao_usuarios WHERE evento_id=?", (eid,))
        conn.executemany("INSERT INTO inventario_comissao_usuarios VALUES (?,?,?)",
                         [(eid, u["id"], u["nome"]) for u in selecionados])
        conn.execute("DELETE FROM inventario_integrantes WHERE evento_id=?", (eid,))
        conn.executemany("INSERT INTO inventario_integrantes VALUES (?,?)",
                         [(eid, n) for n in sorted({u["nome"] for u in selecionados})])
        return usuarios.conceder_funcao(conn, [u["id"] for u in selecionados])


def criar(conn, nome, descricao, ids, salas, abrir: bool = False) -> tuple[int, list[str]]:
    """Cria o evento (fechado, ou aberto se abrir=True) com vínculos e função Inventário na mesma transação.
    Devolve (id do evento, nomes que ganharam a função)."""
    import inventario
    import usuarios
    selecionados = _usuarios_selecionados(conn, ids)
    with conn:
        eid = inventario.criar_evento(conn, nome, descricao, [u["nome"] for u in selecionados],
                                      salas, confirmar=False, abrir=abrir)
        conn.executemany("INSERT INTO inventario_comissao_usuarios VALUES (?,?,?)",
                         [(eid, u["id"], u["nome"]) for u in selecionados])
        concedidos = usuarios.conceder_funcao(conn, [u["id"] for u in selecionados])
    return eid, concedidos


def abrir(conn, nome, descricao, ids, salas) -> int:
    """Cria já com a chave ligada (compatibilidade com os testes e com o caminho antigo)."""
    return criar(conn, nome, descricao, ids, salas, abrir=True)[0]
```

- [ ] **Step 4: Rodar e ver passar; ajustar os testes da regra antiga**

Run: `.venv/bin/pytest -q`
Expected: os novos passam. Um teste antigo em `tests/test_app.py` (`test_inventario_comissao_por_usuarios`, ~linhas 884-903) afirma a regra antiga da mensagem; ajustar assim, mantendo o resto do teste (a asserção `b"Consulta Teste" not in r.data` sobre o formulário fica como está até a Tarefa 5, porque nesta tarefa o formulário de `/inventario` ainda lista só os elegíveis):
- As duas linhas `assert "ao menos um usuário ativo com função de inventário".encode() in r.data` viram `assert "Selecione ao menos um usuário ativo.".encode() in r.data` e o POST que as precede deve usar um ID inexistente (`"usuarios": [999999]`) em vez do `leitor`, porque agora o leitor é aceito. Para o `leitor` aceito, acrescentar logo após: `r = cliente.post(f"/inventario/{eid}/comissao", data={"usuarios": [leitor]}, follow_redirects=True); assert "Função Inventário concedida a: Consulta Teste".encode() in r.data` — esta última asserção só passa depois da Tarefa 4 (flash novo); marcar com `# Tarefa 4` e deixar comentada nesta tarefa, descomentar na Tarefa 4.

Run de novo: `.venv/bin/pytest -q` → tudo PASS.

- [ ] **Step 5: Commit**

```bash
git add usuarios.py comissoes.py tests/test_escopo_inventario.py tests/test_app.py
git commit -m "feat: comissão aceita qualquer usuário ativo e concede a função Inventário"
```

---

### Task 3: Planilha de cadastros com `suspenso_em`

**Files:**
- Modify: `inventario.py` (`ABAS["inv_eventos"]`, `validar_abas`, `substituir_tabelas`)
- Modify: `db.py:1030-1033` (`_ler_aba_cadastro` para `inv_eventos` com `opcionais=("suspenso_em",)`)
- Test: `tests/test_inventario.py`

**Interfaces:**
- Consumes: `inventario.ABERTO_SQL`.
- Produces: aba `inv_eventos` com colunas `["id", "nome", "descricao", "aberto_em", "encerrado_em", "suspenso_em"]`; tuplas de `linhas["inv_eventos"]` com 6 campos.

- [ ] **Step 1: Testes que falham**

Localizar em `tests/test_inventario.py` o teste que confere `wb.sheetnames == [... "inv_eventos" ...]` (~linha 260) e o helper que monta a planilha de cadastros com abas `inv_*` para importação (ler as ~40 linhas ao redor de 260-300 para copiar o jeito). Acrescentar ao fim do arquivo:

```python
def test_exportar_e_importar_cadastros_levam_suspenso_em(dados, tmp_path):
    from openpyxl import load_workbook
    eid = semear_inventario(dados)
    e2 = inventario.criar_evento(dados, "Preparado", "", ["Fulano"])          # fechado
    destino = tmp_path / "cadastros.xlsx"
    db.exportar_cadastros(dados, destino)
    ws = load_workbook(destino)["inv_eventos"]
    cab = [c.value for c in ws[1]]
    assert cab == ["id", "nome", "descricao", "aberto_em", "encerrado_em", "suspenso_em"]
    linhas = {r[0]: r for r in ws.iter_rows(min_row=2, values_only=True)}
    assert linhas[eid][5] is None and linhas[e2][5]                            # aberto sem suspenso_em; fechado com
    db.importar_cadastros(dados, destino)                                       # round-trip mantém os estados
    assert inventario.evento(dados, eid)["estado"] == "aberto" and inventario.evento(dados, e2)["estado"] == "fechado"


def test_importar_cadastros_rejeita_dois_abertos_e_aceita_planilha_antiga(dados, tmp_path):
    from openpyxl import load_workbook
    eid = semear_inventario(dados)
    e2 = inventario.criar_evento(dados, "Preparado", "", ["Fulano"])
    destino = tmp_path / "cadastros.xlsx"
    db.exportar_cadastros(dados, destino)
    wb = load_workbook(destino)
    ws = wb["inv_eventos"]
    for r in ws.iter_rows(min_row=2):
        r[5].value = None                                                       # os dois sem suspenso_em: dois abertos
    wb.save(destino)
    with pytest.raises(db.ImportacaoInvalida, match="mais de um evento aberto"):
        db.importar_cadastros(dados, destino)
    ws.delete_cols(6)                                                           # planilha antiga: sem a coluna
    for r in ws.iter_rows(min_row=2):
        if r[0].value == e2:
            r[4].value = "2026-01-01 10:00:00"                                  # e2 finalizado para sobrar um aberto
    wb.save(destino)
    db.importar_cadastros(dados, destino)
    assert inventario.evento(dados, eid)["estado"] == "aberto" and inventario.evento(dados, e2)["estado"] == "finalizado"
```

Se `db.exportar_cadastros` tiver outra assinatura (conferir com `grep -n "def exportar_cadastros" db.py`), adaptar a chamada mantendo o destino em `tmp_path`.

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_inventario.py -k "suspenso_em" -v`
Expected: FAIL na asserção do cabeçalho (falta `suspenso_em`).

- [ ] **Step 3: Implementar**

`inventario.py`:

```python
ABAS = {
    "inv_eventos": ["id", "nome", "descricao", "aberto_em", "encerrado_em", "suspenso_em"],
    ...
```

Em `validar_abas`, no laço de `inv_eventos`, depois de `encerrado = _data_iso(...)`:

```python
        suspenso = _data_iso(r.get("suspenso_em"), "suspenso_em", rot, problemas, False) if _texto(r.get("suspenso_em")) else None
        if encerrado is None and _texto(r["encerrado_em"]) == "" and suspenso is None:
            abertos += 1
```

(substituindo o `if encerrado is None and _texto(r["encerrado_em"]) == "": abertos += 1` de hoje) e a tupla passa a ser
`(eid, nome, _texto(r["descricao"]) or None, aberto, encerrado, suspenso)`. A mensagem de problema vira
`"inv_eventos: mais de um evento aberto (sem encerrado_em e sem suspenso_em)"`.

Em `substituir_tabelas`:

```python
    conn.executemany("INSERT INTO inventario_eventos (id, nome, descricao, aberto_em, encerrado_em, suspenso_em) VALUES (?,?,?,?,?,?)", linhas["inv_eventos"])
```

Em `db.importar_cadastros`, na compreensão que monta `inv_brutos`, trocar o `opcionais=` por:

```python
                                             opcionais=("foto_url", "fotos_seq") if aba == "inv_leituras" else (("suspenso_em",) if aba == "inv_eventos" else ()))
```

Conferir em `_ler_aba_cadastro` (db.py:996) que coluna opcional ausente vira chave `None` no dict (é o comportamento atual para `foto_url`); se não for, usar `r.get("suspenso_em")` como já está no código acima.

- [ ] **Step 4: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: tudo PASS (os testes existentes de exportação de cadastros que contam colunas de `inv_eventos` precisam ser ajustados para 6 colunas, se houver; procurar com `grep -n "inv_eventos" tests/`).

- [ ] **Step 5: Commit**

```bash
git add inventario.py db.py tests/test_inventario.py
git commit -m "feat: planilha de cadastros exporta e importa suspenso_em dos inventários"
```

---

### Task 4: Tela Administração, rotas da chave, permissões e menu

**Files:**
- Create: `app_admin.py`, `templates/administracao.html`, `templates/_abas_admin.html`
- Modify: `app_inventario.py:74-160` (rotas `eventos_tela`, `abrir`, `encerrar`, `comissao`, `excluir`; novas `abrir_chave`, `fechar`)
- Modify: `app.py:34-36` (registrar blueprint), `:146-153` (contexto do menu: evento corrente)
- Modify: `permissoes.py:60-75`, `menu.py:39-40, 84-88, 96-110`
- Modify: `templates/usuarios/lista.html:5-7`, `templates/inventario_comissao.html:31`, `templates/inventario_excluir.html:11`
- Test: `tests/test_app.py`, `tests/test_permissoes.py`, `tests/test_menu.py`, `tests/test_usuarios_telas.py`, `tests/test_ajuda.py`

**Interfaces:**
- Consumes: `inventario.criar_evento/ligar_chave/desligar_chave/evento_corrente/estado/eventos`, `comissoes.criar/definir`, `usuarios.ativos_para_comissao`.
- Produces: endpoints `admin.tela` (GET `/administracao`, query `aba` ignorada: a aba Usuários é `/usuarios`), `inventario.abrir_chave` (POST `/inventario/<int:id>/abrir`), `inventario.fechar` (POST `/inventario/<int:id>/fechar`); campo de formulário `abrir_agora` (`"1"`, `"0"` ou ausente = automático); flashes "Inventário X aberto.", "Inventário X aberto; Y foi fechado.", "Inventário X fechado.", "Função Inventário concedida a: ..."; template `_abas_admin.html` com variável `aba` (`inventarios` | `usuarios`).

- [ ] **Step 1: Testes que falham**

Acrescentar ao fim de `tests/test_app.py` (o arquivo já importa `db`, `inventario`, `comissoes`? conferir com `grep -n "^import\|^from" tests/test_app.py` e acrescentar `import inventario` e `import comissoes` se faltarem):

```python
def test_administracao_lista_estados_e_acoes(cliente, dados, usuarios_exemplo):
    fulano, beltrana = _ids("Fulano", "Beltrana")
    r = cliente.get("/administracao")
    assert r.status_code == 200 and b"Novo invent" in r.data and b'name="abrir_agora"' in r.data and b"checked" in r.data.split(b'name="abrir_agora"')[1][:80]
    assert b"Consulta Teste" in r.data                                                # qualquer ativo entra na comissão
    r = cliente.post("/inventario/abrir", data={"nome": "Inv A", "usuarios": [beltrana], "escopo": "todas", "abrir_agora": "1"}, follow_redirects=True)
    assert r.request.path == "/administracao" and "Inventário Inv A aberto.".encode() in r.data
    ea = inventario.evento_aberto(dados)["id"]
    r = cliente.post("/inventario/abrir", data={"nome": "Inv B", "usuarios": [beltrana], "escopo": "todas", "abrir_agora": "0"}, follow_redirects=True)
    eb = [e["id"] for e in inventario.eventos(dados) if e["nome"] == "Inv B"][0]
    html = r.get_data(as_text=True)
    assert "Inventário Inv B criado fechado." in html
    assert f'action="/inventario/{eb}/abrir"' in html and f'action="/inventario/{ea}/fechar"' in html
    assert html.count(">aberto<") == 1 and html.count(">fechado<") == 1
    r = cliente.post(f"/inventario/{eb}/abrir", follow_redirects=True)
    assert "Inventário Inv B aberto; Inv A foi fechado.".encode() in r.data and inventario.evento_aberto(dados)["id"] == eb
    r = cliente.post(f"/inventario/{eb}/fechar", follow_redirects=True)
    assert "Inventário Inv B fechado.".encode() in r.data and inventario.evento_aberto(dados) is None
    r = cliente.post(f"/inventario/{ea}/encerrar", data={"confirmar": "1"}, follow_redirects=True)
    assert r.request.path == "/administracao" and b">finalizado<" in r.data
    html = r.get_data(as_text=True)
    assert f'action="/inventario/{ea}/abrir"' not in html and f'href="/inventario/{ea}/relatorio"' in html
    assert cliente.post(f"/inventario/{ea}/abrir", follow_redirects=True).get_data(as_text=True).count("finalizado não pode ser reaberto") == 1


def test_administracao_sem_abrir_agora_abre_so_se_nao_ha_aberto(cliente, dados, usuarios_exemplo):
    beltrana = _ids("Beltrana")[0]
    cliente.post("/inventario/abrir", data={"nome": "Inv A", "usuarios": [beltrana], "escopo": "todas"})
    ea = inventario.evento_aberto(dados)["id"]                                        # sem aberto: abre
    cliente.post("/inventario/abrir", data={"nome": "Inv B", "usuarios": [beltrana], "escopo": "todas"})
    assert inventario.evento_aberto(dados)["id"] == ea                                # já havia aberto: B nasce fechado
    assert inventario.evento(dados, inventario.eventos(dados)[1]["id"])["estado"] == "fechado"


def test_administracao_comissao_concede_funcao_e_funciona_fechado(cliente, dados, usuarios_exemplo):
    beltrana, leitor = _ids("Beltrana", "Consulta Teste")
    cliente.post("/inventario/abrir", data={"nome": "Inv", "usuarios": [beltrana], "escopo": "todas"})
    eid = inventario.evento_aberto(dados)["id"]
    r = cliente.post(f"/inventario/{eid}/comissao", data={"usuarios": [beltrana, leitor]}, follow_redirects=True)
    assert "Função Inventário concedida a: Consulta Teste".encode() in r.data and r.request.path == "/administracao"
    cliente.post(f"/inventario/{eid}/fechar")
    r = cliente.get(f"/inventario/{eid}/comissao")
    assert r.status_code == 200 and b'href="/administracao"' in r.data                # Cancelar volta para a Administração
    r = cliente.post(f"/inventario/{eid}/comissao", data={"usuarios": [leitor]}, follow_redirects=True)
    assert "Comissão atualizada.".encode() in r.data and inventario_do_teste(eid)["integrantes"] == ["Consulta Teste"]


def test_administracao_so_para_admin_e_usuarios_vira_aba(cliente, usuarios_exemplo):
    r = cliente.get("/usuarios")
    assert r.status_code == 200 and b'class="br-tab' in r.data and b'href="/administracao"' in r.data and b">Usu\xc3\xa1rios<" in r.data
    for funcao in ("operador", "inventariante", "consulta"):
        cliente.post("/sair"); logar(cliente, *usuarios_exemplo[funcao])
        assert cliente.get("/administracao").status_code == 403
        assert cliente.post("/inventario/1/abrir").status_code == 403 and cliente.post("/inventario/1/fechar").status_code == 403
```

Em `tests/test_permissoes.py`:
- `_ROTAS_GET`: acrescentar `"/administracao": {"admin": 200, "operador": 403, "inventariante": 403, "consulta": 403},`.
- `test_menu_por_funcao`: trocar as duas ocorrências de `b">Usu\xc3\xa1rios<"` por `b">Administra\xc3\xa7\xc3\xa3o<"`.
- `test_modo_desktop_sem_tela_de_usuarios`: manter o 404 de `/usuarios`; trocar `b">Usu\xc3\xa1rios<" not in m` por `b">Administra\xc3\xa7\xc3\xa3o<" in m` e acrescentar `assert cliente_local.get("/administracao").status_code == 200 and b">Usu\xc3\xa1rios<" not in cliente_local.get("/administracao").data`.

Em `tests/test_menu.py`: nas três listas/asserções com `'Usuários'` (linhas ~61-62 e ~70-71), trocar por `'Administração'`; em `test_modo_local_nao_mostra_usuarios`, o item Administração aparece nos dois modos, então a asserção vira `assert [i['rotulo'] for i in local] == [r['rotulo'] for r in web]` e renomear o teste para `test_modo_local_tambem_tem_administracao`.

Em `tests/test_ajuda.py`: nada muda de ids (a seção `usuarios` continua com esse id); acrescentar em `test_administrador_ve_todas_as_secoes` a linha `assert "Administra".encode() in cliente.get("/ajuda").data`. Em `test_modo_local_nao_traz_usuarios_nem_conta`, a seção `usuarios` (agora "Administração") passa a existir no desktop: as asserções viram `assert 'conta' not in secoes` e `assert secoes == [id for id, *_ in menu.SECOES if id != 'conta']`; renomear para `test_modo_local_nao_traz_conta`.

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_app.py -k administracao -v`
Expected: FAIL com 404 em `/administracao`.

- [ ] **Step 3: Permissões, blueprint e menu**

`permissoes.py`, bloco `_registrar(ADMIN, """...""")`: acrescentar as linhas

```
admin.tela GET
inventario.abrir_chave POST
inventario.fechar POST
```

`app_admin.py` (novo):

```python
"""Tela Administração (só admin): aba Inventários aqui; a aba Usuários é a tela /usuarios de app_usuarios.py."""
from flask import Blueprint, g, render_template

import db
import inventario
import usuarios

admin_bp = Blueprint("admin", __name__)


def _conn():
    if "conn" not in g:
        g.conn = db.conectar()
    return g.conn


def _local() -> bool:
    return g.usuario["id"] is None


def _elegiveis(conn) -> list[dict]:
    """Quem pode compor a comissão: web = todos os ativos; desktop = administrador local + elegíveis por nome."""
    if _local():
        return [dict(g.usuario)] + usuarios.elegiveis_comissao(conn)
    return usuarios.ativos_para_comissao(conn)


@admin_bp.route("/administracao")
def tela():
    conn = _conn()
    eventos = [inventario.evento(conn, e["id"]) for e in inventario.eventos(conn)]
    return render_template("administracao.html", aba="inventarios", eventos=eventos,
                           ha_aberto=any(e["estado"] == "aberto" for e in eventos),
                           salas_ativas=db.localizacoes_ativas(conn), elegiveis=_elegiveis(conn), local=_local(),
                           trilha=[("Administração", None)])
```

`app.py`: junto de `from app_inventario import inventario_bp` (conferir o nome do import no topo), importar `from app_admin import admin_bp` e registrar `app.register_blueprint(admin_bp)` logo após `app.register_blueprint(inventario_bp)`.

No `contexto_dsgov` (app.py ~146-153), trocar

```python
    aberto = next((e for e in visiveis if not e["encerrado_em"]), None)
```

por

```python
    aberto = next((e for e in visiveis if not e["encerrado_em"] and not e["suspenso_em"]), None) \
        or next((e for e in visiveis if not e["encerrado_em"]), None)      # aberto; senão o fechado mais recente
```

(`eventos()` já vem ordenado: aberto, fechados por data desc). Fazer a mesma troca em `_evento_aberto_visivel` (renomear para `_evento_corrente_visivel` e atualizar o único chamador em `home`).

`menu.py`:
- `destino_atual`: trocar `if endpoint in {'usuarios.novo', 'usuarios.incluir', 'usuarios.editar', 'usuarios.nova_senha'}: return 'usuarios.lista', {}` por
  `if endpoint in {'usuarios.lista', 'usuarios.novo', 'usuarios.incluir', 'usuarios.editar', 'usuarios.nova_senha'}: return 'admin.tela', {}`
  e acrescentar `'inventario.comissao', 'inventario.excluir'` ao conjunto que hoje devolve `'inventario.evento_tela'` → passam a devolver `'admin.tela', {}` (a comissão e a exclusão agora são telas da Administração). Manter `inventario.sala_tela` → evento.
- Em `montar`, a lista final `[('Textos', ...), ('Atualizar base', ...), ('Usuários', 'fa-users', 'usuarios.lista'), ('Ajuda', ...)]` vira `[('Textos', ...), ('Atualizar base', ...), ('Administração', 'fa-cogs', 'admin.tela'), ('Ajuda', ...)]` e a condição `(e != 'usuarios.lista' or login_ativo)` sai (Administração aparece nos dois modos).
- `SECOES`: `('usuarios', 'Usuários', 'usuarios.lista', 'GET')` vira `('usuarios', 'Administração', 'admin.tela', 'GET')`. `AJUDA`: acrescentar `'admin.tela': 'usuarios'` e trocar `'usuarios.lista': 'usuarios'` por `'usuarios.lista': 'usuarios', 'inventario.comissao': 'usuarios', 'inventario.excluir': 'usuarios'`. Em `secoes_ajuda`, a exclusão `id not in {'conta', 'usuarios'} or login_ativo` vira `id != 'conta' or login_ativo` (a seção Administração existe no desktop).

- [ ] **Step 4: Rotas em `app_inventario.py`**

Substituir `eventos_tela` e `abrir` por:

```python
@inventario_bp.route("")
def eventos_tela():
    """Só os eventos visíveis a quem pediu: o inventariante vê apenas aqueles de que participa. Administração fica em /administracao."""
    conn = _conn()
    visiveis = [inventario.evento(conn, e["id"]) for e in comissoes.eventos_visiveis(conn, g.usuario)]
    corrente = next((e for e in visiveis if e["estado"] == "aberto"), None) or next((e for e in visiveis if e["estado"] == "fechado"), None)
    return render_template("inventario_eventos.html", corrente=corrente,
                           eventos=[e for e in visiveis if corrente is None or e["id"] != corrente["id"]],
                           pode_administrar=_pode("admin.tela"), pode_relatorios=_pode("inventario.relatorio_tela"),
                           trilha=_trilha())


def _abrir_agora(f) -> bool:
    """Campo abrir_agora: "1" liga a chave, "0" cria fechado; ausente (chamadas antigas) = liga só se não há aberto."""
    valor = f.get("abrir_agora")
    if valor in ("1", "0"):
        return valor == "1"
    return inventario.evento_aberto(_conn()) is None


@inventario_bp.route("/abrir", methods=["POST"])
def abrir():
    """Cria o inventário (fechado, ou já aberto com abrir_agora=1)."""
    conn = _conn()
    f = request.form
    salas = None if f.get("escopo", "todas") == "todas" else f.getlist("salas")
    abrir_agora = _abrir_agora(f)
    concedidos: list = []
    if _local():
        eid = inventario.criar_evento(conn, f.get("nome", ""), f.get("descricao", ""), _nomes_para_comissao(f.getlist("integrantes")),
                                      salas, elegiveis=[u["nome"] for u in _elegiveis(conn)], abrir=abrir_agora)
    else:
        eid, concedidos = comissoes.criar(conn, f.get("nome", ""), f.get("descricao", ""), f.getlist("usuarios"), salas, abrir=abrir_agora)
    nome = inventario.evento(conn, eid)["nome"]
    flash(f"Inventário {nome} aberto." if abrir_agora else f"Inventário {nome} criado fechado.", "success")
    _avisar_concedidos(concedidos)
    return redirect(url_for("admin.tela"))


def _avisar_concedidos(nomes):
    if nomes:
        flash("Função Inventário concedida a: " + ", ".join(nomes) + ".", "info")


@inventario_bp.route("/<int:id>/abrir", methods=["POST"])
def abrir_chave(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    fechado = inventario.ligar_chave(conn, id)
    flash(f"Inventário {e['nome']} aberto; {fechado['nome']} foi fechado." if fechado else f"Inventário {e['nome']} aberto.", "success")
    return redirect(url_for("admin.tela"))


@inventario_bp.route("/<int:id>/fechar", methods=["POST"])
def fechar(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    inventario.desligar_chave(conn, id)
    flash(f"Inventário {e['nome']} fechado.", "success")
    return redirect(url_for("admin.tela"))
```

Em `encerrar`: o redirect sem `confirmar` vai para `url_for("admin.tela", confirmar=id)` e, depois de encerrar, `flash("Inventário finalizado. As leituras ficam congeladas; relatório e planilha continuam disponíveis.", "success")` + `redirect(url_for("admin.tela"))`. A tela Administração mostra o botão "Confirmar finalização" na linha cujo id é igual ao `confirmar` da query (ver template).

Em `comissao`: trocar `inventario._evento_aberto_ou_erro(conn, id)` por `inventario._evento_nao_finalizado_ou_erro(conn, id)`; no POST web, `concedidos = comissoes.definir(conn, id, request.form.getlist("usuarios"))` seguido de `_avisar_concedidos(concedidos)`; os dois `redirect(url_for("inventario.evento_tela", id=id))` viram `redirect(url_for("admin.tela"))`; no GET, `elegiveis=_elegiveis(conn)` continua no desktop e vira `usuarios.ativos_para_comissao(conn)` no web (ajustar `_elegiveis` em `app_inventario.py` para: `return [dict(g.usuario)] + usuarios.elegiveis_comissao(conn) if _local() else usuarios.ativos_para_comissao(conn)`). Trilha: `trilha=[("Administração", url_for("admin.tela")), (f"Comissão de {e['nome']}", None)]`.

Em `excluir`: o redirect de sucesso vira `redirect(url_for("admin.tela"))`; trilha `[("Administração", url_for("admin.tela")), (f"Excluir {e['nome']}", None)]`.

Nos templates `inventario_comissao.html` (linha 31) e `inventario_excluir.html` (linha 11), o link Cancelar passa a `href="{{ url_for('admin.tela') }}"`.

Como `_escopo_do_inventario` em `app.py` exige evento visível para rotas com `id` no blueprint `inventario`, o admin passa (vê tudo); nada a mudar.

- [ ] **Step 5: Templates da Administração**

`templates/_abas_admin.html`:

```jinja
<nav class="br-tab mb-4" aria-label="Administração">
  <div class="tab-nav"><ul>
    <li class="tab-item{% if aba == 'inventarios' %} active{% endif %}"><a href="{{ url_for('admin.tela') }}"{% if aba == 'inventarios' %} aria-current="page" class="text-weight-semi-bold bg-gray-10"{% endif %}><span class="name">Inventários</span></a></li>
    {% if pode('usuarios.lista') and USUARIO.id is not none %}
    <li class="tab-item{% if aba == 'usuarios' %} active{% endif %}"><a href="{{ url_for('usuarios.lista') }}"{% if aba == 'usuarios' %} aria-current="page" class="text-weight-semi-bold bg-gray-10"{% endif %}><span class="name">Usuários</span></a></li>
    {% endif %}
  </ul></div>
</nav>
```

(`USUARIO.id is none` é o modo desktop, onde `/usuarios` responde 404.)

`templates/administracao.html`:

```jinja
{% extends "base.html" %}
{% from '_macros.html' import ajuda_titulo %}
{% from "_macros.html" import cabecalho_tabela %}
{% block titulo %}Inventários · Administração{% endblock %}
{% block conteudo %}
{% include '_abas_admin.html' %}
<div class="d-flex align-items-center mb-3"><h1 class="mb-0">Inventários</h1>{{ ajuda_titulo(AJUDA_ANCORA) }}</div>
<p class="text-gray-70">Só um inventário fica aberto por vez: ligar a chave em um fecha o outro. Finalizar congela as leituras e não tem volta.</p>

<form method="post" action="{{ url_for('inventario.abrir') }}" class="br-card mb-4">{{ csrf_campo() }}<div class="card-content">
  <div class="text-weight-semi-bold text-up-01 mb-3">Novo inventário</div>
  <div class="row">
    <div class="col-md-6 mb-3"><div class="br-input"><label for="nome">Nome</label><input id="nome" name="nome" type="text" placeholder="Ex.: Inventário 2026" required/></div></div>
    <div class="col-md-6 mb-3"><div class="br-input"><label for="descricao">Descrição (portaria, observações)</label><input id="descricao" name="descricao" type="text"/></div></div>
    <div class="col-md-6 mb-3">
      <div class="text-weight-semi-bold mb-2">Comissão</div>
      <p class="text-down-01 text-gray-70">Quem entra sem a função Inventário recebe a função na hora.</p>
      {% for u in elegiveis %}
      {% if local %}
      <div class="br-checkbox mb-1"><input id="int-{{ loop.index }}" name="integrantes" type="checkbox" value="{{ u.nome }}"{% if u.id is none %} checked disabled{% endif %}/><label for="int-{{ loop.index }}">{{ u.nome }} <span class="text-gray-70 text-down-01">· {{ u.funcoes|map('rotulo_funcao')|join(', ') }}</span></label></div>
      {% else %}
      <div class="br-checkbox mb-2"><input id="integrante-{{ u.id }}" name="usuarios" type="checkbox" value="{{ u.id }}"/>
        <label for="integrante-{{ u.id }}">{{ u.nome }} ({{ u.login }}) <span class="text-gray-70 text-down-01">· {{ u.funcoes|map('rotulo_funcao')|join(', ') }}</span></label></div>
      {% endif %}
      {% else %}<p class="text-gray-70">Nenhum usuário ativo.</p>{% endfor %}
    </div>
    <div class="col-md-6 mb-3">
      <div class="text-weight-semi-bold mb-2">Escopo</div>
      <div class="br-radio mb-2"><input id="escopo-todas" name="escopo" type="radio" value="todas" checked/><label for="escopo-todas">Todas as salas com bens ativos ({{ salas_ativas|length }})</label></div>
      <div class="br-radio mb-2"><input id="escopo-escolher" name="escopo" type="radio" value="escolher"/><label for="escopo-escolher">Escolher salas (amostragem)</label></div>
      <div id="lista-salas" hidden>
        {% for s in salas_ativas %}<div class="br-checkbox"><input id="sala-{{ loop.index }}" name="salas" type="checkbox" value="{{ s }}"/><label for="sala-{{ loop.index }}">{{ s }}</label></div>{% endfor %}
      </div>
      <input type="hidden" name="abrir_agora" value="0"/>
      <div class="br-checkbox mt-3"><input id="abrir-agora" name="abrir_agora" type="checkbox" value="1"{% if not ha_aberto %} checked{% endif %}/><label for="abrir-agora">Abrir agora{% if ha_aberto %} (fecha o inventário aberto){% endif %}</label></div>
    </div>
  </div>
  <button class="br-button primary" type="submit"><i class="fas fa-plus mr-1" aria-hidden="true"></i>Criar inventário</button>
</div></form>

{% if eventos %}
{{ cabecalho_tabela('Inventários', 'inventarios') }}
  <thead><tr><th scope="col">Inventário</th><th scope="col">Estado</th><th scope="col">Criado em</th><th scope="col">Finalizado em</th><th scope="col">Comissão</th><th scope="col" class="dsgov-numero">Localizados</th><th scope="col" class="dsgov-acoes">Ações</th></tr></thead>
  <tbody>{% for e in eventos %}
  <tr><td><a href="{{ url_for('inventario.evento_tela', id=e.id) }}">{{ e.nome }}</a>{% if e.descricao %}<div class="text-down-01 text-gray-70">{{ e.descricao }}</div>{% endif %}</td>
    <td>{% if e.estado == 'aberto' %}<span class="br-tag bg-success text-pure-0"><span>aberto</span></span>{% elif e.estado == 'fechado' %}<span class="br-tag bg-warning"><span>fechado</span></span>{% else %}<span class="br-tag bg-gray-20"><span>finalizado</span></span>{% endif %}</td>
    <td>{{ e.aberto_em[8:10] }}/{{ e.aberto_em[5:7] }}/{{ e.aberto_em[:4] }}</td>
    <td>{% if e.encerrado_em %}{{ e.encerrado_em[8:10] }}/{{ e.encerrado_em[5:7] }}/{{ e.encerrado_em[:4] }}{% else %}—{% endif %}</td>
    <td>{{ e.integrantes|join(', ') }}</td>
    <td class="dsgov-numero">{{ e.resumo.lidos }} / {{ e.resumo.bens }}</td>
    <td class="dsgov-acoes">
      {% if e.estado == 'aberto' %}
      <form method="post" action="{{ url_for('inventario.fechar', id=e.id) }}" class="d-inline">{{ csrf_campo() }}<button class="br-button secondary small" type="submit"><i class="fas fa-toggle-on mr-1" aria-hidden="true"></i>Fechar</button></form>
      {% elif e.estado == 'fechado' %}
      <form method="post" action="{{ url_for('inventario.abrir_chave', id=e.id) }}" class="d-inline">{{ csrf_campo() }}<button class="br-button secondary small" type="submit"><i class="fas fa-toggle-off mr-1" aria-hidden="true"></i>Abrir</button></form>
      {% endif %}
      {% if e.estado != 'finalizado' %}
      <form method="post" action="{{ url_for('inventario.encerrar', id=e.id) }}" class="d-inline ml-1">{{ csrf_campo() }}
        {% if confirmar == e.id %}<input type="hidden" name="confirmar" value="1"/><button class="br-button primary small" type="submit"><i class="fas fa-check mr-1" aria-hidden="true"></i>Confirmar finalização</button>
        {% else %}<button class="br-button small" type="submit"><i class="fas fa-lock mr-1" aria-hidden="true"></i>Finalizar</button>{% endif %}
      </form>
      <a class="br-button circle small ml-1" href="{{ url_for('inventario.comissao', id=e.id) }}" aria-label="Comissão de {{ e.nome }}"><i class="fas fa-users" aria-hidden="true"></i></a>
      {% else %}
      {% if pode('inventario.relatorio_tela') %}<a class="br-button circle small" href="{{ url_for('inventario.relatorio_tela', id=e.id) }}" aria-label="Relatório de {{ e.nome }}"><i class="fas fa-table" aria-hidden="true"></i></a>{% endif %}
      {% endif %}
      <a class="br-button circle small danger ml-1" href="{{ url_for('inventario.excluir', id=e.id) }}" aria-label="Excluir {{ e.nome }}"><i class="fas fa-trash" aria-hidden="true"></i></a>
    </td></tr>
  {% endfor %}</tbody>
</table></div>
{% else %}<p class="text-gray-70">Nenhum inventário criado.</p>{% endif %}
{% endblock %}
{% block scripts %}
<script>
document.querySelectorAll('input[name="escopo"]').forEach(function (r) {
  r.addEventListener("change", function () { document.getElementById("lista-salas").hidden = document.getElementById("escopo-escolher").checked === false; });
});
</script>
{% endblock %}
```

Na rota `admin.tela`, passar também `confirmar=request.args.get("confirmar", type=int)` (importar `request`). O bloco `scripts` acima é o mesmo que hoje está em `inventario_eventos.html` (copiar o original de lá, que sai na Tarefa 5).

`templates/usuarios/lista.html`: acrescentar `{% include '_abas_admin.html' %}` logo após `{% block conteudo %}` com `{% set aba = 'usuarios' %}` na linha anterior; trocar o título do `<h1>` por "Usuários" (já é) e `trilha` em `app_usuarios.lista` para `[("Administração", url_for("admin.tela")), ("Usuários", None)]`.

- [ ] **Step 6: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: tudo PASS. Testes antigos que ainda esperam o formulário "Abrir evento" em `/inventario` (`tests/test_app.py` ~507, ~886) seguem passando até a Tarefa 5 porque o formulário ainda está lá; o que muda agora são só redirecionamentos e o texto de confirmação: `tests/test_app.py` ~517 (`b"Confirmar encerramento" in r.data` depois do POST em `/encerrar` sem confirmar) passa a seguir o redirect para `/administracao?confirmar=<id>` e a esperar `b"Confirmar finaliza"`; qualquer teste que cheque `r.request.path == f"/inventario/{eid}"` depois de `/abrir`, `/encerrar` ou `/comissao` passa a esperar `/administracao`; o flash "Evento aberto. Comece pelas salas." não existe mais (nenhum teste o cita). Descomentar a asserção marcada `# Tarefa 4` na Tarefa 2.

- [ ] **Step 7: Commit**

```bash
git add app_admin.py app_inventario.py app.py app_usuarios.py permissoes.py menu.py templates/administracao.html templates/_abas_admin.html templates/usuarios/lista.html templates/inventario_comissao.html templates/inventario_excluir.html tests/test_app.py tests/test_permissoes.py tests/test_menu.py tests/test_ajuda.py
git commit -m "feat: tela Administração com chave aberto/fechado dos inventários e aba Usuários"
```

---

### Task 5: Tela Inventário só para conferência, Início, menu e Ajuda

**Files:**
- Modify: `templates/inventario_eventos.html` (remove o formulário; lista com estado; botão Administrar), `templates/inventario_evento.html:6-30` (remove botões administrativos; mensagem de fechado), `templates/index.html:62-66` (card), `templates/ajuda.html:31-44, 59-62`
- Modify: `app.py:200-208` (`home`: evento corrente com estado)
- Test: `tests/test_app.py`, `tests/test_escopo_inventario.py`

**Interfaces:**
- Consumes: `corrente`, `eventos`, `pode_administrar` de `inventario.eventos_tela`; `e.estado` de `inventario.evento`; `inventario.FECHADO`.
- Produces: contexto `inventario_corrente` em `index.html` (substitui `inventario_aberto`).

- [ ] **Step 1: Testes que falham**

Acrescentar ao fim de `tests/test_app.py`:

```python
def test_tela_inventario_sem_administracao_e_com_estado(cliente, dados, usuarios_exemplo):
    beltrana = _ids("Beltrana")[0]
    cliente.post("/inventario/abrir", data={"nome": "Inv A", "usuarios": [beltrana], "escopo": "todas"})
    ea = inventario.evento_aberto(dados)["id"]
    r = cliente.get("/inventario")
    assert b"Abrir evento" not in r.data and b'name="usuarios"' not in r.data and b'href="/administracao"' in r.data
    assert b"Inv A" in r.data and b">aberto<" in r.data
    r = cliente.get(f"/inventario/{ea}")
    assert b"Encerrar evento" not in r.data and b"Excluir evento" not in r.data and b'href="/administracao"' in r.data
    cliente.post(f"/inventario/{ea}/fechar")
    r = cliente.get(f"/inventario/{ea}")
    assert "Inventário fechado".encode() in r.data
    assert b'placeholder="Aproxime o leitor\xe2\x80\xa6" disabled' in cliente.get(f"/inventario/{ea}/sala/01 - SALA CCI").data   # leitura suspensa
    r = cliente.post(f"/inventario/{ea}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    assert r.status_code == 409 and "fechado" in r.get_json()["erro"]
    r = cliente.get("/inventario")
    assert b">fechado<" in r.data                                                      # comissão vê o fechado
    cliente.post("/sair"); logar(cliente, *usuarios_exemplo["inventariante"])
    r = cliente.get("/inventario")
    assert b'href="/administracao"' not in r.data and b"Inv A" in r.data and b">fechado<" in r.data


def test_inicio_mostra_inventario_fechado(cliente, dados, usuarios_exemplo):
    beltrana = _ids("Beltrana")[0]
    cliente.post("/inventario/abrir", data={"nome": "Inv A", "usuarios": [beltrana], "escopo": "todas"})
    ea = inventario.evento_aberto(dados)["id"]
    assert "Inventário em andamento".encode() in cliente.get("/").data
    cliente.post(f"/inventario/{ea}/fechar")
    html = cliente.get("/").get_data(as_text=True)
    assert "Inventário fechado" in html and "Inventário em andamento" not in html
    m = cliente.get("/").data.split(b'id="main-navigation"')[1].split(b"menu-footer")[0]
    assert b"Inv A" in m                                                                 # menu mostra o corrente
```

Ajustar os testes antigos que dependem da UI administrativa na tela Inventário: `tests/test_app.py` ~507 (`b"Abrir evento" in r.data and b"Nenhum evento aberto"` → `b"Nenhum invent" in r.data and b'href="/administracao"' in r.data`), ~886-888 (o GET do formulário passa a ser `cliente.get("/administracao")` e `b"Consulta Teste" not in r.data` vira `b"Consulta Teste" in r.data`, porque qualquer ativo pode entrar na comissão), ~1065/~1083 (`Excluir evento` no evento → verificar `b"Excluir"` em `cliente.get("/administracao")` para admin e ausência de `href="/administracao"` para o não admin), `tests/test_escopo_inventario.py:229` (`"Abrir evento" not in html` → `'href="/administracao"' not in html`). Rodar e ler cada falha antes de mudar: só trocar a expectativa, nunca a lógica testada.

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_app.py -k "sem_administracao or inventario_fechado" -v`
Expected: FAIL (`Abrir evento` ainda presente; "Inventário fechado" ausente).

- [ ] **Step 3: Templates e `home`**

`templates/inventario_eventos.html`, bloco `conteudo` inteiro:

```jinja
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Inventário</h1>{{ ajuda_titulo(AJUDA_ANCORA) }}
  {% if pode_administrar %}<a class="br-button secondary ml-auto" href="{{ url_for('admin.tela') }}"><i class="fas fa-cogs mr-1" aria-hidden="true"></i>Administrar inventários</a>{% endif %}</div>
{% if corrente %}
<div class="br-card mb-4">
  <div class="card-header"><div class="text-weight-semi-bold text-up-02">{{ corrente.nome }}
      {% if corrente.estado == 'fechado' %}<span class="br-tag bg-warning ml-2"><span>fechado</span></span>{% else %}<span class="br-tag bg-success text-pure-0 ml-2"><span>aberto</span></span>{% endif %}</div>
    <div class="text-down-01 text-gray-70">Criado em {{ corrente.aberto_em[8:10] }}/{{ corrente.aberto_em[5:7] }}/{{ corrente.aberto_em[:4] }}{% if corrente.descricao %} · {{ corrente.descricao }}{% endif %} · Comissão: {{ corrente.integrantes|join(', ') }}</div></div>
  <div class="card-content">
    {% set r = corrente.resumo %}
    <div class="row">
      {% for valor, legenda in [(r.salas_iniciadas ~ ' / ' ~ r.salas, 'salas iniciadas'), (r.lidos ~ ' / ' ~ r.bens, 'bens localizados'), (r.divergentes, 'divergentes'), (r.pendentes, 'pendentes'), (r.sobras, 'sobras')] %}
      <div class="col-6 col-md mb-3"><div class="text-up-03 text-weight-bold">{{ valor }}</div><div class="text-down-01 text-gray-70">{{ legenda }}</div></div>
      {% endfor %}
    </div>
    {% if corrente.estado == 'fechado' %}
    <div class="br-message warning mb-3"><div class="icon"><i class="fas fa-pause-circle fa-lg" aria-hidden="true"></i></div>
      <div class="content"><span class="message-title">Inventário fechado.</span><span class="message-body"> A leitura está suspensa até o administrador reabrir; consulta liberada.</span></div></div>
    {% else %}
    <div class="br-message info mb-3"><div class="icon"><i class="fas fa-info-circle fa-lg" aria-hidden="true"></i></div>
      <div class="content"><span class="message-body">{{ r.pct_bens }}% dos bens localizados.</span></div></div>
    {% endif %}
    <a class="br-button primary" href="{{ url_for('inventario.evento_tela', id=corrente.id) }}"><i class="fas fa-door-open mr-1" aria-hidden="true"></i>{% if corrente.estado == 'aberto' and pode('inventario.ler', 'POST') %}Salas e leitura{% else %}Salas{% endif %}</a>
    {% if pode_relatorios %}<a class="br-button secondary ml-2" href="{{ url_for('inventario.relatorio_tela', id=corrente.id) }}"><i class="fas fa-table mr-1" aria-hidden="true"></i>Relatório</a>{% endif %}
  </div>
</div>
{% else %}
<p class="text-gray-70">{% if eventos %}Nenhum inventário aberto ou fechado no momento.{% elif pode_administrar %}Nenhum inventário criado. Crie um em Administrar inventários.{% else %}Nenhum inventário atribuído a você. Quando a comissão de um evento incluir seu usuário, ele aparece aqui.{% endif %}</p>
{% endif %}
{% if eventos %}
{{ cabecalho_tabela('Outros inventários', 'eventos') }}
  <thead><tr><th scope="col">Inventário</th><th scope="col">Estado</th><th scope="col">Criado em</th><th scope="col">Finalizado em</th>{% if pode_relatorios %}<th scope="col" class="dsgov-acoes">Relatório</th>{% endif %}</tr></thead>
  <tbody>{% for e in eventos %}
  <tr><td><a href="{{ url_for('inventario.evento_tela', id=e.id) }}">{{ e.nome }}</a></td>
    <td>{% if e.estado == 'aberto' %}<span class="br-tag bg-success text-pure-0"><span>aberto</span></span>{% elif e.estado == 'fechado' %}<span class="br-tag bg-warning"><span>fechado</span></span>{% else %}<span class="br-tag bg-gray-20"><span>finalizado</span></span>{% endif %}</td>
    <td>{{ e.aberto_em[8:10] }}/{{ e.aberto_em[5:7] }}/{{ e.aberto_em[:4] }}</td><td>{% if e.encerrado_em %}{{ e.encerrado_em[8:10] }}/{{ e.encerrado_em[5:7] }}/{{ e.encerrado_em[:4] }}{% else %}—{% endif %}</td>
    {% if pode_relatorios %}<td class="dsgov-acoes"><a class="br-button circle small" href="{{ url_for('inventario.relatorio_tela', id=e.id) }}" aria-label="Relatório de {{ e.nome }}"><i class="fas fa-table" aria-hidden="true"></i></a></td>{% endif %}</tr>
  {% endfor %}</tbody>
</table></div>
{% endif %}
```

Remover o bloco `scripts` desse template (o JS do escopo foi para `administracao.html`).

`templates/inventario_evento.html`, linhas 6-30: no `<div class="ml-auto">` deixar só Painel, Relatório e, se `pode('admin.tela')`, `<a class="br-button secondary ml-2" href="{{ url_for('admin.tela') }}"><i class="fas fa-cogs mr-1" aria-hidden="true"></i>Administrar</a>`; remover Comissão, Encerrar e Excluir. Depois da mensagem "Evento encerrado", acrescentar:

```jinja
{% elif e.estado == 'fechado' %}
<div class="br-message warning mb-3"><div class="icon"><i class="fas fa-pause-circle fa-lg" aria-hidden="true"></i></div>
  <div class="content"><span class="message-title">Inventário fechado.</span><span class="message-body"> A leitura está suspensa até o administrador reabrir; consulta liberada.</span></div></div>
```

e trocar `{% set conferir = na_comissao and not e.encerrado_em and pode('inventario.ler', 'POST') %}` por `{% set conferir = na_comissao and e.estado == 'aberto' and pode('inventario.ler', 'POST') %}`. Em `templates/inventario_sala.html`, linha 6, `{% set fechado = e.encerrado_em is not none or not na_comissao %}` vira `{% set fechado = e.estado != 'aberto' or not na_comissao %}` (o campo de leitura fica `disabled` com o evento fechado, como já fica para quem não é da comissão); na linha 18, o `{% if e.encerrado_em %}` da mensagem ganha um `{% elif e.estado == 'fechado' %}` com a mesma mensagem "Inventário fechado." usada em `inventario_evento.html`.

`app.py` `home`: `inventario_aberto` vira `inventario_corrente` (evento corrente visível, com `estado`), e `index.html`:

```jinja
    {% if inventario_corrente %}<a class="br-card h-100 dsgov-kpi{% if inventario_corrente.estado == 'fechado' %} dsgov-kpi-alerta{% endif %}" href="{{ url_for('inventario.evento_tela', id=inventario_corrente.id) }}"><div class="card-content">
      <div class="valor">{{ inventario_corrente.resumo.pct_bens }}%</div>
      <div class="text-down-01 text-gray-70">{% if inventario_corrente.estado == 'fechado' %}Inventário fechado{% else %}Inventário em andamento{% endif %} · {{ inventario_corrente.resumo.divergentes }} divergentes</div></div></a>
```

`templates/ajuda.html`: na seção `inventario`, o parágrafo `{% if pode('inventario.abrir','POST') %}...{% endif %}` vira "Evento fechado pelo administrador aceita consulta, mas não leitura, até ser reaberto." (sem `if`). A seção `usuarios` ganha, antes dos parágrafos atuais: `<p>Em <strong>Administração › Inventários</strong> o administrador cria inventários, define a comissão e o escopo de salas, liga e desliga a chave Abrir/Fechar (só um fica aberto; ligar em um fecha o outro), finaliza e exclui. Finalizar congela as leituras e não tem volta. Quem entra na comissão sem a função Inventário recebe a função na hora.</p>`.

- [ ] **Step 4: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: tudo PASS depois dos ajustes de expectativa listados no Step 1.

- [ ] **Step 5: Commit**

```bash
git add templates/inventario_eventos.html templates/inventario_evento.html templates/inventario_sala.html templates/index.html templates/ajuda.html app.py tests/test_app.py tests/test_escopo_inventario.py
git commit -m "feat: tela Inventário só para conferência; Início e Ajuda com inventário fechado"
```

---

### Task 6: README, spec e publicação

**Files:**
- Modify: `README.md` (seção que descreve inventário/usuários; tabela "Arquivos": `app_admin.py`)
- Modify: `docs/superpowers/specs/2026-09-18-administracao-inventarios-design.md` (§10 e **Estado**)

- [ ] **Step 1: README**

Na seção de uso do inventário do README (`grep -n "Inventário\|Usuários" README.md`), descrever em um parágrafo: Administração (menu, só admin) com abas Inventários e Usuários; chave Abrir/Fechar com um aberto por vez; Finalizar permanente; comissão concede a função Inventário. Acrescentar `| app_admin.py | Tela Administração (aba Inventários; a aba Usuários é app_usuarios.py) |` à tabela "Arquivos".

- [ ] **Step 2: Suíte, merge e publicação (controlador, com autorização do usuário)**

```bash
.venv/bin/pytest -q
git checkout main && git merge --ff-only admin-inventarios && git push
docker compose up -d --build
docker compose exec -T web python -c "import db; c=db.conectar(); print(db._colunas(c, 'inventario_eventos'))"
```

Expected: a lista inclui `suspenso_em`; o site responde; o menu do admin mostra "Administração".

- [ ] **Step 3: Evidências**

Na spec, **Estado** → "implementado e publicado em <data>; ver §10" e §10 com: contagem do `pytest -q`, commits, confirmação da coluna no container, e o que foi visto na tela (Administração com o inventário atual, chave, aba Usuários).

```bash
git add README.md docs/superpowers/specs/2026-09-18-administracao-inventarios-design.md
git commit -m "docs: administração de inventários publicada"
git push
```
