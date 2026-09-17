# Fase 1 — Blindagem barata: plano de implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fechar os 14 defeitos confirmados e baratos do `SYSTEM_IMPROVEMENT_PLAN.md` (segredo, contexto Docker, atomicidade, importação parcial, dedup, entradas inválidas, XLSX literal, dois bugs de JS) sem mudar fluxo, telas, esquema ou dependências.

**Architecture:** Flask + SQLite (`db.py` é a fachada de dados; `app.py` rotas principais; `app_inventario.py` blueprint do inventário; `textos.py` dizeres; `config.py` caminhos). Cada tarefa é um teste que reproduz o defeito, a menor mudança que o corrige e um commit. Nenhum módulo novo.

**Tech Stack:** Python 3.12 (web) / 3.14 (local), Flask 3.1, openpyxl 3.1.5, pytest. Suíte atual: 174 testes em ~4 s.

**Spec:** `docs/superpowers/specs/2026-09-16-fase1-blindagem-design.md`

## Global Constraints

- Contratos públicos mantidos: assinaturas de `db.*` e `textos.salvar/restaurar`, nomes de rotas, formatos de XLSX/JSON, `app:app`.
- Nenhuma dependência nova; nenhuma alteração de esquema (`db.ESQUEMA`).
- Comando de teste: `.venv/bin/python -m pytest -q -p no:cacheprovider` (na raiz do repositório). A suíte usa `TERMOS_DADOS` temporário; nunca apontar para `dados/` real.
- Não fazer merge, push, `docker compose up` nem tocar em `dados/termos.db`; a entrega em produção é a Tarefa 12 e depende de autorização do usuário.
- Textos de interface em português, vocabulário existente ("Carregar", "Baixar").
- Mensagens de commit no padrão do repositório: primeira linha em português, descrevendo o efeito.

---

### Tarefa 0: Branch de trabalho

**Files:** nenhum.

- [ ] **Passo 1: Criar o branch a partir de `main`**

```bash
git -C /opt/web/termos-responsabilidade checkout -b fase1-blindagem main
```

- [ ] **Passo 2: Confirmar a baseline**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider`
Expected: `174 passed`

---

### Tarefa 1: `.dockerignore` sem segredos, bancos e planilhas

**Files:**
- Modify: `.dockerignore`

**Interfaces:** nenhuma.

- [ ] **Passo 1: Substituir o conteúdo do `.dockerignore`**

```
.git
.venv
.pytest_cache
.superpowers
__pycache__
dados
dist
build
docs
tests
secrets
*.spec
*.bat
*.env
*.db
*.db-*
*.xlsx
*.gz
*.md
backup.sh
```

- [ ] **Passo 2: Provar que a imagem não leva arquivos locais**

Run:
```bash
cd /opt/web/termos-responsabilidade && docker build -q -t termos-fase1-teste . && docker run --rm --entrypoint sh termos-fase1-teste -c 'ls -a /app; ls /app/secrets 2>&1; ls /app/*.xlsx 2>&1'; docker rmi termos-fase1-teste
```
Expected: `ls -a /app` mostra `app.py`, `templates`, `static`, `timbrado.docx`…; as duas últimas linhas dizem `No such file or directory`. A imagem de teste é removida ao final.

- [ ] **Passo 3: Commit**

```bash
git add .dockerignore
git commit -m "Docker: contexto sem secrets, bancos, planilhas e documentos"
```

---

### Tarefa 2: Segredo de sessão por instalação

**Files:**
- Modify: `config.py` (fim do arquivo)
- Modify: `app.py:22`
- Modify: `tests/test_cadastros_ux.py:1-17`
- Test: `tests/test_config.py`

**Interfaces:**
- Produces: `config.chave_secreta() -> str` (64 caracteres hexadecimais, ou o valor de `TERMOS_SEGREDO`).

- [ ] **Passo 1: Escrever os testes que falham**

Acrescentar ao final de `tests/test_config.py`:

```python
def test_chave_secreta_vem_da_variavel_de_ambiente(tmp_path, monkeypatch):
    monkeypatch.setenv("TERMOS_DADOS", str(tmp_path))
    monkeypatch.setenv("TERMOS_SEGREDO", "segredo-da-web")
    import config
    assert config.chave_secreta() == "segredo-da-web"
    assert not (tmp_path / "segredo.txt").exists()


def test_chave_secreta_e_gerada_uma_vez_por_instalacao(tmp_path, monkeypatch):
    monkeypatch.setenv("TERMOS_DADOS", str(tmp_path / "dados"))   # pasta ainda não existe
    monkeypatch.delenv("TERMOS_SEGREDO", raising=False)
    import config
    chave = config.chave_secreta()
    assert len(chave) == 64 and int(chave, 16) >= 0
    assert (tmp_path / "dados" / "segredo.txt").read_text() == chave
    assert config.chave_secreta() == chave                          # segunda chamada lê o arquivo


def test_chave_secreta_respeita_arquivo_ja_criado_por_outro_processo(tmp_path, monkeypatch):
    monkeypatch.setenv("TERMOS_DADOS", str(tmp_path))
    monkeypatch.delenv("TERMOS_SEGREDO", raising=False)
    (tmp_path / "segredo.txt").write_text("abc")
    import config
    assert config.chave_secreta() == "abc"
```

- [ ] **Passo 2: Rodar e ver falhar**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider tests/test_config.py`
Expected: 3 FAILED com `AttributeError: module 'config' has no attribute 'chave_secreta'`

- [ ] **Passo 3: Implementar em `config.py`**

Acrescentar `import secrets` aos imports e, ao final do arquivo:

```python
def chave_secreta() -> str:
    """Assina cookies e tokens de revisão. TERMOS_SEGREDO (web) ou dados/segredo.txt, criado na primeira
    execução. Dois processos ao mesmo tempo não brigam: O_EXCL garante um único criador; o outro lê."""
    if os.environ.get("TERMOS_SEGREDO"):
        return os.environ["TERMOS_SEGREDO"]
    caminho = pasta_dados() / "segredo.txt"
    if not caminho.exists():
        pasta_dados().mkdir(parents=True, exist_ok=True)
        try:
            fd = os.open(caminho, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        except FileExistsError:
            pass
        else:
            with os.fdopen(fd, "w") as f:
                f.write(secrets.token_hex(32))
    return caminho.read_text().strip()
```

- [ ] **Passo 4: Usar em `app.py`**

Trocar a linha 22:

```python
app.secret_key = "termos-cfc-local"  # sessão só guarda seleção de bens; programa roda em 127.0.0.1
```
por
```python
app.secret_key = config.chave_secreta()   # por instalação: TERMOS_SEGREDO ou dados/segredo.txt
```

- [ ] **Passo 5: Isolar o import de `app` em `tests/test_cadastros_ux.py`**

O import no topo do módulo criaria `dados/segredo.txt` na pasta real durante a coleta. Remover a linha `from app import app` do topo e importar dentro da fixture, como `test_app.py` faz:

```python
@pytest.fixture
def cliente(dados):
    semear(dados)
    from app import app
    app.config['TESTING'] = True
    with app.test_client() as client:
        yield client
```

Se alguma função de teste desse arquivo usar `app` diretamente (procurar com `grep -n "app\." tests/test_cadastros_ux.py`), trocar por `cliente.application`.

- [ ] **Passo 6: Rodar a suíte inteira**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider`
Expected: `177 passed`; nenhum arquivo `dados/segredo.txt` novo na raiz do repositório (`git status --ignored | grep segredo` vazio ou já existente antes).

- [ ] **Passo 7: Commit**

```bash
git add config.py app.py tests/test_config.py tests/test_cadastros_ux.py
git commit -m "Segredo de sessão por instalação (TERMOS_SEGREDO ou dados/segredo.txt) no lugar do literal"
```

---

### Tarefa 3: Salvar pessoa em uma transação só

**Files:**
- Modify: `db.py:640-666` (`salvar_pessoa`, `renomear_pessoa`)
- Test: `tests/test_db.py` (após `test_pessoa_email_matricula_e_salvar`)

**Interfaces:**
- Produces: `db._renomear_pessoa(conn, antigo, novo) -> str` (sem commit). `renomear_pessoa` e `salvar_pessoa` mantêm assinatura e retorno.

- [ ] **Passo 1: Escrever o teste que falha**

```python
def test_salvar_pessoa_e_tudo_ou_nada(dados):
    import sqlite3
    semear(dados)
    dados.execute("CREATE TRIGGER falha BEFORE UPDATE OF email ON pessoas BEGIN SELECT RAISE(ABORT, 'falha simulada'); END")
    with pytest.raises(sqlite3.IntegrityError):
        db.salvar_pessoa(dados, "ANA SILVA", {"nome": "ana souza", "email": "a@cfc", "matricula": "1"})
    outra = db.conectar()   # o que outra conexão enxerga = o que foi commitado
    assert db.pessoa(outra, "ANA SILVA") is not None and db.pessoa(outra, "ANA SOUZA") is None
    assert db.pessoa_do_bem(outra, 1002) == "ANA SILVA"
    outra.close()
```

- [ ] **Passo 2: Rodar e ver falhar**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider tests/test_db.py::test_salvar_pessoa_e_tudo_ou_nada`
Expected: FAIL em `assert db.pessoa(outra, "ANA SILVA") is not None` (o rename já foi commitado).

- [ ] **Passo 3: Implementar**

Substituir `salvar_pessoa` e `renomear_pessoa` em `db.py` por:

```python
def salvar_pessoa(conn, antigo: str, dados: dict) -> str:
    """Renomeia (mantendo atribuições e histórico) e atualiza e-mail e matrícula. Tudo ou nada."""
    try:
        novo = _renomear_pessoa(conn, antigo, dados.get("nome"))
        conn.execute("UPDATE pessoas SET email = ?, matricula = ? WHERE nome = ?",
                     (_texto(dados.get("email")) or None, _texto(dados.get("matricula")) or None, novo))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    return novo


def renomear_pessoa(conn, antigo: str, novo: str) -> str:
    novo = _renomear_pessoa(conn, antigo, novo)
    conn.commit()
    return novo


def _renomear_pessoa(conn, antigo: str, novo: str) -> str:
    """Sem commit: quem chama decide quando a transação termina."""
    novo = _obrigatorio(novo, "Nome").upper()
    if antigo not in pessoas(conn):
        raise ErroDeNegocio(f"Pessoa {antigo} não encontrada.")
    if novo == antigo:
        return novo
    if novo in pessoas(conn):
        raise ErroDeNegocio(f"Já existe uma pessoa chamada {novo}.")
    conn.execute("UPDATE pessoas SET nome = ? WHERE nome = ?", (novo, antigo))  # cascateia em atribuicoes
    conn.execute("UPDATE termos_emitidos SET chave = ? WHERE tipo IN ('individual','devolucao') AND chave = ?", (novo, antigo))
    return novo
```

- [ ] **Passo 4: Rodar a suíte**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider`
Expected: `178 passed`

- [ ] **Passo 5: Commit**

```bash
git add db.py tests/test_db.py
git commit -m "Salvar pessoa: nome, e-mail e matrícula numa transação só"
```

---

### Tarefa 4: Salvar todos os textos em uma transação só

**Files:**
- Modify: `textos.py:157-168`
- Modify: `app.py:420-434` (`textos_salvar`)
- Test: `tests/test_textos.py`

**Interfaces:**
- Produces: `textos.salvar_todos(conn, valores: dict[str, str]) -> None` (valida tudo, grava tudo, um commit). `salvar` e `restaurar` inalterados por fora.

- [ ] **Passo 1: Escrever o teste que falha**

Acrescentar a `tests/test_textos.py`:

```python
def test_salvar_todos_e_tudo_ou_nada(dados):
    import sqlite3
    valores = dict(textos.PADRAO, orgao_nome="Conselho X", cidade="Goiânia (GO)")
    dados.execute("CREATE TRIGGER falha BEFORE INSERT ON textos WHEN (SELECT count(*) FROM textos) >= 1 "
                  "BEGIN SELECT RAISE(ABORT, 'falha simulada'); END")
    with pytest.raises(sqlite3.IntegrityError):
        textos.salvar_todos(dados, valores)
    assert dados.execute("SELECT count(*) FROM textos").fetchone()[0] == 0
    dados.execute("DROP TRIGGER falha")
    textos.salvar_todos(dados, valores)
    t = textos.obter(dados)
    assert t["orgao_nome"] == "Conselho X" and t["cidade"] == "Goiânia (GO)"
    assert dados.execute("SELECT count(*) FROM textos").fetchone()[0] == 2   # só o que difere do padrão
```

(`pytest` já é importado na linha 1 do arquivo.)

- [ ] **Passo 2: Rodar e ver falhar**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider tests/test_textos.py::test_salvar_todos_e_tudo_ou_nada`
Expected: FAIL com `AttributeError: module 'textos' has no attribute 'salvar_todos'`

- [ ] **Passo 3: Implementar em `textos.py`**

Substituir `salvar` e `restaurar` por:

```python
def salvar(conn, chave: str, valor: str) -> None:
    _gravar(conn, chave, valor)
    conn.commit()


def restaurar(conn, chave: str) -> None:
    conn.execute("DELETE FROM textos WHERE chave = ?", (chave,))
    conn.commit()


def salvar_todos(conn, valores: dict) -> None:
    """Tela Textos: valida tudo antes de gravar qualquer coisa e commita uma vez. Tudo ou nada."""
    for chave, valor in valores.items():
        validar(chave, valor)
    try:
        for chave, valor in valores.items():
            _gravar(conn, chave, valor)
        conn.commit()
    except Exception:
        conn.rollback()
        raise


def _gravar(conn, chave: str, valor: str) -> None:
    """Sem commit. Igual ao padrão = apaga a linha (só o que difere fica no banco)."""
    validar(chave, valor)
    if valor == PADRAO[chave]:
        conn.execute("DELETE FROM textos WHERE chave = ?", (chave,))
    else:
        conn.execute("INSERT OR REPLACE INTO textos VALUES (?, ?)", (chave, valor))
```

- [ ] **Passo 4: Usar em `app.py`**

Em `textos_salvar`, trocar

```python
    novos = {c: request.form.get(c, "").replace("\r\n", "\n") for c in textos.PADRAO}
    for c, v in novos.items():
        textos.validar(c, v)          # tudo validado antes de gravar qualquer coisa
    for c, v in novos.items():
        textos.salvar(conn, c, v)
```
por
```python
    novos = {c: request.form.get(c, "").replace("\r\n", "\n") for c in textos.PADRAO}
    textos.salvar_todos(conn, novos)   # valida tudo, grava tudo, um commit
```

- [ ] **Passo 5: Rodar a suíte**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider`
Expected: `179 passed`

- [ ] **Passo 6: Commit**

```bash
git add textos.py app.py tests/test_textos.py
git commit -m "Textos: salvar todos numa transação só"
```

---

### Tarefa 5: Importação de cadastros recusa abas `inv_*` incompletas

**Files:**
- Modify: `db.py:898-925` (`importar_cadastros`, logo após o `if problemas: raise`)
- Test: `tests/test_inventario.py` (após `test_planilha_sem_abas_de_inventario_nao_toca_nas_tabelas`)

**Interfaces:** nenhuma nova.

- [ ] **Passo 1: Escrever o teste que falha**

```python
def test_planilha_com_abas_de_inventario_incompletas_e_recusada(dados, tmp_path):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    from openpyxl import load_workbook
    wb = load_workbook(db.exportar_cadastros(dados, tmp_path / "c.xlsx"))
    for aba in ("inv_integrantes", "inv_salas", "inv_leituras", "inv_sobras"):   # sobra só inv_eventos
        wb.remove(wb[aba])
    wb.save(tmp_path / "parcial.xlsx")
    with open(tmp_path / "parcial.xlsx", "rb") as f, pytest.raises(db.ImportacaoInvalida) as e:
        db.importar_cadastros(dados, f)
    assert "incompletas" in str(e.value) and "inv_leituras" in str(e.value)
    assert inventario.resumo(dados, eid)["lidos"] == 1 and inventario.evento(dados, eid) is not None
```

- [ ] **Passo 2: Rodar e ver falhar**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider tests/test_inventario.py::test_planilha_com_abas_de_inventario_incompletas_e_recusada`
Expected: FAIL com `DID NOT RAISE` (hoje a planilha parcial é aceita e apaga leituras).

- [ ] **Passo 3: Implementar**

Em `db.importar_cadastros`, logo depois do bloco

```python
    if problemas:
        raise ImportacaoInvalida("Planilha de cadastros: " + "; ".join(problemas))
```
acrescentar:
```python
    faltam = [aba for aba, v in inv_brutos.items() if v is None]
    if tem_inventario and faltam:
        raise ImportacaoInvalida("Planilha de cadastros: abas de inventário incompletas (faltam: " + ", ".join(faltam)
                                 + "). Envie as 5 abas inv_* ou nenhuma.")
```

- [ ] **Passo 4: Rodar a suíte**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider`
Expected: `180 passed`

- [ ] **Passo 5: Commit**

```bash
git add db.py tests/test_inventario.py
git commit -m "Importar cadastros: abas inv_* só todas juntas ou nenhuma (nada de apagar inventário por aba faltante)"
```

---

### Tarefa 6: Deduplicação de emissão considera o processo SEI

**Files:**
- Modify: `db.py:759` (`registrar_emissao`)
- Test: `tests/test_db.py` (após `test_registrar_emissao_mesmo_dia_mesma_lista_nao_duplica`)

**Interfaces:** nenhuma nova.

- [ ] **Passo 1: Escrever o teste que falha**

```python
def test_registrar_emissao_com_outro_processo_cria_termo_novo(dados):
    semear(dados)
    db.incluir_processo(dados, "ccusto", "Termos 2026", "2222")
    a = db.registrar_emissao(dados, "ccusto", "CCI", db.bens_do_centro(dados, "CCI"))
    db.incluir_processo(dados, "ccusto", "Termos 2026 (novo)", "3333")   # passa a ser o vigente
    b = db.registrar_emissao(dados, "ccusto", "CCI", db.bens_do_centro(dados, "CCI"))
    assert b["id"] != a["id"] and b["numero_sei"] == "3333" and db.termo_emitido(dados, a["id"])["numero_sei"] == "2222"
    assert len(db.termos_emitidos(dados)) == 2
```

- [ ] **Passo 2: Rodar e ver falhar**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider tests/test_db.py::test_registrar_emissao_com_outro_processo_cria_termo_novo`
Expected: FAIL em `b["id"] != a["id"]`

- [ ] **Passo 3: Implementar**

Em `registrar_emissao`, trocar

```python
    if ultimo and ultimo["emitido_em"][:10] == agora[:10] and _numeros_do_termo(conn, ultimo["id"]) == numeros:
```
por
```python
    if (ultimo and ultimo["emitido_em"][:10] == agora[:10] and ultimo["processo_id"] == proc["id"]
            and _numeros_do_termo(conn, ultimo["id"]) == numeros):
```
e atualizar a docstring: `"""Foto do termo. Sem processo vigente do tipo → ErroDeNegocio. No mesmo dia, com o mesmo processo e a mesma lista de bens, só atualiza a hora do registro existente."""`

- [ ] **Passo 4: Rodar a suíte**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider`
Expected: `181 passed`

- [ ] **Passo 5: Commit**

```bash
git add db.py tests/test_db.py
git commit -m "Emissão: trocar o processo SEI gera termo novo mesmo com a mesma lista no mesmo dia"
```

---

### Tarefa 7: Tipo inválido dá 404, HEAD não registra, teste sem ano fixo

**Files:**
- Modify: `app.py` (`_exigir_processo`, `termo_docx`)
- Test: `tests/test_app.py` (linha 71 e novo teste)

**Interfaces:** nenhuma nova.

- [ ] **Passo 1: Escrever o teste que falha e tirar o ano fixo**

Acrescentar a `tests/test_app.py`:

```python
def test_docx_tipo_invalido_da_404_e_head_nao_registra(cliente):
    assert cliente.get("/termo/xyz/CCI/docx").status_code == 404
    assert cliente.get("/termo/xyz/CCI").status_code == 404
    db.incluir_processo(db.conectar(), "ccusto", "Termos", "1111")
    assert cliente.head("/termo/ccusto/CCI/docx").status_code == 200
    assert db.termos_emitidos(db.conectar()) == []
    assert cliente.get("/termo/ccusto/CCI/docx").status_code == 200
    assert len(db.termos_emitidos(db.conectar())) == 1
```

Na linha 71 do mesmo arquivo, trocar

```python
    assert j["id"] and j["emitido_em"][:4] == "2026"
```
por
```python
    assert j["id"] and j["emitido_em"][:4] == str(date.today().year)
```
e acrescentar `from datetime import date` aos imports do topo.

- [ ] **Passo 2: Rodar e ver falhar**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider tests/test_app.py::test_docx_tipo_invalido_da_404_e_head_nao_registra`
Expected: FAIL: `500 == 404` (KeyError em `ROTULO_TIPO[tipo]`).

- [ ] **Passo 3: Implementar em `app.py`**

`_exigir_processo`:

```python
def _exigir_processo(conn, tipo, chave):
    """Sem processo SEI vigente do tipo não há emissão: flash + volta à tela do termo."""
    if tipo not in db.TIPOS_TERMO:
        abort(404)
    if db.processo_vigente(conn, tipo):
        return None
    flash(f"Cadastre um processo SEI vigente para {db.ROTULO_TIPO[tipo]} em Cadastros → Processos SEI.", "error")
    return redirect(url_for("termo", tipo=tipo, chave=chave))
```

`termo_docx`, logo após `if (volta := _exigir_processo(conn, tipo, chave)): return volta`:

```python
    if request.method == "HEAD":     # navegadores/antivírus sondam o link: não gera nem registra
        return "", 200
```

- [ ] **Passo 4: Rodar a suíte**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider`
Expected: `182 passed`

- [ ] **Passo 5: Commit**

```bash
git add app.py tests/test_app.py
git commit -m "Termo: tipo inválido responde 404 e HEAD no .docx não registra emissão"
```

---

### Tarefa 8: JSON precisa ser objeto; limite de upload de 20 MB

**Files:**
- Modify: `app_inventario.py:116-140` (`ler`, `atualizar_leitura`)
- Modify: `app.py` (após `app.register_blueprint`; após `erro_de_negocio`)
- Test: `tests/test_app.py`

**Interfaces:** nenhuma nova.

- [ ] **Passo 1: Escrever os testes que falham**

Acrescentar a `tests/test_app.py`:

```python
def test_leitura_com_json_que_nao_e_objeto_da_400(cliente):
    eid = _abrir(cliente, integrante="Fulano")
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json=[1, 2])
    assert r.status_code == 400 and "objeto JSON" in r.get_json()["erro"]
    r = cliente.post(f"/inventario/{eid}/leitura/1001", json="texto")
    assert r.status_code == 400 and "objeto JSON" in r.get_json()["erro"]


def test_upload_acima_de_20mb_da_mensagem_e_nao_500(cliente):
    grande = io.BytesIO(b"x" * (20 * 1024 * 1024 + 1))
    r = cliente.post("/upload", data={"arquivo": (grande, "export.xlsx")}, content_type="multipart/form-data",
                     headers={"Referer": "http://localhost/upload"}, follow_redirects=True)
    assert r.status_code == 200 and "Arquivo muito grande".encode() in r.data
```

(`_abrir(cliente, integrante="Fulano")` já existe em `tests/test_app.py:409`: abre um evento com todas as salas e seleciona o integrante.)

- [ ] **Passo 2: Rodar e ver falhar**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider tests/test_app.py -k "json_que_nao_e_objeto or acima_de_20mb"`
Expected: 2 FAILED (`500 == 400`; e o upload grande passa pelo importador em vez de ser recusado).

- [ ] **Passo 3: Implementar em `app_inventario.py`**

Acrescentar, perto de `_json_erro`:

```python
def _corpo_json() -> dict | None:
    """Corpo JSON das rotas de leitura: precisa ser um objeto; qualquer outra coisa é 400."""
    dados = request.get_json(silent=True)
    return dados if isinstance(dados, dict) else None
```

Em `ler`, trocar
```python
    numero = _numero_lido((request.get_json(silent=True) or {}).get("numero"))
```
por
```python
    dados = _corpo_json()
    if dados is None:
        return jsonify({"erro": "Envie um objeto JSON."}), 400
    numero = _numero_lido(dados.get("numero"))
```

Em `atualizar_leitura`, trocar
```python
    dados = request.get_json(silent=True) or {}
```
por
```python
    dados = _corpo_json()
    if dados is None:
        return jsonify({"erro": "Envie um objeto JSON."}), 400
```

- [ ] **Passo 4: Implementar em `app.py`**

Após `app.register_blueprint(inventario_bp)`:

```python
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024   # mesmo limite do nginx (client_max_body_size 20m)
```

Após o `erro_de_negocio`:

```python
@app.errorhandler(413)
def arquivo_grande(_e):
    flash("Arquivo muito grande: o limite é 20 MB.", "error")
    return redirect(request.referrer or url_for("home"))
```

- [ ] **Passo 5: Rodar a suíte**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider`
Expected: `184 passed`

- [ ] **Passo 6: Commit**

```bash
git add app.py app_inventario.py tests/test_app.py
git commit -m "Entradas inválidas: JSON não-objeto responde 400; upload acima de 20 MB avisa em vez de falhar"
```

---

### Tarefa 9: Texto literal nas planilhas exportadas

**Files:**
- Modify: `db.py` (novo helper perto de `exportar_cadastros`; `exportar_cadastros`, `exportar_recorte`, `exportar_bens`)
- Modify: `inventario.py` (`exportar_xlsx`, `exportar_abas`)
- Modify: `Termo_de_Responsabilidade.py` (`gerar_planilha_centro`)
- Test: `tests/test_db.py`

**Interfaces:**
- Produces: `db.acrescentar_linha(ws, valores: list) -> None` — `ws.append` que grava strings como texto mesmo quando começam com `=`.

- [ ] **Passo 1: Escrever o teste que falha**

Acrescentar a `tests/test_db.py`:

```python
def test_exportar_grava_texto_literal_e_numeros_como_numeros(dados, tmp_path):
    from openpyxl import load_workbook
    semear(dados)
    dados.execute("INSERT INTO bens VALUES (1005,'ATIVO','=1+1','=SOMA(A1)','MÓVEIS','01 - SALA CCI','01/01/2020',10,9)")
    dados.commit()
    wb = load_workbook(db.exportar_bens(dados, tmp_path / "bens.xlsx"))
    linha = [c for c in wb["base"].iter_rows(min_row=2) if c[0].value == 1005][0]
    assert linha[2].value == "=1+1" and linha[2].data_type == "s"       # texto, não fórmula
    assert linha[3].value == "=SOMA(A1)" and linha[3].data_type == "s"
    assert linha[0].data_type == "n" and linha[8].data_type == "n"       # número continua número
    wb = load_workbook(db.exportar_recorte(dados, {}, tmp_path / "recorte.xlsx"))
    assert all(c.data_type != "f" for linha in wb["recorte"].iter_rows(min_row=2) for c in linha)
```

(`exportar_recorte` aceita `{}` como filtro: `recorte` só usa `f.get(...)`.)

- [ ] **Passo 2: Rodar e ver falhar**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider tests/test_db.py::test_exportar_grava_texto_literal_e_numeros_como_numeros`
Expected: FAIL em `linha[2].data_type == "s"` (hoje é `"f"`).

- [ ] **Passo 3: Implementar o helper em `db.py`**

Antes de `exportar_cadastros`:

```python
def acrescentar_linha(ws, valores: list) -> None:
    """ws.append que grava texto como texto: '=1+1' numa descrição vira fórmula no openpyxl e o Excel
    tenta calcular. Números e datas passam como estão."""
    ws.append(valores)
    for celula in ws[ws.max_row]:
        if isinstance(celula.value, str) and celula.data_type == "f":
            celula.data_type = "s"
```

- [ ] **Passo 4: Trocar `ws.append(...)` de linhas de dados pelo helper**

Em `db.py`: `exportar_cadastros` (`ws.append(list(linha))`), `exportar_recorte` (o `ws.append([b["numero"], ...])`), `exportar_bens` (`ws.append(list(linha))`) → `acrescentar_linha(ws, ...)`. Os cabeçalhos podem continuar com `ws.append`.

Em `inventario.py`: acrescentar `acrescentar_linha` ao `from db import ...` da linha 7; em `exportar_xlsx` trocar os dois `ws.append([x[...` e `ws2.append([s[...` por `acrescentar_linha(ws, [...])` / `acrescentar_linha(ws2, [...])`; em `exportar_abas` trocar `ws.append(list(linha))` por `acrescentar_linha(ws, list(linha))`.

Em `Termo_de_Responsabilidade.py`: `import db` (ao lado de `import config`) e, em `gerar_planilha_centro`, trocar `ws.append([b["numero"], ...])` por `db.acrescentar_linha(ws, [b["numero"], b["descricao"], b["complemento"], b["localizacao"], b["valor_atual"]])`.

Conferir que não sobrou linha de dados sem o helper: `grep -n "ws2\?\.append" db.py inventario.py Termo_de_Responsabilidade.py` deve listar só cabeçalhos e títulos.

- [ ] **Passo 5: Rodar a suíte**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider`
Expected: `185 passed`

- [ ] **Passo 6: Commit**

```bash
git add db.py inventario.py Termo_de_Responsabilidade.py tests/test_db.py
git commit -m "Planilhas: texto que começa com '=' sai como texto, não como fórmula"
```

---

### Tarefa 10: Formulário de sobra fora do `.br-card`; Copiar só registra se copiou

**Files:**
- Modify: `templates/inventario_sala.html:71-86, 97, 181-182`
- Modify: `templates/termo.html:24-27, 38-60`
- Test: `tests/test_app.py` (asserção de marcação) + verificação manual no navegador

**Interfaces:** nenhuma.

- [ ] **Passo 1: Escrever o teste de marcação que falha**

Acrescentar a `tests/test_app.py`:

```python
def test_form_sobra_nao_e_o_proprio_br_card(cliente):
    eid = _abrir(cliente, integrante="Fulano")
    html = cliente.get(f"/inventario/{eid}/sala/01 - SALA CCI").get_data(as_text=True)
    assert 'id="card-sobra"' in html and 'id="form-sobra"' in html
    assert 'class="br-card mt-3" id="form-sobra"' not in html   # BRCard trocaria o id do form
```

- [ ] **Passo 2: Rodar e ver falhar**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider tests/test_app.py::test_form_sobra_nao_e_o_proprio_br_card`
Expected: FAIL em `'id="card-sobra"' in html`

- [ ] **Passo 3: Alterar `templates/inventario_sala.html`**

Linha 71, trocar
```html
<form method="post" action="{{ url_for('inventario.sobra', id=e.id, localizacao=localizacao) }}" enctype="multipart/form-data" class="br-card mt-3" id="form-sobra" hidden><div class="card-content">
```
por
```html
<div class="br-card mt-3" id="card-sobra" hidden><form method="post" action="{{ url_for('inventario.sobra', id=e.id, localizacao=localizacao) }}" enctype="multipart/form-data" id="form-sobra"><div class="card-content">
```
e o fechamento (linha 86) `</div></form>` por `</div></form></div>`.

Linha 97: `formSobra = document.getElementById("form-sobra")` → `cardSobra = document.getElementById("card-sobra"), formSobra = document.getElementById("form-sobra")`.

Linhas 181-182:
```js
  btnSobra.addEventListener("click", function () { if (cardSobra) { cardSobra.hidden = false; formSobra.querySelector("#s-descricao").focus(); } });
  var btnCancelar = document.getElementById("btn-cancelar-sobra"); if (btnCancelar) btnCancelar.addEventListener("click", function () { cardSobra.hidden = true; focar(); });
```

- [ ] **Passo 4: Alterar `templates/termo.html`**

Aviso (linhas 24-27): dar id ao título e trocar a classe conforme o resultado:
```html
<div id="aviso-copiado" class="br-message success" hidden>
  <div class="icon"><i class="fas fa-check-circle fa-lg" aria-hidden="true"></i></div>
  <div class="content" role="alert"><span class="message-title" id="aviso-titulo">Copiado.</span><span class="message-body" id="aviso-texto"> Cole no editor do SEI (Ctrl+V).</span></div>
</div>
```

Script do botão Copiar (substituir o bloco inteiro):
```js
document.getElementById("copiar").addEventListener("click", async function () {
  var corpo = document.getElementById("documento").contentDocument.body;
  var html = corpo.innerHTML, texto = corpo.innerText, copiou = false;
  var aviso = document.getElementById("aviso-copiado"), tituloAviso = document.getElementById("aviso-titulo"), textoAviso = document.getElementById("aviso-texto");
  try {
    await navigator.clipboard.write([new ClipboardItem({
      "text/html": new Blob([html], {type: "text/html"}),
      "text/plain": new Blob([texto], {type: "text/plain"})})]);
    copiou = true;
  } catch (e) {
    var docIframe = corpo.ownerDocument, sel = docIframe.defaultView.getSelection(), range = docIframe.createRange();
    range.selectNodeContents(corpo); sel.removeAllRanges(); sel.addRange(range);
    copiou = docIframe.execCommand("copy") === true; sel.removeAllRanges();
  }
  aviso.className = "br-message " + (copiou ? "success" : "danger");
  tituloAviso.textContent = copiou ? "Copiado." : "Não copiado.";
  if (!copiou) {
    textoAviso.textContent = " Não foi possível copiar. Selecione o texto do documento e copie com Ctrl+C.";
    aviso.hidden = false;
    setTimeout(function () { aviso.hidden = true; }, 8000);
    return;   // sem cópia não há emissão a registrar
  }
  textoAviso.textContent = " Cole no editor do SEI (Ctrl+V).";
  try {
    var r = await fetch(this.dataset.registrar, {method: "POST"});
    var j = await r.json();
    textoAviso.textContent = r.ok ? " Cole no editor do SEI (Ctrl+V). Emissão registrada às " + j.emitido_em.slice(11, 16) + "." : " " + j.erro;
  } catch (e) { textoAviso.textContent += " (não foi possível registrar a emissão)"; }
  aviso.hidden = false;
  setTimeout(function () { aviso.hidden = true; }, 5000);
});
```

- [ ] **Passo 5: Rodar a suíte**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider`
Expected: `186 passed`

- [ ] **Passo 6: Verificar no navegador (sem teste automatizado de JS)**

Subir uma instância descartável: `TERMOS_DADOS=/tmp/fase1-dados TERMOS_PORTA=12399 .venv/bin/python main.py` (cai no modo navegador na VPS), carregar um export do SPW e um inventário de teste. Conferir:
1. Sala do inventário: ler um número inexistente → botão "Sobra" abre o formulário; "Cancelar" fecha. Console sem erro.
2. Tela de um termo com processo: "Copiar para o SEI" mostra "Copiado." e "Emissão registrada às hh:mm"; colar num editor traz o documento.
3. Negar permissão de área de transferência no navegador (ou testar em `http://` sem `localhost`, onde a Clipboard API não existe) e clicar de novo: mensagem "Não copiado." e nenhum registro novo em Termos emitidos.
Parar a instância e apagar `/tmp/fase1-dados`.

- [ ] **Passo 7: Commit**

```bash
git add templates/inventario_sala.html templates/termo.html tests/test_app.py
git commit -m "Sobra: formulário fora do .br-card (BRCard trocava o id); Copiar só registra emissão se a cópia deu certo"
```

---

### Tarefa 11: README e nota de entrega

**Files:**
- Modify: `README.md:51-52`
- Modify: `docs/superpowers/specs/2026-09-16-fase1-blindagem-design.md` (linha "Estado")

- [ ] **Passo 1: Corrigir o README**

Trocar
```
6. **Planilha de cadastros**: Cadastros → *Exportar cadastros* gera `cadastros.xlsx` (4 abas). Edite no
   Excel e importe em *Atualizar base → Importar cadastros* — substitui as 4 tabelas inteiras.
```
por
```
6. **Planilha de cadastros**: Cadastros → *Exportar cadastros* gera `cadastros.xlsx` (4 abas; quando já
   existe inventário, mais 5 abas `inv_*`). Edite no Excel e importe em *Atualizar base → Importar
   cadastros* — substitui as 4 tabelas inteiras; as abas `inv_*` só são aceitas todas juntas (e então
   substituem o inventário inteiro) ou nenhuma (inventário preservado).
```

Acrescentar, na seção do README que descreve `dados/` (procurar `grep -n "dados/" README.md`), uma linha: "`dados/segredo.txt` — chave que assina a sessão do navegador, criada na primeira execução; na web pode vir da variável `TERMOS_SEGREDO`."

- [ ] **Passo 2: Marcar a spec como implementada**

Na spec, trocar `Estado: aprovado em conversa` por `Estado: implementado no branch fase1-blindagem em <data>; aguardando merge e publicação`.

- [ ] **Passo 3: Suíte completa e commit**

Run: `.venv/bin/python -m pytest -q -p no:cacheprovider`
Expected: `186 passed`

```bash
git add README.md docs/superpowers/specs/2026-09-16-fase1-blindagem-design.md
git commit -m "README: abas inv_* tudo ou nada e dados/segredo.txt"
```

---

### Tarefa 12: Entrega (somente com autorização do usuário)

- [ ] **Passo 1: Revisão final do branch** — `git log --oneline main..fase1-blindagem` (11 commits) e `git diff main --stat`.
- [ ] **Passo 2: Merge e push** — `git checkout main && git merge --ff-only fase1-blindagem && git push`.
- [ ] **Passo 3: Publicar na VPS** — `docker compose up -d --build` em `/opt/web/termos-responsabilidade`; conferir `docker logs termos-patrimonio --tail 20` e que `dados/segredo.txt` foi criado. Avisar o usuário que a sessão do navegador foi invalidada (refazer seleção de devolução em andamento, se houver).
- [ ] **Passo 4: Desktop** — o próximo `build.bat` no Windows já leva as mudanças; `dados/segredo.txt` nasce ao lado do `.exe` na primeira abertura.
