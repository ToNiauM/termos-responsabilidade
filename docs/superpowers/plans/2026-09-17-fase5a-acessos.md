# Fase 5A — funções e autorização Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Conceder funções cumulativas e impedir acesso fora das funções e da comissão do usuário.

**Architecture:** Matriz pura em `permissoes.py`, persistência de funções em `usuarios.py`, migração versionada em `migracoes_acesso.py` e escopo de inventário em `comissoes.py`. As rotas usam a política antes de chamar os serviços existentes. Nomes históricos permanecem dados de apresentação; autorização por evento usa IDs.

**Tech Stack:** Python, Flask, SQLite, Jinja/DSGov e pytest.

---

**Spec:** `../specs/2026-09-17-fase5-acessos-design.md`.
**Base:** `c0dd564`. Executar antes de 5B/5C, sem publicação intermediária.

## Arquivos e responsabilidades

| Arquivo | Responsabilidade |
|---|---|
| `permissoes.py` (novo) | funções, rótulos, regras por endpoint/método |
| `migracoes_acesso.py` (novo) | conversão única de perfis e vínculos legados |
| `comissoes.py` (novo) | vínculo por ID, visibilidade de eventos e escopo das fotos |
| `usuarios.py` | contas e funções; proteção do admin; hidratação de usuários |
| `db.py` | esquema e chamada de migração |
| `app.py`, `app_usuarios.py`, `app_inventario.py` | autorização e formulários |
| `inventario.py` | preservar vínculos em importação; serviços de conferência existentes |
| `templates/base.html`, `templates/usuarios/*.html`, `templates/inventario*.html`, `templates/index.html` | funções, controles autorizados e entrada |
| `tests/test_funcoes.py`, `tests/test_migracoes_acesso.py`, `tests/test_escopo_inventario.py` (novos) | novas invariantes |
| testes existentes de usuários, permissões, login e inventário | atualizar contrato antigo e provar regressões |

## Tarefa 1 — matriz explícita por função e método

**Create:** `permissoes.py`, `tests/test_funcoes.py`.

- [ ] Escrever o teste da união sem escalada:

```python
import permissoes

def test_funcoes_somam_sem_conceder_outras():
    f = {"inventariante", "consulta"}
    assert permissoes.permitido(f, "bem")
    assert permissoes.permitido(f, "inventario.ler", "POST")
    assert not permissoes.permitido(f, "inventario.xlsx")
    assert not permissoes.permitido(f, "termo_docx")
    assert not permissoes.permitido({"operador"}, "inventario.ler", "POST")
    assert permissoes.permitido({"consulta_inventarios"}, "inventario.xlsx")
    assert not permissoes.permitido({"consulta_inventarios"}, "bem")
    assert not permissoes.permitido({"admin"}, "rota_inexistente")
    assert not permissoes.permitido({"consulta"}, "termo_devolucao", "POST")
    assert permissoes.permitido({"admin"}, "termo_docx", "HEAD")
    assert not permissoes.permitido({"admin"}, "bem", "DELETE")
```

- [ ] Rodar `.venv/bin/python -m pytest tests/test_funcoes.py -q`; esperado: falha porque o módulo não existe.
- [ ] Criar a matriz completa:

```python
FUNCOES = ("admin", "operador", "consulta", "inventariante", "consulta_inventarios")
ROTULOS = dict(zip(FUNCOES, ("Administrador", "Operador", "Consulta", "Inventário", "Consulta de inventários")))
TODAS = frozenset(FUNCOES)
ACERVO = frozenset({"admin", "operador", "consulta"})
GESTAO = frozenset({"admin", "operador"})
ADMIN = frozenset({"admin"})
CONFERENCIA = frozenset({"admin", "inventariante"})
EVENTOS = frozenset({"admin", "inventariante", "consulta_inventarios"})
RELATORIOS = frozenset({"admin", "consulta_inventarios"})
PERMISSOES = {}

def _registrar(funcoes, linhas):
    for linha in linhas.strip().splitlines():
        endpoint, *metodos = linha.split()
        for metodo in metodos:
            chave = (endpoint, metodo)
            if chave in PERMISSOES:
                raise ValueError(f"Permissão repetida: {chave}")
            PERMISSOES[chave] = funcoes

_registrar(TODAS, """
home GET
ajuda GET
usuarios.login GET POST
usuarios.sair POST
usuarios.senha GET POST
""")
_registrar(ACERVO, """
bem GET
pesquisa GET
recorte GET
recorte_xlsx GET
analise GET
analise_xlsx GET
centro_custos GET
termos_individuais GET
termo GET
termo_documento GET
termo_devolucao GET
termos_emitidos_tela GET
termo_emitido_tela GET
""")
_registrar(GESTAO, """
gerar POST
gerar_individual POST
termo_docx GET
termo_planilha GET
termo_registrar POST
termo_devolucao POST
termo_emitido_documento POST
termo_emitido_email POST
cadastros GET
cadastro_novo GET
responsaveis_incluir POST
responsaveis_editar GET POST
pessoas_incluir POST
pessoas_editar GET POST
pessoas_atribuir POST
pessoas_desatribuir POST
localizacoes_incluir POST
localizacoes_alterar GET
localizacoes_mover POST
processos_incluir POST
processos_vigente POST
processos_encerrar POST
cadastros_exportar GET
textos_tela GET
textos_salvar POST
upload GET POST
bens_exportar GET
importacao_tela GET
""")
_registrar(ADMIN, """
responsaveis_excluir POST
pessoas_excluir POST
localizacoes_excluir POST
processos_excluir POST
importar_cadastros POST
inventario.abrir POST
inventario.encerrar POST
inventario.comissao GET POST
inventario.excluir GET POST
usuarios.lista GET
usuarios.novo GET
usuarios.incluir POST
usuarios.editar GET POST
usuarios.nova_senha POST
""")
_registrar(EVENTOS, """
inventario.eventos_tela GET
inventario.evento_tela GET
inventario.sala_tela GET
""")
_registrar(RELATORIOS, """
inventario.relatorio_tela GET
inventario.painel_tela GET
inventario.xlsx GET
""")
_registrar(CONFERENCIA, """
inventario.ler POST
inventario.atualizar_leitura POST
inventario.lote POST
inventario.foto_leitura POST
inventario.foto_excluir POST
inventario.sobra POST
inventario.sobra_excluir POST
""")

def permitido(funcoes, endpoint, metodo="GET"):
    if isinstance(funcoes, str):
        return False
    concedidas = frozenset(funcoes or ())
    if not concedidas <= TODAS:
        return False
    metodo = "GET" if metodo == "HEAD" else metodo
    return bool(concedidas & PERMISSOES.get((endpoint, metodo), frozenset()))
```

- [ ] Acrescentar parametrização para todas as combinações das cinco funções: o resultado é a união das funções isoladas; incluir conjunto vazio e função desconhecida.
- [ ] Rodar o teste novamente; esperado: passa.
- [ ] Commit: `git add permissoes.py tests/test_funcoes.py` e `git commit -m "feat: define cumulative access functions"`.

## Tarefa 2 — migração única e esquema normalizado

**Create:** `migracoes_acesso.py`, `tests/test_migracoes_acesso.py`.
**Modify:** `db.py:ESQUEMA`, `db.py:criar_esquema`.

- [ ] Criar teste que monta banco legado independente das fixtures atuais:

```python
import sqlite3
from migracoes_acesso import migrar

def test_migracao_nao_reconcede_funcao_nem_homonimo():
    c = sqlite3.connect(":memory:")
    c.executescript("""
      PRAGMA foreign_keys=ON;
      CREATE TABLE usuarios(id INTEGER PRIMARY KEY, nome TEXT, ativo INTEGER,
        perfil TEXT CHECK(perfil IN ('admin','operador','consulta','inventariante')));
      CREATE TABLE inventario_eventos(id INTEGER PRIMARY KEY);
      CREATE TABLE inventario_integrantes(evento_id INTEGER, nome TEXT);
      INSERT INTO usuarios VALUES(1,'Ana',1,'inventariante'),(2,'Ana',1,'inventariante'),
        (3,'Beto',1,'inventariante'),(4,'Cris',1,'operador');
      INSERT INTO inventario_eventos VALUES(1);
      INSERT INTO inventario_integrantes VALUES(1,'Ana'),(1,'Beto'),(1,'Cris');
    """)
    migrar(c)
    assert c.execute("SELECT usuario_id FROM inventario_comissao_usuarios").fetchall() == [(3,)]
    assert c.execute("SELECT funcao FROM usuarios_funcoes WHERE usuario_id=4").fetchall() == [('operador',)]
    c.execute("DELETE FROM usuarios_funcoes WHERE usuario_id=3")
    c.execute("DELETE FROM inventario_comissao_usuarios WHERE usuario_id=3")
    c.commit()
    migrar(c)
    assert not c.execute("SELECT 1 FROM usuarios_funcoes WHERE usuario_id=3").fetchone()
    assert not c.execute("SELECT 1 FROM inventario_comissao_usuarios").fetchone()
    assert 'perfil' not in [r[1] for r in c.execute('PRAGMA table_info(usuarios)')]
```

- [ ] Rodar `.venv/bin/python -m pytest tests/test_migracoes_acesso.py -q`; esperado: falha por módulo ausente.
- [ ] Implementar a migração sem `executescript` dentro da transação:

```python
def migrar(conn):
    with conn:
        conn.execute('BEGIN IMMEDIATE')
        conn.execute("CREATE TABLE IF NOT EXISTS migracoes_acesso (nome TEXT PRIMARY KEY)")
        conn.execute("""CREATE TABLE IF NOT EXISTS usuarios_funcoes (
          usuario_id INTEGER NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
          funcao TEXT NOT NULL CHECK(funcao IN
            ('admin','operador','consulta','inventariante','consulta_inventarios')),
          PRIMARY KEY(usuario_id, funcao))""")
        conn.execute("""CREATE TABLE IF NOT EXISTS inventario_comissao_usuarios (
          evento_id INTEGER NOT NULL REFERENCES inventario_eventos(id) ON DELETE CASCADE,
          usuario_id INTEGER NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
          nome_na_comissao TEXT NOT NULL,
          PRIMARY KEY(evento_id, usuario_id))""")
        if conn.execute("SELECT 1 FROM migracoes_acesso WHERE nome='fase5_funcoes'").fetchone():
            return
        colunas = {r[1] for r in conn.execute('PRAGMA table_info(usuarios)')}
        if 'perfil' in colunas:
            conn.execute("INSERT INTO usuarios_funcoes SELECT id, perfil FROM usuarios")
            conn.execute('ALTER TABLE usuarios DROP COLUMN perfil')
        conn.execute("""INSERT INTO inventario_comissao_usuarios
          SELECT i.evento_id, u.id, i.nome
          FROM inventario_integrantes i JOIN usuarios u ON u.nome=i.nome
          WHERE u.ativo=1
            AND (SELECT count(*) FROM usuarios x WHERE x.nome=i.nome)=1
            AND EXISTS(SELECT 1 FROM usuarios_funcoes f WHERE f.usuario_id=u.id
                       AND f.funcao IN ('admin','inventariante'))""")
        conn.execute("INSERT INTO migracoes_acesso VALUES ('fase5_funcoes')")
```

- [ ] Retirar a linha `perfil TEXT ...` de `ESQUEMA`. Em `criar_esquema`, concluir as migrações anteriores e então chamar a migração 5A, que reserva sua própria transação antes de qualquer DDL:

```python
from migracoes_acesso import migrar
conn.commit()
migrar(conn)
```

- [ ] Acrescentar testes de banco novo, rollback por erro injetado e preservação de hash/IDs em tabela legada completa. O teste de rollback força uma exceção no INSERT da comissão e verifica que a coluna `perfil` e os dados legados continuam, sem marcador de sucesso.
- [ ] Rodar `.venv/bin/python -m pytest tests/test_migracoes_acesso.py tests/test_db.py -q`; esperado: passa. Não usar `dados/termos.db`.
- [ ] Commit: `git add migracoes_acesso.py db.py tests/test_migracoes_acesso.py` e `git commit -m "feat: migrate user functions and commission identities"`.

## Tarefa 3 — contratos de usuários, transação e proteção do administrador

**Modify:** `usuarios.py`, `tests/test_usuarios.py`, `tests/conftest.py`.

- [ ] Escrever teste com duas funções e remoção de admin mantendo outra função:

```python
import pytest
import db
import usuarios

def test_editar_funcoes_e_proteger_ultimo_admin(dados):
    uid = usuarios.criar(dados, 'ana', 'Ana', 'Senha!234', ['admin', 'consulta'])
    assert set(usuarios.por_id(dados, uid)['funcoes']) == {'admin', 'consulta'}
    with pytest.raises(db.ErroDeNegocio, match='último administrador'):
        usuarios.editar(dados, uid, 'Ana', ['consulta'], ativo=True)
    outro = usuarios.criar(dados, 'beto', 'Beto', 'Senha!234', ['admin'])
    usuarios.editar(dados, uid, 'Ana', ['inventariante', 'consulta'], True, logado_id=outro)
    assert set(usuarios.por_id(dados, uid)['funcoes']) == {'inventariante', 'consulta'}
    with pytest.raises(db.ErroDeNegocio):
        usuarios.editar(dados, uid, 'Ana', [], True, logado_id=outro)
```

- [ ] Rodar esse teste; esperado: falha no contrato antigo.
- [ ] Importar/reexportar `FUNCOES`, `ROTULOS`, `PERMISSOES`, `permitido` de `permissoes`. Remover `PERFIS`, `PERFIS_COMISSAO`, `ROTULO_PERFIL`, `_validar_perfil` e a matriz antiga. Retirar `perfil` de `_COLUNAS_LISTA`; trocar o usuário local por `funcoes=('admin',)`.
- [ ] Implementar validação e hidratação:

```python
def _validar_funcoes(funcoes):
    if isinstance(funcoes, str):
        raise ErroDeNegocio('Selecione as funções do usuário.')
    valores = set(funcoes or ())
    if not valores or not valores <= set(FUNCOES):
        raise ErroDeNegocio('Selecione ao menos uma função válida.')
    return tuple(f for f in FUNCOES if f in valores)

def _com_funcoes(conn, u):
    if u is None:
        return None
    atribuida = {r[0] for r in conn.execute(
        'SELECT funcao FROM usuarios_funcoes WHERE usuario_id=?', (u['id'],))}
    return dict(u, funcoes=tuple(f for f in FUNCOES if f in atribuida))

def _gravar_funcoes(conn, uid, funcoes):
    conn.execute('DELETE FROM usuarios_funcoes WHERE usuario_id=?', (uid,))
    conn.executemany('INSERT INTO usuarios_funcoes VALUES (?,?)', [(uid, f) for f in funcoes])

def por_id(conn, id):
    return _com_funcoes(conn, _um(conn, 'SELECT * FROM usuarios WHERE id=?', id)) if id is not None else None

def por_login(conn, login):
    return _com_funcoes(conn, _um(conn, 'SELECT * FROM usuarios WHERE login=?', str(login or '').strip().lower()))

def por_email(conn, email):
    v = str(email or '').strip().lower()
    return _com_funcoes(conn, _um(conn, 'SELECT * FROM usuarios WHERE email=?', v)) if v else None

def criar(conn, login, nome, senha, funcoes, trocar_senha=True, email=None):
    login, nome = _validar_login(login), _obrigatorio(nome, 'Nome')
    senha, funcoes = _validar_senha(senha), _validar_funcoes(funcoes)
    email = _validar_email(email)
    _email_livre(conn, email)
    if por_login(conn, login):
        raise ErroDeNegocio(f'O login {login} já existe.')
    with conn:
        cur = conn.execute('''INSERT INTO usuarios
          (login,email,nome,senha_hash,trocar_senha,criado_em) VALUES (?,?,?,?,?,?)''',
          (login,email,nome,generate_password_hash(senha),int(bool(trocar_senha)),_agora()))
        _gravar_funcoes(conn, cur.lastrowid, funcoes)
    return cur.lastrowid

def _admins_ativos(conn):
    return conn.execute('''SELECT count(*) FROM usuarios u WHERE u.ativo=1
      AND EXISTS(SELECT 1 FROM usuarios_funcoes f WHERE f.usuario_id=u.id AND f.funcao='admin')''').fetchone()[0]

def editar(conn, id, nome, funcoes, ativo, logado_id=None, email=_MANTER):
    funcoes = _validar_funcoes(funcoes)
    nome = _obrigatorio(nome, 'Nome')
    # Reserva a escrita antes de contar admins: duas requisições não podem remover o último.
    conn.execute('BEGIN IMMEDIATE')
    try:
        u = por_id(conn, id)
        if not u:
            raise ErroDeNegocio('Usuário não encontrado.')
        perde_admin = 'admin' in u['funcoes'] and u['ativo'] and ('admin' not in funcoes or not ativo)
        if perde_admin and logado_id is not None and int(logado_id) == int(id):
            raise ErroDeNegocio('Você não pode rebaixar nem inativar a própria conta.')
        if perde_admin and _admins_ativos(conn) <= 1:
            raise ErroDeNegocio('Este é o último administrador ativo: não pode ser rebaixado nem inativado.')
        email = u['email'] if email is _MANTER else _validar_email(email)
        _email_livre(conn, email, excluir_id=id)
        conn.execute('UPDATE usuarios SET nome=?,ativo=?,email=? WHERE id=?', (nome,int(bool(ativo)),email,id))
        _gravar_funcoes(conn, id, funcoes)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
```

`editar` exige conexão sem transação pendente; rotas e fixtures devem confirmar suas
semeaduras antes de chamá-la. A semântica de transação deve ser documentada na função.

- [ ] Implementar listagem e elegibilidade por presença de função:

```python
def listar(conn, busca='', funcao=None, inativos=False):
    sql, p = f'SELECT {_COLUNAS_LISTA} FROM usuarios WHERE 1=1', []
    if not inativos:
        sql += ' AND ativo=1'
    if funcao:
        sql += ' AND EXISTS(SELECT 1 FROM usuarios_funcoes f WHERE f.usuario_id=usuarios.id AND f.funcao=?)'
        p.append(funcao)
    if busca.strip():
        sql += " AND (lower(login) LIKE ? OR lower(nome) LIKE ? OR lower(coalesce(email,'')) LIKE ?)"
        p += [f'%{busca.strip().lower()}%'] * 3
    return [_com_funcoes(conn, u) for u in _todos(conn, sql+' ORDER BY login', *p)]

def elegiveis_comissao(conn):
    return sorted((u for u in listar(conn) if set(u['funcoes']) & {'admin','inventariante'}),
                  key=lambda u: (u['nome'],u['login']))
```

- [ ] Em `criar_admin`, criação usa `['admin']`; atualização remove `perfil='admin'` do SQL e faz `_gravar_funcoes(conn, u['id'], ['admin'])` na mesma transação que reativa e troca a senha. Não alterar autenticação, CSRF ou hash.
- [ ] Atualizar fixtures e chamadas de `usuarios.criar/editar`: o argumento antes denominado `perfil` vira coleção. Nos testes, `"admin"` vira `["admin"]`, e verificações `u['perfil']=='admin'` viram `u['funcoes']==('admin',)`. Atualizar também casos inválidos para coleções inválidas, sem relaxar a validação.
- [ ] Rodar `.venv/bin/python -m pytest tests/test_usuarios.py tests/test_funcoes.py tests/test_migracoes_acesso.py -q`; esperado: passa, inclusive CLI.
- [ ] Commit: `git add usuarios.py tests/test_usuarios.py tests/test_funcoes.py tests/conftest.py` e `git commit -m "feat: persist multiple functions per user"`.

## Tarefa 4 — escopo por ID e ciclo de vida da comissão

**Create:** `comissoes.py`, `tests/test_escopo_inventario.py`.
**Modify:** `app_inventario.py`, `inventario.py`, `app_usuarios.py`, templates da comissão.

- [ ] Testar homônimos antes de ligar o guard:

```python
import comissoes
import inventario
import usuarios
from tests.conftest import semear

def test_escopo_por_id_e_nao_por_nome(dados):
    semear(dados)
    a = usuarios.criar(dados,'ana1','Ana','Senha!234',['inventariante'])
    b = usuarios.criar(dados,'ana2','Ana','Senha!234',['inventariante'])
    eid = inventario.abrir_evento(dados,'Evento','',['Ana'])
    comissoes.definir(dados,eid,[a])
    assert comissoes.visivel(dados,usuarios.por_id(dados,a),eid)
    assert not comissoes.visivel(dados,usuarios.por_id(dados,b),eid)
    assert not comissoes.pode_conferir(dados,usuarios.por_id(dados,b),eid)
```

- [ ] Rodar esse teste; esperado: módulo ausente.
- [ ] Criar o módulo com operações pequenas:

```python
import db
import permissoes

def membro(conn, u, eid):
    if u['id'] is None:
        return 'admin' in u['funcoes'] and bool(conn.execute(
            'SELECT 1 FROM inventario_integrantes WHERE evento_id=? AND nome=?', (eid,u['nome'])).fetchone())
    return bool(conn.execute('SELECT 1 FROM inventario_comissao_usuarios WHERE evento_id=? AND usuario_id=?',
                             (eid,u['id'])).fetchone())

def visivel(conn, u, eid):
    f = set(u['funcoes'])
    return bool(f & {'admin','consulta_inventarios'}) or ('inventariante' in f and membro(conn,u,eid))

def pode_conferir(conn, u, eid):
    return bool(set(u['funcoes']) & permissoes.CONFERENCIA) and membro(conn,u,eid)

def eventos_visiveis(conn, u):
    import inventario
    return [e for e in inventario.eventos(conn) if visivel(conn,u,e['id'])]

def _usuarios_selecionados(conn, ids):
    import usuarios
    try:
        ids = sorted({int(i) for i in ids})
    except (TypeError, ValueError):
        raise db.ErroDeNegocio('Selecione integrantes válidos.')
    elegiveis = {u['id']: u for u in usuarios.elegiveis_comissao(conn)}
    if not ids or not set(ids) <= set(elegiveis):
        raise db.ErroDeNegocio('Selecione ao menos um usuário ativo com função de inventário.')
    return [elegiveis[i] for i in ids]

def definir(conn, eid, ids):
    import inventario
    with conn:
        conn.execute('BEGIN IMMEDIATE')
        inventario._evento_aberto_ou_erro(conn,eid)
        selecionados = _usuarios_selecionados(conn,ids)
        conn.execute('DELETE FROM inventario_comissao_usuarios WHERE evento_id=?',(eid,))
        conn.executemany('INSERT INTO inventario_comissao_usuarios VALUES (?,?,?)',
                         [(eid,u['id'],u['nome']) for u in selecionados])
        conn.execute('DELETE FROM inventario_integrantes WHERE evento_id=?',(eid,))
        conn.executemany('INSERT INTO inventario_integrantes VALUES (?,?)',
                         [(eid,n) for n in sorted({u['nome'] for u in selecionados})])

def abrir(conn, nome, descricao, ids, salas):
    import inventario
    selecionados = _usuarios_selecionados(conn,ids)
    # abrir_evento realiza validações e commit; falha posterior fecha a transação
    # externa somente se o serviço aceitar confirmar=False (adaptação abaixo).
    with conn:
        eid = inventario.abrir_evento(conn,nome,descricao,[u['nome'] for u in selecionados],
                                     salas,confirmar=False)
        conn.executemany('INSERT INTO inventario_comissao_usuarios VALUES (?,?,?)',
                         [(eid,u['id'],u['nome']) for u in selecionados])
    return eid

def atualizar_nome(conn, uid):
    # Só comissão aberta: histórico de leituras/sobras e eventos fechados não muda.
    u = conn.execute('SELECT nome FROM usuarios WHERE id=?',(uid,)).fetchone()
    eventos = [r[0] for r in conn.execute('''SELECT c.evento_id FROM inventario_comissao_usuarios c
      JOIN inventario_eventos e ON e.id=c.evento_id WHERE c.usuario_id=? AND e.encerrado_em IS NULL''',(uid,))]
    with conn:
        for eid in eventos:
            vinculados = {r[0] for r in conn.execute('SELECT nome_na_comissao FROM inventario_comissao_usuarios WHERE evento_id=?',(eid,))}
            nomes = {r[0] for r in conn.execute('SELECT nome FROM inventario_integrantes WHERE evento_id=?',(eid,))}
            legado = nomes-vinculados
            conn.execute('UPDATE inventario_comissao_usuarios SET nome_na_comissao=? WHERE evento_id=? AND usuario_id=?',(u[0],eid,uid))
            atuais = {r[0] for r in conn.execute('SELECT nome_na_comissao FROM inventario_comissao_usuarios WHERE evento_id=?',(eid,))}
            conn.execute('DELETE FROM inventario_integrantes WHERE evento_id=?',(eid,))
            conn.executemany('INSERT INTO inventario_integrantes VALUES (?,?)',[(eid,n) for n in sorted(legado|atuais)])
```

- [ ] Em `inventario.abrir_evento`, acrescentar o argumento final `confirmar=True`; substituir somente seu `conn.commit()` por `if confirmar: conn.commit()`. O caminho local mantém a chamada antiga; o web chama `comissoes.abrir`. Testar rollback injetando falha no vínculo depois de inserir evento.
- [ ] Em `app_inventario.abrir/comissao`, formulários web enviam `usuarios` com IDs (`request.form.getlist('usuarios')`); chamar `comissoes.abrir/definir`. Desktop continua enviando nomes e inclui o Administrador local. Toda ação de comissão continua admin-only. Em `_na_comissao` e `_exigir_comissao`, usar `comissoes.pode_conferir` em vez de comparação por nome.
- [ ] Nos dois templates de comissão, usar a marcação web abaixo; o ramo local conserva os campos por nome:

```jinja
{% for u in elegiveis %}
<div class="br-checkbox mb-2">
  <input id="integrante-{{ u.id }}" name="usuarios" type="checkbox" value="{{ u.id }}"{% if u.id in selecionados %} checked{% endif %}/>
  <label for="integrante-{{ u.id }}">{{ u.nome }} ({{ u.login }})</label>
</div>
{% endfor %}
{% if sem_vinculo %}
<div class="br-message warning"><div class="content" role="alert">
  <span class="message-body">Integrantes anteriores sem conta vinculada: {{ sem_vinculo|join(', ') }}. Selecione acima as contas que poderão conferir os bens.</span>
</div></div>
{% endif %}
```

Fornecer `selecionados` pela tabela de vínculos e `sem_vinculo` pela diferença entre
nomes legados e `nome_na_comissao`. No formulário de abertura ambos são vazios.
Não aceitar IDs ocultos ou nomes enviados fora dos usuários elegíveis.

- [ ] Substituir a chamada de `inventario.renomear_integrante` em `app_usuarios.editar` por `comissoes.atualizar_nome(conn,id)` depois de salvar. O nome anterior não identifica uma conta para autorização.
- [ ] Preservar vínculos em `inventario.substituir_tabelas`: antes dos DELETEs, capturar `(evento_id,usuario_id,nome_na_comissao,nome,aberto_em)` com JOIN dos eventos; depois dos INSERTs, repor apenas com identidade de evento intacta e nome ainda na comissão. Código do filtro:

```python
novos = {r[0]:(r[1],r[3]) for r in linhas['inv_eventos']}
nomes = set(linhas['inv_integrantes'])
preservados = [(eid,uid,nome_comissao) for eid,uid,nome_comissao,nome_evento,aberto_em in vinculos
               if novos.get(eid)==(nome_evento,aberto_em) and (eid,nome_comissao) in nomes]
conn.executemany('INSERT INTO inventario_comissao_usuarios VALUES (?,?,?)',preservados)
```

Esse código integra a transação existente de importação; não chama migração nem
associa novos nomes automaticamente. Verificar a ordem das colunas de `inv_eventos`
em `ABAS` (`id,nome,descricao,aberto_em,encerrado_em`) antes de aplicar.

- [ ] Acrescentar testes de renomeação com homônimos, revogação de função, troca de comissão, evento fechado, importação que reutiliza ID e exclusão em cascata. Confirmar que leituras antigas mantêm `integrante` textual.
- [ ] Rodar `.venv/bin/python -m pytest tests/test_escopo_inventario.py tests/test_inventario.py -q`; esperado: passa, atualizando apenas testes que montam comissões web pelo contrato novo.
- [ ] Commit explícito dos arquivos desta tarefa com mensagem `feat: bind inventory access to user identities`.

## Tarefa 5 — guarda de rotas, destino inicial e controles visíveis

**Modify:** `app.py`, `app_usuarios.py`, `app_inventario.py`, templates de usuários, inventário, Início e base; `tests/test_permissoes.py`, `tests/test_login.py`, `tests/test_app.py`, `tests/test_usuarios_telas.py`.

- [ ] Escrever teste de entrada e de URL direta proibida:

```python
from tests.conftest import logar, SENHA_PADRAO

def test_inventariante_entra_so_no_inventario(cliente, monkeypatch):
    import app as web
    def painel_proibido(*args, **kwargs):
        raise AssertionError('não calcular painel geral')
    monkeypatch.setattr(web.db,'painel',painel_proibido)
    cliente.post('/sair')
    assert logar(cliente,'beltrana',SENHA_PADRAO).headers['Location']=='/inventario'
    assert cliente.get('/').headers['Location']=='/inventario'
    for rota in ('/bem?numero=1001','/pesquisa?q=CADEIRA','/recorte','/recorte/xlsx','/centro-custos'):
        assert cliente.get(rota).status_code==403
    html = cliente.get('/inventario').get_data(as_text=True)
    assert 'Nenhum inventário atribuído a você' in html
    assert 'action="/pesquisa"' not in html
```

- [ ] Trocar a autorização geral para `usuarios.permitido(u['funcoes'], ep, metodo)`. No modo local usar `USUARIO_LOCAL['funcoes']`. Para OPTIONS, verificar que pelo menos um método da `request.url_rule.methods - {'HEAD','OPTIONS'}` está permitido; Flask continua produzindo a resposta automática.
- [ ] Em `app.py`, acrescentar `import comissoes`, `import permissoes` e importar `destino_inicial` de `app_usuarios` junto ao blueprint. Em `app_usuarios.py` e `app_inventario.py`, acrescentar `import comissoes` para os usos desta etapa.
- [ ] Depois de identificar o usuário e autorizar a rota, aplicar o escopo antes de executar views:

```python
if request.blueprint == 'inventario' and 'id' in (request.view_args or {}):
    eid = request.view_args['id']
    if not comissoes.visivel(obter_conn(), g.usuario, eid):
        return _negado()
    if request.endpoint in {
        'inventario.ler','inventario.atualizar_leitura','inventario.lote',
        'inventario.foto_leitura','inventario.foto_excluir','inventario.sobra','inventario.sobra_excluir',
    } and not comissoes.pode_conferir(obter_conn(), g.usuario, eid):
        return _negado()
```

Mensagem HTML/JSON: **“Seu usuário não tem permissão para esta ação.”** Ajustar
testes que esperavam “Seu perfil”. O código acima vale também para o modo local:
não retornar cedo após verificar admin local, pois isso pularia a regra da comissão.

- [ ] Adicionar em `app_usuarios.py` o destino inicial e a validação do retorno; importar `MethodNotAllowed`, `NotFound` e `RequestRedirect` nos tipos de exceção usados:

```python
def destino_inicial(u):
    return url_for('home' if set(u['funcoes']) & {'admin','operador','consulta'} else 'inventario.eventos_tela')

def _proximo_autorizado(valor, u):
    from flask import current_app
    from werkzeug.exceptions import MethodNotAllowed, NotFound
    from werkzeug.routing import RequestRedirect
    import comissoes
    seguro = _proximo_seguro(valor)
    try:
        ep, args = current_app.url_map.bind_to_environ(request.environ).match(urlsplit(seguro).path,method='GET')
    except (NotFound, MethodNotAllowed, RequestRedirect):
        return destino_inicial(u)
    if ep in {'usuarios.login','usuarios.sair','static'} or not usuarios.permitido(u['funcoes'],ep):
        return destino_inicial(u)
    if ep.startswith('inventario.') and 'id' in args and not comissoes.visivel(_conn(),u,args['id']):
        return destino_inicial(u)
    return destino_inicial(u) if ep=='home' else seguro
```

`_proximo_seguro` continua rejeitando URLs externas, barras invertidas e controles.
Em login autenticado, login após POST e troca de senha, usar `destino_inicial` ou
`_proximo_autorizado` conforme exista `proximo`. No início de `home`, antes de `db.painel`,
se o destino não for `url_for('home')`, redirecionar. Expor `URL_INICIAL` no contexto.

- [ ] Em `contexto_dsgov`, `pode` passa a usar `usuario['funcoes']`. Listar o evento aberto somente entre `comissoes.eventos_visiveis`. Separar permissão dos filhos de inventário: Eventos/salas conforme EVENTOS; Painel/Relatório somente RELATORIOS. Mostrar Início apenas a ACERVO. Esta etapa conserva o formato atual de MENU; a árvore final vem em 5C.
- [ ] Em `base.html`, trocar `ROTULO_PERFIL[USUARIO.perfil]` por `USUARIO.funcoes|map('rotulo_funcao')|join(', ')`. Aplicar a mesma apresentação aos rótulos de integrantes, inclusive no ramo desktop. Retirar `ROTULO_PERFIL` do contexto; os formulários/listas recebem `funcoes=usuarios.FUNCOES` e `rotulos=usuarios.ROTULOS`. Registrar o filtro em `app.py`:

```python
app.add_template_filter(permissoes.ROTULOS.__getitem__, 'rotulo_funcao')
```

Proteger gatilho e formulário de pesquisa com `pode('pesquisa')`; breadcrumb inicial usa `URL_INICIAL`.
- [ ] Em `app_inventario.eventos_tela`, usar apenas `comissoes.eventos_visiveis`; calcular resumo somente para eventos visíveis. Dados para criar evento (`salas_ativas`, `elegiveis`) são obtidos somente para admin. Expor `pode_abrir`, `pode_relatorios` e, quando não houver eventos visíveis, mostrar o estado vazio sem aludir a evento alheio.
- [ ] Proteger formulário de abertura, botões de painel/relatório/Excel e links da lista de eventos por `pode`. Em `inventario_sala.html`, `fechado` deve incluir `not pode('inventario.ler','POST')` além de evento encerrado e ausência na comissão. Na tela de evento, usar **Ver sala** para consulta e **Conferir sala** quando escrita estiver habilitada.
- [ ] A ficha do bem filtra fotos no servidor antes de renderizar. `inventario.fotos_do_bem` já devolve `evento_id` em cada grupo. Aplicar:

```python
grupos = inventario.fotos_do_bem(conn, numero)
grupos = [grupo for grupo in grupos if comissoes.visivel(conn,g.usuario,grupo['evento_id'])]
```

- [ ] Em `templates/index.html`, aplicar `pode` aos atalhos e aos KPIs de inventário/importação já nesta etapa para evitar links negados. O cálculo de evento aberto em `home` deve respeitar `eventos_visiveis`; manter o restante do layout até 5C.
- [ ] Formulário de usuário: usar `request.form.getlist('funcoes')` em criação/edição; repassar a coleção em erros e na senha temporária. Nenhum campo `perfil` persistido. Listagem usa `funcao` e filtro pela presença. Trocar radios pelo fieldset:

```jinja
<fieldset class="mb-3">
  <legend class="text-weight-semi-bold">Funções do usuário</legend>
  <p class="text-down-01 text-gray-70">Selecione somente as funções necessárias. Os acessos selecionados se somam.</p>
  {% for f in funcoes %}<div class="br-checkbox mb-1">
    <input id="funcao-{{ f }}" name="funcoes" type="checkbox" value="{{ f }}"{% if f in valores.funcoes %} checked{% endif %}/>
    <label for="funcao-{{ f }}">{{ rotulos[f] }}</label>
  </div>{% endfor %}
</fieldset>
```

Validar coleção não vazia no servidor; não marcar `required` em cada checkbox.
Lista e cabeçalho exibem todas as funções; filtro usa macro DSGov `select` existente.
`usuarios.novo` fornece `funcoes=[]`, e o reset fornece as funções atuais sem mudá-las.

- [ ] Atualizar cobertura de rotas para chaves `(endpoint,metodo)`, sem fallback. Ajustar semeaduras web de eventos para IDs; não conceder Consulta a inventariante para manter expectativas antigas.
- [ ] Em `tests/test_usuarios_telas.py`, trocar os payloads `perfil` por `funcoes` em lista e o filtro `?perfil=` por `?funcao=`. Testar criação sem função pré-marcada, preservação da seleção após erro e concessão múltipla:

```python
def test_criar_com_varias_funcoes_e_filtrar_sem_duplicar(cliente,dados):
    r=cliente.post('/usuarios/incluir',data={
        'login':'multiplas','nome':'Múltiplas','senha':'Senha!234','confirmacao':'Senha!234',
        'funcoes':['inventariante','consulta'],
    })
    assert r.status_code==302
    assert set(usuarios.por_login(dados,'multiplas')['funcoes'])=={'inventariante','consulta'}
    html=cliente.get('/usuarios?funcao=consulta').get_data(as_text=True)
    assert html.count('<code>multiplas</code>')==1
    r=cliente.post('/usuarios/incluir',data={
        'login':'semfuncao','nome':'Sem função','senha':'Senha!234','confirmacao':'Senha!234',
    })
    assert r.status_code==200
    assert usuarios.por_login(dados,'semfuncao') is None
```

- [ ] Em `tests/test_app.py`, atualizar o teste de comissão: Operador sozinho não aparece entre elegíveis; um POST de leitura desse usuário retorna 403 por falta da função, em vez do 409 antigo. POSTs web usam IDs e os testes de desktop conservam o ramo por nome. As negativas de vínculo retornam 403; evento fechado conserva o erro de negócio.
- [ ] Rodar `.venv/bin/python -m pytest tests/test_login.py tests/test_permissoes.py tests/test_app.py tests/test_escopo_inventario.py tests/test_usuarios.py tests/test_usuarios_telas.py -q`; esperado: passa.
- [ ] Commit dos arquivos desta tarefa com mensagem `feat: enforce cumulative access across routes and screens`.

## Tarefa 6 — testes de regressão de autorização e entrega de 5A

**Modify:** `tests/test_escopo_inventario.py`, `tests/test_permissoes.py`, `README.md`.

- [ ] Acrescentar caso de revogação com sessão existente:

```python
def test_revogacao_usa_banco_na_proxima_chamada(cliente,dados):
    from tests.conftest import logar, SENHA_PADRAO
    import usuarios
    uid = usuarios.por_login(dados,'beltrana')['id']
    usuarios.editar(dados,uid,'Beltrana',['inventariante','consulta'],True)
    cliente.post('/sair'); logar(cliente,'beltrana',SENHA_PADRAO)
    assert cliente.get('/bem?numero=1001').status_code==200
    usuarios.editar(dados,uid,'Beltrana',['inventariante'],True)
    assert cliente.get('/bem?numero=1001').status_code==403
```

- [ ] Parametrizar todas as mutações do inventário e todos os relatórios: usuário só Consulta, inventariante de outra comissão, homônimo sem vínculo e função retirada recebem 403. Espionar serviços e storage; nenhum é chamado na negativa.
- [ ] Testar Consulta de inventários em evento antigo e em evento criado depois da concessão; GET/HEAD do Excel permitido, POST de leitura negado. Combinação com Inventário só permite escrita no próprio evento.
- [ ] Testar login `proximo` externo, não permitido, evento alheio e destino válido; testar troca obrigatória de senha e modo desktop. Sem evento atribuído, nenhuma contagem/nome alheio aparece no HTML.
- [ ] Testar queda de função sem invalidar dados históricos: nenhum DELETE em leituras, sobras, fotos ou termos.
- [ ] Atualizar README com as cinco funções, combinações e migração restrita; substituir instruções de quatro perfis. Descrever que criar usuário não o inclui automaticamente na comissão.
- [ ] Rodar `.venv/bin/python -m pytest tests/test_funcoes.py tests/test_migracoes_acesso.py tests/test_usuarios.py tests/test_usuarios_telas.py tests/test_login.py tests/test_permissoes.py tests/test_escopo_inventario.py tests/test_inventario.py tests/test_app.py -q`. Esperado: todos passam. Suite completa fica para a integração final de 5C.
- [ ] Conferir `git diff --check` e commit `test: cover least privilege and role combinations` com arquivos explícitos.

**Saída da etapa:** permissões funcionando nas telas atuais. A implementação da
Análise e do menu em árvore segue nos planos 5B e 5C. Nenhuma publicação faz parte
desta etapa.
