# Fase 4: usuários, senha e perfis (painel administrativo no Flask)

**Data:** 2026-09-17
**Estado:** desenho aprovado pelo usuário em 2026-09-17, seção a seção (AskUserQuestion).
**Base:** `main` em `eccd1e6` (Fase 3 publicada).
**Specs anteriores:** `2026-09-17-fase3-fotos-menu-bem-design.md`, `2026-09-16-fase1-blindagem-design.md`.

## 1. Objetivo e decisões

O sistema não tem login próprio. Na web (`patrimonio.sistemascfc.org`) a proteção é o `auth_basic` do nginx
com um usuário compartilhado; no programa Windows não há proteção. O inventário identifica quem lê por um
nome escolhido numa lista digitada ao abrir o evento, sem senha.

O usuário pediu "um painel administrativo ao estilo Django, sem migrar para Django": usuários com senha
para quem opera o sistema ou faz inventário. Decisões dele:

- **Escopo**: só usuários e permissões. Nenhum CRUD genérico de tabelas, nenhum log de ações. Os cadastros
  que existem continuam onde estão.
- **Quatro perfis**: administrador, operador, inventariante, consulta (matriz na seção 3).
- **Integrante do inventário = usuário logado.** A comissão do evento é escolhida pelo administrador entre
  os usuários cadastrados; só quem está na comissão lê. O passo "escolha o integrante" some.
- **Só o administrador** abre e encerra evento, escolhe e edita a comissão, altera perfil de usuário,
  exclui registros (inclusive evento de inventário inteiro, novidade desta fase) e importa planilhas que
  substituem tabelas. Toda exclusão tem trava.
- **Login só no site.** O programa Windows continua entrando direto, como administrador local.
- **Senha esquecida**: o administrador gera uma senha temporária, mostrada uma vez; a pessoa é obrigada a
  trocar no primeiro login. Não há e-mail.
- **Abordagem A**: login próprio, sem biblioteca nova. Tabela no SQLite, `werkzeug.security` para o hash
  (já vem com o Flask), sessão assinada que o Flask já usa, permissões por rota num módulo pequeno.
  Flask-Login descartado (dependência para ganhar `current_user` e "lembrar-me", que não queremos).
  Vários `auth_basic` por caminho no nginx descartado (não identifica quem lê nem tem perfil).
- **CSRF entra agora.** Recusado na fase 1 porque o `auth_basic` bastava; com login por cookie em site
  público, um POST disparado por outro site passaria a valer. Token na sessão, campo oculto nos
  formulários, cabeçalho nos POSTs via `fetch`.
- **Termos emitidos continuam sem exclusão** (decisão de 2026-09-15 mantida: histórico é auditoria).

Fora desta fase: log de quem fez o quê (colunas `emitido_por`, `importado_por`), "lembrar-me", redefinição
de senha por e-mail, política de senha além do mínimo, login no programa Windows, CRUD genérico de tabelas,
exclusão de termos emitidos e de importações.

---

## 2. Dados (`db.py`)

Tabela nova em `ESQUEMA`, criada por `criar_esquema` como as demais (o programa Windows migra sozinho na
primeira abertura; na VPS, `db.inicializar()` no `CMD` do container):

```sql
CREATE TABLE IF NOT EXISTS usuarios (
  id            INTEGER PRIMARY KEY,
  login         TEXT NOT NULL UNIQUE,        -- curto, minúsculo, sem espaço: [a-z0-9._-]{2,30}
  nome          TEXT NOT NULL,               -- como aparece no cabeçalho e nas leituras do inventário
  senha_hash    TEXT NOT NULL,               -- werkzeug.security.generate_password_hash (scrypt)
  perfil        TEXT NOT NULL CHECK (perfil IN ('admin','operador','inventariante','consulta')),
  ativo         INTEGER NOT NULL DEFAULT 1,
  trocar_senha  INTEGER NOT NULL DEFAULT 0,  -- 1 = obrigado a trocar no próximo acesso
  falhas        INTEGER NOT NULL DEFAULT 0,  -- tentativas erradas seguidas
  bloqueado_ate TEXT,                        -- ISO; 5 falhas = 15 minutos
  criado_em     TEXT NOT NULL,
  ultimo_acesso TEXT
);
```

Regras:

- **Login** é imutável depois de criado; nome, perfil e ativo mudam. `login` é normalizado para minúsculas
  e validado pela expressão acima; `nome` é obrigatório, espaços colapsados, sem exigência de unicidade
  (mas ver §5 sobre o nome na comissão).
- **Não se exclui usuário, só se inativa**: o nome dele está em `inventario_leituras.integrante` e
  `inventario_sobras.integrante`. Inativo não entra, não aparece para compor comissão e some da lista
  padrão de usuários (filtro "mostrar inativos").
- **Último administrador ativo** não pode ser inativado nem ter o perfil rebaixado (`ErroDeNegocio`).
  O próprio usuário logado não pode se inativar nem se rebaixar (mesma trava, mensagem própria).
- `inventario_integrantes(evento_id, nome)` **não muda**: continua guardando o `nome` do usuário. Leituras
  antigas ficam como estão.

---

## 3. Permissões (`usuarios.py`)

Nega por padrão: toda rota exige login, exceto `usuarios.login` e `static`.
A matriz é um dicionário `PERMISSOES` de `endpoint → perfis que podem`, com `metodo` quando GET e POST
diferem. Uma rota que não estiver no dicionário devolve 403 para todo mundo, inclusive admin, e o teste da
seção 10 falha, o que impede rota nova sem regra.

```python
def permitido(perfil: str, endpoint: str, metodo: str) -> bool
```

Resumo por área ("ver" = abrir tela e baixar `.xlsx`; "emitir" = copiar ou baixar termo, o que registra
emissão):

| Área | admin | operador | inventariante | consulta |
|---|---|---|---|---|
| Início, Pesquisa, ficha do bem, Recorte (e `.xlsx`) | tudo | tudo | ver | ver |
| Termo por centro, individual, devolução; Termos emitidos | tudo | tudo | não | ver, sem emitir |
| Termos emitidos: anotar documento SEI, registrar e-mail | sim | sim | não | não |
| Cadastros: ver, incluir, editar, atribuir, mover localizações | sim | sim | não | não |
| Cadastros: excluir centro, pessoa, localização, processo SEI | sim | não | não | não |
| Textos | sim | sim | não | não |
| Atualizar base (export do SPW), exportar bens, exportar cadastros | sim | sim | não | não |
| Importar cadastros (`cadastros.xlsx`, substitui tabelas e inventário) | sim | não | não | não |
| Inventário: eventos, evento, sala (ver), painel, relatório, `.xlsx` | ver | ver | ver | ver |
| Inventário: abrir evento, comissão, encerrar, excluir evento | sim | não | não | não |
| Inventário: ler, lote, desmarcar, editar leitura, foto, sobra | se na comissão | se na comissão | se na comissão | não |
| Usuários (listar, criar, editar, perfil, ativar/inativar, nova senha) | sim | não | não | não |
| Trocar a própria senha, sair | sim | sim | sim | sim |

Notas:

- Para o **consulta**, "ver sem emitir" significa: `centro_custos`, `termos_individuais`, `termo_devolucao`
  (GET), `termo`, `termos_emitidos_tela`, `termo_emitido` abrem; `gerar`, `gerar_individual`,
  `termo_documento`, `termo_docx`, `termo_planilha`, `termo_registrar`, e os POSTs de `termo_devolucao`
  devolvem 403. Na tela do termo os botões "Copiar para o SEI" e "Baixar .docx" não aparecem para ele
  (mesma condição no template).
- **Comissão** é checagem adicional em `inventario.ler`, `ler_lote`, `registrar_sobra`, edição de leitura e
  fotos: além do perfil, `g.usuario["nome"]` tem que estar em `inventario_integrantes` do evento. A checagem
  que já existe ("Escolha o integrante da comissão antes de ler") vira "Você não faz parte da comissão deste
  evento". Ela vale para o admin também.
- O **menu** (`contexto_dsgov`) mostra só itens cujo endpoint o perfil pode abrir. "Usuários" entra no menu
  para admin, ícone `fa-users`, depois de "Textos".
- **403** renderiza `403.html` na própria tela (com menu), mensagem "Seu perfil não tem acesso a isso", sem
  redirecionar. Requisições JSON (leitor de plaquetas) recebem `{"erro": ...}` com 403.

---

## 4. Entrar, sair, senha (`app_usuarios.py`, blueprint `usuarios`)

- **`GET /login`**: tela DSGov sem menu nem cabeçalho de navegação (template próprio, `login.html`), campos
  usuário e senha, parâmetro `proximo`. Se a tabela `usuarios` estiver vazia, a tela mostra "Nenhum usuário
  cadastrado" e o comando da seção 8 em vez do formulário. Usuário já logado é redirecionado para `/`.
- **`POST /login`**: `usuarios.autenticar(conn, login, senha)`:
  1. Usuário inexistente, inativo ou senha errada → mesma mensagem "Usuário ou senha inválidos"; se o
     usuário existe, `falhas += 1`; na 5ª, `bloqueado_ate = agora + 15 min`.
  2. `bloqueado_ate` no futuro → "Muitas tentativas. Aguarde 15 minutos." mesmo com a senha certa; não
     incrementa falhas.
  3. Acerto → `falhas = 0`, `bloqueado_ate = NULL`, `ultimo_acesso = agora`; `session.clear()` seguido de
     `session["usuario_id"] = id` (troca de identidade não herda `bens_selecionados` etc.);
     `session.permanent = True`; redireciona para `proximo` se for caminho relativo (`/...`, sem `//` nem
     esquema), senão para `/`.
- **Sessão**: `PERMANENT_SESSION_LIFETIME = 12 h`, `SESSION_REFRESH_EACH_REQUEST = True` (padrão do Flask,
  renova a cada requisição), `SESSION_COOKIE_HTTPONLY = True`, `SESSION_COOKIE_SAMESITE = "Lax"`,
  `SESSION_COOKIE_SECURE` quando `TERMOS_LOGIN` está ligado (o site é só https). No desktop nada disso
  importa: não há login.
- **`POST /sair`**: `session.clear()`, redireciona para `/login`. Botão "Sair" no bloco `header-login` do
  DSGov, com o nome do usuário, o perfil por extenso e o link "Trocar senha" (para `/senha`). (POST, não GET, para não sair por link acidental e
  para ficar coberto pelo CSRF.)
- **`GET/POST /senha`**: senha atual, nova, confirmação. Regras: nova com no mínimo 8 caracteres, diferente
  da atual, confirmação igual. Sucesso grava o hash e `trocar_senha = 0`. Enquanto `trocar_senha = 1`, o
  `before_request` redireciona qualquer outra rota (exceto `sair`, `static`) para `/senha` com aviso
  "Defina uma nova senha para continuar".
- **Usuários** (`/usuarios`, só admin):
  - Lista com busca por login/nome, filtro de perfil e "mostrar inativos"; colunas login, nome, perfil,
    situação, último acesso, ações (Editar).
  - `GET /usuarios/novo`, `POST /usuarios/incluir`: login, nome, perfil, senha inicial (digitada duas vezes,
    mínimo 8) e caixa "obrigar troca no primeiro acesso" (marcada por padrão).
  - `GET/POST /usuarios/<id>/editar`: nome, perfil, ativo. Login exibido, não editável.
  - `POST /usuarios/<id>/nova-senha`: gera 10 caracteres de `ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789`
    (`secrets.choice`; sem 0/O, 1/l/I), grava o hash, `trocar_senha = 1`, `falhas = 0`,
    `bloqueado_ate = NULL`; a resposta renderiza a tela de edição com a senha em claro dentro de um
    `br-message` de aviso ("Anote agora; não será mostrada de novo"). Não usa `flash` (que sobreviveria a
    um redirect e poderia aparecer em outra tela) e nada fica gravado em claro.
  - Formulários seguem o padrão de `cadastros/formulario.html`: erro mantém os dados; salvar volta para a
    lista preservando busca e filtros.

---

## 5. Inventário (`inventario.py`, `app_inventario.py`)

- **Abrir evento**: o campo de texto "integrantes" vira lista de caixas de seleção com os usuários ativos de
  perfil admin, operador ou inventariante (`usuarios.elegiveis_comissao(conn)`), mostrando nome e perfil.
  `abrir_evento` continua recebendo `integrantes: list[str]` (nomes); a rota valida que cada nome veio de
  um usuário elegível. Rota `abrir` só admin.
- **Comissão editável**: `GET/POST /inventario/<id>/comissao` (só admin, só evento aberto) com a mesma
  lista; `inventario.editar_comissao(conn, evento_id, nomes)` substitui `inventario_integrantes` do evento.
  Não mexe em leituras já feitas: quem sai da comissão deixa de poder ler, o que leu fica. Botão "Comissão"
  na tela do evento, ao lado de "Encerrar".
- **Integrante = logado**: `ler`, `lote`, `atualizar_leitura`, `foto_leitura`, `foto_excluir`, `sobra`,
  `sobra_excluir` usam `g.usuario["nome"]`. Rota `/<id>/integrante`, `session["integrante"]` e o seletor
  de integrante nas telas de evento e sala são removidos. O aviso ao abrir evento vira "Evento aberto.
  Comece pelas salas."
- **Nome do usuário x nome nas leituras**: o relatório e o painel agrupam por texto, como hoje. Para o
  evento aberto em produção ("Inventário 2026", 3 integrantes), o administrador cria os 3 usuários com
  exatamente os nomes já gravados e os marca na comissão; o README registra isso como passo da
  publicação (§8). Renomear um usuário depois não altera leituras antigas.
- **Desktop**: sem `TERMOS_LOGIN`, `g.usuario["nome"]` é "Administrador local" (§7). Ao abrir evento no
  desktop, esse nome entra automaticamente na comissão, além dos marcados, para que dê para ler.

### 5.1 Excluir evento (novo, só admin)

- `GET /inventario/<id>/excluir`: página de confirmação (`inventario_excluir.html`) com o nome do evento, a
  situação (aberto/encerrado) e a contagem do que vai sumir (`inventario.contagem_para_exclusao`: leituras,
  fotos, sobras, integrantes, bens congelados). Campo "Digite o nome do evento para confirmar".
- `POST /inventario/<id>/excluir` com `nome` igual ao do evento (comparação exata após colapsar espaços):
  1. Apaga no bucket todas as fotos do evento (leituras e sobras), na ordem, com o mesmo `fotos.apagar` de
     hoje. **Qualquer falha interrompe antes de tocar no banco**: nada é apagado, a tela avisa "Não foi
     possível apagar as fotos no bucket; o evento foi mantido" (mesmo padrão do commit `eccd1e6`).
  2. Numa transação: `inventario_fotos`, `inventario_leituras`, `inventario_sobras`,
     `inventario_integrantes`, `inventario_bens_encerrados`, `inventario_salas` e `inventario_eventos` do id.
  3. Redireciona para a lista de eventos com "Evento X excluído".
- Nome errado → volta à confirmação com erro, nada apagado. Não há lixeira nem desfazer; o caminho de volta
  é o backup diário (README).
- Botão "Excluir evento" (vermelho, `br-button danger`) na tela do evento, visível só para admin, tanto
  aberto quanto encerrado.

---

## 6. CSRF (`app.py`)

- Token: `session["csrf"]`, gerado com `secrets.token_urlsafe(32)` na primeira vez que a sessão é usada
  numa resposta HTML (no `context_processor`, se ausente). `session.clear()` no login e no sair gera um
  novo.
- Função Jinja global `csrf_campo()` → `<input type="hidden" name="csrf" value="...">`. **Todo**
  `<form method="post">` dos templates recebe `{{ csrf_campo() }}` (cerca de 35 formulários, listados pelo
  teste de §10). Um `<meta name="csrf" content="...">` em `base.html` alimenta os `fetch` POST de
  `inventario_sala.html` e `termo.html`, que passam o cabeçalho `X-CSRF`.
- `before_request`: para `POST`, `request.form.get("csrf")` ou `request.headers.get("X-CSRF")` tem que ser
  igual a `session["csrf"]` (`hmac.compare_digest`). Diferente ou ausente → 400 com "Sessão expirada ou
  formulário inválido. Recarregue a página e tente de novo." (HTML) ou `{"erro": ...}` (JSON). `POST /login`
  também exige o token (a página de login já carrega a sessão).
- No desktop a checagem também vale (mesmo código; o token vive na sessão do WebView). Só o login é que não
  existe lá.

---

## 7. `before_request` e modo desktop (`app.py`, `config.py`)

```python
# config.py
def exigir_login() -> bool:
    return os.environ.get("TERMOS_LOGIN") == "1"     # compose.yml define; desktop e testes por padrão não
```

`before_request` em `app.py`, nesta ordem:

1. `request.endpoint` em (`static`,) → segue.
2. Resolve `g.usuario`:
   - `exigir_login()` falso → `{"id": None, "login": "local", "nome": "Administrador local", "perfil": "admin"}`.
   - Verdadeiro → `usuarios.por_id(conn, session.get("usuario_id"))`, exigindo `ativo = 1`; ausente ou
     inativo → `session.clear()` e redirect para `/login?proximo=<caminho>` (endpoint `usuarios.login`
     passa direto).
3. CSRF para POST (§6).
4. `trocar_senha` → redirect para `/senha` salvo em `usuarios.senha` e `usuarios.sair`.
5. `permitido(g.usuario["perfil"], request.endpoint, request.method)` falso → 403.

No desktop: sem tela de login, sem "Sair", sem item "Usuários" no menu (o `context_processor` omite quando
`g.usuario["id"] is None`), e `/usuarios` devolve 404 para não confundir. Tudo o mais funciona como admin.

`compose.yml` ganha `environment: TERMOS_LOGIN=1`.

---

## 8. Primeiro administrador e publicação

- `python usuarios.py criar-admin <login> "<Nome>"`: pede a senha duas vezes no terminal (`getpass`), cria
  o usuário com perfil admin e `trocar_senha = 0`. Se o login já existir, redefine a senha dele e reativa
  (`ativo = 1`, `falhas = 0`, `bloqueado_ate = NULL`) — é o socorro para admin trancado. Usa
  `config.caminho_db()`, portanto respeita `TERMOS_DADOS`.
- Na VPS: `docker compose exec web python usuarios.py criar-admin antonio "Antônio ..."`.
- Passos de publicação (README, seção "Servir na web"):
  1. `git pull` e `docker compose up -d --build` (o esquema cria `usuarios`).
  2. Criar o admin pelo comando.
  3. Entrar, criar os usuários da comissão do evento aberto com os nomes exatos das leituras, abrir
     "Comissão" no evento e marcá-los.
  4. Remover `auth_basic` e `auth_basic_user_file` do vhost nginx; `nginx -t && systemctl reload nginx`.
  Enquanto o passo 4 não é feito, o site pede as duas senhas (nginx e sistema), o que é inofensivo.

---

## 9. Telas e arquivos

Novos: `usuarios.py`, `app_usuarios.py`, `templates/login.html`, `templates/senha.html`,
`templates/usuarios/lista.html`, `templates/usuarios/formulario.html`, `templates/403.html`,
`templates/inventario_comissao.html`, `templates/inventario_excluir.html`, `tests/test_usuarios.py`,
`tests/test_permissoes.py`, `tests/test_csrf.py`.

Alterados: `db.py` (tabela), `config.py` (`exigir_login`), `app.py` (config de sessão, `before_request`,
`csrf_campo`, menu por perfil, `header-login`, registro do blueprint), `app_inventario.py` e
`inventario.py` (§5), `templates/base.html` (meta CSRF, bloco de usuário, menu), todos os templates com
`<form method="post">` (campo CSRF), `inventario_evento.html`, `inventario_eventos.html`,
`inventario_sala.html` (sem seletor de integrante; cabeçalho nos `fetch`), `termo.html` (cabeçalho no
`fetch`; botões de emissão condicionados), `compose.yml`, `README.md`, `tests/conftest.py`.

DSGov: `login.html` usa o padrão de página de entrada (card centralizado com `br-input` e `br-button
primary`); cabeçalho usa `header-login` com `br-sign-in` (avatar com iniciais, nome, perfil) e o botão
"Sair"; 403 usa `br-message danger`. Visual só; nenhuma dependência nova.

---

## 10. Testes

`tests/conftest.py`: a fixture `app`/`cliente` passa a ligar `TERMOS_LOGIN=1`, semear um admin
(`admin`/`Senha!234`) e logar antes de devolver o cliente, para que os 216 testes atuais continuem valendo
sem mudança. Um helper `logar(cliente, login, senha)` e uma fixture `usuarios_exemplo` (um de cada perfil)
servem aos testes novos. Para os POSTs, o cliente de teste é uma subclasse de `FlaskClient` que injeta o token
CSRF da sessão em todo POST (campo `csrf` no formulário ou cabeçalho `X-CSRF`), então os 153 `cliente.post(...)`
existentes não mudam; o teste do CSRF usa um cliente cru para provar o 400.

- **Matriz de permissões** (`test_permissoes.py`): para cada `(endpoint, metodo)` de `app.url_map` e cada
  perfil, a resposta é 200/302 (permitido) ou 403 (negado) conforme `PERMISSOES`; endpoint fora da matriz
  faz o teste falhar com o nome dele. Menu de cada perfil contém exatamente os itens permitidos. Consulta
  não vê botões de emissão na tela do termo.
- **Login** (`test_usuarios.py`): acerto redireciona e grava `ultimo_acesso`; erro genérico; inativo não
  entra; 5ª falha bloqueia e a senha certa falha durante o bloqueio; após 15 min (relógio simulado via
  `monkeypatch` em `usuarios._agora`) entra e zera falhas; `proximo` externo (`//x`, `http://x`) cai em
  `/`; sair limpa `bens_selecionados`; sem login redireciona para `/login?proximo=...`.
- **Senha**: senha temporária aparece uma vez na resposta e não em requisição seguinte; com
  `trocar_senha = 1` qualquer rota redireciona para `/senha`; regras (mínimo 8, diferente, confirmação);
  troca zera `trocar_senha`.
- **Usuários**: login duplicado, login inválido, último admin não inativa nem rebaixa, o logado não se
  inativa, inativo não aparece em `elegiveis_comissao`; `criar-admin` cria e, repetido, redefine a senha.
- **CSRF** (`test_csrf.py`): POST sem token → 400; com token → passa; JSON com `X-CSRF` → passa; varredura
  dos templates: todo `<form` com `method="post"` (case-insensitive) contém `csrf_campo()`; `base.html`
  tem a `<meta name="csrf">`.
- **Inventário**: abrir/encerrar/comissão/excluir só admin; caixa de comissão só lista elegíveis; leitura
  grava o nome do logado; fora da comissão → 409 com a mensagem nova; editar comissão não altera leituras;
  excluir evento apaga tudo (contagem confere antes e depois) e chama `fotos.apagar` para cada foto;
  falha do bucket → nada apagado, aviso; nome errado → nada apagado.
- **Desktop**: sem `TERMOS_LOGIN`, `/` abre sem login, `g.usuario` é o admin local, `/usuarios` → 404, menu
  sem "Usuários" e sem "Sair"; abrir evento inclui "Administrador local" na comissão.

---

## 11. Riscos e cuidados

- **Nome divergente na comissão** do evento em produção deixa o inventariante sem poder ler ("não faz parte
  da comissão"). Mitigação: passo 3 da publicação e a tela "Comissão" mostrar, ao lado de cada usuário, se
  já há leituras com aquele nome no evento.
- **Publicar sem criar o admin** deixa o site na tela "Nenhum usuário cadastrado". Inofensivo; o comando
  resolve.
- **Trancar todos os admins** (inativar por engano é impedido; esquecer todas as senhas não): `criar-admin`
  com um login existente redefine.
- **Sessões existentes** no dia da publicação não têm `usuario_id`: todo mundo cai no login, como esperado.
- Alterar `PERMISSOES` sem atualizar o teste é impossível pela construção do teste; alterar rota sem
  atualizar `PERMISSOES` derruba a suíte.
