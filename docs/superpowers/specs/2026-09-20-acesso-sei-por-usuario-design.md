# Acesso ao SEI por usuário — cada emissão em nome de quem clicou

**Data:** 2026-09-20 (tarde, depois da publicação do envio ao SEI).
**Estado:** desenho aprovado em conversa; pronto para o plano.
**Base:** `main`, commit `503f9f8`.
**Plano:** `../plans/2026-09-20-acesso-sei-por-usuario.md`.

## 1. Resultado e limites

Hoje o robô entra no SEI sempre com a credencial de `secrets/sei.env` (a do Antônio): todo termo emitido por qualquer
operador nasce no nome dele. Esta entrega faz cada pedido usar **a credencial do SEI de quem clicou**: o usuário cadastra
login, senha e unidade do SEI numa tela própria ("Meus acessos"); a senha fica cifrada no banco; o trabalhador
decifra ao atender o pedido e o robô entra como aquela pessoa, na unidade dela. Quem não cadastrou não emite.

Não entra: instâncias separadas do trabalhador; usuário de serviço no SEI (se um dia existir, basta cadastrá-lo como
acesso de um usuário do sistema); admin definindo senha de terceiros; mudança no fluxo do SPW; modo desktop (continua sem
os botões). Manter Flask, SQLite, Jinja e DSGov.

## 2. Decisões do usuário (2026-09-20)

| Pergunta | Decisão |
|---|---|
| Como variar o usuário do SEI | Credencial por usuário do sistema (não instâncias) — opção 1 |
| Sem acesso cadastrado | Falha antes de enfileirar: "Cadastre seu acesso ao SEI em Meus acessos para emitir." — sem reserva no `sei.env` |
| Quem cadastra | Só o próprio usuário; admin vê se está cadastrado e pode apagar, nunca vê nem define senha |
| Unidade | Por usuário (obrigatória); o robô troca para ela após o login; os blocos "Termos {CC}" precisam existir nessa unidade |

## 3. Dados

Colunas novas em `usuarios` (migração guardada em `criar_esquema`, como as demais): `sei_login TEXT`, `sei_senha TEXT`
(cifrada, nunca em claro), `sei_unidade TEXT`, `sei_atualizado_em TEXT`. Acesso "cadastrado" = as três primeiras
preenchidas.

Cifra: **Fernet** (`cryptography`), chave em `secrets/chaves.env` (`CHAVE_SENHAS=<chave Fernet>`), lida por
`segredos.ler_env`. Módulo novo `cofre.py`: `gerar_chave() -> str`, `cifrar(texto) -> str`, `decifrar(token) -> str`;
sem a chave → `ErroDeNegocio("secrets/chaves.env não encontrado ou incompleto …")` no site e erro legível no pedido.
A chave é a mesma para `web` e `robo` (os dois montam `secrets/` — o `web` ganhou esse mount na leva de correção
do code review final). `cryptography` entra no Dockerfile (alvo `web`,
herdado pelo `robo`) e em `requirements.txt`.

Funções em `usuarios.py`:
- `acesso_sei(conn, id) -> dict | None`: `{login, unidade, atualizado_em}` — **sem** a senha.
- `salvar_acesso_sei(conn, id, login, senha, unidade)`: `login` e `unidade` obrigatórios (unidade em maiúsculas,
  sem espaços); `senha` vazia mantém a atual (erro se não há atual); grava `sei_senha = cofre.cifrar(senha)` e
  `sei_atualizado_em`.
- `apagar_acesso_sei(conn, id)`: zera as quatro colunas.
- `credencial_sei(conn, login_do_sistema) -> dict | None`: `{SEI_USUARIO, SEI_SENHA (decifrada), SEI_UNIDADE}` —
  usada **só** pelo trabalhador.
- `listar` passa a trazer `sei_login` e `sei_atualizado_em` (para a coluna da lista).

`secrets/sei.env` passa a ter só `SEI_LOGIN_URL` e `SEI_ORGAO` (`robo_sei.CHAVES_ENV` idem). `SEI_USUARIO`, `SEI_SENHA`,
`SEI_UNIDADE` deixam de ser lidos (podem ficar no arquivo; README manda apagar).

## 4. Telas

- **`/meus-acessos`** (`usuarios.acessos`, GET/POST, qualquer usuário logado; só com login ativo): campos
  *Usuário do SEI*, *Senha do SEI* (`type=password`, vazio = manter; texto "Senha cadastrada em dd/mm/aaaa" quando há),
  *Unidade no SEI* (sigla); botões *Salvar* e *Apagar meu acesso* (POST `/meu-acesso-sei/apagar`). Ajuda curta: "O termo
  é criado no SEI com este usuário e nesta unidade; os blocos 'Termos {sigla}' precisam existir nela." Link no cabeçalho
  ao lado de *Trocar senha*: **Meus acessos**. Entrada "Sua conta" da Ajuda menciona a tela.
- **Usuários** (admin): coluna "Acesso ao SEI" — `sim · dd/mm` ou `—`; botão *Apagar acesso* (POST
  `/usuarios/<id>/apagar-acesso-sei`, `usuarios.apagar_acesso_sei`, com confirmação JS simples) quando há acesso.
- **Página do termo emitido**: sem mudança de layout; mensagens novas em §6.

## 5. Fluxo

1. `_enfileirar_emissao` (app.py): antes de `preparar_envio_sei`, `usuarios.acesso_sei(conn, g.usuario["id"])`; se
   `None` → `ErroDeNegocio("Cadastre seu acesso ao SEI em Meus acessos para emitir.")` (nada é registrado —
   a checagem entra junto com a da unidade, antes de `registrar_emissao`). `criado_por` continua sendo o login do sistema.
2. `atender_pedidos.executar_sei(conn, pedido)` (novo executor para `sei`): monta `env` =
   `segredos.ler_env(sei.env, (SEI_LOGIN_URL, SEI_ORGAO))` + `usuarios.credencial_sei(conn, pedido["criado_por"])`;
   sem credencial (usuário apagou o acesso depois de clicar, ou pedido antigo sem `criado_por`) → marca `erro` com
   "Quem pediu a emissão não tem acesso ao SEI cadastrado; cadastre em Meus acessos e emita de novo." e não abre
   navegador; chave ausente → `erro` com a mensagem do cofre. Com credencial → `robo_sei.enviar_termo(conn, pedido,
   env=env)`. `enviar_termo` não lê mais `sei.env` sozinho: `env` passa a ser obrigatório (erro claro se faltar).
3. `robo_sei`: sem mudança de lógica (já troca de unidade por `SEI_UNIDADE` e já trata login recusado). A mensagem de
   login recusado passa a "O SEI recusou seu usuário ou senha; atualize em Meus acessos."
4. *Atualizar com SPW*: inalterado (usa `spw.env`).

## 6. Mensagens

| Situação | Mensagem |
|---|---|
| Emitir sem acesso | "Cadastre seu acesso ao SEI em Meus acessos para emitir." |
| Pedido sem credencial ao ser atendido | "Quem pediu a emissão não tem acesso ao SEI cadastrado; cadastre em Meus acessos e emita de novo." |
| Chave ausente | "secrets/chaves.env não encontrado ou incompleto (…): falta CHAVE_SENHAS" |
| SEI recusou | "O SEI recusou seu usuário ou senha; atualize em Meus acessos." |
| Salvar sem senha e sem anterior | "Informe a senha do SEI." |
| Unidade vazia | "Informe a sigla da unidade no SEI." |

## 7. Permissões

`permissoes.py`: `usuarios.acesso_sei GET POST` e `usuarios.apagar_meu_acesso_sei POST` em TODAS (qualquer função);
`usuarios.apagar_acesso_sei POST` em ADMIN. Rotas só com login ativo (`_so_com_login_ligado`, como `/senha`); no desktop
não existem.

## 8. Testes

- `tests/test_cofre.py`: ida-e-volta, tokens diferentes para o mesmo texto, erro sem chave/chave inválida (chave via
  `monkeypatch` em `segredos.PASTA` ou variável).
- `tests/test_usuarios.py`: salvar/manter/apagar acesso, `acesso_sei` sem senha, `credencial_sei` decifra, validações.
- `tests/test_usuarios_telas.py`: tela do próprio usuário (salvar, senha nunca no HTML, apagar); lista do admin com a
  coluna e o botão; admin não consegue definir senha de outro (não há rota).
- `tests/test_app.py`: emitir sem acesso → flash e nenhum registro; com acesso → pedido criado.
- `tests/test_atender_pedidos.py`: `executar_sei` monta o env de `criado_por` (SEI falso recebe o login certo),
  erro sem credencial, erro sem chave.
- `tests/test_robo_sei.py`: `enviar_termo` sem `env` → erro legível; mensagem nova de login recusado.
- `tests/test_permissoes.py`: rotas novas na matriz.

## 9. Publicação

1. `.venv/bin/python -c "import cofre; print(cofre.gerar_chave())"` → `secrets/chaves.env` (`CHAVE_SENHAS=…`, chmod 600).
2. `secrets/sei.env`: manter só `SEI_LOGIN_URL` e `SEI_ORGAO`.
3. `docker compose up -d --build` (dependência nova nas duas imagens).
4. Cada operador: *Acesso ao SEI* no cabeçalho → cadastrar. Blocos "Termos {CC}" na unidade de cada um.
5. README: seção "Acesso ao SEI por usuário" (chave, backup da chave junto com o banco, o que fazer ao trocar a senha do
   SEI); Ajuda atualizada.

## 10. SPW pelo mesmo princípio (pedido do usuário, 2026-09-20 à tarde)

**"Atualizar com SPW" pelo site usa a credencial do SPW de quem clicou**; a atualização automática do cron da madrugada
(`atualizar_base.sh` → `docker compose exec -T robo python importar_spw.py`) continua com `secrets/spw.env`.

- Colunas em `usuarios`: `spw_login`, `spw_senha` (cifrada), `spw_atualizado_em`.
- A tela passa a chamar-se **"Meus acessos"** (`/meus-acessos`, endpoint `usuarios.acessos`), com dois blocos: **SEI**
  (usuário, senha, unidade) e **SPW** (usuário, senha), cada um com *Salvar* e *Apagar*; link do cabeçalho **"Meus acessos"**.
  Funções: `acesso_spw(conn, id)`, `salvar_acesso_spw(conn, id, login, senha)`, `apagar_acesso_spw(conn, id)`,
  `credencial_spw(conn, login_sistema) -> {SPW_USUARIO, SPW_SENHA} | None`. Lista de usuários: coluna **"Acessos"** com
  "SEI" e/ou "SPW" (ou —) e um botão *Apagar acessos* (apaga os dois).
- `base_atualizar_spw` (site): sem acesso ao SPW → `ErroDeNegocio("Cadastre seu acesso ao SPW em Meus acessos para
  atualizar.")` antes de enfileirar. `executar_spw` (trabalhador): pedido com `criado_por` → `credencial_spw` +
  `SPW_LOGIN_URL`/`SPW_CONSULTA_URL` do `spw.env`; sem credencial → `erro` "Quem pediu a atualização não tem acesso ao
  SPW cadastrado; cadastre em Meus acessos e peça de novo."; `importar_spw.executar(conn, env=...)` passa o `env` a
  `baixar_export`. Pedido sem `criado_por` (não existe pelo site) e o cron continuam com `spw.env` completo.
- `importar_spw.CHAVES_ENV` continua exigindo as quatro chaves (o cron precisa delas); só o caminho do pedido
  substitui usuário/senha.
- Mensagens do robô do SPW com senha recusada: a existente ("credenciais recusadas") ganha o sufixo "; atualize em Meus
  acessos" quando o pedido veio do site.

## 11. Evidências

Publicado em 2026-09-20 (~14:52), main=`1591b5f` (merge ff do ramo `acesso-sei`, apagado depois).

- Commits: `08dd025` cofre · `5a20c65` acesso por usuário em `usuarios.py` · `56883f4` tela Meus acessos · `607c800` fluxo
  (emissão/atualização usam a credencial de quem clicou) · `7495fcf` sufixo "atualize em Meus acessos" no login recusado
  do SPW · `90374e4` README · `1591b5f` correções da revisão final (web monta `./secrets:ro`, mensagem do cofre neutra,
  README do backup da chave, `sei_unidade` fora de `_COLUNAS_CONTA`, `MSG_LOGIN_RECUSADO`).
- Suíte: `.venv/bin/pytest -q` → 3381 passed (antes: 3171).
- Publicação: `secrets/chaves.env` gerado (chmod 600); `secrets/sei.env` podado para `SEI_LOGIN_URL` e `SEI_ORGAO`
  (cópia anterior em `secrets/sei.env.antes-acessos-2026-09-20`, chmod 600 — apagar depois que o acesso do Antônio
  estiver cadastrado e uma emissão real tiver funcionado); `./backup.sh` ok (`termos-2026-09-20-1450.db.gz`);
  `docker compose up -d --build`; `/app/secrets/chaves.env` visível nos dois containers e `cofre.cifrar/decifrar`
  funcionando em `web` e em `robo`; `/login` 200; `/meus-acessos` sem sessão → 302; `dados/robo_pedidos.log`
  "trabalhador iniciado" 14:52:36; colunas `sei_*`/`spw_*` criadas em `usuarios`.
- Revisão final (opus) achou o crítico que o plano não previa: o serviço `web` não montava `secrets/` — corrigido
  antes do merge (§3 atualizado). `backup.sh` **não** copia a chave (decisão: chave no mesmo bucket do banco cifrado
  anularia a cifra); guardá-la à parte fica com o administrador.

Pendências (do usuário):
- Cada operador cadastra o próprio acesso em *Meus acessos* (Antônio: `antonio.junior`/GELIC e o do SPW). Até lá,
  Emitir Termo no SEI e Atualizar com SPW pelo site recusam com a mensagem de orientação.
- Blocos "Termos {CC}" precisam existir na unidade de cada emissor.
- Guardar `secrets/chaves.env` fora do servidor (gerenciador de senhas).
