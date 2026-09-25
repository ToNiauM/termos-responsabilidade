# ASTRA — plano completo de correções (Fases 0 a 8)

**Data:** 2026-09-21.
**Estado:** spec; nada implementado.
**Base:** `main`, commit `f28cd8f` (hiperlinks do SEI), 3500 testes passando.
**Origem:** `ASTRA.md` (auditoria técnica de 21/09/2026). Os IDs desta spec (B01, S02, U05, …) são os da auditoria;
a tabela de rastreabilidade está no §15.
**Planos:** um por fase, em `../plans/2026-09-2X-astra-fase-N.md`, escritos quando a fase for começar.

## 1. Resultado e limites

Ao fim das nove fases o sistema deixa de ter os defeitos de integridade e concorrência demonstrados pela auditoria
(planilha vazia apagando o acervo, número de termo duplicado, foto sobrescrevendo foto, leitura depois de
finalizar), passa a guardar o termo emitido como documento imutável, registra quem mudou o quê no patrimônio,
tem build reproduzível com testes rodando em CI, e tem backup e restauração comprovados. As telas não são
redesenhadas: a F5 corrige foco, refresh e confirmação; a F6 corrige o que está quebrado no celular e em telas
menores, dentro do DSGov.

Continua valendo tudo o que já foi decidido: Flask + SQLite + Jinja + DSGov, um servidor, um trabalhador,
modo desktop preservado. Não entra nada do §14 (lista do que a auditoria mandou **não** fazer).

A ordem das fases é a da auditoria e cada fase só começa com a anterior mesclada. A Fase 0 é a única que
não pode esperar; as demais são independentes o bastante para serem reordenadas se a comissão pedir, exceto
onde a coluna "pré-condições" disser o contrário.

## 2. Regras válidas para todas as fases

1. **Regressão antes da correção.** Cada bug da auditoria vira um teste que reproduz o defeito no código atual
   (a receita está em `ASTRA.md` §19), é visto falhando, e só então o código muda. Os testes de concorrência usam
   `threading.Barrier`, duas conexões e monkeypatch do ponto de sincronização — nunca `sleep`.
2. **Só banco sintético.** `tests.conftest.semear` sobre `db.criar_esquema()` em pasta temporária. Nenhum teste
   toca `dados/`, `secrets/`, SEI, SPW ou R2; integrações usam os doubles que já existem (`SEIFalso`, callbacks de
   `enviar`/`apagar`).
3. **Backup antes de dado real.** Toda fase que muda esquema ou saneia dados (F0, F3, F4) começa, na produção,
   por `scripts/backup.sh` com restore ensaiado (F0 entrega o ensaio mínimo; F7 completa).
4. **Legado não se reescreve.** Registro antigo sem snapshot completo fica marcado como "histórico parcial"; nunca
   se reconstrói conteúdo antigo a partir do cadastro de hoje e se apresenta como original. Duplicado não se
   apaga por script: o administrador reconcilia pela tela, com o diagnóstico em mãos.
5. **Sem refatoração ampla no caminho de um P0.** Correção pequena no lugar do defeito; a reorganização
   (transações, migrations) é a Fase 4.
6. **Desktop e web.** Mudança em rota, config ou download passa nos testes dos dois modos (`TERMOS_LOGIN=1` e sem).
7. **Definição de pronto por fase:** regressões passando, suíte inteira passando, evidência no README da fase
   (`docs/superpowers/notes/…`), publicada com `git push` + rebuild, e o item correspondente do §15 marcado.

## 3. Decisões que dependem do usuário

A auditoria marcou como **hipótese a validar** pontos que são política, não código. A spec assume a coluna
"Assumido" até o usuário dizer o contrário; a fase que depende de cada uma está indicada.

| # | Pergunta | Assumido nesta spec | Fase |
|---|---|---|---|
| D1 | Trocar a senha desconecta as outras sessões do mesmo usuário? | Sim (logout global na troca, no reset e na inativação) | F2 |
| D2 | Pedido já na fila de um usuário que foi inativado ou perdeu a função: executa ou cancela? | Cancela, com mensagem "solicitante sem acesso" no pedido | F2 |
| D3 | Bem BAIXADO pode ser atribuído a pessoa ou entrar em termo novo? | Não atribui nem emite; leitura de inventário continua aceita | F3 |
| D4 | Devolução por terceiro (selecionar bens de outra pessoa)? | Não: o servidor confere que cada bem selecionado está atribuído à pessoa do termo | F3 |
| D5 | Importar inventário finalizado (export → import) continua permitido? | Sim, só para o admin, e a leitura de bem que já saiu do acervo é aceita quando está no snapshot | F3 |
| D6 | Editar documento/bloco SEI de termo já concluído pelo robô? | Só admin, com motivo obrigatório, gravado na auditoria | F2 |
| D7 | Frequência e alvo do backup (RPO/RTO) | Diário 3h, retenção 30 dias no R2; RTO aceitável de 4 horas úteis | F7 |
| D8 | Redução grande de acervo no upload manual (mais de 10 % a menos) | Bloqueia e mostra a lista; o admin marca "aceito a remoção de N bens" e reenvia | F0 |

## 4. Fase 0 — Proteção imediata (cinco entregas pequenas)

Pré-condições: nenhuma além de backup. Cada entrega é um commit próprio com sua regressão; podem ser
publicadas juntas ou uma por vez.

### 4.1 E0.1 — Importação segura (B01, B02, B11)

**Hoje:** `db.importar_bens` (`db.py:325`) monta `linhas` a partir da planilha e substitui `bens` tudo-ou-nada,
mas aceita planilha só com cabeçalho (resultado: zero bens) e converte `1001.9` em `1001` pelo `int()`. O robô
(`importar_spw.py:101`) recusa quando vem menos de 90 % do total atual; o upload manual (`app.upload`, `app.py:591`)
não. `importar_planilhas.migrar` (`importar_planilhas.py:22`) faz `DELETE` das quatro tabelas e `commit` antes de
ler o export.

**Mudança em `db.importar_bens`:**

- Número: `_numero()` devolve float; agora exige `num == int(num)`; caso contrário `ImportacaoInvalida(
  "Linha N: número patrimonial 1001.9 não é inteiro.")`. Mesma checagem para vazio/texto que virou `None`
  em linha que tem outras colunas preenchidas (hoje é pulada em silêncio: passa a ser erro com a linha).
- Vazia: zero linhas válidas → `ImportacaoInvalida("A planilha não tem nenhum bem.")`, antes de qualquer DELETE.
- Redução: parâmetro novo `aceitar_reducao: bool = False`. Com `atuais > 0` e `len(linhas) < atuais * 0.9` e
  `aceitar_reducao` falso → `ImportacaoInvalida` cuja mensagem traz `atuais`, `novos` e a quantidade que sairia;
  a exceção ganha atributo `removidos: list[int]` (até 200 números) para a tela listar. A regra dos 90 % sai de
  `importar_spw.py` e passa a viver aqui; o robô chama com `aceitar_reducao=False` (comportamento igual ao de hoje).
- Toda validação acontece antes do primeiro DML; a substituição continua sob a mesma transação de hoje.

**Tela de upload (`upload.html`, `app.upload`):** quando a exceção tiver `removidos`, mostra a mensagem, a lista
(número + descrição atual) e uma caixa "Aceito a remoção destes N bens" que reenvia o mesmo arquivo com
`aceitar_reducao=1`. Só admin vê a caixa (`pode('upload','POST')` já limita a rota; a caixa checa `admin`).

**`importar_planilhas.migrar`:** lê e valida `acervo` e `geral` inteiros em memória (reaproveitando as funções de
validação de `importar_bens` e `_linhas`) antes de tocar no banco; recusa banco que já tenha `pessoas` ou
`atribuicoes` a menos que `--forcar`; DELETE + INSERT numa única transação com rollback em qualquer exceção.

**Testes (novos em `test_db.py`, `test_app.py`, `test_importar_planilhas.py`):** planilha só com cabeçalho deixa
todas as tabelas com o mesmo hash lógico (SELECT * ordenado) de antes; `1001.9` é recusado com a linha na mensagem;
redução de 10 %+ recusada sem a caixa e aceita com ela; robô continua recusando; `migrar` com `geral` inválido
preserva pessoas/atribuições/bens; `migrar` em banco cheio sem `--forcar` recusa. Regressão: baixa em massa
legítima (redução de 5 %) passa sem caixa.

### 4.2 E0.2 — Número de termo único (B03)

**Hoje:** `proximo_numero_termo` (`db.py:997`) lê o maior sequencial e `preparar_envio_sei` (`db.py:1011`) grava,
sem trava de escrita entre a leitura e o UPDATE; duas emissões simultâneas recebem `01/2026`. Não há constraint.

**Mudança:**

- `preparar_envio_sei` abre `BEGIN IMMEDIATE` antes de ler o contador e comita depois do UPDATE; `proximo_numero_termo`
  passa a ser chamada só dentro dessa transação (documentado na docstring).
- `salvar_numero_termo` (manual) também sob `BEGIN IMMEDIATE`, e recusa número já usado na mesma
  unidade/tipo/ano com `ErroDeNegocio("Número 03/2026 já é do termo #17.")`.
- Índice `CREATE UNIQUE INDEX IF NOT EXISTS termos_numero_unico ON termos_emitidos(unidade_sei, tipo, numero_termo)
  WHERE numero_termo IS NOT NULL`, criado em `criar_esquema` **dentro de try**: se falhar por duplicado já existente,
  o app sobe normalmente, grava aviso em log e o card do Início do admin mostra "Há números de termo duplicados:
  ver diagnóstico". O índice é tentado de novo a cada startup até a base estar limpa.
- `scripts/diagnostico_termos.py` (somente leitura): lista duplicados por unidade/tipo/número com id, data,
  documento SEI e situação do pedido. Reconciliação é manual, pelo campo de número da tela do termo, comparando
  com o SEI. Nada apaga.

**Testes:** duas threads com barreira dentro de `proximo_numero_termo` (receita §19) recebem números distintos e o
banco tem os dois; número manual repetido é recusado; `criar_esquema` num banco com duplicado não levanta e o
diagnóstico lista o par; depois de corrigido, o índice existe. Regressão: numeração manual em unidade que já tinha
termos à mão; termos de tipos diferentes com o mesmo número convivem.

### 4.3 E0.3 — Chave única da foto e recuperação de falha (B07)

**Hoje:** `inventario.adicionar_foto` (`inventario.py:409`) calcula `nfoto = max(fotos_seq, MAX(nfoto)) + 1`, chama
`enviar(chave)` (rede) e só então INSERT; duas chamadas simultâneas calculam o mesmo `nfoto`, enviam para a mesma
chave no R2 (a segunda sobrescreve a primeira) e uma delas quebra no INSERT.

**Mudança, em três passos curtos:**

1. Reserva: `BEGIN IMMEDIATE`; checa evento aberto; lê `fotos_seq` e `MAX(nfoto)`; `nfoto = maior + 1`;
   `UPDATE inventario_leituras SET fotos_seq = nfoto`; `COMMIT`. A transação dura microssegundos e não envolve rede.
2. Envio: `enviar(chave_bem(pasta, nfoto, numero))` fora de transação. Se falhar, o `nfoto` fica queimado (regra
   que já existe: nunca reaproveitar) e a exceção sobe; nada a compensar.
3. Gravação: `BEGIN IMMEDIATE`; checa evento aberto de novo; INSERT em `inventario_fotos`; `COMMIT`. Se o INSERT
   falhar (evento fechou no meio, por exemplo), chama `apagar(url)` em best-effort e relança `ErroDeNegocio`.

Chaves antigas não mudam: `chave_bem` e `_PREFIXO_ANTIGO` ficam como estão; a ordem de exibição continua por `nfoto`.

**Testes:** duas conexões com barreira no callback `enviar` recebem chaves distintas e as duas fotos ficam no banco;
`enviar` que levanta não grava e a próxima foto pula o número; INSERT que falha dispara `apagar` com a URL enviada.
Regressão: `apagar_foto` da última não reusa o número; fotos ordenadas na tela.

### 4.4 E0.4 — Estado e gravação atômicos no inventário (B08)

**Hoje:** `_evento_aberto_ou_erro` (`inventario.py:78`) é um SELECT sem trava; entre ele e o INSERT de `ler`, outra
conexão pode `encerrar_evento`. A leitura entra num evento finalizado e fica fora do snapshot.

**Mudança:** context manager `_escrita_no_evento(conn, evento_id)` em `inventario.py`: `BEGIN IMMEDIATE`, chama
`_evento_aberto_ou_erro` **dentro** da transação, cede o controle, comita ao sair, rollback em exceção. Todas as
mutações de evento passam a usá-lo: `ler`, `editar_leitura`, `desfazer_leituras`, `adicionar_foto` (passos 1 e 3),
`apagar_foto`, `registrar_sobra`, `editar_sobra`, `excluir_sobra`, `apagar_foto_sobra`. `encerrar_evento`,
`suspender`/`reabrir` (chave do admin) e `excluir_evento` também abrem `BEGIN IMMEDIATE` antes de checar o estado.
`_evento_nao_finalizado_ou_erro` (comissão) segue o mesmo padrão.

`ler_lote` continua comitando por bem (contrato atual), mas passa a devolver também `gravados: list[int]` e, se uma
leitura falhar por evento fechado, interrompe e devolve o que já entrou; a tela mostra "N gravados; o evento foi
fechado antes dos demais".

Como `db.conectar` usa o `isolation_level` padrão do sqlite3 (transação implícita no primeiro DML), o helper
executa `BEGIN IMMEDIATE` explicitamente e as funções dentro dele não chamam `commit`. Onde hoje há `with conn:`,
trocar pelo helper.

**Testes:** monkeypatch em `_evento_aberto_ou_erro` com barreira; segunda conexão finaliza; a leitura retomada é
recusada com "Evento finalizado" e o banco não tem a linha (o critério é ausência da linha, não timeout). O mesmo
esquema para foto, sobra e desfazer. Regressão: fluxo abrir → ler → fechar → reabrir → ler → finalizar; comissão
editável com evento fechado.

### 4.5 E0.5 — Scripts operacionais seguros (R02, R03)

- `scripts/atualizar_base.sh --teste`: troca `cp dados/termos.db` por `sqlite3 dados/termos.db ".backup <destino>"`
  (ou `python -c` com `conn.backup()` se o binário não existir na imagem). O modo `--teste` continua fazendo login
  real no SPW — isso é documentado, não "corrigido".
- `scripts/zerar_banco.sh` e `scripts/apagar_termos_emitidos.sh`: só rodam se existir `dados/AMBIENTE_DESCARTAVEL`
  (arquivo criado à mão, nunca pelo compose); param `web` e `robo` (`docker compose stop`) antes de qualquer
  operação e religam ao final; fazem `.backup` antes; imprimem o caminho do banco e pedem confirmação digitada
  "APAGAR". Em `apagar_termos_emitidos.sh`, aviso explícito: "Os documentos no SEI não são apagados; a numeração
  recomeça e pode repetir número já usado no SEI".
- `scripts/backup.sh` entra no Git (hoje está ignorado): credenciais e destino vêm de `secrets/backup.env`;
  o script não muda de comportamento nesta fase (R01 é F7). Ensaio mínimo de restore: `scripts/restaurar_teste.sh`
  restaura o último `.gz` numa pasta temporária, roda `integrity_check`, `foreign_key_check` e imprime contagens
  de bens, pessoas, termos e eventos. Rodado uma vez na publicação da F0 e o resultado vai no README da fase.

**Testes:** `test_scripts.py` novo com `bash -n` nos três scripts e execução de `restaurar_teste.sh` sobre um
banco sintético gzipado; os scripts destrutivos rodados sem o marcador saem com código 2 e sem alterar nada.

**DoD da Fase 0:** as seis reproduções (B01, B02, B03, B07, B08, B11) não reproduzem o dano; índice de número
existe ou o diagnóstico explica por quê; restore ensaiado; suíte passa.

## 5. Fase 1 — Rede de segurança e build reproduzível

Pré-condições: F0 mesclada.

### 5.1 Dependências fixadas (S07 ferramenta, D03 parte)

- `requirements.txt` (web), `requirements-robo.txt` (Playwright, xlrd) e `requirements-desktop.txt` novo
  (pywebview, pyinstaller) com `==versão` para tudo, gerados de `pip freeze` filtrado; `waitress` entra no web.
- `Dockerfile` passa a `COPY requirements*.txt` + `pip install -r`, com `pip` atualizado na imagem para versão
  que cubra os seis avisos do OSV (≥ 26.2.0 conforme `ASTRA.md` S07); `FROM python:3.12-slim@sha256:…` com digest
  fixado e comentário de como atualizar.
- `README.md`: seção "Atualizar dependências" (regenerar freeze, rodar suíte, trocar digest); remove a frase
  "dependências fixas" onde não eram; corrige "backup = copiar dados/" apontando para `backup.sh`.
- `.env.exemplo` sem segredos, com todas as variáveis lidas por `config.py`.

### 5.2 CI

`.github/workflows/testes.yml`: em push e PR, Python 3.12, `pip install -r requirements.txt -r requirements-robo.txt`,
`pytest -q`. Sem rede de integração e sem segredos — a suíte já não precisa. Falha bloqueia merge (proteção de
branch configurada à mão no GitHub, anotada no README).

### 5.3 Testes de rede de segurança

- `tests/test_matriz_rotas.py`: percorre `app.url_map`; toda rota+método tem entrada em `permissoes` (ou está na
  lista de exceções: login, logout, static); método fora da matriz responde 405/403 com cliente cru.
- `tests/test_esquema_legado.py`: fixtures de banco nas versões anteriores (DDL congelado em arquivos
  `tests/esquemas/v*.sql`) passam por `criar_esquema` e chegam ao esquema atual com os dados preservados.
- Smoke de inicialização: `import app`, `import atender_pedidos`, `import main` com `TERMOS_LOGIN=1` e sem;
  geração de um DOCX de cada tipo.
- E2E local: `tests/e2e/` com Playwright (já está na venv do robô) contra o Flask em thread, banco sintético,
  marcado `@pytest.mark.e2e` e fora do default (`-m e2e` para rodar; roda no CI num job separado). Nesta fase só
  a infra e um teste (login → Início). Os casos de foco/clipboard/câmera entram na F5.

**Regressões:** build da imagem web e robô; `scripts/build.bat` (PyInstaller) com o requirements novo; Chromium
do Playwright instala na imagem.

**DoD:** clone limpo + README reproduz o ambiente; CI verde; `pip audit`/OSV sem aviso aberto para os três
alvos (registrado no README da fase com data); contagem de testes anotada.

## 6. Fase 2 — Sessão e correções funcionais delimitadas

Pré-condições: F1; decisões D1, D2, D6.

### 6.1 Sessão e método HTTP (S01, S02, S04)

**S01 — versão de sessão.** Coluna `usuarios.sessao_versao INTEGER NOT NULL DEFAULT 1`. `app_usuarios.login`
grava `session["sessao_versao"]`; `resolver_usuario` compara com o banco a cada request (já consulta o usuário)
e, se diferir, `session.clear()` + redirect para login com "Sua sessão foi encerrada". Incrementam:
`trocar_senha`, `redefinir_senha` (reset temporário), `editar` quando `ativo` passa a 0, e um botão novo
"Encerrar minhas outras sessões" em `/senha`. Logout continua local (D1: só a troca de senha é global).
`PERMANENT_SESSION_LIFETIME` de 12 h continua; `SESSION_REFRESH_EACH_REQUEST` fica `False` para o prazo ser
absoluto (hoje renova a cada request).

**S02 — emitir só por POST.** `termo_docx` (`app.py:406`) deixa de chamar `registrar_emissao`. A rota fica
`methods=["POST"]` para emitir (o botão "Baixar DOCX" vira um `<form method=post>` com CSRF; a resposta continua
sendo o arquivo) e ganha irmã `GET /termos-emitidos/<id>/docx`, que gera o DOCX **a partir do registro**
(`_html_do_registro` já faz isso para HTML; o gerador DOCX recebe os bens do snapshot), sem gravar nada. HEAD segue
respondendo vazio. Desktop: pywebview trata download de POST igual ao de GET (testado em `test_app.py` com
`TERMOS_LOGIN` ausente e no ensaio manual do README).

**S04 — modo explícito.** `config.exigir_login()` passa a ler `TERMOS_MODO`: `web` (exige login) ou `desktop`.
`main.py` define `desktop`; `compose.yml` define `web`; `conftest` define por fixture. Valor ausente ou diferente
→ `RuntimeError("Defina TERMOS_MODO=web ou desktop")` na importação de `app`. `TERMOS_LOGIN=1` continua aceito
por um release como sinônimo de `web`, com aviso em log. Modo `web` também exige `SECRET_KEY` de arquivo
(já é assim) e recusa `debug`.

### 6.2 Estado do documento SEI e formulário atômico (S03, B06, S09)

**B06.** `termo_emitido_documento` (`app.py:558`) valida `numero_termo`, `documento_sei` e `bloco_sei` antes de
qualquer escrita e grava os três numa única transação (`db.salvar_documento_sei` recebe também o número e
não comita; a rota comita). O `finally` sai. Erro → flash e nenhuma coluna muda.

**S03 — transições no servidor.** Estados do termo em relação ao SEI, derivados do pedido ativo
(`robo_pedidos`) e de `documento_sei`:

| Estado | Condição | Número | Documento/bloco |
|---|---|---|---|
| rascunho | sem documento, sem pedido ativo | editável | editável |
| na fila / andamento | pedido `aguardando`…`bloco` | bloqueado | bloqueado |
| erro_bloco | pedido `erro` com documento gravado | bloqueado | só bloco |
| concluído | documento gravado, pedido concluído ou nenhum | bloqueado | só admin com motivo (D6) |

`db.salvar_documento_sei` recebe `usuario`, `motivo` e valida a transição; edição em `concluído` grava linha em
`auditoria` (tabela criada aqui, §6.4). Apagar documento não volta o número para editável (fecha a brecha de
"apagar para renumerar"). O template continua escondendo o que o servidor recusa.

**S09 — revogação alcança a fila (D2).** `atender_pedidos.atender_um`, depois do `flock` e antes de abrir o
navegador, relê o pedido e chama `usuarios.pode_executar(conn, criado_por, tipo)`: usuário ativo e com a função
exigida (`emitir_sei` / `atualizar_spw`). Caso contrário o pedido vai para `passo='cancelado'`,
`mensagem='Solicitante sem acesso'`, sem efeito externo. `credencial_sei`/`credencial_spw` passam a exigir `ativo=1`.

### 6.3 Retorno, e-mail, paginação, URLs, limites (S08, B13/U08, B14/U07, S05, S06)

- **S08.** `app._retorno_local()` novo: aceita só path relativo da própria aplicação (`urlsplit` sem scheme/netloc,
  path que resolve num endpoint do `url_map`); handlers de `ErroDeNegocio` e 413 usam `_retorno_local(request.referrer)`
  com fallback `home`. `app_cadastros.retorno` fica como está (já é mais estrito).
- **B13/U08.** O JS que dispara POST 500 ms após o clique no `mailto:` sai. O botão passa a "Abrir e-mail" (só o
  `mailto:`) e, ao lado, "Marcar como enviado" (POST explícito). Rótulo gravado: "Envio marcado em …" em vez de
  "E-mail enviado". Indicador "Bens iguais aos de hoje" vira "Mesmos números de hoje" até a F3 trocar por hash.
- **B14/U07.** `db.termos_emitidos` ganha `pagina`/`por_pagina` (200) e devolve `total`; `termos_emitidos.html`
  ganha paginação DSGov e filtros (tipo, chave, período) na query. A ficha (`/bem`) e o termo emitido recebem
  `retorno=` construído por `_retorno_local` e o link "Voltar" o usa; breadcrumb mantém a rota simples.
- **S05.** `inventario.validar_abas` só aceita `foto_url` `https://` com host igual ao do bucket configurado
  (`fotos.host_publico()`), ou com o prefixo legado `inventario/`; rejeita o resto com a linha. `fotos.apagar`
  recebe também `pasta_esperada` e recusa URL cuja chave não comece por ela (ou pelo prefixo legado + nome do
  evento). Templates de foto passam a `href`/`src` só de URL que passou na validação (`fotos.url_segura`), senão
  mostram "foto indisponível".
- **S06.** `importar_bens` e `importar_cadastros`: limite de 100 000 linhas e aborto se `ws.max_row` passar disso;
  `defusedxml` no requirements (openpyxl o usa quando presente). `fotos.validar`: `Image.MAX_IMAGE_PIXELS = 40_000_000`
  explícito e `DecompressionBombError` → `ErroDeNegocio("Imagem grande demais")`; `comprimir` dentro do mesmo
  tratamento. Campos livres de leitura/sobra/textos com `maxlength` no servidor (2 000 caracteres).
  Contador de falhas do login: `UPDATE usuarios SET falhas = falhas + 1 WHERE id = ?` atômico e mensagem
  de bloqueio igual para conta existente e inexistente.

### 6.4 Tabela `auditoria` (base do O01)

Criada aqui porque S03 e D6 precisam dela; F3 a alimenta nas demais operações.

```sql
CREATE TABLE IF NOT EXISTS auditoria (
  id          INTEGER PRIMARY KEY,
  em          TEXT NOT NULL,
  usuario_id  INTEGER,            -- NULL no desktop
  usuario     TEXT NOT NULL,      -- login ou 'local'
  origem      TEXT NOT NULL,      -- 'web' | 'robo' | 'script'
  entidade    TEXT NOT NULL,      -- 'termo' | 'bem' | 'pessoa' | 'centro' | 'localizacao' | 'texto' | 'importacao' | 'usuario'
  entidade_id TEXT NOT NULL,
  acao        TEXT NOT NULL,
  antes       TEXT,               -- JSON
  depois      TEXT,               -- JSON
  motivo      TEXT,
  termo_id    INTEGER,
  importacao_id INTEGER
);
CREATE INDEX IF NOT EXISTS auditoria_entidade ON auditoria(entidade, entidade_id, em);
```

`db.auditar(conn, **campos)` grava **sem** commit, dentro da transação do caso de uso. Nunca recebe senha, cookie
ou credencial cifrada (teste garante que `antes`/`depois` de `usuario` não têm `senha_hash`).

**Testes da F2:** cookie guardado antes da troca de senha é recusado em segundo cliente; reset → troca → cookie
antigo recusado; GET do DOCX não altera `termos_emitidos`; POST emite uma vez; `TERMOS_MODO` ausente derruba a
importação; POST com número inválido não muda nenhuma coluna (comparar `SELECT *` antes/depois); transições da
tabela §6.2 uma a uma; worker com solicitante inativo cancela sem chamar `SEIFalso`; Referer externo cai no Início;
201 termos aparecem em duas páginas; `javascript:` e host estranho recusados na importação; ZIP com 200 000 linhas
recusado; imagem 100 MP recusada; falhas de login contadas sob 10 threads.

**DoD:** os controles são do servidor (testes com cliente cru, sem template); nenhuma elevação de permissão;
E2E de emitir-por-POST no desktop e web.

## 7. Fase 3 — Integridade histórica e regras patrimoniais

Pré-condições: F0–F2; decisões D3, D4, D5; backup + restore ensaiado antes de sanear `atribuicoes`.
Ordem interna obrigatória: 7.1 → 7.2 → 7.3 → 7.4.

### 7.1 Uma carga por bem (B10)

`scripts/diagnostico_atribuicoes.py` (leitura) lista números com mais de uma pessoa. Reconciliação manual pela tela
Cadastros → Pessoas (já tem desatribuir com revisão). Depois: `CREATE UNIQUE INDEX IF NOT EXISTS atribuicoes_numero
ON atribuicoes(numero)` no mesmo esquema de "tenta a cada startup, avisa se falhar" do §4.2. `importar_cadastros` e
`importar_planilhas.migrar` validam unicidade antes de gravar. `db.atribuir` sob `BEGIN IMMEDIATE` (já é) e com D3:
recusa bem BAIXADO com `ErroDeNegocio`.

### 7.2 Termo emitido imutável (B04, B05, B12, O01 parte)

**Identidade estável.** `pessoas.uid TEXT` e `centros.uid TEXT` (UUID4, preenchidos na migração para as linhas
existentes e em toda criação). Renomear pessoa/centro continua com `ON UPDATE CASCADE`, mas o histórico aponta pelo
`uid`.

**Snapshot completo.** Colunas novas em `termos_emitidos`: `responsavel_uid`, `responsavel_nome`,
`responsavel_extra` (JSON: cargo, e-mail, sigla/nome do centro, matrícula — o que o texto usar), `textos` (JSON
dos textos configurados na hora), `html` (corpo renderizado, o mesmo que o Copiar cola), `hash` (SHA-256 de
tipo+chave+processo+responsável+textos+bens com descrição/complemento/localização/valor). `registrar_emissao`
preenche tudo e **deduplica por `hash` no mesmo dia** (não mais por números): conteúdo diferente = termo novo;
igual = atualiza `emitido_em` como hoje. `_html_do_registro` devolve `html` gravado; para registros anteriores
(`html IS NULL`) reconstrói como hoje e a tela mostra tarja "Registro anterior a <data da F3>: reconstruído a partir
do cadastro atual". `situacao_termo` compara por `hash` (rótulo "Vigente" só quando o hash do termo bate com o que
seria emitido agora; "Conteúdo mudou" caso contrário).

**Prévia = registro (B12).** A tela do termo calcula o `hash` no servidor e o coloca no `data-hash` do iframe; o
Copiar envia `hash` no POST `termo_registrar`; o servidor recalcula e, se diferir, responde 409 "O termo mudou
desde que a tela foi aberta; recarregue". DOCX (POST da F2) e enqueue usam o registro cujo `html`/`hash` já está
gravado; `robo_pedidos.html` deixa de duplicar o corpo e passa a ler `termos_emitidos.html` (coluna mantida por
um release para pedidos antigos).

**Auditoria nas operações patrimoniais (O01).** `db.auditar` dentro da transação de: `atribuir`, `desatribuir`,
`salvar_centro`, `_renomear_pessoa`, `salvar_pessoa`, `excluir_pessoa`, mover localização, `salvar_responsavel`,
`salvar_texto`, `importar_bens` (ator + importacao_id), `importar_cadastros`, `registrar_emissao`,
`salvar_documento_sei`, `salvar_numero_termo`, encerrar/excluir evento. `historico_do_bem` passa a juntar a
auditoria (por `entidade='bem'` e por `termo_id`) e a ficha mostra "quem, quando, antes → depois, motivo".
`_mudancas` da importação registra também valor, descrição, classificação e complemento. Leituras guardam
`usuario_id` do integrante além do nome.

**Devolução (D4).** `termo_devolucao` valida no servidor que cada bem selecionado está em `atribuicoes` para a
pessoa do termo; a seleção na sessão fica sob a chave `bens_selecionados[nome]` para duas abas não se misturarem.

### 7.3 Inventário histórico e importação (B15, B16, B09)

- `inventario_leituras` ganha `descricao`, `localizacao_cadastro`, `situacao` copiados de `bens` no momento da
  leitura. `relatorio` e `encerrar_evento` fazem `LEFT JOIN bens` e caem no snapshot da leitura quando o bem saiu.
- `importar_bens`: bem com leitura em evento aberto que sairia do acervo entra na lista de `removidos` da
  `ImportacaoInvalida` (§4.1) com marca "lido no inventário X"; só passa com `aceitar_reducao`.
- `validar_abas` (D5): leitura de número ausente em `bens` é aceita se o evento importado está finalizado e o
  número está em `inventario_bens_encerrados` da mesma planilha.
- **B09.** `inventario_fotos.excluida_em TEXT` e `inventario_sobras_fotos` idem. Excluir foto, desfazer leitura e
  excluir evento marcam `excluida_em` e comitam **antes** de chamar `apagar`; cada `apagar` que falhar deixa a
  linha marcada e a mensagem diz "N fotos ficaram pendentes de exclusão no bucket". `scripts/limpar_fotos.py`
  (reexecutável, idempotente) tenta de novo as pendentes e apaga a linha quando o bucket confirma. Telas ignoram
  fotos com `excluida_em`.

### 7.4 Fila atômica (riscos "prováveis" do §6 da auditoria)

- `termo_emitido_enviar_sei` e `termo_enviar_sei`: `registrar_emissao` + `preparar_envio_sei` + `enfileirar_pedido`
  numa transação só (as três funções ganham `commit=False`; a rota abre `BEGIN IMMEDIATE` e comita). Falha em
  qualquer uma → nada gravado.
- `atender_um`: relê o pedido depois do `flock` (`SELECT … WHERE id = ? AND passo = 'aguardando'`) e só então
  marca `iniciado_em`; `robo_pedidos.pulsacao TEXT` atualizado a cada passo do robô;
  `pedidos_orfaos_para_aguardando` usa `pulsacao` (10 min sem pulso), não `iniciado_em`.

**Testes da F3:** UNIQUE por inserção direta e por migração; renomear pessoa após emissão não muda `html` nem
`responsavel_nome`; alterar valor mantendo números cria termo novo e o antigo fica igual (comparar `SELECT *`);
POST com hash defasado responde 409; DOCX do registro bate com o `html` (mesmos bens/valores); pedido lê
`termos_emitidos.html`; registro antigo sem `html` mostra tarja; auditoria tem antes/depois em cada operação da
lista e nunca `senha_hash`; devolução com bem de outra pessoa é recusada; leitura de bem removido aparece no
relatório e no snapshot de encerramento; import de inventário finalizado com bem inexistente passa; exclusão com
`apagar` falhando deixa `excluida_em` e o script limpa depois; enqueue com falha na terceira etapa não deixa termo
sem pedido; dois workers com o mesmo pedido (barreira antes do flock) executam só um.

**DoD:** constraints ativas ou diagnóstico explicando; termo emitido é imutável e o documento baixado/copiado/
enviado é o mesmo registro; toda operação da lista tem auditoria; legado rotulado; conciliação de duplicados
aprovada pela comissão antes da publicação.

## 8. Fase 4 — Arquitetura, esquema e performance proporcional

Pré-condições: F3 mesclada; medição de baseline com fixture de 8 000 bens, 300 pessoas, 40 centros, 2 000 termos.

### 8.1 Propriedade das transações (D01)

Regra escrita em `db.py` (docstring de módulo) e verificada por teste: **função pública de caso de uso comita uma
vez; função com `_` na frente nunca comita; função pública chamada por outra recebe `commit=False`**. Casos
reorganizados: emissão → preparo → enqueue (já na F3), edição de usuário → nome na comissão, importação → auditoria,
`ler_lote` (decisão do §4.4 mantida: por bem). Teste `test_transacoes.py` faz monkeypatch de `commit` e conta
chamadas por caso de uso. Não se extrai camada nova.

### 8.2 Esquema versionado (D02)

`schema_versao (versao INTEGER PRIMARY KEY, aplicada_em TEXT)`. `db.MIGRACOES = [(1, fn), (2, fn), …]`: cada
função roda numa transação, é idempotente, e a lista começa pelo que `criar_esquema` faz hoje por inspeção
(ALTERs, índices tentados, limpezas). `criar_esquema` = DDL base + `aplicar_migracoes`. Só o serviço `web` migra;
`atender_pedidos` chama `db.esperar_esquema(versao_esperada, timeout=120)` e sai com erro claro se não chegar.
Limpezas históricas que hoje rodam em silêncio passam a imprimir relatório ("N linhas removidas de X") e a exigir
que o backup do dia exista (`dados/backups/` com arquivo de hoje) antes de rodar. Testes com os bancos legados da
F1 (`tests/esquemas/v*.sql`) sobem até a versão atual e voltam a subir sem efeito.

### 8.3 Performance medida (P01–P05, índices)

1. `tests/bench/` com a fixture grande e `pytest --bench` que registra tempo e contagem de queries
   (`sqlite3.Connection.set_trace_callback`) para: painel, `bens_da_sala` com 300 bens/fotos, `usuarios.listar`,
   análise com export, busca global, relatório de inventário. Baseline vai para o README da fase.
2. Só o que o baseline mostrar acima de 300 ms ou com N+1 acima de 50 queries é corrigido: `situacoes_centros/
   pessoas` com uma consulta agregada de carga e outra de últimos termos (`GROUP BY`), fotos de `bens_da_sala` em
   lote (`WHERE numero IN (…)`), funções de `usuarios.listar` em lote, `recorte` separando linhas de gráficos
   (`exportar` não calcula dimensões).
3. Índices candidatos (`atribuicoes(numero)` já vem do UNIQUE; `bens(localizacao, situacao)`,
   `importacoes_mudancas(numero, importacao_id)`, `termos_emitidos(tipo, chave, emitido_em, id)`,
   `termos_emitidos_bens(numero, termo_id)`, `inventario_leituras(evento_id, localizacao)`) criados um a um com
   `EXPLAIN QUERY PLAN` antes/depois anexado; os que não mudarem o plano não entram.
4. WAL: teste de carga (4 threads Waitress + robô) com e sem `journal_mode=WAL` e `backup.sh` rodando; adota-se só
   se reduzir `database is locked` e o backup continuar íntegro. Decisão registrada, não presumida.
5. Dinheiro: `valor_total` e somas passam a arredondar com `round(x, 2)` numa função única `db.soma_reais`;
   armazenamento continua REAL (mudar para centavos inteiros só se o bench de reconciliação mostrar diferença).
   Datas: `_data_br()` valida `DD/MM/AAAA` real na importação e rejeita impossíveis com a linha; `_agora()` usa
   `ZoneInfo("America/Sao_Paulo")` explícito e testes de virada de ano.

**DoD:** regras de negócio inalteradas (suíte igual); bench antes/depois no README; migrador único; robô espera
o esquema; rollback de release testado com backup do esquema anterior.

## 9. Fase 5 — UI/UX e acessibilidade

Pré-condições: F1 (E2E) e F3 (estados estáveis). Sem redesenho; DSGov continua só visual.

| ID | Mudança | Teste E2E |
|---|---|---|
| U01 | `termo_emitido.html` e `upload.html` trocam `meta refresh` por `fetch` a cada 5 s em `GET /termos-emitidos/<id>/estado` e `GET /upload/estado` (JSON), atualizando só o bloco de status | arquivo escolhido, foco e scroll sobrevivem a 3 ciclos |
| U02 | `inventario_sala.html`: cada `.campo` com `<label>` (visível ou `sr-only`) "Conservação do bem 1001"; upload de foto com botão real + `aria-describedby`; `focar()` só quando o campo ativo não é editável | tab percorre a sala inteira; leitor de tela (axe) sem violação crítica |
| U03/F06 | Rota `/inventario/<evento>/bem/<numero>`: ficha reduzida (dados do bem, leitura, fotos todas, divergência) para quem tem acesso ao evento; sem link para `/bem` global | inventariante puro abre; outro evento → 403 |
| U04 | Confirmação DSGov (modal) em excluir foto, excluir sobra e foto de sobra; texto avisa que a exclusão no bucket é pendente se falhar (B09) | cancelar não exclui |
| U05 | Autosave com três estados por campo (pendente/salvo/erro) via classe + ícone; sequência por campo (`seq` incrementado no cliente, servidor ignora resposta mais antiga que a última aplicada); retry único | rede lenta simulada: último valor vence; erro visível sem "salvo" |
| U06 | Botão "Importar cadastros" só com `pode('importar_cadastros','POST')` | operador não vê; POST direto 403 |
| U07 | Ficha e termo emitido com "Voltar" preservando `retorno` (F2) | análise filtrada → ficha → voltar mantém filtros e página |
| U08 | Rótulos finais: "Emitido em", "Enviado ao SEI em", "Envio de e-mail marcado em", "Conteúdo igual/mudou" (hash da F3) | textos presentes |

Também: contraste dos estados de tag/alerta medido com axe no E2E (limiar AA); viewport 390 px sem scroll
horizontal nas telas de inventário, termo e cadastros; tamanho de alvo ≥ 40 px nos botões de leitura.

**Regressões:** BRCard/BRTable inicializam; câmera (`html5-qrcode`) continua abrindo no mobile (ensaio manual no
celular registrado com captura); clipboard do Copiar; menus cumulativos por função.

**DoD:** E2E dos fluxos críticos verdes em desktop e 390 px; nenhum acesso novo a acervo/evento; relatório axe
sem críticos; capturas antes/depois no README da fase.

## 10. Fase 6 — Responsividade e mobile

Pré-condições: F5 (a mesma infra E2E e os mesmos templates). Pedido do usuário em 2026-09-21: "tem muita coisa
quebrada em tamanho de tela, especialmente mobile; melhorar sem comprometer o DSGov". A fase é só de layout —
nenhuma regra, rota ou permissão muda — e usa exclusivamente os utilitários e componentes do DSGov (grid
`col-sm/md/lg`, `br-table` responsivo, `br-list`, `d-none d-sm-block`, `br-button block`) mais o CSS próprio já
existente em `static/dsgov/css/dsgov.css` (hoje com 5 `@media`). Nada de framework CSS novo nem de tema paralelo.

### 10.1 Levantamento (primeira sessão, sem código)

O que se sabe do código hoje: o grid usa quase só `col-md-*` (um breakpoint em 768 px), então abaixo dele tudo
empilha em 100 % e acima dele vira desktop de uma vez; tabelas HTML aparecem em `_macros.html`, `cadastros/_lista.html`,
`inventario_sala.html`, `termo_devolucao.html` e nos termos emitidos; `analise.html` tem 11 colunas `col-md-3` de
filtros e gráficos ECharts; o cabeçalho é o `br-header` sticky com menu lateral.

Entregável: `docs/superpowers/notes/2026-XX-XX-responsividade/levantamento.md` com uma captura por tela em três
larguras — **390 px** (celular), **768 px** (tablet) e **1280 px** — feitas pelo E2E da F1 (`page.screenshot`) sobre
banco sintético, e uma tabela tela × largura × defeito (texto cortado, scroll horizontal, botão fora da área,
tabela ilegível, toque menor que 40 px, modal maior que a tela, cabeçalho cobrindo conteúdo). O usuário marca
nessa tabela o que mais dói no uso real da comissão com o celular; essa marcação define a ordem do 10.2.

### 10.2 Correções por grupo (uma sessão por grupo, ordem definida pelo levantamento)

| Grupo | Telas | O que muda |
|---|---|---|
| **Inventário no celular** (o uso mobile de verdade) | `inventario_sala.html`, `inventario_evento.html`, `inventario_painel.html` | Tabela da sala vira lista de cards abaixo de 768 px (`br-table` com `data-responsive`/cards do DSGov, um bem por card: número, descrição, conservação, foto, ações); campo de leitura e botão "Ler" fixos no rodapé da viewport (`position: sticky; bottom: 0`), botão da câmera com 48 px; fotos em grade 3 colunas com miniaturas quadradas; ações de lote em `br-button block`; painel do evento com contadores em 2 colunas em vez de 4 |
| **Listas e tabelas** | `cadastros/_lista.html`, `_macros.html`, `termos_emitidos.html`, `termos_individuais.html`, `centro_custos.html`, `usuarios/` | Toda tabela dentro de `.br-table` com `overflow-x: auto` (nunca a página); abaixo de 768 px colunas secundárias escondidas (`d-none d-md-table-cell`) e as essenciais (identificador + ação) visíveis; paginação e filtros empilhados; botões de ação por linha viram menu `br-button circle` + dropdown quando houver mais de dois |
| **Formulários** | `cadastros.html`, `termo_devolucao.html`, `upload.html`, `usuarios/`, `acessos.html`, `senha.html`, `login.html` | Grid com `col-sm-12 col-md-6 col-lg-4` em vez de só `col-md`; campos com `width: 100%`; botões de submissão em bloco no mobile; `input type=number` com `inputmode=numeric` para patrimônio; datas com `type=date`; `select` do DSGov substituído pelo nativo abaixo de 768 px onde o do DSGov quebra |
| **Termos e documentos** | `termo.html`, `termo_base.html`, `termo_emitido.html`, `termos_emitidos.html` | Prévia (iframe) com altura calculada pela viewport e barra de ações (Copiar/DOCX/Enviar) sticky no topo do conteúdo; a tabela de bens do termo com `overflow-x` próprio; `termo_emitido.html` com os `col-md-2/3` reorganizados em duas linhas no mobile |
| **Análise e painéis** | `analise.html`, `index.html`, `inventario_relatorio.html`, `pesquisa.html` | Filtros num `br-collapse` fechado por padrão no mobile; gráficos ECharts com `resize()` no `window.resize` e altura mínima de 240 px; cards do Início em uma coluna abaixo de 576 px; KPIs em 2×2 |
| **Cabeçalho, menu e base** | `base.html`, `dsgov.css`, `menu-estado.js` | Menu lateral do DSGov fechando ao tocar fora e ao navegar; título do sistema encurtado abaixo de 576 px; breadcrumb com só o último nível no mobile (`br-breadcrumb` já suporta); flash messages não cobrindo o cabeçalho sticky; `html { -webkit-text-size-adjust: 100% }`; fonte mínima 16 px em inputs (evita zoom automático no iOS) |

Regras da fase:

- Cada correção é um seletor ou classe do DSGov aplicada ao template, ou um bloco `@media` em `dsgov.css` com
  comentário apontando a tela. Nada de `!important` sobre o `core.min.css`; se o componente do DSGov não responder,
  troca-se de componente (tabela → lista), não se sobrescreve o componente.
- Desktop não muda de aparência: capturas em 1280 px antes/depois comparadas no E2E (`expect(page).toHaveScreenshot`
  com tolerância pequena) para as telas tocadas.
- Zoom até 200 % sem scroll horizontal (WCAG 1.4.10) nas telas do grupo Inventário e Formulários.
- Ensaio manual no celular do usuário (Android/Chrome e, se houver, iOS/Safari) a cada grupo, com capturas no
  README da fase — o E2E não substitui a câmera nem o teclado virtual.

### 10.3 Testes

- `tests/e2e/test_responsivo.py`: para cada tela e cada largura (390, 768, 1280): `document.documentElement.scrollWidth
  <= innerWidth` (sem scroll horizontal), nenhum elemento interativo com `boundingBox` menor que 40×40 px,
  nenhum texto com `overflow: hidden` cortando conteúdo (`scrollWidth > clientWidth` em `.br-table td` fora do
  contêiner rolável), captura salva.
- Fluxo mobile completo em 390 px: login → Início → Inventário → sala → ler 3 bens → foto (upload simulado) → lote
  → voltar → painel. Já é o fluxo da F5, só que na largura menor.
- Teste unitário de template: as tabelas listadas em 10.2 renderizam dentro de `.br-table` (assert no HTML).

**DoD:** tabela do levantamento com todos os itens marcados "corrigido" ou "aceito assim" (com motivo); E2E
responsivo verde; capturas em três larguras antes/depois no README; desktop visualmente igual; ensaio manual no
celular registrado.

## 11. Fase 7 — Observabilidade e recuperação completa

Pré-condições: auditoria da F3, scripts da F0; decisão D7.

### 11.1 Logs, saúde e alertas (O02, F05)

- `logging` padrão do Python com `RotatingFileHandler` (10 MB × 5) em `dados/logs/web.log` e `robo.log`; linha
  `em nivel request_id usuario endpoint duracao_ms resultado mensagem`. `request_id` gerado em `before_request`,
  devolvido no header `X-Request-Id` e gravado na `auditoria` (coluna nova `request_id`). `_registrar()` do
  trabalhador vira `logging` e não engole `OSError` (loga em stderr). Exceção inesperada loga o traceback
  **fora** da transação que falhou.
- `GET /saude` (sem login, só de `127.0.0.1` ou com token em `secrets/saude.token`): JSON com `esquema_versao`,
  `fila_aguardando`, `fila_mais_antigo_min`, `ultimo_backup_horas`, `ultima_importacao_spw_horas`, sem nomes,
  contagens de acervo ou segredos. Card no Início do admin com os mesmos números e alerta amarelo quando fila
  > 60 min, backup > 26 h ou importação > 3 dias úteis.
- Alerta externo: `scripts/verificar_saude.sh` no cron (a cada hora) lê `/saude` e, fora dos limites, manda e-mail
  pelo `mail` do host (o servidor já tem cron; sem SMTP no app). Destinatário em `secrets/backup.env`.
- Capturas de erro do robô (`dados/robo/…png`) com retenção de 30 dias e permissão 600; aviso no README de que
  contêm dados pessoais.

### 11.2 Backup completo e restore verificado (R01)

- `backup.sh`: além do banco, empacota `timbrado/` customizado e `manifesto.json` (versão do esquema, commit,
  hash do banco, contagens). Chave Fernet e `SECRET_KEY` **não** entram no pacote: ficam em custódia separada
  (documentada em `docs/operacao/custodia.md` com quem guarda, sem o valor). Retenção 30 dias no R2 (D7) e cópia
  mensal para um segundo bucket com credencial sem permissão de apagar (imutabilidade prática).
- Fotos: versionamento de objetos ligado no bucket R2 (configuração fora do código, registrada) e
  `scripts/inventario_fotos.py` que compara chaves do banco com o bucket e lista faltantes/sobrando.
- `restaurar_teste.sh` da F0 vira procedimento completo: restaura em host/pasta limpa sem R2/SEI/SPW e sem
  trabalhador; `integrity_check`, `foreign_key_check`, contagens do manifesto, amostra de 5 termos com `html`,
  descriptografia de uma credencial sintética, referências de fotos; mede tempo total e idade do backup; escreve
  `dados/restore-YYYY-MM-DD.txt`. Rodado na publicação da F7 e a cada trimestre (cron `@monthly` roda só o
  `integrity_check` do último backup).
- Reconciliação de pedidos depois de restore: `scripts/reconciliar_pedidos.py` lista pedidos `aguardando` no banco
  restaurado cujo termo pode já ter documento no SEI (data do pedido anterior ao backup) e os coloca em `pausado`
  até um admin decidir. Rollback de release: README com o passo "restaurar backup do esquema anterior + imagem
  identificada pelo commit"; testado uma vez com a imagem da F5.

**DoD:** RPO/RTO medidos (D7) e aprovados pelo usuário; evidência de restore no README; custódia definida;
alerta disparado de propósito e recebido; logs rotacionando; `/saude` sem informação sensível (teste).

## 12. Fase 8 — Evolução funcional justificada

Pré-condições: nenhum P0/P1 do §15 aberto; F3 e F7 mescladas. Cada item só entra com demanda registrada
(pedido da comissão ou problema da auditoria) e critério de aceitação mensurável.

- **F03 — Relatório de inconsistências** (`/analise/inconsistencias`, só leitura): duplicados de atribuição (deve
  ser zero após F3), bem BAIXADO com carga, sala sem centro mapeado, termo cujo hash difere do que seria emitido
  hoje, foto com `excluida_em` pendente, pedido parado > 24 h, termo com número mas sem documento há > 7 dias.
  Cada linha aponta a tela de correção. Não corrige nada sozinho.
- **F04 — Retificação de termo:** `termos_emitidos.substitui_id` + `motivo_retificacao`. Botão "Retificar" no termo
  emitido cria nova emissão (fluxo normal da F3) referenciando a anterior; a anterior mostra "Retificado pelo
  termo #N" e mantém seu conteúdo; o SEI recebe documento novo pelo fluxo normal; a auditoria registra.
- **F01 — Linha do tempo do bem:** a ficha (`/bem`) ganha aba "Histórico" com a auditoria da F3 (responsável,
  localização, valor, termos, leituras) em ordem cronológica; registros anteriores à F3 aparecem com a marca
  "histórico parcial". Nada de UI ampla de versões além disso.
- **F07 — Assinatura integrada:** **não especificada**. Só entra se houver mecanismo oficial do SEI para
  confirmar assinatura e a comissão pedir; até lá o sistema não afirma que um termo foi assinado.
- QR de consulta, API externa, anexos genéricos, notificações: fora, salvo caso de uso concreto.

**DoD:** cada item com teste, demanda anotada na spec do item e sem depender de vulnerabilidade pendente.

## 13. Matriz de aceitação transversal

Resumo do §10 da auditoria; cada fase marca as linhas que fecha. Todas precisam estar verdes ao fim da F8.

| Área | Resultado exigido | Fases |
|---|---|---|
| Termos | Mesmo conteúdo/hash em prévia, DOCX, registro e fila; versão nova quando o conteúdo muda | F2, F3 |
| Numeração | Número único por unidade/tipo/ano; sem registro parcial | F0, F3 |
| Responsabilidade | Uma carga por bem; antes/depois/ator; termos anteriores intactos | F3 |
| Exclusão | Bloqueios corretos; nenhuma exclusão externa não rastreável | F3 |
| Permissões | Todo endpoint/método na matriz; inativo/revogado sem efeito; cookies revogados | F1, F2 |
| Importação | Vazio/fracionário/redução recusados sem alterar nada; histórico real | F0, F3 |
| Exportação | Sem fórmulas; 201+ termos; restore de histórico válido | F2, F3 |
| Inventário | Nenhuma gravação após finalizar; lote parcial explícito | F0 |
| Fotos | Chaves únicas; compensação observável; URL validada | F0, F2, F3 |
| SEI/SPW | Fila sem duplicar; revogação; dois trabalhadores | F2, F3 |
| UI | Sem perda de seleção/filtro/foco; erro não finge salvamento; usável no celular | F5, F6 |
| Recovery | Restore íntegro com fotos/chaves; pedidos reconciliados | F0, F7 |

## 14. O que não entra (em nenhuma fase)

Django, microsserviços, SPA, Kubernetes, event sourcing, transação distribuída, PostgreSQL, Redis/Celery,
CSP antes de inventariar scripts inline, prova de assinatura sem mecanismo oficial, reescrita do DSGov/CSS,
remoção do modo desktop, reconstrução de histórico com dados atuais, saneamento automático de duplicados,
execução de scripts de spike contra SEI/banco real na suíte, remoção das dependências desktop.
Também fora: a "auditoria visual ampliada" (P4) e tudo do §12 sem demanda.

## 15. Rastreabilidade (ID da auditoria → fase → seção)

| ID | Prioridade | Fase | Seção |
|---|---|---|---|
| B01, B02, B11 | P0/P1 | 0 | 4.1 |
| B03 | P0 | 0 | 4.2 |
| B07 | P0 | 0 | 4.3 |
| B08 | P0 | 0 | 4.4 |
| R02, R03 | P1 | 0 | 4.5 |
| S07 (ferramenta), D03 | P1/P2 | 1 | 5.1–5.2 |
| Matriz de rotas, esquema legado, E2E infra | — | 1 | 5.3 |
| S01, S02, S04 | P1 | 2 | 6.1 |
| S03, B06, S09 | P1 | 2 | 6.2 |
| S08, B13/U08, B14/U07, S05, S06, contador de login | P2/P1 | 2 | 6.3 |
| Tabela `auditoria` | — | 2 | 6.4 |
| B10 | P1 | 3 | 7.1 |
| B04, B05, B12, O01, D4 | P1 | 3 | 7.2 |
| B15, B16, B09 | P1/P2 | 3 | 7.3 |
| Fila atômica, heartbeat | prov. | 3 | 7.4 |
| D01 | P2 | 4 | 8.1 |
| D02 | P2 | 4 | 8.2 |
| P01–P05, índices, WAL, dinheiro, datas | P2/P3 | 4 | 8.3 |
| U01–U08, F06 | P2 | 5 | 9 |
| Responsividade/mobile (pedido do usuário, 2026-09-21) | P2 | 6 | 10 |
| O02, F05 | P2 | 7 | 11.1 |
| R01 | P1 | 7 | 11.2 |
| F01, F03, F04, F07 | P1–P3 | 8 | 12 |
| D04 (JS/CSS residual) | P3 | — | só depois de inventário de usos; sem fase reservada |

## 16. Ordem de publicação e riscos de cada fase

| Fase | Muda esquema? | Risco principal | Mitigação |
|---|---|---|---|
| 0 | índice tentado | índice falhar por duplicado existente | tenta a cada startup; app sobe; diagnóstico |
| 1 | não | build quebrar com pins | CI antes do rebuild em produção |
| 2 | `sessao_versao`, `auditoria`, `cancelado` | todo mundo deslogado na publicação; `TERMOS_MODO` faltando | avisar a comissão; compose atualizado no mesmo commit |
| 3 | uid, snapshot, leituras, `excluida_em`, `pulsacao`, índice | conciliação de duplicados; tarja em termos antigos | backup + restore ensaiado; comissão aprova a conciliação antes |
| 4 | `schema_versao`, índices | migrador único: robô esperar | `esperar_esquema` com timeout claro |
| 5 | não | quebrar câmera/clipboard no celular | ensaio manual no celular antes do push |
| 6 | não | regressão visual no desktop ao mexer no CSS | capturas antes/depois em 3 larguras; E2E em 390 px |
| 7 | `request_id` | segundo bucket, custódia | tudo documentado sem valores |
| 8 | `substitui_id` | escopo crescer | cada item com demanda anotada |

Estimativa relativa (mesma escala das fases anteriores do projeto): F0 uma sessão; F1 uma; F2 duas; F3 três a
quatro; F4 duas; F5 duas; F6 duas; F7 duas; F8 uma por item.
