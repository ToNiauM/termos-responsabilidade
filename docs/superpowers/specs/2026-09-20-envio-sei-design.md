# Envio ao SEI — documento no processo, bloco de assinatura e "Atualizar agora"

**Data:** 2026-09-20.
**Estado:** desenho aprovado; spike da §10 concluído em 2026-09-20 (PASS); prioridade máxima do usuário — pronto para o plano.
**Base:** `main`, commit `20fac60`, depois da spec da administração de inventários (ainda não implementada).
**Plano:** `../plans/2026-09-20-envio-sei.md` (a escrever).
**Spike:** `../notes/2026-09-20-spike-sei-escrita/README.md` (21 termos reais criados e incluídos em bloco no processo de rascunho).

## 1. Resultado e limites

Hoje, depois de *Copiar para o SEI*, o operador ainda faz à mão: incluir o documento no processo, colar o
termo no editor, incluir o documento no bloco de assinatura do setor e anotar no sistema o número do
documento e do bloco. Esta entrega troca isso por um botão **Emitir Termo no SEI** na página do termo: o sistema
registra a emissão (como o *Copiar* já faz), numera o termo (`NN/AAAA` por unidade e ano) e enfileira um
pedido; um robô no host abre o SEI com Playwright, cria o documento no processo vigente do tipo, cola o
termo, inclui o documento no bloco **"Termos {UNIDADE}"** (criado à mão por vocês) e grava `documento_sei`
e `bloco_sei` no registro. A partir daí o pedido de assinatura por e-mail (`mailto:`) aparece como hoje.

O mesmo mecanismo de fila atende um botão **Atualizar com SPW** na tela *Atualizar base*, que roda a importação do
SPW sob demanda sem terminal. Nas telas, nada de "robô": os textos falam do que acontece ("Emitindo no SEI…",
"Atualizando com o SPW…"); o nome só existe no código, nos logs e neste documento.

Não entra: criar ou disponibilizar bloco no SEI; assinar; enviar e-mail pelo SEI ou por SMTP; assistente
com LLM (decidido explicitamente: o fluxo é fixo, um modelo probabilístico só acrescentaria custo e falhas);
web service do SEI (a TI não libera); Playwright dentro da imagem Docker; vários pedidos em paralelo.
Manter Flask, SQLite, Jinja e os componentes DSGov. O modo desktop (sem login) não mostra os botões — não há
robô lá.

## 2. Decisões do usuário (2026-09-20)

| Pergunta | Decisão | Motivo |
|---|---|---|
| "Já temos a API" | É a da OpenAI, não do SEI | — |
| Acesso ao SEI | Robô de tela (Playwright), como o robô do SPW e o do PCA | TI não libera o SEI WS |
| Determinístico × LLM | Botão determinístico; LLM fica para depois, se ainda fizer sentido | Sequência fixa de passos |
| Credencial | A do próprio Antônio (`antonio.junior`), em `secrets/sei.env` | Já provada no robô do PCA, sem captcha/2FA |
| Bloco de assinatura | Criado à mão, um por setor, nome **`Termos {UNIDADE}`**; o robô só inclui o documento | Tira do robô o passo mais arriscado |
| Disponibilizar bloco | Continua manual (SEPAT) | Idem |
| Unidade SEI do centro de custo | A própria sigla (`ccustos`), com campo opcional para exceção | Siglas já batem com o SEI |
| Unidade SEI da pessoa | Campo novo `unidade_sei` na pessoa, obrigatório para enviar | Pessoa não tem setor no cadastro |
| Devolução | Mesma regra de responsabilidade (unidade de quem devolve) | — |
| Tipo de documento no SEI | Um por tipo de termo, nome guardado em Textos; o processo já é o vigente do tipo | — |
| Conteúdo | O mesmo HTML do *Copiar*, colado no editor | — |
| Nome na árvore | "Termo de Responsabilidade NN/AAAA - UNIDADE"; sequência **por unidade e ano**, gerada pelo sistema | Robô não lê a árvore para "descobrir" o último |
| Nível de acesso | Sempre público | — |
| Depois de salvar | Nada: só incluir no bloco; quem assina é o responsável | — |
| E-mail | O `mailto:` de hoje basta | Sem dependência nova |
| Atualizar base pelo site | Sim, botão "Atualizar agora" pela mesma fila | Quase de graça com a fila pronta |
| Arquitetura | Fila em `termos.db` + trabalhador no host (`.venv-robo`, systemd) | Banco e venv já compartilhados; sem Chromium na imagem |
| Conteúdo no editor | Substitui **todo** o modelo do tipo (o SEI já traz um texto padrão) pelo termo gerado | Decisão do usuário durante o spike |
| Formatação | Parágrafos justificados (recuo 1,25 cm; abertura/centro/direita sem recuo) e tabela com **90%** de largura, tudo em `style=` inline no gerador | O editor do SEI descarta CSS e classes; decisão do usuário (2026-09-20) |
| Bloco: estado e unidade | Irrelevantes; o robô casa só o nome na lista `#selBloco` | Decisão do usuário |
| Nomes dos botões | **"Atualizar com SPW"**, **"Emitir Termo no SEI"**, **"Enviar email"**; a interface **nunca usa a palavra "robô"** nem explica como faz | Decisão do usuário (2026-09-20) |
| Título do termo | Menor e centralizado (`h1` inline: 14pt, `text-align:center`), no HTML gerado | Decisão do usuário (2026-09-20) |
| Tipo do termo de devolução | "Termo de Devolução" existe, mas só na lista completa: clicar `#ancExibirSeries` ("Exibir todos os tipos") antes de procurar | Provado no spike |
| Prioridade | Este fluxo antes da Administração de inventários | Decisão do usuário (2026-09-20) |

## 3. Dados (`db.py`, migrações em `criar_esquema`)

Colunas novas:

- `responsaveis.unidade_sei TEXT` — opcional; vazio significa a própria sigla `ccustos`.
- `pessoas.unidade_sei TEXT` — opcional no cadastro, **obrigatório para enviar**.
- `termos_emitidos.numero_termo TEXT` — `NN/AAAA` (dois dígitos no mínimo; `100/2026` se passar de 99).
- `termos_emitidos.unidade_sei TEXT` — a unidade usada no envio, congelada no registro.

Tabela nova:

```sql
CREATE TABLE IF NOT EXISTS robo_pedidos (
  id           INTEGER PRIMARY KEY,
  tipo         TEXT NOT NULL CHECK (tipo IN ('sei','spw')),
  termo_id     INTEGER REFERENCES termos_emitidos(id) ON DELETE CASCADE,   -- só para 'sei'
  html         TEXT,                                                       -- corpo do termo, só para 'sei'
  criado_em    TEXT NOT NULL,
  criado_por   TEXT,                                                       -- login do usuário; NULL no desktop
  iniciado_em  TEXT,
  terminado_em TEXT,
  passo        TEXT NOT NULL DEFAULT 'aguardando'
               CHECK (passo IN ('aguardando','login','documento','bloco','rodando','concluido','erro')),
               -- 'login'/'documento'/'bloco' são do pedido 'sei'; 'rodando' é do pedido 'spw'
  mensagem     TEXT
);
CREATE UNIQUE INDEX IF NOT EXISTS robo_pedidos_ativo_termo ON robo_pedidos(termo_id)
  WHERE tipo = 'sei' AND passo NOT IN ('concluido','erro');
CREATE UNIQUE INDEX IF NOT EXISTS robo_pedidos_ativo_spw ON robo_pedidos(tipo)
  WHERE tipo = 'spw' AND passo NOT IN ('concluido','erro');
```

Funções novas em `db.py`:

- `unidade_sei(conn, tipo, chave) -> str`: `ccusto` → `responsaveis.unidade_sei or ccustos`; `individual` e
  `devolucao` → `pessoas.unidade_sei` ou `ErroDeNegocio("Cadastre a unidade SEI de <nome> em Cadastros → Pessoas.")`.
- `proximo_numero_termo(conn, unidade, ano) -> str`: `1 + max(sequencial)` entre os `termos_emitidos` com a
  mesma `unidade_sei` e `numero_termo LIKE '%/AAAA'`; nunca reaproveita (o número fica no registro mesmo se o
  envio falhar).
- `salvar_numero_termo(conn, id, numero)`: só enquanto não há `documento_sei`; valida `^\d{2,}/\d{4}$`.
- `enfileirar_pedido(conn, tipo, termo_id=None, html=None, criado_por=None) -> int`: `ErroDeNegocio` se o
  índice parcial recusar ("Já há um envio deste termo em andamento." / "O robô do SPW já está rodando.").
- `pedido_do_termo(conn, termo_id) -> dict | None`: o mais recente do termo.
- `pedido_spw_ativo(conn) -> dict | None`.
- `proximo_pedido(conn) -> dict | None`: o mais antigo em `aguardando`.
- `marcar_passo(conn, id, passo, mensagem=None)`: grava `passo`; `iniciado_em` no primeiro passo depois de
  `aguardando`; `terminado_em` em `concluido`/`erro`. Commit imediato (o site lê enquanto o robô anda).
- `salvar_documento_sei` (existe) passa a aceitar gravar só o documento, mantendo o bloco vazio.

Textos (`textos.py`): grupo novo "Envio ao SEI" com `sei_tipo_ccusto`, `sei_tipo_individual`,
`sei_tipo_devolucao` — nome exato do tipo de documento na lista do SEI (padrão: "Termo de Responsabilidade"
para os dois primeiros e "Termo de Devolução" para o terceiro; o spike confirma os nomes reais). Editável em
Textos como os demais.

Gerador (`termos_html.py`): `_p`, `_abertura` e as tabelas passam a emitir estilo inline — parágrafo comum
`text-align:justify;text-indent:1.25cm;margin:0 0 7pt`, `semrecuo` idem sem recuo, `centro`/`direita`/`assinatura`
como no CSS de hoje — o `h1` sai com `text-align:center;font-size:14pt;margin:10pt 0 12pt` (menor e centralizado,
como o usuário pediu ao ver o resultado no SEI) e as três tabelas ficam com `width:90%`. O `termo_base.html` continua com o CSS (o `.docx`
e a tela não mudam de aparência); só deixa de ser a única fonte do alinhamento.

Cadastros: `unidade_sei` entra no formulário de Responsáveis e de Pessoas e como coluna opcional na
importação de planilha (`importar_planilhas.py`), sem quebrar planilhas antigas.

Segredos: `secrets/sei.env` com `SEI_USUARIO`, `SEI_SENHA`, `SEI_LOGIN_URL`, `SEI_ORGAO`, lido por um `ler_env`
comum (o de `importar_spw.py` vira função compartilhada que recebe o caminho e a lista de chaves).

## 4. Fluxo no site

### 4.1 Emitir Termo no SEI

`POST /termo/<tipo>/<chave>/enviar-sei` (novo, ao lado de `termo_registrar`):

1. Exige processo vigente do tipo (mesma mensagem do *Copiar*) e `unidade_sei` (§3).
2. `registrar_emissao` com os bens de hoje e o `corpo_html` de `_bens_do_termo` — o mesmo que o *Copiar* usa.
3. Atribui `numero_termo` e `unidade_sei` ao registro; enfileira o pedido `sei` com o HTML.
4. Redireciona para `/termos-emitidos/<id>`.

Na página do termo o botão **Emitir Termo no SEI** fica ao lado de *Copiar para o SEI*, como `form method="post"` (CSRF como as
demais rotas POST). O *Copiar* continua existindo, para quem quiser colar à mão.

### 4.2 Página do termo emitido (`/termos-emitidos/<id>`)

Estados, decididos pelo pedido mais recente do termo e pelos campos do registro:

| Situação | Mostra |
|---|---|
| Sem pedido e sem `documento_sei` | campos de hoje (documento/bloco à mão) **e** botão **Emitir Termo no SEI** (`POST /termos-emitidos/<id>/enviar-sei`, que enfileira sem registrar de novo) |
| Pedido `aguardando` / `login` / `documento` / `bloco` | "Emitindo no SEI: <passo>…" ("entrando no SEI", "criando o documento", "incluindo no bloco") e `<meta http-equiv="refresh" content="5">`; campos e botões escondidos. Se `aguardando` há mais de 2 min: aviso "A emissão ainda não começou; avise o administrador." |
| Pedido `concluido` | "Emitido no SEI em dd/mm/aaaa hh:mm: documento X, bloco Y" e o botão **Enviar email** (o `mailto:` de hoje, renomeado) |
| Pedido `erro` com `documento_sei` gravado | mensagem do erro, documento X, botão **Incluir no bloco** (enfileira de novo; o robô pula o passo `documento`) |
| Pedido `erro` sem documento | mensagem do erro e botão **Emitir Termo no SEI** de novo |

O número do termo aparece no cabeçalho do registro ("Termo 03/2026 - GESERV") e é editável no formulário
existente enquanto `documento_sei` está vazio — serve para acertar o ponto de partida quando a unidade já
tem termos numerados à mão em 2026.

Enquanto houver pedido ativo, a lista *Termos emitidos* marca a linha com "enviando…".

### 4.3 Atualizar com SPW

Na tela *Atualizar base*, acima da tabela de execuções: botão **Atualizar com SPW**
(`POST /atualizar-base/spw`, só admin). Enfileira um pedido `spw`. Enquanto houver pedido `spw` ativo, a
tela mostra "Atualizando com o SPW…", recarrega a cada 5 s e o botão fica desabilitado. Ao terminar,
o resultado já está na tabela de execuções (`robo_execucoes`, gravada pelo próprio `importar_spw.executar`) e
o card do Início reflete como hoje. O cron das 3h continua.

## 5. Robô do SEI (`robo_sei.py`, raiz, ao lado de `importar_spw.py`)

Roda só no host, no `.venv-robo`. Mesma separação do SPW: Playwright e seletores só na classe `SEI`;
orquestração pura e testável.

### 5.1 Classe `SEI` (seletores provados no spike)

Login, `abrir_processo`, acesso aos frames copiados de `/opt/web/pca-cfc/apps/pca/sei/coletor.py`; o esqueleto
das ações novas está em `../notes/2026-09-20-spike-sei-escrita/sei_acoes.py`. Os iframes do SEI têm URL pouco
confiável: localizar frames **pelo conteúdo** (`frame_com(seletor)`), não pela URL.

- `abrir_processo(numero)`: pesquisa rápida → título `SEI - <numero>` → espera a árvore; se houver pastas fechadas
  (nó `anchorAGUARDE`, acontece a partir de ~20 documentos), clica em `img[title='Abrir todas as Pastas']` e espera a
  contagem de nós estabilizar. Toda busca por rótulo passa por isso.
- `documento_na_arvore(rotulo) -> (id, numero) | None`: casa `^<rotulo>\s*\((\d{6,8})\)$` nos nós da árvore.
- `incluir_documento(tipo_nome, nome_arvore, html) -> numero`: raiz selecionada → `a:has(img[title='Incluir Documento'])`
  → clica `#ancExibirSeries` ("Exibir todos os tipos": 77 → 226 tipos) → lista `a[onclick^='escolher']` com texto exato (`RoboErro` se não houver) → formulário: `#optNenhum` já
  marcado, **`#txtNomeArvore` = "NN/AAAA - UNIDADE"** (o tipo não tem campo Número), Público via
  `label[for=optPublico]` + espera de `#optPublico.checked` → `#btnSalvar` abre popup (`expect_page`) → espera
  `CKEDITOR` e `iframe[title="Corpo do Texto"]` → instância `txaEditor_NNNN` do container desse iframe →
  `setData(html)` + `fire('change')` → `a[title^='Salvar']:visible` → fecha o popup → espera o rótulo na árvore
  (com abertura de pastas) e devolve o número.
- `incluir_em_bloco(numero, nome_bloco) -> numero_bloco`: seleciona o documento na árvore → **espera o
  `ifrVisualizacao` terminar de carregar** (documentos grandes engolem o clique) →
  `a:has(img[title='Incluir em Bloco de Assinatura'])`, repetindo o clique uma vez se `#selBloco` não vier em 20 s →
  opção que termina em `" - " + nome_bloco` (`RoboErro` se não houver; o texto é "69766 - Termos TESTE") → **antes
  de clicar, lê a linha do documento: se já mostra o nº do bloco, devolve sem clicar** → `#sbmIncluir` → confirma
  pela coluna "Blocos" da linha do documento.

Timeout padrão 30 s por operação; uma sessão por pedido; `logout` ao final mesmo em erro; em exceção, fecha popups
sobrando, salva `dados/sei/erro.png` (sobrescrito) e relança. Medido: ~15 s por termo, independente do tamanho
(660 KB de HTML no GESERV).

### 5.2 `enviar_termo(conn, pedido, sei=None, agora=None) -> dict`

Orquestra, gravando `marcar_passo` **antes** de cada passo:

1. `login` — `sei.login(env)`; `autenticado` falso → erro "credenciais recusadas".
2. `documento` — pulado se o termo já tem `documento_sei`. Abre o processo `numero_sei` do registro; título
   diferente do esperado → erro **antes** de criar qualquer coisa. `incluir_documento(...)`; grava
   `documento_sei` **na hora** (commit), antes de seguir.
3. `bloco` — `incluir_em_bloco(documento_sei, f"Termos {unidade}")`; grava `bloco_sei`.
4. `concluido`.

Devolve `{"passo", "mensagem", "documento_sei", "bloco_sei"}`. `sei` é injetável para os testes.

### 5.3 Nunca duplicar documento

Se o passo `documento` estourar timeout **depois** de salvar (a árvore não respondeu), o documento pode
existir sem que o sistema saiba. Na retomada de um pedido em erro sem `documento_sei`, o robô primeiro chama
`documento_na_arvore("Termo de Responsabilidade NN/AAAA - UNIDADE")` (com as pastas abertas): se achar, grava o número e segue para o bloco; só cria
se não achar. O rótulo com número por unidade e ano é o que torna essa checagem determinística.

## 6. Trabalhador (`atender_pedidos.py`) e serviço

- Laço: a cada 3 s, `db.proximo_pedido`; `sei` → `robo_sei.enviar_termo`; `spw` → `importar_spw.executar`
  (que já grava `robo_execucoes`), com `marcar_passo` `rodando`/`concluido`/`erro` só para a tela
  acompanhar. Um pedido por vez. Exceção não prevista → `erro` com a mensagem curta e traceback em
  `dados/robo_pedidos.log`; o laço continua.
- Lock por arquivo `dados/robo.lock` (`fcntl.flock`), compartilhado com `atualizar_base.sh` e o cron do SPW:
  quem não pega o lock espera (cron) ou sai avisando (script à mão).
- `termos-robo.service` (systemd, `Restart=always`, usuário ToNiauM, `WorkingDirectory` no checkout,
  `ExecStart=.venv-robo/bin/python atender_pedidos.py`), instalado à mão pelo README. Sem ele o site funciona:
  pedidos ficam `aguardando` e a tela avisa depois de 2 min. Reinício do serviço no meio de um pedido:
  o pedido em passo intermediário há mais de 10 min volta a `aguardando` ao subir (a retomada da §5.3 cobre o
  documento já criado).
- `atualizar_base.sh` continua, agora respeitando o lock.

## 7. Erros que o usuário vê

Sempre uma frase em `mensagem`, mostrada na página do termo ou em Atualizar base:

| Causa | Mensagem |
|---|---|
| `secrets/sei.env` ausente/incompleto | "secrets/sei.env não encontrado ou incompleto: falta …" |
| Login recusado | "O SEI recusou usuário ou senha." |
| Processo não abre / título diferente | "Processo <nº> não abriu no SEI; nada foi criado." |
| Tipo de documento inexistente | "Tipo de documento '<x>' não existe no SEI; corrija em Textos." |
| Bloco inexistente | "Bloco 'Termos <UNIDADE>' não existe no SEI; crie o bloco e clique em Incluir no bloco." |
| Timeout | "O SEI não respondeu a tempo ao <passo em palavras>." |
| Pessoa sem unidade | (antes de enfileirar) "Cadastre a unidade SEI de <nome> em Cadastros → Pessoas." |

## 8. Permissões e menu

`permissoes.py`: `termo_enviar_sei` (POST nas duas rotas de emissão e em "Incluir no bloco") = admin e operador,
igual a `termo_docx`; `base_atualizar_spw` = admin. `pode('termo_enviar_sei')` esconde os botões; com
`TERMOS_LOGIN` desligado (desktop) os botões não aparecem. Nenhum item de menu novo.

## 9. Testes

Sem SEI nem SPW reais; Playwright nunca importado na suíte (imports lazy, como no SPW).

- `tests/test_db.py`: `unidade_sei` (sigla, exceção, pessoa sem unidade), `proximo_numero_termo` (por
  unidade e ano, zero-padding, não reaproveita), `salvar_numero_termo` (recusa depois do documento), índices
  parciais (segundo pedido do mesmo termo → `ErroDeNegocio`; `spw` duplicado idem), `marcar_passo` e datas.
- `tests/test_app.py`: `enviar-sei` registra emissão + número + pedido e redireciona; recusa sem processo
  vigente e sem unidade; estados da página do termo emitido (5 linhas da tabela §4.2, meta refresh só nos
  ativos, aviso de 2 min); "Incluir no bloco"; `atualizar-base/agora` e a tela com pedido ativo; botões por
  perfil (`test_permissoes.py`) e escondidos no desktop.
- `tests/test_robo_sei.py`: `enviar_termo` com `SEI` falso — sequência de passos gravada; falha no bloco deixa
  `documento_sei` e passo `erro`; retomada pula `documento`; retomada sem documento consulta a árvore e não
  cria se já existe; processo com título errado não chama `incluir_documento`; logout sempre chamado.
- `tests/test_atender_pedidos.py`: laço com executores falsos — ordem FIFO, um por vez, exceção vira `erro`
  e não derruba o laço, pedido órfão volta a `aguardando`.

## 10. Spike (concluído em 2026-09-20 — PASS)

Relatório em `../notes/2026-09-20-spike-sei-escrita/README.md`, com scripts e evidências. Rodou com o login do
usuário no processo de rascunho `90796110000022.000059/2026-88` e no bloco "Termos TESTE" (nº 69766): criou os
**21 termos por centro de custo** (rótulos "01/2026 - <CC>") e incluiu todos no bloco, sem duplicata mesmo com
repetições após erro. As seis perguntas estão respondidas na §5.1; os incidentes (árvore em pastas, documento
grande, confirmação por estado, recuperação após erro) viraram regras da §5.1 e §5.3.

Complemento (mesmo dia): a lista inicial mostra só 77 tipos; `#ancExibirSeries` ("Exibir todos os tipos") expande
para 226 e **"Termo de Devolução" existe** (`evidencias/tipos-de-documento-todos.json`). O robô sempre expande antes
de procurar.

## 11. Publicação

1. `docker compose up -d --build` (migração roda no primeiro start).
2. `secrets/sei.env` (senha atual do SEI; o usuário disse que vai trocar as credenciais expostas).
3. `sudo systemctl enable --now termos-robo` (unidade em `ops/termos-robo.service`).
4. No SEI: criar os blocos "Termos {UNIDADE}" das unidades que recebem termo.
5. Em Cadastros: `unidade_sei` das pessoas que recebem termo individual; exceções em Responsáveis.
6. Em Textos: conferir os nomes dos tipos de documento.
7. Evidência: um termo real enviado, com print da árvore do SEI e do bloco, anexados nesta spec (§12).

## 12. Evidências

(a preencher na publicação)
