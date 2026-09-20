# Spike — escrever no SEI (incluir documento, colar termo, incluir em bloco)

**Data:** 2026-09-20. **Veredito: PASS.** Código descartável: os scripts aqui são a prova, não o robô.
**Spec:** `../../specs/2026-09-20-envio-sei-design.md` (§5.1 e §10 ajustadas com o que se provou aqui).

Rodou com o login do Antônio (`antonio.junior`, unidade GELIC) contra o processo de rascunho
`90796110000022.000059/2026-88` e o bloco "Termos TESTE" (nº 69766), criado à mão por ele.
Resultado: **21 termos por centro de custo criados e incluídos no bloco, sem duplicata**, mais três de teste
(01, 02 e 03/2026 - TESTE). Ambiente: `.venv-robo` do projeto (Playwright + Chromium headless).

## Respostas às perguntas da §10 da spec

| # | Pergunta | Resposta provada |
|---|---|---|
| 1 | Incluir documento | Selecionar a raiz na árvore (`#anchor<id>` em `ifrArvore`) → no frame com a barra de ações (`ifrConteudoVisualizacao`), `a:has(img[title='Incluir Documento'])` → tela "Gerar Documento" com a lista de tipos: `a[onclick^='escolher']` filtrado pelo texto exato ("Termo de Responsabilidade" = `escolher(334)`) |
| 2 | Rótulo na árvore | O tipo **não mostra o campo Número**; o rótulo inteiro vai em **`#txtNomeArvore` = "01/2026 - GESERV"** e a árvore exibe "Termo de Responsabilidade 01/2026 - GESERV (1557099)". Texto inicial "Nenhum" (`#optNenhum`) já vem marcado. Nível Público: clicar em **`label[for=optPublico]`** (o radio fica coberto pelo label) e esperar `#optPublico.checked` |
| 3 | Editor | `#btnSalvar` abre um **popup** (`context.expect_page`) com três CKEditors; o do corpo é o que contém `iframe[title="Corpo do Texto"]` (id `cke_txaEditor_NNNN` → instância `txaEditor_NNNN`). `CKEDITOR.instances[x].setData(html)` **substitui o modelo inteiro do tipo** (cabeçalho do CFC fica, pois é outro editor) e `a[title^='Salvar']:visible` grava; fechar o popup. Reaberto, o documento mostra nosso termo |
| 4 | Número SEI | Sai do rótulo na árvore: `^(?P<rotulo>.*?)\s*\((?P<numero>\d{6,8})\)$` |
| 5 | Bloco | Selecionar o documento na árvore → barra do documento → `a:has(img[title='Incluir em Bloco de Assinatura'])` → `#selBloco` com opções "**69766 - Termos TESTE**" (casar por `endswith(" - " + nome)`) → `#sbmIncluir`. **Confirmação = a linha do documento na tabela passa a mostrar o nº do bloco** na coluna "Blocos" (e perde a caixa de seleção). Há também `#sbmIncluirDisponibilizar`, que não usamos. Estado/unidade do bloco não importam: o robô só casa o nome (decisão do usuário) |
| 6 | Tempo e instabilidade | ~10 s para criar + ~5 s para o bloco = **~15 s por termo**, independente do tamanho (COMUNICA 100 KB e GESERV 660 KB de HTML levaram o mesmo). Instabilidades encontradas e resolvidas abaixo |

## O que quebrou e como o robô deve tratar

1. **Árvore em pastas.** Com mais de ~20 documentos o SEI agrupa a árvore em "Pasta I, II…" fechadas, com um nó
   `anchorAGUARDE`. Qualquer busca por rótulo precisa antes clicar em `img[title='Abrir todas as Pastas']` e esperar
   a contagem de nós estabilizar (mesmo truque do coletor do PCA). Sem isso, "documento não apareceu na árvore" —
   e o documento existia. Um processo real de termos chega nisso rápido.
2. **Documento grande no visualizador.** No GESERV (1.464 linhas), o clique em "Incluir em Bloco" foi engolido
   enquanto o `ifrVisualizacao` ainda renderizava. Esperar `wait_for_load_state("load")` do frame antes de
   clicar e repetir o clique uma vez se `#selBloco` não vier em 20 s.
3. **Confirmar pelo estado, não pela exceção.** O GESERV *estava* no bloco apesar do timeout. Antes de incluir, ler
   a linha do documento: se já traz o nº do bloco, não clicar. Isso torna o passo idempotente.
4. **Recuperação após erro.** Fechar popups sobrando (`ctx.pages[1:]`), reabrir o processo pela pesquisa rápida e
   esperar a árvore; sem isso os pedidos seguintes falham em cascata (`arvore()` vazio).
5. **Idempotência do documento.** `no_da_arvore(rótulo)` antes de criar: nas repetições nenhum dos 21 foi duplicado.

## Formatação (pedido do usuário durante o spike)

O editor do SEI descarta `<style>` e `class=`; só respeita `style=` inline. O gerador atual usa classes
(`semrecuo`, `centro`, `direita`, `assinatura`) para justificar/recuar, então no SEI o texto saía sem
justificar. Provado com "03/2026 - TESTE": `style="text-align:justify;text-indent:1.25cm"` nos parágrafos e
`width:90%` na tabela **são preservados** (`evidencias/formatacao-inline.png`). Decisão do usuário: **texto
justificado e tabela com 90%**, feito no gerador do sistema (`termos_html.py`) — vale para o *Copiar* também.

## Pontos em aberto

- A lista "Escolha o Tipo do Documento" mostrou 77 tipos e **não tem "Termo de Devolução"** (há "Termo", "Termo de
  Responsabilidade", "Termo de Cessão de Uso"…). O botão "+" ao lado do título pode listar todos os tipos do
  órgão; a lista está em `evidencias/tipos-de-documento.json`. Decisão do usuário: qual tipo o termo de devolução usa.
- Os 24 documentos de teste ficam no processo de rascunho; o bloco "Termos TESTE" não foi disponibilizado.

## Arquivos

- `sei_base.py` — login, abrir processo, árvore (com abertura de pastas), frames por conteúdo.
- `sei_acoes.py` — `criar_documento`, `incluir_em_bloco`, `no_da_arvore`: o esqueleto do futuro `robo_sei.py`.
- `passo1_abrir.py` … `passo5_bloco.py` — exploração passo a passo; `passo7_todos.py` — o lote dos 21;
  `passo10_formatacao.py` — teste de estilo inline.
- `evidencias/` — prints (árvore com os 21, bloco incluído, editor com o modelo do tipo, formatação) e os JSON dos lotes.
