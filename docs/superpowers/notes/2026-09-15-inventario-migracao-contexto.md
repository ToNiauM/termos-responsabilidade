# Migração do módulo de inventário — contexto para retomar (2026-09-15)

Se a conversa foi limpa (/clear), leia este arquivo e o levantamento
`docs/superpowers/notes/2026-09-15-inventario-existente.md` antes de qualquer coisa. Depois siga o
fluxo brainstorming → spec → plano → subagent-driven-development, como nas rodadas anteriores.

## Pedido do usuário (palavras dele, resumidas)
- Migrar o **módulo de inventário** do sistema existente (sistemadeinventario.com.br = `/opt/web/sga/cfc`,
  Flask, roda a partir de planilha, container `sga-cfc` na porta 11001) para ESTE sistema
  (`/opt/web/termos-responsabilidade`), em SQLite e com a cara do gov.br (DSGov).
- "A mudança é só essa. Mais simples possível." O código e as regras já existem no outro projeto:
  reaproveitar a lógica, não reinventar.
- Fluxo: escolhe uma **sala** → carrega os bens daquela sala → leitura por **código de barras** →
  bem registrado na sala = **confirmado**; bem lido que não é da sala = **divergente**.
  Um **evento de inventário** deve ser registrado.
- Migrar também a **regra das fotos** e a **exportação** do sistema existente.
- O módulo deve ficar **em separado do restante** (menu/telas próprias) para poder iniciar.
- Novas tabelas no `termos.db` (só acréscimos, como sempre).

## Estado deste repositório
- `main` em af16a39 (rodada "processos SEI, histórico, painel e recorte" concluída, publicada e com push).
- Base de dados: `dados/termos.db` (produção do site; bind mount do container `termos-patrimonio`, porta 12012).
- `bens` (7.003; 3.101 ativos) vem do export do SPW; `bens.localizacao` é a "sala" (100 localizações ativas);
  `localizacoes` → `responsaveis` (centro de custo); `pessoas`/`atribuicoes`; `processos_sei`,
  `termos_emitidos(+_bens)`, `importacoes(+_mudancas)`.
- Padrões: `db.py` (funções com `conn` primeiro, sem Flask), `app.py` (rotas), `painel.py`/`graficos.py`
  (ECharts), templates DSGov com macros em `templates/_macros.html`, testes em `tests/` (122 passando),
  `.venv/bin/pytest -q`. Docker: `docker compose up -d --build`. Commits com rodapé Co-Authored-By/Claude-Session.

## Preferências firmes do usuário
- Simplicidade acima de tudo; sem Django/Postgres; SQLite; DSGov só visual.
- Gráficos: ≤5 rosca, 6–10 barras, 11–20 colunas, tudo clicável.
- Bem atribuído a pessoa tem responsável (não é "sem centro").
- Não excluir registros de histórico (auditoria).

## Próximos passos
1. Ler o levantamento e fazer as perguntas de desenho (uma por vez): o que é o "evento" (por sala? por
   campanha anual?), o que fazer com divergentes (só registrar? propor mover a localização?), fotos
   (por bem? obrigatórias?), exportação (quais colunas), quem opera (celular com câmera? leitor USB?).
2. Spec em `docs/superpowers/specs/2026-09-15-inventario-design.md`; plano em `docs/superpowers/plans/`.
3. Rodada de backup (0h e 12h, `.backup` + rclone → R2; WAL) continua pendente e por último.

## Respostas do usuário (2026-09-15)
- Quem lê: membro da comissão de inventário nomeada, ou responsável pela atividade de patrimônio.
- Com quê: leitor de código de barras USB/Bluetooth ligado ao notebook ou ao celular (age como teclado),
  OU a câmera do celular lendo a plaqueta.
- UX: o teclado virtual do celular NÃO pode aparecer a cada leitura — campo com `inputmode="none"` e foco
  mantido; botão "Câmera" abre leitura por câmera na página (biblioteca JS embutida, offline); botão
  "Digitar" para plaqueta ilegível (só aí o teclado aparece).

## Decisões de desenho (2026-09-15, aprovadas em conversa)
- Evento = campanha de inventário: escopo = todas as salas com bens ativos, ou subconjunto (amostragem).
  Abrir outro evento reinicia o dever de conferir. Um evento aberto por vez.
- Divergente = "local sistema" (bens.localizacao, do SPW) ≠ "local inventário" (sala onde foi lido).
  `bens` NÃO muda por aqui; correção é no SPW e aparece na próxima importação como "movido".
- Fotos: bucket R2 como hoje (boto3, env R2_* via secrets/.env), WebP q85 até 1920×1080 (Pillow).
- Sobras (bem sem cadastro): tabela própria `inventario_sobras`, nunca em `bens`; foto obrigatória.
- Exportação: só .xlsx (aba Bens + aba Sobras). Sem PDF.
- Leitura grava direto (sem staging/auto-flush). Integrante escolhido por sessão (sem login).
- Versão COMPLETA (com câmera e fotos) nesta rodada. Módulo separado: `inventario.py` (dados) +
  blueprint `app_inventario.py` (rotas) + templates `inventario_*.html`.
- Spec: docs/superpowers/specs/2026-09-15-inventario-design.md
- Ajustes finais do usuário (2026-09-15): SEM "concluir sala" (evento fica aberto até encerrar; progresso = % de
  bens localizados); bem baixado lido fica registrado e consultável, continua baixado; um evento aberto por vez;
  código em arquivos próprios; bibliotecas embutidas. Plano: docs/superpowers/plans/2026-09-15-inventario.md.

## Retomada da execução (se a conversa for limpa no meio)
Estado vivo: branch `inventario`; ledger em `.superpowers/sdd/2026-09-15-inventario/progress.md` (tarefas com
linha `Task N: complete` estão prontas; briefs em `task-N-brief.md`, relatórios em `task-N-report.md`).
Instrução: "Leia este arquivo e o ledger; use superpowers:subagent-driven-development para continuar o plano
`docs/superpowers/plans/2026-09-15-inventario.md` a partir da primeira tarefa sem `complete`; ao final,
revisão final, merge na main, push e `docker compose up -d --build`."

## CONCLUÍDO (2026-09-15, noite)
Módulo de inventário entregue: merge na `main` (b06fba6), push no GitHub, container reconstruído, 142 testes.
Pendências do usuário: bucket R2 + `secrets/.env` com R2_* (fotos), teste real no celular/leitor, rodada de backup.
