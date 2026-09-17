# Revisão do SYSTEM_IMPROVEMENT_PLAN — 2026-09-16

Fonte revisada: `SYSTEM_IMPROVEMENT_PLAN.md` (raiz do repositório, não versionado, 1.010 linhas, referência
`9950dc1`). Cada achado citado abaixo foi conferido no código antes desta triagem.

## 1. Veredito

O diagnóstico do plano é bom: os defeitos principais existem. A execução proposta é desproporcional ao
sistema (Flask de 4 mil linhas, um mantenedor, uso interno de um setor, site atrás de auth_basic):
24 tarefas, 31 commits, testes de concorrência com barreiras, Playwright, CI, logs estruturados, ledger de
migrações e refatoração do `db.py`. Critério de Antônio: simplicidade e custo-benefício.

Dois achados do plano estão superestimados:

- **F24 (backup)**: `backup.sh` já roda no cron (10h e 18h, seg–sex), usa `.backup` do SQLite, confere
  `integrity_check`, envia para o R2 via rclone e tem retenção. Não há problema a corrigir.
- **F19 (desempenho)**: 185 SELECTs no painel levam 0,09 s. Não há problema a corrigir.

## 2. Triagem

### Vale a pena agora — Fase 1 (aprovada em conversa)

| Achado | Confirmado em | Correção |
|---|---|---|
| F11 segredo de sessão literal | `app.py:22` | segredo por instalação (`TERMOS_SEGREDO` ou `dados/segredo.txt`) |
| F23 `.dockerignore` incompleto | `.dockerignore`, `Dockerfile` (`COPY . .`) | excluir `secrets/`, `*.env`, `*.db`, `*.xlsx`, `*.md`, `backup.sh` |
| F02 abas `inv_*` parciais apagam o inventário | `db.py:908-909` (`tem_inventario = any(...)`) | exigir as 5 abas ou nenhuma |
| F04 rename de pessoa commita antes do e-mail/matrícula | `db.py:640-666` | uma transação por operação |
| F05 textos salvos com commit por chave | `textos.py:157-168`, `app.py:420-434` | `textos.salvar_todos` com um commit |
| F06 dedup ignora o processo SEI | `db.py:759` | comparar também `processo_id` |
| F13 tipo inválido em `/docx` dá 500 | `app.py:_exigir_processo` (`ROTULO_TIPO[tipo]`) | 404 |
| F13 JSON não-objeto em `/ler` dá 500 | `app_inventario.py:119,133` | 400 |
| F14 sem limite de upload no Flask | nginx já limita em 20m | `MAX_CONTENT_LENGTH` 20 MB + mensagem |
| F12 HEAD em `/docx` registra emissão | Flask atende HEAD com a view GET | HEAD não registra |
| F15 texto `=...` vira fórmula no XLSX | todos os `ws.append` | helper que grava texto como texto |
| F21 fallback do Copiar ignora o retorno | `templates/termo.html:49` | só registra se copiou |
| F22 BRCard troca o id `form-sobra` | `templates/inventario_sala.html:71`, mesmo mecanismo já corrigido nos cadastros | form dentro do card |
| F25 ano 2026 fixo no teste; README diz 4 abas | `tests/test_app.py:71`, `README.md:51` | corrigir |

Spec: `../specs/2026-09-16-fase1-blindagem-design.md`. Plano: `../plans/2026-09-16-fase1-blindagem.md`.

### Decidido em conversa: não fazer agora

- **CSRF (T11)**: 30 formulários e 4 fetch; site atrás de auth_basic, uso interno. Fica como pendência
  conhecida. Se um dia for feito: Flask-WTF `CSRFProtect`, campo oculto nos forms, cabeçalho nos fetch.
- **Emissão via GET → POST (T11-D)**: o download em GET registrar a emissão é comportamento escolhido; só o
  HEAD deixa de registrar.
- **Marcação de e-mail "enviado" no clique (T20-D)**: decisão anterior do usuário (mailto sem infraestrutura).

### Não vale a pena (custo sem retorno para este sistema)

- T01 CLI de diagnóstico do banco (o caso observado era o banco de testes do próprio usuário).
- T03–T06 infraestrutura de testes: contratos HTTP, comparação semântica de DOCX, concorrência com
  barreiras, Playwright. A suíte de 174 testes em 4 s já protege o que importa.
- T09 ledger de migrações: `criar_esquema` com `IF NOT EXISTS` e checagem de colunas resolve.
- T14 compensação de fotos no R2, chaves com UUID: 3 integrantes, uploads raros.
- T15 constraints e locks de concorrência: 4 threads do waitress, 3 usuários.
- T16 snapshot versionado de emissão com hash do documento.
- T17 otimização de consultas e índices.
- T18/T19 refatoração do `db.py` e dos geradores: churn sem ganho funcional.
- T21 CI, cobertura, logs estruturados, request-id.
- T02 pin de dependências e sandbox Docker de teste: pode ser feito num commit avulso se incomodar.

## 3. Fase 2 — completar a migração do inventário (próxima rodada de brainstorming)

Pedido do usuário (2026-09-16): "traz o sistema de inventário, regras, dashboard e tudo para esse
sistema, no design system do gov". O que o sistema antigo (`/opt/web/sga/cfc`, levantamento em
`2026-09-15-inventario-existente.md`) tem e o módulo atual ainda não tem:

- **Dashboard do evento**: KPIs (ativos, localizados, pendentes, divergentes, % conclusão), contagem por
  integrante e por conservação, progresso por andar e por sala dentro do andar. Aqui: só o card do painel
  inicial e os contadores por sala. Encaixa no padrão do Recorte (ECharts, regra de tipo de gráfico por nº
  de itens, tudo clicável).
- **Relatório**: filtros por integrante, conservação e "com/sem foto", busca textual sem acento, ordenação
  por coluna, miniatura da foto com modal. Aqui: filtros só por sala e situação.
- **Exportação**: cabeçalho com título, data, filtros ativos e total; opção "incluir fotos" com `=IMAGEM()`;
  PDF (ReportLab). PDF estava fora do escopo na spec de 2026-09-15; confirmar se continua fora.
- **Leitura**: marcar/desmarcar localizado em lote; alternar status manualmente; "quem usa" já existe.
- **Congelamento do relatório no encerramento** (F07 do plano): hoje o relatório de evento encerrado é
  recalculado contra `bens` atual; ao carregar o SPW de 2027, o relatório de 2026 muda. Snapshot dos bens
  do escopo no encerramento (~60 linhas + 1 tabela). Decidir junto com a Fase 2.
- **Desligar o sistema antigo** (`sga-cfc`, porta 11001, sistemadeinventario.com.br): depende de tudo
  acima estar no ar e validado no celular.

Não migrar (já decidido em 2026-09-15): login próprio, multi-datasource, PWA, CRUD de usuários,
`/movimentacao` (não implementado lá), `terms.py` (código morto).

Perguntas para o brainstorming da Fase 2: PDF sim ou não; andar extraído do nome da localização (antes do
primeiro `-`) ou cadastrado; snapshot no encerramento sim ou não; filtros do relatório que ele usa de fato.
