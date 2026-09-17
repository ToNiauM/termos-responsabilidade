# Fase 2 do inventário: painel, relatório completo, exportação com fotos, lote e snapshot

**Data:** 2026-09-17
**Estado:** desenho aprovado pelo usuário em 2026-09-17 (execução autorizada por subagentes).
**Base:** `main` em `e07e67f` (Fase 1 "blindagem" publicada).
**Origem:** seção 3 de `docs/superpowers/notes/2026-09-16-revisao-system-improvement-plan.md` (o que o sistema
antigo `sga-cfc` tem e o módulo atual ainda não tem). Spec do módulo: `2026-09-15-inventario-design.md`.

## 1. Objetivo e decisões

Completar a migração do inventário para que o `sga-cfc` possa ser desligado. Entram quatro frentes e um
ajuste de modelo; tudo só acréscimo, `bens` continua espelho do SPW e nunca muda por aqui.

Decisões do usuário (AskUserQuestion, 2026-09-17):

- Entram: painel do evento, relatório completo, exportação com fotos, lote na leitura.
- Snapshot dos bens no encerramento: **sim**.
- Andar: **extraído do nome da sala** (texto antes do primeiro ` - `; sem hífen → "Sem andar").
- PDF: **não** (continua fora, só `.xlsx`).
- Painel em **página própria** (`/inventario/<id>/painel`); a tela do evento continua leve para o celular.
- Lote com **"Desmarcar"** (apaga a leitura do bem neste evento, como o "alternar status" do sistema antigo).

Fora: desligar o `sga-cfc` (passo manual do usuário depois de validar no celular), login/perfis, PDF.

---

## 2. Modelo

### 2.1 Tabela nova (acrescentada a `db.ESQUEMA`)

```sql
CREATE TABLE IF NOT EXISTS inventario_bens_encerrados (
  evento_id     INTEGER NOT NULL REFERENCES inventario_eventos(id) ON DELETE CASCADE,
  numero        INTEGER NOT NULL,
  situacao      TEXT,
  descricao     TEXT,
  complemento   TEXT,
  classificacao TEXT,
  localizacao   TEXT,
  PRIMARY KEY (evento_id, numero)
);
```

`CREATE TABLE IF NOT EXISTS` basta como migração (`db.criar_esquema` roda o script inteiro; bancos antigos
ganham a tabela vazia).

### 2.2 Snapshot no encerramento

`encerrar_evento(conn, id)` passa a, na mesma transação que grava `encerrado_em`, copiar para
`inventario_bens_encerrados` (colunas `numero, situacao, descricao, complemento, classificacao, localizacao`
de `bens`):

- todos os bens **ATIVO** cujo `localizacao` está em `inventario_salas` do evento, e
- todo bem com leitura em `inventario_leituras` do evento (qualquer situação, dentro ou fora do escopo).

Idempotente: se o evento já estava encerrado, não faz nada (nem regrava o snapshot).

### 2.3 Fonte dos bens por evento

Toda consulta do módulo que hoje lê `bens` (`salas`, `resumo`, `bens_da_sala`, `relatorio`, `painel`, e a
`UNION` de lidos fora do escopo) passa a usar a **fonte do evento**:

```python
def _fonte_bens(conn, evento_id: int) -> str:
    """'bens' se o evento está aberto ou não tem snapshot (encerrado antes desta versão);
    senão a subconsulta do snapshot. evento_id é int: pode ir inline no SQL."""
```

Devolve `"bens"` ou
`"(SELECT numero, situacao, descricao, complemento, classificacao, localizacao FROM inventario_bens_encerrados WHERE evento_id = {evento_id})"`,
usado como `FROM {fonte} b` nos SQLs existentes. Efeito: o relatório e o painel de um evento encerrado não
mudam quando o export do SPW de 2027 for carregado. `ler` e `bens_da_sala` de evento aberto seguem em `bens`.
Um evento encerrado sem linhas no snapshot (não há nenhum em produção; só testes antigos) cai em `bens`.

### 2.4 Aba `inv_bens_encerrados` na planilha de cadastros

`ABAS` ganha `"inv_bens_encerrados": ["evento_id", "numero", "situacao", "descricao", "complemento",
"classificacao", "localizacao"]` (exportada por último, tabela `inventario_bens_encerrados`). Validação em
`validar_abas`: `evento_id` existe em `inv_eventos` **e** está encerrado (tem `encerrado_em`); `numero` inteiro;
`(evento_id, numero)` único. Não exige que o bem exista em `bens` (é justamente o que o snapshot preserva).
`substituir_tabelas` apaga e regrava a tabela junto com as outras cinco (regra "tudo ou nada" inalterada).

---

## 3. Funções de dados (`inventario.py`)

Assinaturas que mudam ou nascem; o resto fica.

```python
ANDAR_SEM = "Sem andar"

def andar(localizacao: str) -> str
    # texto antes do primeiro " - " (strip); sem " - " → ANDAR_SEM. "07 - COAD - SALA" → "07"

def encerrar_evento(conn, id) -> None                      # + snapshot (2.2)

def ler_lote(conn, evento_id, localizacao, numeros: list[int], integrante) -> dict
    # chama a regra de `ler` para cada número (mesmas validações; bem inexistente é pulado e listado);
    # um só commit ao final. Devolve {"lidos": n, "nao_encontrados": [numeros]}

def desfazer_leituras(conn, evento_id, numeros: list[int]) -> list[str]
    # evento aberto obrigatório; DELETE das leituras (evento_id, numero) informadas; devolve as foto_url
    # que existiam (a rota apaga no bucket). Números sem leitura são ignorados.

FILTROS_RELATORIO = ("localizacao", "situacao", "integrante", "conservacao", "foto", "busca", "ordem", "dir")
CONSERVACAO_VAZIA = "-"    # valor do filtro para "Não informada"

def relatorio(conn, evento_id, filtros: dict | None = None) -> list[dict]
    # filtros: localizacao, situacao (localizado|divergente|pendente), integrante, conservacao (valor de
    # CONSERVACAO ou "-" = sem conservação), foto ("com"|"sem"), busca (texto), ordem (nome de coluna de
    # COLUNAS_ORDEM), dir ("asc"|"desc", padrão asc). Filtro em Python sobre o resultado do SQL atual.
    # Busca: normaliza (NFD, sem acento, casefold) e exige cada palavra em algum de
    # numero, descricao, complemento, quem_usa, observacao, local_sistema, local_inventario.
    # Ordem padrão: local_sistema, numero (a de hoje). Chave desconhecida em ordem → padrão.

COLUNAS_ORDEM = ("numero", "descricao", "local_sistema", "local_inventario", "situacao_inv", "conservacao",
                 "quem_usa", "integrante", "lido_em")

def descrever_filtros(filtros: dict) -> str
    # "Todas as salas · Situação Divergente · Integrante Fulano · Conservação Ruim · Com foto · Busca "cadeira""
    # (só os filtros presentes; sem nenhum → "Todas as salas"). Usada na tela e no cabeçalho do xlsx.

def contar_fotos(linhas) -> int                             # linhas com foto_url http(s)

def exportar_xlsx(conn, evento_id, destino, filtros=None, fotos=False)
    # Aba "Bens": linha 1 nome do evento; linha 2 "Gerado em dd/mm/aaaa hh:mm"; linha 3 descrever_filtros;
    # linha 4 "Total: N bens"; linha 5 COLUNAS_XLSX; dados a partir da 6ª. Mesmos filtros/ordem do relatório.
    # fotos=True: célula Foto = fórmula '=_xlfn.IMAGE("url")' gravada direto em ws.cell (não por
    # acrescentar_linha, que força texto), altura da linha 60 pt; sem url http(s) → "-". fotos=False: URL como hoje.
    # Aba "Sobras" como hoje (com o filtro de sala se houver); com fotos=True também usa IMAGE.

def painel(conn, evento_id, andar_sel: str | None = None) -> dict
    # {"resumo": resumo(...),
    #  "situacao": [{"chave": "localizado", "rotulo": "Localizado", "quantidade": n}, ... divergente, pendente],
    #  "integrantes": [{"chave": nome, "rotulo": nome, "quantidade": leituras}]  (todas as leituras do evento, maior→menor),
    #  "conservacao": [{"chave": "Bom"|...|"-", "rotulo": "Bom"|...|"Não informada", "quantidade": n}] (na ordem de CONSERVACAO, "-" por último; só os > 0),
    #  "andares": [{"andar": "02", "total": n, "localizados": n, "pendentes": n, "divergentes": n, "salas": k}] (ordem do andar),
    #  "salas_do_andar": [{"localizacao": ..., "total", "localizados", "pendentes", "divergentes"}] se andar_sel, senão []}
    # Tudo a partir de salas(conn, evento_id) (que já usa a fonte do evento) e de inventario_leituras.
```

`_xlfn.IMAGE`: é o nome canônico que o Excel pt-BR exibe como `=IMAGEM()`; gravar `IMAGEM` literal dá `#NOME?`
(comentário em `/opt/web/sga/cfc/app/services/report_exports.py:243-249`). O `acrescentar_linha` do Fase 1
força texto em strings que começam com `=`, por isso a fórmula é escrita em `ws.cell(row, col).value`.

---

## 4. Rotas (`app_inventario.py`)

| rota | método | faz |
|---|---|---|
| `/inventario/<id>/painel` | GET | `?andar=` opcional; monta cards com `painel_inventario.cards(...)` (seção 5) e renderiza `inventario_painel.html` |
| `/inventario/<id>/relatorio` | GET | lê os `FILTROS_RELATORIO` de `request.args`; passa `filtros`, `descricao`, `n_fotos`, opções dos selects (salas do evento, integrantes, CONSERVACAO + "Não informada") |
| `/inventario/<id>/xlsx` | GET | mesmos args + `fotos=1`; nome do arquivo como hoje |
| `/inventario/<id>/sala/<loc>/lote` | POST (form) | `acao` = `marcar` \| `desmarcar`, `numeros` (checkboxes, lista); `marcar` → `ler_lote` com `session["integrante"]`; `desmarcar` → `desfazer_leituras` e `fotos.apagar` em cada url devolvida; flash com o resultado ("12 bens marcados como localizados; 2 não encontrados: 1, 2" / "3 leituras desfeitas"); redirect para a sala. Evento encerrado ou sem integrante → `ErroDeNegocio` (flash + redirect, handler global). Sem número marcado → flash "Selecione ao menos um bem." |

Rota `atualizar_leitura` e demais ficam.

---

## 5. Cards do painel: `painel_inventario.py` (novo, ~80 linhas)

Monta os cards no formato da macro `grafico` reaproveitando `graficos.py` e `painel._grafico` (tipo pelo nº
de itens, cores) — o mesmo padrão do Recorte, sem duplicar a regra.

```python
def cards(dados: dict, evento_id: int, andar_sel: str | None) -> list[dict]
```

Cards, nesta ordem (cada um com `tabela` "Ver dados" e clique em todo item):

1. **Bens por situação** — `graficos.rosca` com `status={"Localizado": "sucesso", "Divergente": "alerta", "Não localizado": "neutro"}`, total no centro; url → `relatorio?situacao=<chave>`.
2. **Leituras por integrante** — `_grafico` (≤5 rosca, senão barras); url → `relatorio?integrante=<nome>`.
3. **Conservação informada** — `_grafico`; url → `relatorio?conservacao=<chave>`.
4. **Progresso por andar** — `graficos.colunas(andares, {"Localizados": [...], "Pendentes": [...]}, empilhado=True, rotulos=True, status={"Localizados": "sucesso", "Pendentes": "pendente"}, urls={"Localizados": [...], "Pendentes": [...]})`; url → `painel?andar=<andar>`. Subtítulo: "clique no andar para ver as salas".
5. **Salas do andar X** — só quando `andar_sel`; mesmo montador, uma coluna por sala (rótulo = nome da sala sem o prefixo do andar), `col-12` e rótulos a 45° quando > 10 salas; url → `sala_tela`. Subtítulo com botão/link "todos os andares" (volta ao painel sem `andar`).

`painel._grafico` depende de `url_recorte` (filtros do Recorte), então **não** é reaproveitado.
`painel_inventario` tem seu próprio `_grafico(itens, urls)` com a mesma regra reduzida: ≤ 5 itens → `graficos.rosca`
(total no centro), senão `graficos.barras_horizontais` dos 20 maiores (`escala=False`, subtítulo "20 maiores no
gráfico; todos na tabela" quando passa de 20; altura `alto` acima de 8, `extra` acima de 15). Os cards 1, 4 e 5
chamam `graficos.rosca`/`graficos.colunas` diretamente (tipo fixo). `tabela` de cada card via `graficos.tabela_dados`.

**Template `inventario_painel.html`:** cabeçalho com nome do evento e botões Relatório / Salas; linha de
cards KPI (bens no escopo, localizados com %, divergentes, pendentes, sobras, salas iniciadas/total — mesmo
HTML dos cards do Recorte); `{% for g in cards %}{{ grafico(g) }}{% endfor %}`; scripts do ECharts como em
`recorte.html`. Evento encerrado: mesma página, aviso "Evento encerrado" no topo.

---

## 6. Telas

### 6.1 Evento (`inventario_evento.html`)
Botão **Painel** (`fa-chart-pie`) antes de "Relatório". Nada mais muda.

### 6.2 Relatório (`inventario_relatorio.html`)
- Formulário GET em `br-card`, como o Recorte: Sala, Situação, Integrante, Conservação (+ "Não informada"),
  Foto (Todas / Com foto / Sem foto), Busca (`br-input` texto), botões Aplicar / Limpar. `ordem`/`dir` viajam
  em `input hidden` para sobreviver ao Aplicar.
- O mesmo formulário tem a caixa **Incluir fotos** (`br-checkbox`, `name="fotos" value="1"`) e o botão
  **Exportar .xlsx** como `<button type="submit" formaction="{{ url_for('inventario.xlsx', id=e.id) }}">`:
  os filtros e a caixa viajam juntos, sem JS. Abaixo do formulário: `descricao` do filtro em
  `text-gray-70` e "N linha(s)"; aviso `br-message warning` "Relatório pesado: N fotos" quando `n_fotos >= 50`.
- Cabeçalhos ordenáveis: cada `th` de `COLUNAS_ORDEM` é um link para a mesma URL com `ordem=<col>` e `dir`
  alternado (asc→desc quando já é a coluna ativa); ícone `fa-sort-up`/`fa-sort-down` na ativa.
- Foto: `<img class="dsgov-miniatura" data-foto="url">`; clique abre um `br-modal` único no fim da página
  (`<img>` grande + link "Abrir em nova aba"); JS de ~15 linhas, sem biblioteca. Sem foto → "—".
- Coluna Complemento entra (hoje não aparece na tela, só no xlsx).

### 6.3 Sala (`inventario_sala.html`)
- Tabela "Bens da sala" com `cabecalho_tabela(..., selecao=True)`, `th_selecao`/`td_selecao` (checkbox por
  bem, `name="numeros"`, `form="form-lote"`), só quando o evento está aberto.
- Acima da tabela, `form id="form-lote" method="post" action=".../lote"` com dois botões: **Marcar
  selecionados como localizados** (`name="acao" value="marcar"`, primary) e **Desmarcar** (`value="desmarcar"`,
  secondary, `data-confirmar`: o JS pede `confirm("Desfazer N leituras?")` antes de enviar). Desabilitados
  sem integrante.
- Linhas inseridas pelo JS após uma leitura (função `aplicar`) também ganham a célula de seleção.

---

## 7. Regras migradas / preservadas

- Marcar em lote = leitura sem plaqueta: mesma regra de `ler` (divergente é aceito, não ativo é aceito,
  inexistente é pulado). Sistema antigo: `marcar_lote` (`inventory.py:900-948`).
- Desmarcar = `alternar_status` do sistema antigo; aqui apaga a leitura do evento (a foto vai junto).
- `_xlfn.IMAGE` e altura 60 pt: `report_exports.py` do sistema antigo. Aviso de relatório pesado a partir
  de 50 fotos: `reports.py:334-340`.
- Andar heurístico: `_extract_andar` (`dashboard.py:23-29`), simplificado para o separador ` - `.
- Busca sem acento: mesma normalização de `db.py:369-373`.

---

## 8. Testes

- `tests/test_inventario.py`: `andar`; `encerrar_evento` grava o snapshot (ativos do escopo + lidos fora)
  e é idempotente; `salas`/`relatorio` de evento encerrado não mudam quando `bens` muda depois; evento
  encerrado sem snapshot cai em `bens`; `ler_lote` (lidos, não encontrados, integrante inválido, encerrado);
  `desfazer_leituras` (devolve urls, ignora sem leitura, encerrado); `relatorio` com cada filtro, busca sem
  acento e ordenação asc/desc; `descrever_filtros`; `exportar_xlsx` com cabeçalho de 4 linhas, filtros e
  `fotos=True` (`=_xlfn.IMAGE(`, altura 60) / `fotos=False` (URL); `painel` (situação, integrantes,
  conservação com "Não informada", andares, salas_do_andar); aba `inv_bens_encerrados` exporta/importa e
  validações (evento aberto recusado, numero inválido, repetido).
- `tests/test_painel_inventario.py`: `cards` — tipos, urls, card de salas só com andar, rótulo sem prefixo.
- `tests/test_app.py`: `/painel` 200 com JSON dos gráficos e com `?andar=`; `/relatorio` com filtros e
  cabeçalho ordenável; `/xlsx?fotos=1`; `/lote` marcar/desmarcar (flash, redirect, encerrado → erro, sem
  seleção → aviso); modal de foto presente.
- Nada de rede: `fotos.apagar` substituído por monkeypatch onde o lote desmarca com foto.

## 9. Entrega

Branch `fase2-inventario`; commits atômicos por tarefa; ao final, revisão, merge na `main`, push e
`docker compose up -d --build` na VPS (banco de produção ganha a tabela nova vazia no primeiro start).
Depois do usuário validar o painel, o relatório e o lote no celular, ele desliga o `sga-cfc`.
