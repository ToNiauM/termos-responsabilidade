# Processos SEI, histórico de termos e importações, painel e recorte

**Data:** 2026-09-14
**Estado:** aprovado em conversa; aguardando revisão do texto.
**Base:** `main` em `aae233f` (SQLite + DSGov; pesquisa rápida no cabeçalho; site em
patrimonio.sistemascfc.org via Docker).

## 1. Objetivo

Três melhorias, implementadas nesta ordem, num único ciclo:

1. **Processos SEI e registro de termos** — só se emite termo com um processo SEI vigente do tipo
   certo; toda cópia ou download grava uma "foto" do termo (data, bens, valores), para saber
   exatamente quando foi a última emissão de cada área ou pessoa.
2. **Rastreio das importações** — cada importação do export do SPW registra o que mudou (bens novos,
   removidos, movidos, mudança de situação). Com a foto do último termo, o sistema aponta quais termos
   estão desatualizados.
3. **Painel e recorte** — a tela inicial vira um painel com cards e tabelas; uma tela de recorte
   filtra bens por faixa de valor e período de entrada, com gráfico e exportação.

Princípio inalterado: simplicidade. Só acréscimos ao esquema; nada do que existe muda de forma.
Geradores `.docx`, `termos_html.py` e `textos.py` ficam intocados. Sem biblioteca nova: os gráficos
são barras em HTML/CSS.

Fora do escopo: backup automático (próxima rodada), web service do SEI, e-mail, login individual.
Os processos SEI **não** entram na planilha de cadastros (4 abas).

---

## 2. Parte 1 — Processos SEI e termos emitidos

### 2.1 Esquema

```sql
CREATE TABLE IF NOT EXISTS processos_sei (
  id         INTEGER PRIMARY KEY,
  tipo       TEXT NOT NULL CHECK (tipo IN ('ccusto','individual','devolucao')),
  descricao  TEXT NOT NULL,
  numero_sei TEXT NOT NULL,
  vigente    INTEGER NOT NULL DEFAULT 0,
  criado_em  TEXT NOT NULL              -- ISO 'YYYY-MM-DD HH:MM:SS'
);
CREATE UNIQUE INDEX IF NOT EXISTS processos_sei_vigente ON processos_sei(tipo) WHERE vigente = 1;

CREATE TABLE IF NOT EXISTS termos_emitidos (
  id            INTEGER PRIMARY KEY,
  tipo          TEXT NOT NULL,          -- ccusto | individual | devolucao
  chave         TEXT NOT NULL,          -- sigla do centro ou nome da pessoa
  processo_id   INTEGER NOT NULL REFERENCES processos_sei(id),
  documento_sei TEXT,                   -- número do documento no SEI; preenchido depois
  emitido_em    TEXT NOT NULL,          -- ISO
  quantidade    INTEGER NOT NULL,
  valor_total   REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS termos_emitidos_bens (   -- a foto
  termo_id    INTEGER NOT NULL REFERENCES termos_emitidos(id) ON DELETE CASCADE,
  numero      INTEGER NOT NULL,
  descricao   TEXT, complemento TEXT, localizacao TEXT, valor_atual REAL,
  PRIMARY KEY (termo_id, numero)
);
```

`termos_emitidos_bens` não referencia `bens` de propósito: `bens` é substituída a cada importação.
`renomear_centro` e `renomear_pessoa` passam a atualizar também `termos_emitidos.chave`.

### 2.2 Regras

- **Um vigente por tipo.** `marcar_vigente(id)` desmarca o vigente atual do mesmo tipo e marca o
  novo, na mesma transação. `encerrar(id)` só zera `vigente`. Excluir processo só sem termos
  registrados (`ErroDeNegocio` caso contrário).
- **Sem processo vigente do tipo, não há emissão.** Em `db.processo_vigente(conn, tipo)` → `None`
  faz a tela do termo mostrar `br-message danger` ("Cadastre um processo SEI vigente para termos
  <tipo> em Cadastros → Processos SEI", com link) e esconder Copiar e Baixar. As rotas
  `/termo/<tipo>/<chave>/documento` e `/docx` também recusam (flash + redirect para a tela do termo),
  para não escapar pela URL. O iframe com a prévia continua aparecendo (só visualiza).
- **Copiar ou Baixar registra.** `db.registrar_emissao(conn, tipo, chave, bens)` grava
  `termos_emitidos` + foto. Baixar: a rota `/docx` registra antes de enviar. Copiar: o botão faz
  `fetch POST /termo/<tipo>/<chave>/registrar` após a cópia dar certo; a resposta é JSON
  `{"emitido_em": ...}` e o aviso passa a dizer "Copiado e registrado às HH:MM".
- **Sem duplicata no mesmo dia.** Se já existir registro do mesmo `tipo`+`chave` no mesmo dia com a
  mesma lista de números, só `emitido_em` é atualizado. Lista diferente (a base mudou entre uma cópia
  e outra) gera registro novo.
- **Baixar planilha** (centro de custo) não registra: não é o termo.
- **Devolução** registra no processo de tipo `devolucao`, com `chave` = nome e a foto = bens
  selecionados na sessão. Devolução não tem "desatualizado".
- **Documento SEI** é opcional e editável a qualquer momento na tela de termos emitidos.

### 2.3 Funções em `db.py`

```python
def processos(conn) -> list[dict]                         # todos, vigentes primeiro, depois por criado_em desc
def processo_vigente(conn, tipo) -> dict | None
def incluir_processo(conn, tipo, descricao, numero_sei, vigente=True) -> int
def marcar_vigente(conn, id) -> None
def encerrar_processo(conn, id) -> None
def excluir_processo(conn, id) -> None                    # ErroDeNegocio se houver termos
def registrar_emissao(conn, tipo, chave, bens) -> dict    # ErroDeNegocio sem processo vigente; devolve o registro
def termos_emitidos(conn, tipo=None, chave=None, limite=200) -> list[dict]   # com descricao/numero do processo
def termo_emitido(conn, id) -> dict | None                # registro + lista de bens da foto
def salvar_documento_sei(conn, id, documento) -> None
def ultimo_termo(conn, tipo, chave) -> dict | None
def situacao_termo(conn, tipo, chave, bens_atuais) -> dict
    # {"estado": "sem_termo"|"vigente"|"desatualizado", "ultimo": dict|None, "entraram": int, "sairam": int}
```

`situacao_termo` compara o conjunto de números da foto do último termo com `bens_atuais`
(`bens_do_centro` ou `bens_da_pessoa`). Só entrada/saída conta; valor e descrição não.

### 2.4 Telas

- **Cadastros → aba "Processos SEI"** (`cadastros.html`, aba nova; rota `/cadastros/processos`):
  tabela (tipo, descrição, número SEI, vigente como `br-tag`, criado em, ações *Marcar vigente* /
  *Encerrar* / *Excluir*) e formulário de inclusão (tipo, descrição, número; "vigente" marcado por
  padrão). Rotas POST `/cadastros/processos/incluir`, `/vigente`, `/encerrar`, `/excluir`.
- **Tela do termo** (`termo.html`): acima da prévia, uma linha de situação:
  - sem termo: "Nenhum termo registrado para GEX-LIC."
  - vigente: "Último termo registrado em 03/09/2026 14:12 (SEI 1234567). Bens iguais aos de hoje."
  - desatualizado: "... Desde então entraram 3 bens e saiu 1." com link para o detalhe do registro.
  - Devolução mostra só o último registro.
- **Menu → "Termos emitidos"** (`termos_emitidos.html`, rota `/termos-emitidos`): filtros por tipo e
  por chave (texto, "contém"); tabela (data/hora, tipo, chave, quantidade, valor, processo, documento
  SEI) com link para o detalhe. Detalhe (`/termos-emitidos/<id>`): dados do registro, campo
  "Documento SEI" com botão Salvar, e tabela dos bens da foto. Botão para reabrir a tela do termo
  daquela chave.
- **Listas de centros e de pessoas** (`centro_custos.html`, `termos_individuais.html`): além do
  `select` atual, uma tabela com situação do termo (`br-tag`: *sem termo* cinza, *vigente* verde,
  *desatualizado* amarelo com "+3 −1") e link direto para o termo. A tabela usa a busca do `br-table`.

---

## 3. Parte 2 — Rastreio das importações

### 3.1 Esquema

```sql
CREATE TABLE IF NOT EXISTS importacoes (
  id            INTEGER PRIMARY KEY,
  importado_em  TEXT NOT NULL,
  arquivo       TEXT,
  total INTEGER NOT NULL, ativos INTEGER NOT NULL,
  novos INTEGER NOT NULL, removidos INTEGER NOT NULL, movidos INTEGER NOT NULL, situacao INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS importacoes_mudancas (
  importacao_id INTEGER NOT NULL REFERENCES importacoes(id) ON DELETE CASCADE,
  numero        INTEGER NOT NULL,
  tipo          TEXT NOT NULL CHECK (tipo IN ('novo','removido','movido','situacao')),
  de            TEXT, para TEXT,
  descricao     TEXT                 -- cópia, para o detalhe não depender de `bens`
);
```

### 3.2 Regras

- `importar_bens` lê a tabela atual num dicionário `{numero: (situacao, localizacao, descricao)}`
  antes do `DELETE`, e depois do `INSERT` calcula: `novo` (número só no export), `removido` (só na
  tabela antiga), `movido` (mesma situação, localização diferente: `de`/`para` = localizações),
  `situacao` (situação diferente: `de`/`para` = situações). Um bem pode gerar `movido` e `situacao`.
  Tudo na mesma transação; falha em qualquer ponto desfaz importação e log.
- A primeira importação com a tabela vazia registra tudo como `novo`; isso é o esperado.
- O resumo devolvido ganha `novos`, `removidos`, `movidos`, `situacao`, `importacao_id`.
- `importacoes(conn, limite=20)` e `importacao(conn, id)` (registro + mudanças) para as telas.
- `historico_do_bem(conn, numero)` → mudanças do bem (com data da importação) e termos em que
  apareceu (data, tipo, chave), para a ficha.

### 3.3 Telas

- **Atualizar base** (`upload.html`): flash pós-importação passa a dizer "3.101 bens importados
  (3.101 ativos): 12 novos, 3 removidos, 7 movidos, 40 mudaram de situação"; abaixo do formulário,
  tabela das últimas importações com link para o detalhe.
- **Detalhe** (`/importacoes/<id>`, `importacao.html`): cards com os quatro contadores e tabela
  (número, descrição, tipo da mudança como `br-tag`, de, para) com a busca do `br-table`.
- **Ficha do bem** (`bem.html`): seção "Histórico" com movimentações e termos em que apareceu.

---

## 4. Parte 3 — Painel e recorte

### 4.1 Painel (tela inicial, `index.html`)

O card de pesquisa continua no topo. Abaixo:

**Cards** (`br-card` com número grande e legenda), em `db.painel(conn)`:

| card | cálculo |
|---|---|
| Bens ativos | `count` de `situacao='ATIVO'` |
| Valor dos ativos (sem imóveis) | soma de `valor_atual` dos ativos com classificação fora de `{SEDE, TERRENOS}` |
| Imóveis | quantidade e valor dos ativos em `{SEDE, TERRENOS}` |
| Sem centro de custo | ativos cuja localização não está em `localizacoes` |
| Termos a emitir | centros com situação `sem_termo` ou `desatualizado` (individuais idem, em linha menor) |
| Última importação | data/hora e "N novos, N removidos"; "nenhuma" se não houver |

**Tabelas** (cada uma com a busca do `br-table`; valores em R$ formatados):

- por centro de custo: sigla, responsável, quantidade, valor, situação do termo (`br-tag`), link
  para o termo. Uma linha "sem centro" ao final.
- por classificação contábil: classe, quantidade, valor. Imóveis aparecem, mas em linha separada
  ao final para não distorcer a leitura.
- por localização: localização, centro (ou "—"), quantidade, valor.
- por faixa de idade (pela `data_entrada`, `DD/MM/AAAA`): até 5 anos, 5 a 10, 10 a 20, mais de 20,
  sem data; quantidade e valor.

Aviso `br-message warning` quando houver ativos com `valor_atual` nulo ou zero ("35 bens ativos sem
valor; não entram nas somas"). Cada tabela tem um link "ver no recorte" que abre a tela de recorte
já filtrada (por classificação, centro, faixa).

### 4.2 Recorte (`/recorte`, `recorte.html`, menu "Recorte")

Formulário `GET` com: valor de / até (R$), entrada de / até (data, `input type=date`), classificação
(`select`, opcional), centro de custo (`select`, opcional), situação (padrão ATIVO). Tudo opcional;
sem filtro nenhum, mostra os ativos.

`db.recorte(conn, filtros) -> dict` devolve `bens` (lista, limite 1.000 com aviso), `quantidade`,
`valor_total`, `por_ano` (`[(ano, quantidade, valor)]`) e `por_faixa_valor` (faixas fixas: até 100,
100–500, 500–1.000, 1.000–5.000, 5.000–20.000, acima de 20.000; quantidade e valor). Datas são
comparadas convertendo `DD/MM/AAAA` para `AAAA-MM-DD` na consulta (`substr`), sem alterar a tabela.

Tela: os filtros no topo; dois cards (quantidade, valor total); dois gráficos de barras em HTML/CSS
(largura proporcional, rótulo com quantidade e valor): **por ano de entrada** e **por faixa de
valor**; tabela dos bens (número com link para a ficha, descrição, complemento, localização, centro,
classificação, entrada, valor) com a busca do `br-table`; botão **Exportar .xlsx** (`/recorte/xlsx`,
mesmos filtros, gerado em memória com openpyxl, sem limite de 1.000).

As barras são `div`s com `style="width: N%"`: é o segundo lugar do projeto com inline style, além de
`termos_html.py`, e por motivo análogo (valor calculado por linha). O CSS fica em `dsgov.css`.

---

## 5. Menu

Ordem nova: Início, Termo por centro de custo, Termo individual, Termo de devolução, **Termos
emitidos**, **Recorte**, Cadastros, Textos, Atualizar base.

## 6. Testes

- Esquema: tabelas novas criadas; índice de um vigente por tipo (segundo `vigente=1` do mesmo tipo
  via `marcar_vigente` desmarca o anterior; `INSERT` direto duplicado falha).
- Processos: incluir/marcar/encerrar/excluir; excluir com termos → `ErroDeNegocio`.
- Emissão: sem vigente → `ErroDeNegocio` e rotas `/documento` e `/docx` redirecionam com flash;
  com vigente, `/docx` registra e `/registrar` responde JSON; duplicata no mesmo dia só atualiza a
  hora; lista diferente cria novo registro; devolução registra no processo `devolucao`.
- `situacao_termo`: sem termo, vigente, desatualizado com contadores; renomear centro/pessoa leva o
  histórico junto.
- Importação: novos/removidos/movidos/situacao contados e gravados; falha desfaz o log;
  `historico_do_bem`.
- Painel: cards e tabelas com a base de teste (incluindo imóveis separados, sem centro, idade).
- Recorte: filtros de valor e datas, faixas, limite, xlsx com as colunas e sem limite.
- Rotas: cada tela nova responde 200 e traz o conteúdo esperado.
