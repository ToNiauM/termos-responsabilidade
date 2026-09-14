# Textos editáveis, cadastros completos e planilha de cadastros

**Data:** 2026-09-14
**Estado:** aprovado em conversa; aguardando revisão do texto.
**Base:** `main` em `4140c36` (sistema já em SQLite + DSGov; ver spec de 2026-09-14 "sqlite-desktop").

## 1. Objetivo

Três melhorias, implementadas nesta ordem, num único ciclo:

1. **Textos do termo editáveis** — tudo que hoje é texto fixo nos três geradores `.docx` e em
   `termos_html.py` passa a vir de uma única fonte, editável no app, com marcadores (`{nome}`,
   `{ccustos}`…) que o sistema substitui.
2. **Cadastros totalmente editáveis + De-Para** — editar responsável em página própria, mover uma ou
   várias localizações para outro centro de custo de uma vez, editar nome de pessoa, e a trava de
   exclusão de centro passa a ser "tem bens sob guarda".
3. **Exportar / importar cadastros em planilha** — um `.xlsx` de 4 abas no formato do banco;
   exportar = estado atual; importar = substitui as 4 tabelas, tudo ou nada.

Princípio inalterado: simplicidade. Adaptar o que existe; nada de tela genérica de CRUD.

## 2. Modelo (recapitulação; não muda)

Centro de custo (1) → localizações (N). Cada centro tem um responsável (nome, função, matrícula,
e-mail, tratamento). Bem em `atribuicoes` responde pela pessoa, não pelo setor. `responsaveis` é a
lista de centros; `localizacoes.ccustos` referencia-a com `ON UPDATE CASCADE`.

---

## 3. Parte 1 — Textos do termo

### 3.1 Armazenamento

```sql
CREATE TABLE IF NOT EXISTS textos (
  chave TEXT PRIMARY KEY,
  valor TEXT NOT NULL
);
```

Só guarda o que foi alterado. O padrão (o texto de hoje) vive em `textos.py`:

```python
PADRAO = {chave: valor_padrao, ...}
MARCADORES = {chave: {"nome", ...}, ...}     # marcadores permitidos por chave
def obter(conn) -> dict                      # PADRAO sobreposto pelo que há na tabela
def validar(chave, valor) -> None            # levanta ErroDeNegocio se marcador desconhecido ou chave { mal formada
def salvar(conn, chave, valor) -> None       # valida e grava (INSERT OR REPLACE)
def restaurar(conn, chave) -> None           # DELETE
```

`validar` usa `string.Formatter().parse(valor)` para extrair os campos; campo fora de
`MARCADORES[chave]` → erro "marcador {x} não existe neste bloco"; `ValueError` do parser → erro
"chave { sem fechar". Um texto vazio é permitido (bloco some do documento).

### 3.2 Chaves, valores padrão e marcadores

**Gerais** (sem marcadores):

| chave | padrão |
|---|---|
| `orgao_nome` | Conselho Federal de Contabilidade |
| `orgao_sigla` | CFC |
| `cidade` | Brasília (DF) |
| `assinatura_eletronica` | Assinado eletronicamente via SEI |
| `recebedor_nome` | ANTÔNIO RODRIGUES DE SOUSA JÚNIOR |
| `recebedor_cargo` | Supervisor de Patrimônio |

**Termo individual** (marcadores: `{nome}`, `{orgao_sigla}`):

| chave | padrão |
|---|---|
| `individual_titulo` | TERMO DE RESPONSABILIDADE |
| `individual_abertura` | Pelo presente termo, eu, {nome}, declaro que o(s) equipamento(s) abaixo discriminado(s) se encontra(m) sob a minha guarda e responsabilidade. |
| `individual_compromissos_intro` | Comprometo-me a: |
| `individual_compromissos` | as 5 linhas atuais, "CFC" trocado por `{orgao_sigla}`; **uma linha por item** |
| `individual_ciencia` | Declaro estar ciente das responsabilidades mencionadas acima e assumo total responsabilidade pelos bens listados. |

**Termo por centro de custo** (marcadores: `{responsavel}`, `{matricula}`, `{funcao}`, `{ccustos}`, `{orgao_sigla}`):

| chave | padrão |
|---|---|
| `ccusto_titulo` | Termo de Responsabilidade - {ccustos} |
| `ccusto_paragrafos` | os 7 parágrafos atuais, "CFC" trocado por `{orgao_sigla}`; **linha em branco separa parágrafos** |
| `ccusto_assinatura` | `{responsavel}` ⏎ `{funcao} do(a) {ccustos} do {orgao_sigla}` (duas linhas) |

**Termo de devolução** (marcadores: `{nome}`, `{orgao_sigla}`, `{cidade}`, `{data}`):

| chave | padrão |
|---|---|
| `devolucao_titulo` | TERMO DE DEVOLUÇÃO |
| `devolucao_abertura` | Pelo presente termo, eu, {nome}, declaro que devolvi ao Setor de Patrimônio o(s) bem(ns) abaixo discriminado(s), que se encontrava(m) sob minha guarda e responsabilidade: |
| `devolucao_data` | {cidade}, {data} |
| `devolucao_recebimento` | Declaro que recebi o(s) bem(ns) acima especificado(s): |

`{data}` é "14 de setembro de 2026" (o formato por extenso continua no código). Os marcadores
gerais (`{orgao_sigla}`, `{cidade}`) são resolvidos a partir das chaves gerais.

### 3.3 Renderização

- **Negrito do nome:** nos blocos `*_abertura`, o marcador `{nome}` é renderizado em negrito. O
  renderizador divide o texto em "antes / nome / depois" (`str.partition("{nome}")`) — no `.docx`
  três `runs`, no HTML `<b>`. Só `{nome}` tem esse tratamento; os demais marcadores são texto comum.
- **Blocos multilinha:** `individual_compromissos` → um parágrafo por linha não vazia;
  `ccusto_paragrafos` → parágrafos separados por linha em branco (linhas consecutivas sem linha em
  branco entre elas viram um parágrafo só, unidas por espaço); `ccusto_assinatura` → uma linha por
  `\n`, no mesmo parágrafo (quebra de linha), como hoje.
- Os geradores `.docx` e `termos_html.py` recebem `textos: dict` como parâmetro (obtido pela rota
  com `textos.obter(conn)`) e **não têm mais texto próprio**. A formatação Word (fontes, tabela,
  larguras) não muda. Substituição com `str.format_map` sobre um dicionário com todos os
  marcadores do bloco.
- A tabela de bens, o TOTAL, a data por extenso e as larguras 80 %/100 % continuam no código.

### 3.4 Tela "Textos" (menu, entre Cadastros e Atualizar base)

- Rota `GET /textos`: um formulário só, quatro `br-card` (Gerais, Termo individual, Termo por centro
  de custo, Termo de devolução). Chaves gerais e títulos em `br-input`; os demais em `br-textarea`
  (6–10 linhas). Abaixo de cada campo, ajuda com os marcadores permitidos e a regra de linhas.
- `POST /textos`: grava todas as chaves do formulário (`textos.salvar` em cada uma; valor igual ao
  padrão → `restaurar`, para a tabela não acumular cópias do padrão). Primeiro erro de validação →
  flash com a chave e o motivo, nada é gravado (validação de todas as chaves antes de qualquer
  gravação).
- "Restaurar padrão" por bloco: `<button class="br-button" name="restaurar" value="<chave>">` no
  mesmo formulário; a rota, ao ver `restaurar`, apaga só aquela chave e ignora o resto do POST.
- Um `br-button primary` ("Salvar") na página; os "Restaurar padrão" são terciários.

---

## 4. Parte 2 — Cadastros completos e De-Para

### 4.1 Responsáveis (centros de custo)

- Linha da tabela: ações **editar** (lápis) e **excluir**. O formulário "Renomear" da linha sai.
- `GET|POST /cadastros/responsaveis/<ccustos>/editar` → `editar_responsavel.html`: formulário
  `col-md-8` com sigla, tratamento, responsável, função, matrícula, e-mail, e abaixo a lista das
  localizações do centro (só leitura, com link para a aba Localizações). Salvar (primary) e
  Cancelar (terciário).
- POST: se a sigla mudou, `db.renomear_centro(conn, antigo, novo)` (cascade, como hoje); depois
  `db.atualizar_responsavel(conn, ccustos, dados)` (UPDATE dos demais campos; responsável
  obrigatório). Erros → `ErroDeNegocio` → flash e volta ao formulário.
- **Excluir:** `db.excluir_responsavel` muda de regra — bloqueia se `bens_do_centro(conn, ccustos)`
  não for vazio ("SEPAT tem 37 bens ativos sob guarda; mova as localizações antes de excluir").
  Sem bens: apaga os mapeamentos do centro (as localizações voltam a "pendentes") e o centro.
  Pede confirmação pelo padrão de dois POSTs (`confirmar=1`), como excluir pessoa.

### 4.2 Localizações (De-Para)

- A tabela ganha a **seleção nativa do `br-table`** (`data-selection`, coluna de checkboxes, "selecionar
  todas"), dentro de um `<form method="post" action="/cadastros/localizacoes/mover">`.
- Abaixo da tabela: `br-select` "Mover selecionadas para" (centros cadastrados) + botão
  **Mover** (secondary — o primary da aba continua sendo "Mapear").
- `POST /cadastros/localizacoes/mover` com `localizacoes` (múltiplos) e `ccustos` →
  `db.mover_localizacoes(conn, lista, ccustos)` (um `UPDATE ... WHERE localizacao IN (...)`, uma
  transação). Nenhuma selecionada → flash de erro. Sucesso: "N localização(ões) movida(s) para X".
- A macro `cabecalho_tabela` ganha o parâmetro `selecao=False`; com `True` emite `data-selection` e
  a célula de cabeçalho com o checkbox "selecionar todas"; a linha usa a célula canônica
  `column-checkbox` com `name="localizacoes" value="<localizacao>"`. Só a aba Localizações usa.
- Incluir/excluir mapeamento continuam como hoje.

### 4.3 Pessoas

- Ao lado do nome selecionado: **editar** (lápis) → `GET|POST /cadastros/pessoas/<nome>/editar`
  (`editar_pessoa.html`: um campo "Nome", Salvar/Cancelar). POST → `db.renomear_pessoa(conn,
  antigo, novo)`: normaliza (maiúsculas, espaços), recusa nome já existente, `UPDATE pessoas`
  (cascade em `atribuicoes`). Redireciona para `/cadastros/pessoas?nome=<novo>`.

### 4.4 O que não muda

Ficha do bem, termos, regra "setor OU pessoa", upload de bens, renomear com cascade.

---

## 5. Parte 3 — Planilha de cadastros

### 5.1 Formato (`cadastros.xlsx`, 4 abas, cabeçalho na linha 1)

| aba | colunas |
|---|---|
| `responsaveis` | ccustos, tratamento, responsavel, email, matricula, funcao |
| `localizacoes` | localizacao, ccustos |
| `pessoas` | nome |
| `atribuicoes` | nome, numero |

É o formato das tabelas, nomes idênticos. Bens e textos **não** entram.

### 5.2 Exportar

- Botão **Exportar cadastros** (terciário, no cabeçalho da tela Cadastros) → `GET
  /cadastros/exportar` → `db.exportar_cadastros(conn, destino)` grava `dados/saida/cadastros.xlsx`
  com openpyxl e a rota devolve com `send_file`. Linhas ordenadas pela chave.

### 5.3 Importar

- Na tela **Atualizar base**, um segundo formulário "Importar cadastros" (upload `.xlsx`, botão
  secondary — o primary da página continua sendo o upload de bens), `POST /importar-cadastros` →
  `db.importar_cadastros(conn, arquivo) -> {"responsaveis": n, "localizacoes": n, "pessoas": n,
  "atribuicoes": n}`.
- **Substitui as 4 tabelas inteiras**, numa transação, na ordem `atribuicoes, pessoas, localizacoes,
  responsaveis` (apagar) e `responsaveis, localizacoes, pessoas, atribuicoes` (inserir). A planilha
  é a verdade: o que não estiver nela some. Fluxo esperado: exportar → editar → importar.
- **Validação antes de gravar** (qualquer problema → `ImportacaoInvalida`, nada muda, mensagem lista
  até 20 problemas com aba e linha):
  - aba ausente ou coluna ausente;
  - `responsaveis`: sigla vazia, responsável vazio, sigla repetida;
  - `localizacoes`: localização vazia ou repetida; `ccustos` que não está na aba `responsaveis`;
  - `pessoas`: nome vazio ou repetido;
  - `atribuicoes`: `nome` que não está na aba `pessoas`; `numero` não numérico ou que não existe em
    `bens`; `numero` repetido (um bem, uma pessoa).
- **Normalização** (as mesmas regras do app): sigla e nome de pessoa em maiúsculas; espaços
  colapsados (`db._texto`); matrícula como texto; `numero` inteiro.
- Após importar, a tela mostra as contagens e, como no upload de bens, as localizações ativas sem
  centro.

---

## 6. Estrutura de arquivos

```
textos.py                    (novo) PADRAO, MARCADORES, obter, validar, salvar, restaurar
db.py                        + tabela textos no ESQUEMA; atualizar_responsavel, mover_localizacoes,
                               renomear_pessoa, exportar_cadastros, importar_cadastros;
                               excluir_responsavel com a regra nova
app.py                       + rotas: /textos (GET/POST), editar responsável, mover localizações,
                               editar pessoa, exportar, importar-cadastros
Script_Termo_Individual.py   recebem `textos`; sem texto próprio
Termo_de_Responsabilidade.py
termo_devolucao.py
termos_html.py               idem
templates/textos.html        (novo)
templates/editar_responsavel.html, editar_pessoa.html   (novos)
templates/cadastros.html     lápis nas linhas; seleção + mover na aba Localizações; botão Exportar
templates/upload.html        segundo formulário (importar cadastros)
templates/_macros.html       cabecalho_tabela(selecao=False)
templates/base.html          item de menu "Textos"
tests/test_textos.py         (novo); test_db.py, test_app.py, test_docx.py, test_termos_html.py ampliados
```

## 7. Testes

- `textos`: `obter` devolve o padrão; sobreposição do banco; `validar` recusa marcador desconhecido
  e chave mal formada; `salvar` com valor igual ao padrão não deixa linha na tabela.
- Geradores/HTML: com um texto alterado (ex.: `individual_abertura` contendo "TESTE {nome}"), o
  `.docx` e o HTML trazem "TESTE" e o nome em negrito; compromissos com 3 linhas → 3 parágrafos;
  `ccusto_paragrafos` com 2 parágrafos → 2.
- Rotas: `POST /textos` grava e reflete no documento; marcador inválido → flash e nada gravado;
  `restaurar` apaga só a chave.
- Cadastros: editar responsável (com e sem mudança de sigla); excluir centro com bens → bloqueado;
  sem bens → centro some e localização vira pendente; mover 2 localizações; renomear pessoa mantém
  atribuições; nome repetido → erro.
- Planilha: exportar gera 4 abas com as contagens do banco; importar a própria exportação é
  idempotente; cada regra de validação tem um caso que falha e deixa o banco intacto; importar com
  centro novo + localização movida reflete na ficha do bem.

## 8. Fora de escopo

Histórico/auditoria de alterações; importação parcial ("mesclar"); edição de bens; textos por
idioma; ordem/colunas da tabela do termo; múltiplos responsáveis por centro.
