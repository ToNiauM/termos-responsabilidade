# Termos de Responsabilidade — SQLite + programa de desktop

**Data:** 2026-09-14
**Estado:** aprovado em conversa; aguardando revisão do texto.

## 1. Objetivo

Trocar as duas planilhas Excel (`acervo.xlsx`, `geral.xlsx`) por um banco SQLite mantido pelo próprio
programa, e entregar o sistema como um programa de desktop Windows que abre com dois cliques — sem
servidor, sem Render, sem infraestrutura. Adicionar a emissão dos termos em HTML com botão
"Copiar para o SEI". Adotar o visual do Design System gov.br (só o visual).

**Princípio que governa toda decisão deste documento: simplicidade.** Adaptar o código atual, não
reescrever. Cada item abaixo que puder ser feito com menos código deve ser.

## 2. O que fica igual

- Flask, Jinja2, python-docx e os três geradores de `.docx` (`Script_Termo_Individual.py`,
  `Termo_de_Responsabilidade.py`, `termo_devolucao.py`). A lógica de formatação do documento Word
  não muda; só a **entrada** deles muda (ver §4.3).
- Os textos legais dos termos.
- `timbrado.docx`.

## 3. O que sai

- `pandas` (openpyxl lê o Excel; SQL faz agrupamento, soma e junção).
- `gunicorn`, `Procfile`, qualquer referência ao Render.
- `acervo.xlsx` e `geral.xlsx` do repositório, depois da migração inicial (§4.4).
- `static/style.css` e os blocos `<style>` dos templates (substituídos pelo DS, §6).
- Os blocos `if __name__ == "__main__"` dos geradores.
- O pseudo-centro-de-custo `TERMOS INDIVIDUAIS` (a regra §5.2 o torna desnecessário).

## 4. Dados

### 4.1 Esquema (`termos.db`, SQLite, 5 tabelas)

```sql
CREATE TABLE bens (
  numero        INTEGER PRIMARY KEY,
  situacao      TEXT NOT NULL,          -- ATIVO, BAIXADO, DOADO, INSERVÍVEL
  descricao     TEXT NOT NULL,
  complemento   TEXT,
  classificacao TEXT,
  localizacao   TEXT,
  data_entrada  TEXT,                   -- como vem do export (dd/mm/aaaa)
  valor_compra  REAL,
  valor_atual   REAL
);

CREATE TABLE responsaveis (             -- é também a lista de centros de custo
  ccustos     TEXT PRIMARY KEY,         -- sigla: CCI, DECOM, GESERV...
  tratamento  TEXT,                     -- Prezado / Prezada
  responsavel TEXT NOT NULL,
  email       TEXT,
  matricula   TEXT,
  funcao      TEXT
);

CREATE TABLE localizacoes (
  localizacao TEXT PRIMARY KEY,         -- "14 - CASA DE MÁQUINAS"
  ccustos     TEXT NOT NULL REFERENCES responsaveis(ccustos) ON UPDATE CASCADE
);

CREATE TABLE pessoas (
  nome TEXT PRIMARY KEY
);

CREATE TABLE atribuicoes (
  nome   TEXT    NOT NULL REFERENCES pessoas(nome) ON UPDATE CASCADE ON DELETE CASCADE,
  numero INTEGER NOT NULL REFERENCES bens(numero) DEFERRABLE INITIALLY DEFERRED,
  PRIMARY KEY (nome, numero)
);
```

`PRAGMA foreign_keys = ON` em toda conexão. A FK `atribuicoes.numero → bens` é **deferida**: o
`DELETE FROM bens` da reimportação (§4.2) não falha no meio; a verificação acontece no `COMMIT`, e o
app checa antes disso para devolver a lista de números órfãos (§4.2, passo 3).

Origem de cada tabela hoje:

| Tabela | Vem de | Manutenção |
|---|---|---|
| `bens` | `geral.xlsx/base` (export do sistema de patrimônio) | upload de Excel |
| `responsaveis` | `acervo.xlsx/responsavel` | app (aba Responsáveis) |
| `localizacoes` | `acervo.xlsx/ccustos` | app (aba Localizações) |
| `pessoas` | `geral.xlsx/nomes` ∪ nomes de `geral.xlsx/dados` | app (aba Pessoas) |
| `atribuicoes` | `geral.xlsx/dados` (Nome, Patrimônio) | app (aba Pessoas) |

`acervo.xlsx/acervo` não vira tabela: é o mesmo export de `base` filtrado por `ATIVO`, tirado em
outra data. Fica uma única tabela `bens`.

### 4.2 Importação de bens (tela "Atualizar base")

Entrada: um `.xlsx` no formato do export do sistema (aba com cabeçalho *Número Bem, Situação,
Descrição, Complemento, Classificação Contábil, Localização, Data Entrada, Valor Compra, Valor
Atual*). O app usa a primeira aba cujo cabeçalho contenha "Número Bem"; nome do arquivo não importa.

Comportamento, numa única transação:
1. Valida o cabeçalho; se faltar coluna, aborta com mensagem e nada muda.
2. `DELETE FROM bens`; insere todas as linhas (número não nulo).
3. Se algum `atribuicoes.numero` não existir mais em `bens`, a transação é revertida e a mensagem
   lista os números — o usuário remove a atribuição ou usa um export completo.
4. Mensagem de sucesso: "N bens importados (M ativos)".
5. Se houver localizações de bens ATIVOS sem linha em `localizacoes`, a tela lista essas
   localizações com um link para a aba Localizações. Não é erro; é o que hoje vira "Sem termo"
   sem ninguém ver.

As outras quatro tabelas não são tocadas pelo upload.

### 4.3 Consultas que substituem as planilhas

Todas em `db.py`, devolvendo listas de `dict` (via `sqlite3.Row`). Os geradores de `.docx` passam a
receber essas listas em vez de `DataFrame`; a troca é `bem['campo']` por `bem['campo']` — só sai o
`iterrows()`/`isna()`.

- **Termo por centro de custo** (`bens_do_centro(ccustos)`): bens `ATIVO` cuja `localizacao` mapeia
  para `ccustos` **e que não estão em `atribuicoes`** (regra §5.2), ordenados por número. Cabeçalho
  vem de `responsaveis`.
- **Termo individual** (`bens_da_pessoa(nome)`): `atribuicoes JOIN bens`, ordenados por número.
  Descrição e valor vêm sempre de `bens` (hoje dependem de VLOOKUP na planilha).
- **Termo de devolução**: mantém o fluxo atual (pessoa + números digitados um a um, guardados na
  sessão); a busca do bem é `SELECT ... FROM bens WHERE numero = ?`. O campo de nome é `br-select`
  com as `pessoas`.
- **Consulta de bem** (`ficha_do_bem(numero)`): dados do bem + `ccustos` e responsável do setor (via
  `localizacoes` e `responsaveis`) + pessoa (via `atribuicoes`), se houver.

### 4.4 Migração inicial (uma vez)

`importar_planilhas.py`: lê `acervo.xlsx` e `geral.xlsx` atuais e popula as cinco tabelas.
Detalhes que a checagem dos dados exige:
- `bens` ← `geral.xlsx/base` (7003 linhas; é o export mais completo).
- `pessoas` ← união de `geral.xlsx/nomes` e dos nomes distintos de `geral.xlsx/dados` (43 nomes de
  `dados` não estão em `nomes`). Nomes são normalizados com `strip()`; não se tenta corrigir grafia
  (ex.: "RICARDO DA SILVA CARVALHO," entra como está, e o usuário corrige na aba Pessoas).
- `localizacoes` ← `acervo.xlsx/ccustos`, **exceto** linhas com `ccustos = 'TERMOS INDIVIDUAIS'`.
  Se algum `ccustos` do mapa não tiver linha em `responsaveis`, o script cria a linha com
  `responsavel = '(preencher)'` e avisa.
- Depois de rodar e conferir, as planilhas saem do repositório.

O script fica no repositório (é pequeno) e é executado uma vez pelo desenvolvedor, não pelo
programa.

### 4.5 Onde os arquivos ficam

Uma única variável `PASTA_DADOS` em `config.py` decide tudo. Windows empacotado: pasta `dados/` ao
lado do `.exe`. Desenvolvimento: `./dados/`. Docker (fora de escopo, mas previsto): um volume.

```
TermosCFC.exe
dados/
  termos.db
  timbrado.docx
  saida/          ← .docx e .html gerados (hoje caem na raiz do projeto)
```

Backup do sistema = copiar `dados/`.

## 5. Regras de negócio

### 5.1 Centro de custo: renomear e "de/para"

`responsaveis` é a lista de centros de custo. Renomear (`COLOG → GESERV`) é um
`UPDATE responsaveis SET ccustos = ?` que cascateia para `localizacoes`. Uma transação; nada fica
órfão.

Reorganização em que uma área vira duas (COLOG → parte GESERV, parte outra): renomeia COLOG→GESERV,
cadastra a outra área na aba Responsáveis, e na aba Localizações troca o centro de custo das salas
que mudaram. Não há tela específica de "de/para"; são as duas abas.

A aba Localizações só oferece centros de custo já cadastrados (dropdown). Não existe mais
"Sem termo" implícito.

### 5.2 Responsabilidade: setor OU pessoa

Um bem que está em `atribuicoes` **não entra** no termo do centro de custo onde está fisicamente —
a pessoa responde por ele. Consequência: para tirar um bem do termo do setor, basta atribuí-lo;
para devolvê-lo ao setor, basta remover a atribuição.

Ao atribuir um bem que já está com outra pessoa, o app avisa ("14359 está com FULANO — transferir?")
e só transfere com confirmação explícita (segundo POST com `confirmar=1`).

Um bem pode estar com uma pessoa só (a PK permite mais, mas a tela impede; se um dia for necessário
compartilhar, é só liberar a tela).

### 5.3 Consulta de bem

Campo de busca por número na tela inicial → `/bem/<numero>`: descrição, situação, localização, centro
de custo, responsável do setor e, se houver, pessoa responsável, com links para o termo de cada um.
É o "no bem, meu nome". O "no meu nome, os bens" é a aba Pessoas.

## 6. Interface — DSGov só como visual

O skill `/dsgov` gera projetos Django; **nada disso entra**. Entra apenas a identidade visual, vendorizada
(offline, funciona no `.exe` sem internet) em `static/dsgov/`:
`core-tokens.css`, `core.min.css`, `fontes.css` + Rawline/Raleway, Font Awesome 5, `core.min.js`,
`dsgov.js` (inicializador dos componentes). ~1,3 MB. Fica de fora: Django, Postgres, login, HTMX,
ECharts, dashboard, paginação server-side, verificador.

`templates/base.html` segue a estrutura fixa do DS: skiplink → `br-header` (logo CFC · "Termos de
Responsabilidade" · subtítulo "Setor de Patrimônio") → `br-menu` lateral → `br-breadcrumb` →
mensagens (`br-message`, substitui os flashes coloridos) → `{% block conteudo %}` → `br-footer`.

Componentes: `br-table` nas listagens (busca do próprio DS no cliente; tabelas têm 30–150 linhas),
`br-input` / `br-select` nos formulários (o DS proíbe `<select>` nativo), `br-button` (um `primary` por
tela), `br-tab` nas abas de Cadastros. Sem `<style>`, sem hex, sem CSS próprio além do vendorizado —
com a única exceção do §7.

Menu: Início · Termo por centro de custo · Termo individual · Termo de devolução · Cadastros ·
Atualizar base.

### 6.1 Tela Cadastros (`/cadastros/<aba>`), três abas

Padrão único: tabela dos registros + formulário de incluir + botão excluir por linha. Sem edição
in-place (exclui e inclui de novo; são tabelas pequenas).

1. **Responsáveis** — ccustos, tratamento, responsável, e-mail, matrícula, função. Botão
   **Renomear** por linha (campo inline com a nova sigla; §5.1). Excluir só é permitido se nenhuma
   localização apontar para o centro de custo (mensagem explica).
2. **Localizações** — `br-select` com as localizações distintas de `bens` ATIVO que ainda não estão
   mapeadas + `br-select` com os centros de custo; lista as já mapeadas com excluir. Localizações
   pendentes (ativas sem mapa) aparecem destacadas no topo.
3. **Pessoas** — `br-select` de pessoa (ou campo para cadastrar nova); ao escolher, lista os bens
   atribuídos com remover; campo "nº do patrimônio" + botão "Atribuir": o app busca em `bens`, mostra
   descrição/valor e, se o bem estiver com outra pessoa, pede confirmação (§5.2). Excluir pessoa
   remove as atribuições (cascade) e pede confirmação pelo mesmo padrão de dois POSTs de §5.2 — o DS
   não usa modal de confirmação.

Rotas: `GET /cadastros/<aba>`, `POST /cadastros/<aba>/incluir`, `POST /cadastros/<aba>/excluir`,
`POST /cadastros/responsaveis/renomear`, `POST /cadastros/pessoas/atribuir`,
`POST /cadastros/pessoas/desatribuir`.

## 7. Termo em HTML com "Copiar para o SEI"

Gerar um termo passa a abrir a **página do termo** em vez de baixar o `.docx`:
`/termo/ccusto/<sigla>`, `/termo/individual/<nome>`, `/termo/devolucao/<nome>` (esta última após o
POST "gerar" do fluxo atual).

A página (DSGov) tem:
- `br-button primary` **Copiar para o SEI** — Clipboard API com `text/html` (+ `text/plain` como
  fallback); ao colar no editor do SEI a tabela chega como tabela. Mostra `br-message` "Copiado" por
  3 s. Funciona no WebView2 (Chromium) e em `http://127.0.0.1`, que conta como contexto seguro.
- `br-button secondary` **Baixar .docx** — chama o gerador python-docx atual e devolve o arquivo.
- O documento, dentro de um `<iframe>` que carrega `/termo/<tipo>/<chave>/documento`.

O iframe isola o documento do CSS do DS. Seu conteúdo é HTML autônomo gerado por `termos_html.py`,
**no padrão do `/gelic`**: `templates/termo_base.html` copia a folha de estilo de documento oficial
de `gelic/assets/base.html` (Times New Roman 11.5pt, A4, parágrafos justificados com recuo, tabela
com borda colapsada e cabeçalho cinza), e o corpo é montado por funções Python pequenas
(`esc`, `tabela`, `assinatura`, `corpo_ccusto`, `corpo_individual`, `corpo_devolucao`) com o mesmo
texto e a mesma ordem do `.docx` correspondente.

Larguras da tabela: **80 % centralizada** no termo individual e no de devolução; **100 %** no termo por
centro de custo.

**Exceção deliberada ao DS:** `table`, `th` e `td` do documento levam `style=` inline (borda, largura,
padding, alinhamento), porque é isso que a área de transferência carrega para o SEI — a folha de
estilo do `<style>` não vai junto. É o único inline style do projeto e fica confinado a
`termos_html.py`.

O HTML gerado também é gravado em `dados/saida/` ao lado do `.docx` (é uma linha; serve de registro).

## 8. Programa de desktop (Windows)

`main.py` (~20 linhas): sobe o Flask numa thread em `127.0.0.1:5000` e abre uma janela `pywebview`
("Termos de Responsabilidade – CFC", 1100×750, ícone do CFC). Fechar a janela encerra o programa. Se
a criação da janela falhar (WebView2 ausente), abre o navegador padrão em `http://127.0.0.1:5000` e
mantém o processo vivo até Ctrl+C. Porta ocupada: mensagem clara e sai.

`config.py`: `PASTA_DADOS` = pasta do executável quando congelado (`sys.frozen`), senão `./dados`.
Na primeira execução, se `termos.db` não existir, cria o esquema vazio e copia `timbrado.docx` do
pacote para `dados/`.

`build.bat`: um comando PyInstaller `--onedir --windowed --add-data static;static
--add-data templates;templates --add-data timbrado.docx;.` → `dist/TermosCFC/`. `--onedir` abre mais
rápido e incomoda menos o antivírus que `--onefile`. Opcional, depois: `instalador.iss` (Inno Setup)
→ `Setup-TermosCFC.exe` com atalho no menu Iniciar. O build roda no Windows do usuário; nesta VPS
Linux o mesmo `main.py` abre uma janela GTK/Qt e serve para validar.

Zip do `dist/` com `dados/` vazio ao lado é a forma de entrega mínima.

## 9. Estrutura final de arquivos

```
app.py                      rotas Flask (Excel → db.py; rotas novas de §6.1, §5.3, §7)
config.py                   PASTA_DADOS e caminhos derivados
db.py                       conexão, esquema, importação de bens, consultas de §4.3, operações de §6.1
termos_html.py              corpo HTML dos três termos (padrão gelic)
Script_Termo_Individual.py  gerador docx (entrada: lista de dicts)
Termo_de_Responsabilidade.py
termo_devolucao.py
main.py                     janela pywebview + thread Flask
importar_planilhas.py       migração inicial (executado uma vez)
build.bat
requirements.txt            flask, openpyxl, python-docx, pywebview, pyinstaller
templates/
  base.html                 estrutura DSGov
  index.html                busca de bem + atalhos
  centro_custos.html, termos_individuais.html, termo_devolucao.html, upload.html
  cadastros.html            três abas
  bem.html                  ficha do bem
  termo.html                página do termo (botões + iframe)
  termo_base.html           folha de estilo de documento (do gelic)
static/dsgov/               CSS/JS/fontes vendorizados
static/logo.png
dados/                      (não versionado) termos.db, timbrado.docx, saida/
tests/
docs/superpowers/specs/
```

## 10. Testes (pytest)

Sem UI. Banco em memória ou em `tmp_path`.

- `test_db.py`: esquema cria; importação de um `.xlsx` de 5 linhas (gerado no teste com openpyxl)
  substitui `bens`; importação com cabeçalho errado não altera nada; importação que removeria bem
  atribuído é revertida com a lista de números; `bens_do_centro` exclui bem atribuído (§5.2);
  renomear centro cascateia para `localizacoes`; excluir centro com localizações falha.
- `test_app.py` (Flask test client): cada rota de termo responde 200 e o documento HTML contém as
  linhas esperadas e a largura da tabela certa (80 %/100 %); `/bem/<n>` mostra pessoa e setor;
  atribuir bem já atribuído exige confirmação; upload de xlsx válido reporta contagem.
- `test_docx.py`: cada gerador produz um `.docx` com N+2 linhas na tabela (cabeçalho + N + total) a
  partir de listas de dicts — só garante que a troca de DataFrame por dict não quebrou nada.

Os testes existentes: não há. Os geradores de `.docx` já estão validados na prática pelo usuário; não
se testa formatação.

## 11. Fora de escopo

Docker/Linux como serviço (previsto por `PASTA_DADOS`, não implementado); assinatura digital do `.exe`;
edição in-place nas tabelas; login/permissões; histórico de termos emitidos; múltiplas pessoas por bem;
Inno Setup (opcional, depois do primeiro `.exe` funcionar).
