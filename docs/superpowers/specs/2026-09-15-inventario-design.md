# Módulo de inventário

**Data:** 2026-09-15
**Estado:** aprovado pelo usuário em 2026-09-15 (execução autorizada, commit e push ao final).
**Base:** `main` em `2cf7ff0`+ (rodada "processos SEI, histórico, painel e recorte" concluída e publicada).
**Origem das regras:** sistema antigo em `/opt/web/sga/cfc` (sistemadeinventario.com.br), levantado em
`docs/superpowers/notes/2026-09-15-inventario-existente.md`. Contexto e decisões em
`docs/superpowers/notes/2026-09-15-inventario-migracao-contexto.md`.

## 1. Objetivo

Trazer o inventário físico para este sistema, em SQLite e DSGov, como **módulo separado** (menu
"Inventário", arquivos próprios), migrando as regras do sistema antigo e acrescentando o que ele não
tinha: o **evento de inventário** como entidade, com histórico por evento e por sala.

Fluxo: abre-se um evento (campanha) → cada sala do escopo é conferida lendo as plaquetas (leitor
USB/Bluetooth, câmera do celular ou digitação) → bem lido na sala cadastrada = **localizado**; lido em
sala diferente = **divergente**; bem da sala não lido = **não localizado**; bem sem cadastro =
**sobra** → relatório e `.xlsx` do evento → encerramento.

Princípios: simplicidade; só acréscimos ao esquema; `bens` continua espelho do SPW e **nunca** muda por
aqui; nada de histórico é apagado (sobra é a única exceção, como no sistema antigo).

Fora do escopo: login/perfis (site segue com a senha do nginx), PDF, dashboard por andar, migração das
leituras antigas da planilha, o legado `terms.py` do sistema antigo.

---

## 2. Modelo

### 2.1 Esquema (acrescentado a `db.ESQUEMA`)

```sql
CREATE TABLE IF NOT EXISTS inventario_eventos (
  id           INTEGER PRIMARY KEY,
  nome         TEXT NOT NULL,                 -- "Inventário 2026"
  descricao    TEXT,                          -- portaria, comissão, observações
  aberto_em    TEXT NOT NULL,                 -- ISO
  encerrado_em TEXT                           -- NULL = aberto
);
-- "um evento aberto por vez" é garantido em abrir_evento() (índice único em coluna NULL não serve: NULLs são distintos)

CREATE TABLE IF NOT EXISTS inventario_integrantes (
  evento_id INTEGER NOT NULL REFERENCES inventario_eventos(id) ON DELETE CASCADE,
  nome      TEXT NOT NULL,
  PRIMARY KEY (evento_id, nome)
);
CREATE TABLE IF NOT EXISTS inventario_salas (
  evento_id    INTEGER NOT NULL REFERENCES inventario_eventos(id) ON DELETE CASCADE,
  localizacao  TEXT NOT NULL,                 -- valor de bens.localizacao
  PRIMARY KEY (evento_id, localizacao)
);
CREATE TABLE IF NOT EXISTS inventario_leituras (
  id          INTEGER PRIMARY KEY,
  evento_id   INTEGER NOT NULL REFERENCES inventario_eventos(id) ON DELETE CASCADE,
  numero      INTEGER NOT NULL,               -- bens.numero (sem FK: bens é substituída na importação)
  localizacao TEXT NOT NULL,                  -- sala onde foi lido ("local inventário")
  lido_em     TEXT NOT NULL,
  integrante  TEXT NOT NULL,
  conservacao TEXT CHECK (conservacao IN ('Bom','Regular','Ruim','Inservível')),
  quem_usa    TEXT,
  observacao  TEXT,
  foto_url    TEXT,
  UNIQUE (evento_id, numero)                  -- um bem, uma leitura por evento (reler atualiza)
);
CREATE TABLE IF NOT EXISTS inventario_sobras (
  id          INTEGER PRIMARY KEY,
  evento_id   INTEGER NOT NULL REFERENCES inventario_eventos(id) ON DELETE CASCADE,
  localizacao TEXT NOT NULL,
  descricao   TEXT NOT NULL,
  complemento TEXT,
  observacao  TEXT NOT NULL,
  foto_url    TEXT NOT NULL,
  integrante  TEXT NOT NULL,
  criado_em   TEXT NOT NULL
);
```

### 2.2 Situação de um bem no evento (calculada, nunca gravada)

Para a sala S do evento E, os bens considerados são os **ativos** com `bens.localizacao = S` (o "local
sistema"), mais os lidos em S que pertencem a outra sala:

| situação | regra |
|---|---|
| localizado | leitura em E com `localizacao = bens.localizacao` |
| divergente | leitura em E com `localizacao ≠ bens.localizacao` (mostra "cadastrado em X") |
| não localizado | bem ativo da sala sem leitura em E |
| sobra | linha em `inventario_sobras` |

Reler um bem no mesmo evento atualiza sala, data/hora e integrante da leitura; conservação, quem usa,
observação e foto ficam (não se perde dado; a leitura anterior é mostrada na tela). Ao registrar sobra, o
integrante também tem de estar na comissão do evento.

Bem BAIXADO/DOADO/INSERVÍVEL lido: aceito e registrado (é informação útil), com aviso "bem não ativo";
não entra na contagem da sala.

### 2.3 Módulo `inventario.py` (dados; `conn` primeiro, sem Flask, mesmo padrão de `db.py`)

```python
CONSERVACAO = ("Bom", "Regular", "Ruim", "Inservível")

def evento_aberto(conn) -> dict | None
def eventos(conn) -> list[dict]                                  # aberto primeiro, depois por aberto_em desc
def evento(conn, id) -> dict | None                              # + integrantes (lista) + resumo (ver abaixo)
def abrir_evento(conn, nome, descricao, integrantes: list[str], salas: list[str] | None = None) -> int
    # ErroDeNegocio se já houver aberto, nome vazio, sem integrante ou sem sala;
    # salas=None → todas as localizações com bens ATIVO (db.localizacoes_ativas)
def encerrar_evento(conn, id) -> None                            # grava encerrado_em; idempotente
def salas(conn, evento_id) -> list[dict]
    # por sala: localizacao, ccustos (via localizacoes), total (ativos), localizados, divergentes, pendentes
    # (não existe "concluir sala": o evento fica aberto até ser encerrado; a sala está "completa" quando pendentes = 0)
def bens_da_sala(conn, evento_id, localizacao) -> dict
    # {"bens": [bem + leitura + situacao], "trazidos": [lidos aqui mas de outra sala], "sobras": [...]}
def ler(conn, evento_id, localizacao, numero: int, integrante: str) -> dict
    # ErroDeNegocio: evento encerrado, sala fora do escopo, integrante fora da lista, bem inexistente
    # ("Bem não encontrado" → a tela oferece registrar sobra). INSERT ou UPDATE (UNIQUE evento+numero).
    # devolve {"situacao": "localizado"|"divergente", "bem": {...}, "cadastrado_em": str, "reler": bool,
    #          "leitura_anterior": {...} | None}
def atualizar_leitura(conn, evento_id, numero, conservacao=None, quem_usa=None, observacao=None, foto_url=None) -> None
def registrar_sobra(conn, evento_id, localizacao, descricao, complemento, observacao, foto_url, integrante) -> int
    # descricao, observacao e foto_url obrigatórios (ErroDeNegocio)
def excluir_sobra(conn, evento_id, sobra_id) -> dict            # devolve a sobra (para apagar a foto); só sobras
def relatorio(conn, evento_id, localizacao=None, situacao=None) -> list[dict]
    # uma linha por bem do escopo (todas as salas ou uma), com situação, leitura e dados do bem
def exportar_xlsx(conn, evento_id, destino, localizacao=None) -> destino   # abas "Bens" e "Sobras"
def resumo(conn, evento_id) -> dict
    # {"salas": n, "salas_iniciadas": n, "bens": n, "lidos": n, "divergentes": n, "pendentes": n, "sobras": n, "pct_bens": float}
```

`db.py` ganha só `localizacoes_ativas(conn) -> list[str]` (localizações distintas de bens ATIVO,
ordenadas) e o esquema. `situacao` de leitura é calculada com um `LEFT JOIN bens` (`localizacao ≠
bens.localizacao`).

### 2.4 Fotos: módulo `fotos.py`

Transcrição do sistema antigo (`app/services/storage.py` e `images.py`), sem Flask:

```python
def configurado() -> bool                       # todas as env R2_ACCESS_KEY_ID, R2_SECRET_ACCESS_KEY, R2_ENDPOINT_URL, R2_BUCKET_NAME presentes
def validar(arquivo) -> None                    # ErroDeNegocio: extensão fora de {.jpg,.jpeg,.png,.webp}, > 5 MB, Pillow verify falha
def comprimir(arquivo) -> bytes                 # WebP q85, máx 1920×1080 mantendo proporção, EXIF orientation aplicada
def enviar(nome: str, dados: bytes) -> str      # boto3 put_object em R2_BUCKET_NAME, chave "inventario/<nome>", ContentType image/webp;
                                                # devolve URL pública R2_PUBLIC_URL/inventario/<nome> (ou endpoint/bucket/... sem PUBLIC_URL)
def apagar(url: str) -> None                    # delete_object pela chave extraída da URL; erro só loga
def nome_bem(evento_id, numero) -> str          # f"INV{evento_id}_BEM_{numero}_{YYYYMMDDHHMMSS}.webp"
def nome_sobra(evento_id, sobra_id) -> str      # f"INV{evento_id}_SOBRA_{sobra_id}_{YYYYMMDDHHMMSS}.webp"
```

Credenciais: as mesmas variáveis do sistema antigo, lidas de `secrets/.env` pelo `compose.yml`
(`env_file: - path: secrets/.env, required: false`); `secrets/` no `.gitignore`. Sem configuração
(programa Windows offline), `configurado()` é falso: os botões de foto ficam desabilitados com o aviso
"Fotos desativadas: bucket não configurado" e **sobra passa a exigir foto só quando fotos estão
ativas** (senão a sobra é aceita sem foto, com a observação obrigatória). Tests usam um `fotos.enviar`
substituído por monkeypatch; nada de rede nos testes.

---

## 3. Rotas: blueprint `app_inventario.py`

Registrado em `app.py` com `app.register_blueprint(inventario_bp)`; prefixo `/inventario`. Usa
`obter_conn()` e o `errorhandler` de `ErroDeNegocio` já existentes (o handler é global). O integrante
escolhido fica em `session["integrante"]`.

| rota | método | faz |
|---|---|---|
| `/inventario` | GET | eventos: aberto em destaque (resumo, barra de progresso) + histórico; formulário "Abrir evento" |
| `/inventario/abrir` | POST | nome, descricao, integrantes (textarea, um por linha), escopo: `todas` ou lista de salas marcadas |
| `/inventario/<id>` | GET | salas do evento (tabela com busca, contadores, tag concluída, link Ler), resumo, botões Relatório / Encerrar |
| `/inventario/<id>/encerrar` | POST | confirmação por campo `confirmar=1`; grava `encerrado_em` |
| `/inventario/<id>/integrante` | POST | grava `session["integrante"]` (select entre os integrantes do evento) e volta |
| `/inventario/<id>/sala/<localizacao>` | GET | tela de leitura |
| `/inventario/<id>/sala/<localizacao>/ler` | POST (JSON) | `{"numero": "…"}` → `inventario.ler(...)`; 200 com o dict; 404 `{"erro": "Bem não encontrado", "numero": n}`; 409 evento encerrado/sem integrante |
| `/inventario/<id>/leitura/<numero>` | POST (JSON) | atualizar conservação / quem usa / observação (auto-save ao sair do campo) |
| `/inventario/<id>/leitura/<numero>/foto` | POST (multipart) | valida, comprime, envia, grava `foto_url`; JSON `{"foto_url"}` |
| `/inventario/<id>/leitura/<numero>/foto/excluir` | POST | apaga no bucket e zera `foto_url` |
| `/inventario/<id>/sala/<localizacao>/sobra` | POST (multipart) | descrição, complemento, observação, foto → cria sobra; se o upload falhar, não cria (ordem: valida → comprime → cria linha → envia → grava url; falha no envio apaga a linha) |
| `/inventario/<id>/sobra/<sobra_id>/excluir` | POST | apaga sobra e sua foto |
| `/inventario/<id>/relatorio` | GET | tabela filtrável (sala, situação) |
| `/inventario/<id>/xlsx` | GET | `.xlsx` em memória (`_baixar`), mesmos filtros |

Evento encerrado: toda rota de escrita responde `ErroDeNegocio("Evento encerrado")` (flash ou JSON 409).

---

## 4. Telas (templates `inventario_eventos.html`, `inventario_evento.html`, `inventario_sala.html`, `inventario_relatorio.html`)

**Eventos.** Card do evento aberto: nome, aberto em, integrantes, resumo em cards pequenos (salas
iniciadas/total, bens lidos/total, divergentes, sobras) e `br-progress`-like (barra DSGov). Formulário
"Abrir evento" só quando não há aberto: nome, descrição, integrantes (textarea), escopo (`br-radio`
"Todas as salas com bens ativos" / "Escolher salas" → lista de checkboxes das `localizacoes_ativas`).
Histórico: tabela dos encerrados com link.

**Evento.** Tabela das salas (macro `cabecalho_tabela`, busca): sala, centro de custo, bens, lidos,
divergentes, pendentes, situação (`br-tag` completa (pendentes = 0) / em andamento / não iniciada), botão **Ler**.
Barra superior: select de integrante (o da sessão, marcado), botões Relatório e Encerrar (com
confirmação em duas etapas como a exclusão de centro).

**Leitura da sala** (a tela que importa; funciona no celular):

- Cabeçalho: sala, centro de custo, integrante (link para trocar), contadores ao vivo (lidos/total,
  divergentes, pendentes).
- Campo de leitura: `<input id="leitura" inputmode="none" autocomplete="off" enterkeyhint="done">`,
  autofocus e **refocus** após cada leitura e ao clicar fora (o leitor USB/Bluetooth digita e envia
  Enter; `inputmode="none"` evita o teclado virtual). Botões: **Câmera** (abre `html5-qrcode`
  embutido — `static/dsgov/vendor/html5-qrcode/html5-qrcode.min.js`, versão 2.3.8, baixada do
  unpkg na Task de infra e commitada — num `br-modal`; ao ler, preenche o campo e envia; o scanner fica
  aberto para leituras seguidas até fechar) e **Digitar** (troca `inputmode` para `numeric` e foca, para
  plaqueta ilegível; volta a `none` após enviar).
- Envio por `fetch` JSON; a página nunca recarrega numa leitura. Resposta atualiza a linha do bem na
  lista (ou insere em "Trazidos de outra sala"), toca um `br-message` curto e some em 3 s:
  - localizado: verde "Bem 14359 localizado".
  - divergente: amarelo "Bem 14359 cadastrado em 03 - CGTI; registrado aqui".
  - reler: se o bem já tinha leitura neste evento, a resposta traz `reler=true` e a leitura anterior; a
    tela mostra "Já lido em 12/09 por Fulano em 02 - CCOM. Atualizado." (a regravação é feita direto,
    sem modal: no sistema antigo o modal existia porque a planilha era o registro mestre; aqui o histórico
    de leituras é do evento e a última leitura vale).
  - não encontrado: vermelho "Bem 99999 não está na base" + botão **Registrar sobra** que abre o
    formulário (descrição, complemento, observação, foto com `capture="environment"`).
  - bem não ativo: azul "Bem 1003 está BAIXADO; leitura registrada".
- Lista de bens da sala (tabela DSGov): número, descrição, complemento, chip de situação (Localizado
  verde / Divergente amarelo com "cadastrado em …" / Pendente cinza), conservação (`select` DSGov de 4
  opções), quem usa (input), observação (input), foto (miniatura 56 px, clique abre grande com
  Excluir; botão Câmera abre `<input type="file" accept="image/*" capture="environment">`). Campos
  salvam ao `change` via `fetch`. Pendentes primeiro, depois localizados, depois divergentes.
- Seções "Trazidos de outras salas" e "Sobras desta sala" (com Excluir).
- Rodapé: voltar ao evento. Não há "concluir sala": o evento fica aberto até ser encerrado.
- Evento encerrado: tudo somente leitura, campo de leitura desabilitado, aviso no topo.

**Relatório.** Filtros (sala, situação), cards (lidos, divergentes, pendentes, sobras), tabela com as
colunas do xlsx, botão **Exportar .xlsx**.

**Painel (tela inicial).** Card "Inventário em andamento: N% dos bens · M divergentes" (link para o
evento) quando há evento aberto; senão "Nenhum inventário aberto" (link para Inventário).

**Menu.** "Inventário" (ícone `fa-clipboard-check`) entre "Recorte" e "Cadastros".

---

## 5. Exportação `.xlsx` (openpyxl, em memória)

Aba **Bens** (uma linha por bem do escopo, ou da sala filtrada): Patrimônio, Descrição, Complemento,
Classificação, Local sistema, Local inventário, Situação (Localizado/Divergente/Não localizado),
Conservação, Quem usa, Observação, Integrante, Data/hora, Foto (URL). Aba **Sobras**: Sala, Descrição,
Complemento, Observação, Integrante, Data/hora, Foto (URL). Primeira linha: título do evento, data de
geração e filtro. Nome do arquivo: `inventario_<id>_<sala ou tudo>.xlsx`.

---

## 6. Dependências e infra

- `requirements.txt`: `boto3`, `Pillow` (o Dockerfile instala os mesmos pacotes explicitamente; o
  `build.bat`/PyInstaller do desktop passa a embutir ambos).
- `static/dsgov/vendor/html5-qrcode/html5-qrcode.min.js` + LICENSE (Apache-2.0), carregado só na tela de
  leitura. Sem CDN (offline).
- `compose.yml`: `env_file: [{path: secrets/.env, required: false}]`; `.gitignore` ganha `secrets/`.
- `nginx`: `client_max_body_size 20m` já cobre fotos de 5 MB.

## 7. Regras migradas do sistema antigo (com origem)

- Número lido: só dígitos após remover zeros à esquerda e espaços (`inventory.py:39-41`; aqui a base é
  numérica, então letras/pontos viram "não encontrado").
- Conservação em `('Bom','Regular','Ruim','Inservível')` (`validation.py:55-61`).
- Foto: extensão, 5 MB, `Image.verify()`, WebP q85 1920×1080 (`validation.py:64-104`, `images.py`).
- Sobra exige descrição, observação e foto; rollback se o upload falhar (`inventory.py:688-746`).
- Só sobra pode ser excluída (`inventory.py:803-817`).
- Divergente sempre calculado, nunca gravado (`reports.py:112-132`).
- Bem lido em qualquer sala é aceito e sinalizado; nunca bloqueado (`inventory.py:334-337`).
- Colunas do relatório (`reports.py:26-45`), sem PDF.

## 8. Testes

- `tests/test_inventario.py` (dados): abrir evento (todas as salas / amostra / erro com aberto /
  validações), `salas` com contadores, `ler` nos quatro casos (localizado, divergente, reler atualiza,
  não encontrado, bem não ativo, evento encerrado, sala fora do escopo, integrante inválido),
  `atualizar_leitura`, `registrar_sobra` (obrigatórios; sem foto quando fotos desativadas), `excluir_sobra`,
   `relatorio` e `resumo`, `exportar_xlsx` (abas, cabeçalhos, linhas).
- `tests/test_fotos.py`: `validar` (extensão, tamanho, conteúdo), `comprimir` (WebP, ≤1920), nomes,
  `configurado` com/sem env; `enviar`/`apagar` com `boto3` substituído por um cliente falso.
- `tests/test_app.py`: rotas 200/JSON (ler 200/404/409), fluxo abrir → ler → sobra (com `fotos.enviar`
  monkeypatched) → relatório → xlsx → encerrar → escrita bloqueada; card do painel; menu.

---

## 9. Planilha de cadastros: abas de inventário (migração de inventários antigos)

A exportação em Cadastros → *Exportar cadastros* passa a gerar, além das 4 abas atuais, cinco abas
opcionais com o **estado inteiro** das tabelas de inventário, no formato do banco (uma coluna por
campo, cabeçalho = nome da coluna):

| aba | colunas |
|---|---|
| `inv_eventos` | id, nome, descricao, aberto_em, encerrado_em |
| `inv_integrantes` | evento_id, nome |
| `inv_salas` | evento_id, localizacao |
| `inv_leituras` | evento_id, numero, localizacao, lido_em, integrante, conservacao, quem_usa, observacao, foto_url |
| `inv_sobras` | evento_id, localizacao, descricao, complemento, observacao, foto_url, integrante, criado_em |

Importar (*Atualizar base → Importar cadastros*): as abas `inv_*` são **opcionais**. Se nenhuma existir,
as tabelas de inventário não são tocadas (planilhas antigas continuam válidas). Se **qualquer** uma
existir, as cinco tabelas são substituídas pelo conteúdo das abas presentes (aba ausente = tabela
vazia), tudo ou nada junto com as 4 abas de cadastro. Validações (mesmo estilo das atuais, linha e
motivo): `id` de evento único e inteiro; `evento_id` das outras abas existe em `inv_eventos`; no máximo
um evento sem `encerrado_em`; `numero` de leitura existe em `bens`; `conservacao` vazia ou em
`CONSERVACAO`; `(evento_id, numero)` de leitura único; datas em ISO `YYYY-MM-DD HH:MM:SS` (ou
`YYYY-MM-DD`, completada com `00:00:00`); sobra com descricao, observacao, integrante.

É assim que um inventário feito no sistema antigo (planilha com "Local Inventariado", "Conservação",
"Integrante", "Data/Hora Inventário", "Usuário do Bem", "Observação", "Foto") migra: o usuário monta
`inv_eventos` com uma linha (o evento antigo, já encerrado), `inv_salas` com as salas, e `inv_leituras`
com uma linha por bem lido, copiando as colunas da planilha antiga. Os ids de evento ficam como estão
na aba (a importação preserva o `id`).

`db.exportar_cadastros` / `db.importar_cadastros` são estendidos; o módulo `inventario.py` fornece
`exportar_abas(conn, wb)` e `validar_abas(conn, brutos) -> (linhas_por_tabela, problemas)` para que a
regra de inventário não vaze para `db.py`.
