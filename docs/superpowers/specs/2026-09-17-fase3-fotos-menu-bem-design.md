# Fase 3 do inventário: várias fotos por bem, chaves simples no R2, submenus e fotos no cadastro do bem

**Data:** 2026-09-17
**Estado:** desenho aprovado pelo usuário em 2026-09-17 (execução autorizada por subagentes).
**Base:** `main` em `fa52d4f` (Fase 2 publicada).
**Spec anterior:** `2026-09-17-fase2-inventario-design.md`; spec do módulo: `2026-09-15-inventario-design.md`.

## 1. Objetivo e decisões

Três pedidos do usuário, todos sobre o inventário:

1. **Fotos no R2 com chave simples e minúscula**: `inventario2026/1-12334.webp` (pasta = nome do evento
   normalizado; `1` = número da foto do bem no evento; `12334` = número do bem). Continua usando o bucket e as
   variáveis `R2_*` de hoje.
2. **Menu**: "Inventário" vira um grupo com "Eventos" e, quando há evento aberto, as telas desse evento
   (Salas, Painel, Relatório).
3. **Cadastro do bem** (`/bem?numero=`): as fotos que o bem recebeu em cada evento de inventário, agrupadas
   por evento, na ordem em que foram tiradas.

Decisões do usuário (AskUserQuestion, 2026-09-17):

- **Várias fotos por bem no mesmo evento** (hoje é uma só, em `inventario_leituras.foto_url`). Exige tabela
  nova.
- **Relatório e xlsx mostram só a primeira foto**; na tela, "+N" quando houver mais. As demais aparecem na
  sala e no cadastro do bem.
- Pasta pelo **nome do evento**, tudo minúsculo, sem acento, sem espaço nem pontuação.

Fora: PDF, login, limite de fotos por bem, reenviar fotos antigas com o nome novo, editar nome de evento.

---

## 2. Chaves no bucket (`fotos.py`)

```python
def pasta(nome: str, evento_id: int) -> str
    # NFD, remove marcas combinantes, casefold, mantém só [a-z0-9]. "Inventário 2026" → "inventario2026".
    # Resultado vazio → f"evento{evento_id}".

def chave_bem(pasta: str, nfoto: int, numero: int) -> str      # f"{pasta}/{nfoto}-{numero}.webp"
def chave_sobra(pasta: str, sobra_id: int) -> str               # f"{pasta}/sobra-{sobra_id}.webp"

def enviar(chave: str, dados: bytes) -> str
    # put_object com a chave exatamente como recebida (PREFIXO "inventario/" deixa de existir); devolve _url(chave).

def apagar(url: str | None) -> None
    # chave = url sem o prefixo _url("") (R2_PUBLIC_URL/ ou endpoint/bucket/). Se a URL não começa por essa base,
    # tenta o formato antigo (procura "inventario/" na URL, como hoje). Nada encontrado → ignora.
```

`nome_bem`/`nome_sobra`/`_carimbo`/`PREFIXO` saem. `configurado`, `validar`, `comprimir`, `_cliente`, `_url` ficam.

- **Fotos antigas** (`inventario/INV1_BEM_..._carimbo.webp`) continuam válidas: a URL gravada não muda e
  `apagar` ainda reconhece o formato antigo. Nada é copiado nem renomeado no bucket.
- **Sobra** continua com uma foto: chave `{pasta}/sobra-{id}.webp`. Sobra apagada e recriada tem id novo,
  então não há reaproveitamento de chave.
- **Contador `nfoto`** por `(evento, numero)`: sempre `MAX(nfoto) + 1` (1 quando não há nenhuma). Nunca
  reaproveitado: apagar a foto 2 e tirar outra gera a 3. Assim a URL nova nunca coincide com uma já servida
  pelo cache do navegador ou do R2.
- **Pasta única por evento**: como `nome` não é único em `inventario_eventos`, `abrir_evento` recusa nome
  cuja `pasta` coincida com a de qualquer evento existente: `ErroDeNegocio("Já existe um evento com esse
  nome (pasta de fotos 'inventario2026'); escolha outro nome.")`. `validar_abas` faz a mesma checagem entre
  as linhas de `inv_eventos` ("pasta de fotos repetida"). O nome do evento não é editável hoje, logo a pasta é
  estável.

---

## 3. Modelo

### 3.1 Tabela nova (em `db.ESQUEMA`)

```sql
CREATE TABLE IF NOT EXISTS inventario_fotos (
  evento_id INTEGER NOT NULL,
  numero    INTEGER NOT NULL,
  nfoto     INTEGER NOT NULL,
  url       TEXT NOT NULL,
  criado_em TEXT NOT NULL,
  PRIMARY KEY (evento_id, numero, nfoto),
  FOREIGN KEY (evento_id, numero) REFERENCES inventario_leituras(evento_id, numero) ON DELETE CASCADE
);
```

`inventario_leituras` perde a coluna `foto_url` no `CREATE TABLE` e ganha `fotos_seq INTEGER NOT NULL DEFAULT 0`
(maior `nfoto` já usado nessa leitura; decisão tomada na execução, 2026-09-17: só `MAX(nfoto) + 1` reaproveitaria
o número depois de apagar a última foto). `adicionar_foto` usa `nfoto = max(fotos_seq, MAX(nfoto)) + 1` e grava
`fotos_seq = nfoto`; o `MAX` cobre bancos migrados e leituras recriadas por importação. A chave estrangeira composta aponta para
`UNIQUE (evento_id, numero)` de `inventario_leituras`; com `PRAGMA foreign_keys = ON` (já ligado em
`db.conectar`), apagar a leitura apaga as fotos. A rota recolhe as URLs **antes** do DELETE para apagar no
bucket, como o Desmarcar já faz.

### 3.2 Migração (em `db.criar_esquema`, depois do `executescript`)

```python
if "foto_url" in _colunas(conn, "inventario_leituras"):
    conn.execute("""INSERT OR IGNORE INTO inventario_fotos (evento_id, numero, nfoto, url, criado_em)
                    SELECT evento_id, numero, 1, foto_url, lido_em FROM inventario_leituras
                    WHERE foto_url IS NOT NULL AND foto_url <> ''""")
    conn.execute("ALTER TABLE inventario_leituras DROP COLUMN foto_url")
```

Roda uma vez por banco (produção no primeiro start do container; banco local no primeiro `db.conectar`).
Segue o padrão das migrações já existentes na função (`tratamento`, `email`, `bloco_sei`).

### 3.3 Planilha de cadastros (`inventario.ABAS`)

- `inv_leituras` deixa de ter `foto_url` e ganha `fotos_seq` como última coluna (opcional na importação, padrão 0;
  inteiro ≥ 0), para o contador sobreviver a exportar/importar.
- Aba nova, exportada por último: `"inv_fotos": ["evento_id", "numero", "nfoto", "url", "criado_em"]`
  (tabela `inventario_fotos`). Como `inv_bens_encerrados`, é **opcional**: ausente → a regra abaixo; presente →
  substituída junto com as outras. A regra "as 5 abas originais juntas ou nenhuma" não muda.
- **Planilha antiga** (exportada antes desta versão) tem `foto_url` em `inv_leituras` e não tem `inv_fotos`.
  `db._ler_aba_cadastro` ganha o parâmetro `opcionais: tuple = ()` — colunas ausentes do cabeçalho entram
  como `None` em vez de virar problema. `db.importar_cadastros` lê **só** `inv_leituras` com
  `opcionais=("foto_url", "fotos_seq")`; as outras abas mantêm a checagem estrita de colunas (`inv_sobras.foto_url`
  continua obrigatória).
  `validar_abas`: quando `brutos["inv_fotos"]` está **vazio** (aba ausente ou sem linhas — `validar_abas` não
  distingue e não precisa), cada `foto_url` não vazia de `inv_leituras` vira uma linha
  `(evento_id, numero, 1, url, lido_em)` em `linhas["inv_fotos"]`; quando há linhas em `inv_fotos`, a coluna
  `foto_url` de `inv_leituras` é ignorada. Em `db.importar_cadastros`, a lista `faltam` (abas obrigatórias)
  passa a excluir `inv_fotos` além de `inv_bens_encerrados`.
- Validação de `inv_fotos`: `evento_id` existe em `inv_eventos`; `(evento_id, numero)` existe entre as
  leituras válidas da planilha; `nfoto` inteiro ≥ 1; `url` não vazia; `(evento_id, numero, nfoto)` único.
- `substituir_tabelas`: `DELETE` de `inventario_fotos` junto com as outras (ordem reversa de `ABAS`) e
  `INSERT` depois de `inventario_leituras`. `inv_fotos` ausente e sem `foto_url` antiga → tabela fica vazia
  (a planilha é a fonte de verdade quando as abas `inv_*` vêm).

Snapshot do encerramento (`inventario_bens_encerrados`) não muda: fotos já são por evento.

---

## 4. Funções de dados (`inventario.py`)

```python
def pasta_do_evento(conn, evento_id: int) -> str        # fotos.pasta(evento.nome, evento_id)

def fotos_do_bem_no_evento(conn, evento_id: int, numero: int) -> list[dict]
    # [{"nfoto", "url", "criado_em"}] ORDER BY nfoto

def adicionar_foto(conn, evento_id: int, numero: int, enviar) -> list[dict]
    # evento aberto obrigatório; leitura (evento_id, numero) obrigatória ("Leia o bem antes de fotografar.");
    # nfoto = COALESCE(MAX(nfoto), 0) + 1; chave = fotos.chave_bem(pasta, nfoto, numero);
    # url = enviar(chave)  (callable recebido da rota: fotos.enviar com os bytes já comprimidos — mantém o
    # módulo de dados sem rede e testável). enviar roda ANTES do INSERT: se levantar exceção, nada é gravado
    # e a exceção propaga para a rota. Depois INSERT; commit; devolve fotos_do_bem_no_evento(...).

def apagar_foto(conn, evento_id: int, numero: int, nfoto: int) -> str | None
    # evento aberto obrigatório; DELETE da linha; devolve a url apagada (None se não existia). commit.

def desfazer_leituras(conn, evento_id, numeros) -> tuple[list, int]
    # como hoje, mas as urls vêm de inventario_fotos (todas as fotos dos bens informados).

def fotos_do_bem(conn, numero: int) -> list[dict]
    # Para o cadastro do bem: [{"evento_id", "evento": nome, "aberto_em", "encerrado_em", "lido_em",
    #   "fotos": [{"nfoto", "url"}]}] — só eventos em que o bem tem ao menos uma foto;
    #   eventos do mais recente para o mais antigo (aberto_em DESC, id DESC); fotos por nfoto.
```

`atualizar_leitura` deixa de aceitar `foto_url` (permitidos: `conservacao`, `quem_usa`, `observacao`).
`_LEITURA` troca `r.foto_url` por duas subconsultas correlacionadas:

```sql
(SELECT url FROM inventario_fotos f WHERE f.evento_id = r.evento_id AND f.numero = r.numero ORDER BY f.nfoto LIMIT 1) AS foto_url,
(SELECT COUNT(*) FROM inventario_fotos f WHERE f.evento_id = r.evento_id AND f.numero = r.numero) AS n_fotos
```

Assim `bens_da_sala`, `relatorio`, `exportar_xlsx`, `contar_fotos`, `_tem_foto` e o filtro `foto=com|sem`
continuam funcionando sem mudança ("primeira foto"). `bens_da_sala` acrescenta em cada bem/trazido a lista
`fotos` (`fotos_do_bem_no_evento`) para a tela da sala desenhar todas as miniaturas — uma consulta por bem
lido é aceitável (salas têm dezenas de bens; painel com 185 SELECTs leva 0,09 s).

`abrir_evento`: checagem de pasta única (seção 2). `registrar_sobra`/`definir_foto_sobra` ficam; só a rota
muda a chave.

---

## 5. Rotas (`app_inventario.py`)

| rota | método | faz |
|---|---|---|
| `/inventario/<id>/leitura/<numero>/foto` | POST (multipart `foto`) | `_foto_processada()`; `inventario.adicionar_foto(conn, id, numero, lambda chave: fotos.enviar(chave, dados))`; falha de rede → `_json_erro("Falha ao enviar a foto.")` sem gravar; resposta `{"fotos": [{"nfoto", "url"}]}` |
| `/inventario/<id>/leitura/<numero>/foto/<int:nfoto>/excluir` | POST (form `volta`) | `inventario.apagar_foto` → `fotos.apagar(url)`; redirect para a sala (`volta`), como hoje |
| `/inventario/<id>/sala/<loc>/lote` (desmarcar) | POST | sem mudança de contrato: `desfazer_leituras` já devolve todas as urls |
| sobra (`/inventario/<id>/sala/<loc>/sobra`) | POST | chave `fotos.chave_sobra(inventario.pasta_do_evento(conn, id), sid)` |

A rota antiga `/foto/excluir` (sem `nfoto`) sai. `foto_leitura` sem leitura prévia responde 400 com a
mensagem de `adicionar_foto` (hoje o botão fica desabilitado até a leitura, e continua).

---

## 6. Telas

### 6.1 Sala (`inventario_sala.html`)

Célula Foto de cada bem (tabela "Bens da sala" e "Trazidos"):

- Uma miniatura por foto, na ordem de `nfoto`, cada uma como link `target="_blank"` (como hoje); com evento
  aberto, cada miniatura tem a lixeira (form POST para `.../foto/<nfoto>/excluir`, `volta` = sala).
- Depois das miniaturas, o botão da câmera (`label.br-button.circle.small` + `input.foto-input`) **sempre**
  presente quando `fotos_ativas and not fechado`, desabilitado enquanto o bem não foi lido. Hoje o botão some
  quando já há foto; passa a ficar.
- JS: `preencherFoto(td, fotos, numero)` recebe a lista e redesenha a célula inteira (miniaturas + lixeiras +
  câmera); a resposta do upload é `res.j.fotos`. `URL_FOTO_EXCLUIR` passa a ter dois marcadores
  (`/0/foto/0/excluir` → número e nfoto). O texto do `confirm` do Desmarcar diz "as fotos são apagadas".

### 6.2 Relatório (`inventario_relatorio.html`)

Só a primeira foto, como hoje; quando `x.n_fotos > 1`, um `<span class="br-tag small">+{{ x.n_fotos - 1 }}</span>`
ao lado da miniatura. O modal continua abrindo a primeira. xlsx não muda.

### 6.3 Cadastro do bem (`bem.html`, rota `bem` em `app.py`)

Card novo "Fotos do inventário" depois de "Histórico", ocupando a linha (`col-12`), só quando
`fotos` (de `inventario.fotos_do_bem`) não é vazio. Para cada evento: linha com o nome do evento, "aberto em
dd/mm/aaaa" (e "encerrado em" quando houver), "lido em dd/mm/aaaa hh:mm"; abaixo, as miniaturas
(`dsgov-miniatura`, link em nova aba, `alt="Foto N do bem X no evento Y"`) em `d-flex flex-wrap`. Sem fotos em
nenhum evento: card não aparece (a tela fica como hoje).

### 6.4 Menu (`app.py` `contexto_dsgov`, `base.html`)

`MENU` passa a ser lista de tuplas `(rotulo, icone, url, filhos)`; `filhos` é lista de `(rotulo, url)`, vazia
para todos menos "Inventário". Para "Inventário":

- sempre: `("Eventos", url_for("inventario.eventos_tela"))`;
- com evento aberto `e` (`inventario.evento_aberto`, já consultado na `home`): `(e.nome, url_for("inventario.evento_tela", id=e.id))`,
  `("Painel", url_for("inventario.painel_tela", id=e.id))`, `("Relatório", url_for("inventario.relatorio_tela", id=e.id))`.

`base.html`: item sem filhos renderiza como hoje; com filhos vira **grupo aberto** do DSGov:

```html
<div class="menu-folder">
  <div class="menu-item"><span class="icon"><i class="fas {{ icone }}"></i></span><span class="content">{{ rotulo }}</span></div>
  <ul role="group">{% for r, u in filhos %}<li><a class="menu-item" href="{{ u }}" role="treeitem"><span class="content">{{ r }}</span></a></li>{% endfor %}</ul>
</div>
```

O título do grupo não é `<a>`: o `core.min.js` do DSGov trata `.menu-folder:not(.drop-menu)` como grupo
sempre expandido (sem clique para abrir), o que evita um toque a mais no celular. A entrada "Inventário" deixa
de ser link; "Eventos" faz esse papel.

---

## 7. Regras preservadas

- Foto só depois da leitura (botão desabilitado + `adicionar_foto` recusa).
- Sobra exige foto quando as fotos estão ativas; uma só.
- Evento encerrado: nenhuma foto entra ou sai (`_evento_aberto_ou_erro`).
- Compressão WebP 1920×1080 q85 e validação de 5 MB não mudam.
- Sem bucket configurado, tudo funciona sem foto (programa Windows offline).

---

## 8. Testes

- `tests/test_fotos.py` (novo, sem rede): `pasta` (acento, maiúscula, espaço, pontuação, vazio → `evento<id>`),
  `chave_bem`/`chave_sobra`, `apagar` extrai a chave da URL nova e da antiga (monkeypatch em `_cliente`).
- `tests/test_inventario.py`: `adicionar_foto` numera 1, 2 e, após apagar a 2, 3; recusa sem leitura e em
  evento encerrado; a chave passada ao `enviar` é `inventario2026/1-<numero>.webp`; `apagar_foto` devolve a
  url e ignora inexistente; `desfazer_leituras` devolve todas as urls; `bens_da_sala`/`relatorio` trazem
  `foto_url` = primeira e `n_fotos`; `fotos_do_bem` agrupa por evento do mais recente para o mais antigo;
  `abrir_evento` recusa pasta repetida; migração em `criar_esquema` (banco com `foto_url` preenchida →
  `inventario_fotos` nfoto 1 e coluna removida; rodar duas vezes não duplica); aba `inv_fotos` exporta e
  importa; planilha antiga (com `foto_url`, sem `inv_fotos`) vira fotos nfoto 1; validações (nfoto inválido,
  leitura inexistente, repetido, pasta repetida em `inv_eventos`).
- `tests/test_app.py`: POST foto duas vezes devolve lista com 2 (`fotos.enviar` monkeypatched); excluir uma
  (`fotos.apagar` chamado com a url certa); sala mostra as duas miniaturas e a câmera; relatório mostra "+1";
  `/bem?numero=` com e sem card de fotos; menu com grupo "Inventário" com "Eventos" só (sem evento aberto)
  e com nome do evento, Painel e Relatório (com evento aberto).
- README: chave das fotos, aba `inv_fotos` (7 abas), grupo do menu.

## 9. Entrega

Branch `fase3-fotos`; commits atômicos por tarefa; revisão; merge na `main`; push;
`docker compose up -d --build` na VPS (a migração roda no primeiro start). Nenhum passo manual no bucket: as
fotos antigas ficam com as chaves antigas. Depois, o usuário valida no celular: tirar duas fotos de um bem,
apagar uma, ver o bem em `/bem`, abrir o menu.
