# Levantamento do módulo de inventário existente (sga/cfc) — 2026-09-15

Fonte: `/opt/web/sga/cfc` (branch `googlesheets`, HEAD `059af0d`), servido em
sistemadeinventario.com.br, container `sga-cfc`. **Nada foi alterado nesse projeto** — este
arquivo é só leitura/registro. Para retomar depois de um `/clear`, leia primeiro
`docs/superpowers/notes/2026-09-15-inventario-migracao-contexto.md` e depois este arquivo.

## 1. Visão geral

- **Stack**: Flask 3.1 (factory pattern, blueprints), Python. `app/__init__.py:13-108`.
  Sessão server-side em arquivo (`flask_session`, `SESSION_TYPE=filesystem`) — necessário para
  múltiplos workers Gunicorn não perderem dados de bipagem em andamento
  (`README.MD:70-90`, `config.py:32-36`).
- **Como sobe**: Docker. `docker-compose.yml` builda a imagem local, container `sga-cfc`,
  publica `127.0.0.1:11001 -> 3002` (interno `PORT=3002`). Volumes: `./secrets:/app/secrets:ro`
  (credenciais + `datasource.yaml`) e `./data:/app/data` (xlsx, backups, pendências).
  `docker-compose.yml:1-30`. Nginx faz proxy de `sistemadeinventario.com.br` (porta 443) para
  `127.0.0.1:11001` — `/etc/nginx/conf.d/sistemadeinventario.conf`. Confirmado rodando:
  `docker ps` mostra `sga-cfc  127.0.0.1:11001->3002/tcp  Up 6 days (healthy)`. (Esse mesmo
  domínio nginx também tem `crcrr.sistemadeinventario.com.br` e `cfc-db.sistemadeinventario.com.br`
  como vhosts separados — outros tenants, fora do escopo.)
- **De onde vêm os dados**: fonte de dados plugável — Google Sheets (`gspread`, modo padrão em
  produção) **ou** planilha Excel local `.xlsx`, escolhida por `DATASOURCE_TYPE` (env) ou por um
  arquivo de estado persistido (`data/datasource.state`, alterável pela tela de admin,
  `app/blueprints/admin.py:79-104`). Camada de abstração `DataSource` (protocolo em
  `app/services/datasources/base.py`) com duas implementações — `google_sheets.py` e
  `excel.py` — escolhidas por `app/services/datasources/factory.py:22-38,109-131`.
  O mapeamento de **qual coluna da planilha é qual campo lógico** não é hardcoded: vem de um
  descritor YAML (`config/datasource.yaml`, gitignored; exemplo em
  `config/datasource.example.yaml`), resolvido por texto do cabeçalho (case/acento-insensível),
  não por posição — `app/services/datasources/config.py:118-198`.
  - Nome da planilha (Sheets): `INVENTARIO_2025`; aba de dados `INVENTARIO`; aba de usuários
    `USUARIOS`; cabeçalho na linha 1; dados a partir da linha 2
    (`config/datasource.example.yaml:12-20`).
  - **Colunas da aba INVENTARIO** (texto do cabeçalho → campo lógico), de
    `config/datasource.example.yaml:22-40`:
    `N. BEM→numero_bem`, `SITUACAO→situacao`, `DESCRICAO→descricao`,
    `COMPLEMENTO→complemento`, `CLASSIFICACAO CONTABIL→classificacao_contabil`,
    `LOCAL→local_sistema`, `LOCAL VERIFICADO→local_verificado`, `DATA ENTRADA→data_entrada`,
    `VALOR COMPRA→valor_compra`, `VALOR ATUAL→valor_atual`, `ESTADO→estado_conservacao`,
    `USUARIO INVENTARIO→usuario_inventario`, `DATA→data_inventario`, `USUARIO BEM→usuario_bem`,
    `FOTO→foto`, `OBSERVACAO→observacao` (campo opcional — planilhas antigas sem essa coluna
    continuam funcionando, só não ganham "Inserir sobra", ver seção 3).
  - **Colunas da aba USUARIOS**: `Nome→nome`, `Email→email`, `Senha→senha`, `Role→role`
    (opcional, ausente/vazio = `comum`), `Status→status` (opcional, ausente/vazio = `ativo`).
  - `data/inventario.xlsx` existe no volume local (932KB, dono root, não lido diretamente por
    mim — só os nomes de coluna acima, vindos do descritor YAML, foram inspecionados; **não
    abri o conteúdo/linhas da planilha**, por instrução). Há 1 backup em `data/backups/`.
  - `data/pending/*.json`: um arquivo por usuário (nome = e-mail sanitizado) com a lista de bens
    "em conferência" ainda não sincronizados com a fonte — staging durável que sobrevive a
    reinício do servidor e troca de dispositivo (`app/services/pending_store.py:1-8,41-66`).
    Inspecionei só a estrutura de um registro (ver seção 2, "bem em sessão").
- **Autenticação**: própria (não é SSO), Flask-Login. E-mail + senha contra a aba
  USUARIOS/planilha. Verificação **híbrida**: aceita hash Werkzeug (`scrypt:`/`pbkdf2:`) ou,
  para compatibilidade com linhas legadas, texto puro comparado direto
  (`app/models/user.py:47-71`). Senha temporária padrão do sistema é `12345678`; login com
  senha em texto puro OU com o hash de `12345678` força troca obrigatória antes de qualquer
  outra rota (`app/services/user_source.py:35-49`, `app/blueprints/auth.py:54-58`,
  bloqueio global em `app/__init__.py:112-133`). Papéis: `admin` / `comum`
  (`app/models/user.py:19-22`); e-mail listado em `ADMIN_EMAIL` sempre vira admin
  independentemente da coluna (`app/services/user_source.py`, comentário no topo). Rate limit
  de login: 5 tentativas/15min por IP+e-mail (`app/blueprints/auth.py:10-17`).
- **Quem usa**: equipe de patrimônio do CFC fazendo a conferência física de bens sala a sala
  (ver nomes de arquivo em `data/pending/*.cfc.org.br.json` — não abri conteúdo). Multi-tenant:
  o mesmo código-base (`sga`) serve outros clientes em containers separados (`sga-crcrr`,
  `sgadb-web`, `sgadb-ng-web`) — cada um com seu próprio volume/planilha; **não é multi-tenant
  dentro do mesmo processo**.

## 2. Modelo de dados do inventário

Não há um "modelo" com tabelas separadas — é **uma linha por bem em uma única aba/planilha**,
mais uma aba de usuários. Os "eventos" e "leituras" não são persistidos como entidades — o que
existe é o próprio bem sendo atualizado in-place, mais um staging efêmero em sessão/JSON local.

### 2.1 Bem (linha da planilha) — via `row_to_asset` / `row_to_inventory_record`
`app/services/inventory_rows.py:169-260`. Campos lógicos e labels amigáveis
(`app/services/inventory_rows.py:17-46`):

| campo lógico | label PT | onde é setado |
|---|---|---|
| `numero_bem` | Patrimônio | cadastro (import), chave de busca |
| `situacao` | Situação | cadastro (ex.: ATIVO) — só itens ATIVO entram no fluxo de local (`inventory_rows.py:180`) |
| `descricao` | Descrição | cadastro |
| `complemento` | Complemento | cadastro; editável em sobras |
| `classificacao_contabil` | (sem label específico) | cadastro, só aparece em relatório/registro dashboard |
| `local_sistema` | Local do sistema | cadastro (onde o bem "deveria" estar) |
| `local_verificado` | Local inventariado | **escrito pelo inventário** quando o bem é lido/confirmado |
| `estado_conservacao` | Estado de Conservação | editável no inventário (Bom/Regular/Ruim/Inservível) |
| `usuario_inventario` | Usuário inventário | quem bipou — **carimbado automaticamente** com o nome do usuário logado ao salvar |
| `data_inventario` | Data do inventário | **carimbado automaticamente** (`datetime.now()`) ao salvar |
| `usuario_bem` | Quem usa | quem usa o bem no dia a dia — editável |
| `foto` | Foto | URL pública no Cloudflare R2 |
| `observacao` | Observação | campo opcional; usado para a justificativa da sobra |

Campos calculados em memória, nunca persistidos: `status` (`Localizado`/`Não localizado`/
`Nao aplicavel`/`Excluido`/`Sobra`), `status_inventario` no dashboard, `Local Divergente`
(calculado comparando `local_sistema` × `local_verificado`, nunca gravado como valor —
`app/services/reports.py:112-132`).

### 2.2 "Sala" = `local_sistema` / `local_verificado`
Não existe uma tabela de salas. `get_locations()` varre a coluna `local_sistema` **e**
`local_verificado` de todas as linhas e devolve a união distinta como lista de strings
(`app/services/datasources/excel.py:177-211`). Ou seja, o conjunto de "salas" é simplesmente o
conjunto de valores livres já usados na planilha — sem cadastro, sem hierarquia, sem
metadados (não há capacidade, responsável, centro de custo associado à sala dentro deste
módulo — isso existe no legado `terms.py`, ver seção 8).

### 2.3 "Bem em sessão" (staging da conferência em andamento)
Estrutura por item, construída em `row_to_asset` e persistida em `data/pending/<email>.json`
(lista de dicts) + espelhada na sessão Flask (efêmera, 1h):
```
row_idx, numero_bem, situacao, descricao, complemento, local_sistema, local_verificado,
status ('Localizado'|'Não localizado'|'Excluido'|'Sobra'), conservacao, foto, usuario_bem,
usuario_inventario, data_inventario, observacao, alterado (bool — pendente de sync),
tipo ('sobra', só quando aplicável), campos_fonte (lista de todos os campos brutos da
planilha, usada pelo modal de ficha completa — `montar_campos_fonte`, inventory_rows.py:51-98)
```
Exemplo real de 1 registro (`data/pending/teste@teste.com.json`, li o conteúdo pois é um
registro de teste, sem dado pessoal sensível):
```json
{"row_idx": 2725, "numero_bem": "10622", "situacao": "ATIVO",
 "complemento": "PURIFICADOR, FR 600, IBBL, 220 V SERIE 739P090771",
 "local_sistema": "03 - CORREDOR", "local_verificado": "03 - CORREDOR",
 "status": "Localizado", "conservacao": "Bom", "foto": "", "usuario_bem": "", "alterado": true}
```
Os demais arquivos `data/pending/*.json` são listas de 1 a 357 itens — só contei o tamanho, não
abri o conteúdo (nomes de arquivo = e-mails de usuários reais do CFC).

### 2.4 Estado de UI por usuário (não durável)
`session['estado_usuario'] = {'local_atual': ..., 'last_user_input': ...}` — sala selecionada e
último nome digitado, guardado só na sessão Flask (expira em 1h) —
`app/blueprints/inventory.py:30-37`.

### 2.5 Usuário
`nome, email (id), senha (hash ou legado texto puro), role (admin|comum), status (ativo|inativo)`
— `app/models/user.py:5-13`, `app/services/user_source.py`.

### 2.6 Persistência
Tudo isso vive na própria planilha (Sheets ou `.xlsx`) + 2 diretórios auxiliares:
`data/pending/` (staging por usuário) e `data/backups/` (snapshot antes de cada escrita no
modo Excel, retenção configurável, `excel.py:275-295`). **Não há banco relacional, não há
tabela de eventos/leituras/divergências separada** — cada bem carrega seu próprio "último
estado de inventário" nas mesmas colunas que os dados cadastrais.

## 3. Fluxo do inventário passo a passo

Rotas em `app/blueprints/inventory.py`, blueprint registrado como `system` (prefixo vazio).

1. **Tela `/inventario`** (`inventario()`, `inventory.py:168-191`) — GET. Mostra select de
   "Local atual" (= salas, vindas de `get_locations()`, cacheado 10min), campo de bipagem, e a
   tabela dos bens já carregados na sessão do usuário (staging de `data/pending/`).
2. **Escolher a sala**: `POST /carregar_bens` (`inventory.py:193-260`). Valida que o local
   existe na lista de locais conhecidos (`valida_local`, `app/services/validation.py:46-52`).
   Lê **toda** a planilha (`get_all_inventory_data()`, cacheado 5min), converte todas as linhas
   e filtra em memória as que têm `local_sistema == local` OU `local_verificado == local`
   (`inventory.py:227`) — ou seja, a sala carrega tanto os bens que "deveriam" estar lá quanto
   os que já foram verificados como estando lá (útil se um bem já foi movido antes). O
   resultado vira o staging do usuário (`set_user_bens`) e o local vira `estado_usuario.local_atual`.
3. **Leitura por código de barras**: `POST /marcar_bem` (`inventory.py:262-374`). O campo texto
   `#numero_bem` é um input simples (`autocomplete="off"`, autofocus quando a sala está
   carregada) que funciona tanto para **leitor USB (wedge/teclado)** — o leitor "digita" o
   código + Enter, que dispara o submit do form — quanto para digitação manual
   (`app/templates/inventario.html:233-238`). Além disso há um **botão de câmera** que abre um
   modal com a biblioteca JS `html5-qrcode` (CDN `unpkg.com/html5-qrcode`,
   `inventario.html:7,641-728`), lendo QR/código de barras pela câmera do celular
   (`Html5QrcodeScanner`, `fps:10`, `qrbox 250x250`) e preenchendo o mesmo campo. **Não há
   prefixo/dígito verificador nem formato fixo de código** — a única validação é
   `valida_numero_bem`: 1–50 caracteres, regex `^[A-Z0-9.\-]+$` (letras, números, ponto, hífen)
   — `app/services/validation.py:8,14-22`. O número lido é normalizado tirando zeros à
   esquerda (`normalizar_numero_bem`, `inventory.py:39-41`).
   - **Bem já no staging da sala** (já bipado nesta sessão): marca `status='Localizado'`,
     `alterado=True` (`inventory.py:300-306`).
   - **Bem não está no staging**: busca na planilha inteira por `numero_bem`
     (`service.find_asset`, varre a coluna até achar — `excel.py:151-175`). Se achar em
     **qualquer** local (não só na sala escolhida): converte a linha, marca `Localizado`, e — se
     `local_sistema` do bem for diferente da sala atual — devolve a mensagem
     `"(Trazido de {local_sistema})"` (`inventory.py:334-337`); no front, essa linha fica com
     chip **"Divergente"** em vez de "Localizado" quando `local_sistema != local_sel`
     (`inventario.html:312-317,348-351`). **Não existe bloqueio**: qualquer bem cadastrado pode
     ser lido em qualquer sala, o sistema só sinaliza visualmente a divergência — não impede
     nem pede confirmação adicional além do "já localizado, sobrescrever?" (ver abaixo).
   - **Bem não encontrado em lugar nenhum**: erro "Bem não encontrado" (`inventory.py:338-349`)
     — não cria registro automaticamente (isso só acontece via "sobra", fluxo manual).
   - Se o bem já estava **Localizado** (confirmado) e é lido de novo, o JS mostra um **modal de
     confirmação de sobrescrita** com local/usuário/data anteriores antes de re-submeter
     (`inventario.html:664-728`, função `checarSobrescrita`).
   - **Auto-flush**: a cada 20 itens alterados pendentes de envio, o sistema salva
     automaticamente o lote na fonte de dados sem o usuário precisar clicar em nada
     (`AUTO_FLUSH_LIMIAR=20`, `inventory.py:19,356-359`).
4. **Bens da sala não lidos** ("faltantes"): não são marcados/gravados como nada especial — a
   linha continua na tabela com status calculado `'Não localizado'` (chip vermelho "Pendente")
   simplesmente porque nunca tiveram `alterado=True`/nunca passaram por `marcar_bem`
   (`inventory.py:296-306` e template `inventario.html:356-357`). Não há persistência de "não
   encontrado" na planilha — é um estado só visual até o operador agir.
5. **Bens sem cadastro (sobras)**: `POST /cadastrar_sobra` (`inventory.py:672-778`). Cria uma
   linha **nova** na planilha (sem `numero_bem`) via `append_asset`, com foto **obrigatória**
   (server-side), descrição obrigatória e observação obrigatória
   (`inventory.py:688-703`). Fluxo com rollback: 1) cria a linha, 2) faz upload síncrono da
   foto, 3) se o upload falhar, **apaga a linha criada** (`delete_row`) e cancela
   (`inventory.py:717-746`, comentado como "D-10/D-11"). Sobra fica com `status='Sobra'` (chip
   azul "Sobra") e só pode ser identificada por `row_idx` (nunca por `numero_bem`, que fica
   vazio) — edição/exclusão física só é permitida se `numero_bem==''` e `tipo=='sobra'`,
   checado no servidor (`inventory.py:781-841`, "política somente-sobra").
6. **Encerramento do inventário**: **não existe um botão/ação de "encerrar" ou "fechar
   evento"**. O que existe é `POST /enviar_inventario` (`inventory.py:420-475`), que só
   sincroniza o staging pendente do usuário com a planilha (grava `local_verificado`,
   `estado_conservacao`, `usuario_inventario`, `data_inventario`, `usuario_bem` de cada linha
   alterada via `save_batch`) e zera as flags `alterado`. Isso pode ser chamado quantas vezes o
   usuário quiser, para qualquer sala, a qualquer momento — **não há conceito de evento/campanha
   com início e fim, nem trava depois de "fechado"**.
7. **"Evento de inventário"**: **não existe essa entidade no código.** Não há tabela, id,
   data de abertura/fechamento, nem agrupamento das leituras por campanha. Cada bem carrega
   apenas seu **último** estado (`data_inventario`, `usuario_inventario`) — uma nova conferência
   sobrescreve a anterior, sem histórico. Isso é uma lacuna real se o sistema novo precisa de
   "evento" como entidade (ver seção 9, dúvidas).
8. Outras ações de staging: `POST /alternar_status` (liga/desliga manualmente Localizado↔Não
   localizado, sem tocar a planilha até salvar — `inventory.py:872-898`), `POST /marcar_lote` /
   `POST /excluir_lote` (ação em massa sobre selecionados via checkbox —
   `inventory.py:900-948`), `POST /excluir_item` (marca para exclusão do staging, não da
   planilha — `inventory.py:844-870`), `POST /atualizar_info_bem` (endpoint JSON, salva
   imediatamente conservação/"quem usa"/observação — auto-save ao sair do campo, sem esperar o
   "Salvar alterações" geral — `inventory.py:478-593`).

## 4. Fotos

- **Por bem** (bens regulares e sobras) — não há foto por leitura nem por sala; uma foto por
  bem, substituível.
- **Captura**: `<input type="file" accept="image/*" capture="environment">`
  (`inventario.html:413,514`) — abre a câmera traseira no celular; em desktop funciona como
  seletor de arquivo comum. Botão de câmera abre esse input escondido via JS
  (`ativarCamera`/`tirarFotoSobraAtual`).
- **Validação server-side**: `valida_arquivo_upload` — extensão em
  `{.jpg,.jpeg,.png,.webp}`, tamanho ≤ `MAX_CONTENT_LENGTH` (padrão 5MB,
  `config.py:41`), conteúdo realmente é imagem válida via Pillow (`Image.verify()`), formato
  final em `{JPEG,PNG,WEBP}` (`app/services/validation.py:64-104`). Rate limit: 20
  uploads/hora por usuário (`inventory.py:598`).
- **Compressão/redimensionamento**: toda foto enviada é convertida para **WebP** qualidade 85,
  redimensionada (mantendo proporção) para no máximo **1920×1080**
  (`app/services/images.py:14-24,71-93`), antes do upload — "~88% de redução" segundo o README.
- **Onde ficam gravadas**: Cloudflare R2 (S3-compatível via `boto3`), não em disco local
  (`app/services/storage.py:1-64`). URL pública devolvida e gravada na coluna `foto` da
  planilha (`update_photo_url`/`update_foto_row`).
- **Nomes de arquivo**: `BEM_{numero_bem}_{YYYYMMDDHHMMSS}.jpg` para bens regulares
  (`inventory.py:649`), `SOBRA_{row_idx}_{YYYYMMDDHHMMSS}.jpg` para sobras
  (`inventory.py:623,738`) — a extensão final no bucket vira `.webp` (troca automática em
  `storage.py:44-45`).
- **Exibição**: thumbnail 56×56 (28×28 no mobile) na tabela, clicável, abre modal com a
  imagem grande e opção de excluir (`abrirFotoModal`, `inventario.html:333-340`, CSS
  `:61-66,150-155`). Excluir foto (`POST /excluir_foto`) apaga o objeto do bucket R2 **e** limpa
  a URL na planilha (`inventory.py:980-1009`).
- Upload de bem regular roda a atualização da planilha **em thread separada** (não bloqueia a
  resposta — `inventory.py:654`); upload de sobra é **síncrono** porque precisa do rollback se
  falhar (`inventory.py:736-749`).

## 5. Exportações e relatórios

- **Tela `/relatorios`** (`dashboard.py:129-150`) — grid com todos os bens ativos, filtros por
  usuário/local/status/conservação/foto, busca textual tolerante a acento
  (`normalizar_busca`, `reports.py:84-96`), ordenação por coluna (numérica para `numero_bem`).
  Tudo calculado/filtrado no **backend** a partir do cache de 5min (`_load_inventory_rows`,
  `dashboard.py:31-46`) — o filtro por coluna vem como JSON (`column_filters`) do front.
- **Coluna calculada "Local Divergente"**: no relatório (não na leitura), quando
  `status=='Localizado'` e `local_sistema != local_verificado` (ambos não-vazios) —
  `reports.py:112-132`.
- **Exportação** `POST /relatorios/exportar` (`dashboard.py:232-311`) — aplica os mesmos
  filtros/ordenação da tela (reconstruídos no backend, nunca confia no DOM) e gera:
  - **Excel (.xlsx)**, via `openpyxl` — `app/services/report_exports.py`. Cabeçalho com título,
    data/hora, usuário emissor, resumo textual dos filtros ativos, total de registros
    (linhas 1–4); estilo com cor institucional; larguras de coluna por campo; opção "incluir
    fotos" grava fórmula `=IMAGEM("URL")` na célula (sem embutir binário) e aumenta a altura da
    linha; coluna Foto mostra `-` quando não há URL válida.
  - **PDF** (ReportLab, A4 paisagem) — mesma lógica de colunas/filtros/resumo, com thumbnails
    quando "incluir fotos" está marcado; falha de 1 foto não derruba o PDF inteiro (mostra "Sem
    foto").
  - **Colunas exportáveis** (ordem canônica, `reports.py:26-45`): Foto, Patrimônio, Situação,
    Descrição, Complemento, Classificação Contábil, Local Sistema, Local Inventariado, Data
    Entrada, Valor Compra, Valor Atual, Conservação, Integrante (=usuario_inventario),
    Data/Hora Inventário, Usuário do Bem, Observação (opcional, não pré-selecionada).
  - Aviso de "relatório pesado" quando ≥ 50 fotos válidas no conjunto filtrado (`contar_fotos`,
    `reports.py:334-340`).
- **Dashboard `/dashboard`** (`dashboard.py:48-127`) — KPIs (total ativos, localizados,
  pendentes, divergentes, % conclusão), contagem por usuário/conservação, progresso por andar
  (extraído heuristicamente do texto de `local_sistema` antes do primeiro `-`/`/`, ver
  `_extract_andar`, `dashboard.py:23-29`) e por sala dentro do andar.
- Não há exportação CSV nem exportação específica "por sala" ou "por evento" — só o relatório
  geral filtrável.

## 6. Telas e rotas

| Rota | Método | Função | Template | O que mostra |
|---|---|---|---|---|
| `/` | GET | `home` | `home.html` | Página inicial pós-login |
| `/login` | GET/POST | `login` | `login.html` | Autenticação |
| `/logout` | GET | `logout` | — | Encerra sessão |
| `/alterar_senha` | GET/POST | `alterar_senha` | `alterar_senha.html` | Troca voluntária |
| `/alterar_senha_obrigatoria` | GET/POST | `alterar_senha_obrigatoria` | `alterar_senha.html` | Troca forçada (senha legada/temp) |
| `/inventario` | GET | `inventario` | `inventario.html` | Tela principal de bipagem |
| `/carregar_bens` | POST | `carregar_bens` | `inventario.html` | Carrega bens da sala escolhida |
| `/marcar_bem` | POST | `marcar_bem` | `inventario.html` | Registra leitura de 1 bem |
| `/salvar_inventario` | POST | `salvar_inventario` | `inventario.html` | Aplica edições do form ao staging (não persiste na fonte) |
| `/enviar_inventario` | POST | `enviar_inventario` | `inventario.html` | Sincroniza staging pendente com a planilha |
| `/atualizar_info_bem` | POST (JSON) | `atualizar_info_bem` | — | Auto-save de conservação/quem usa/observação/sobra |
| `/upload_foto` | POST | `upload_foto` | redirect | Upload de foto de bem/sobra |
| `/cadastrar_sobra` | POST | `cadastrar_sobra` | redirect | Cria bem "sobra" sem número |
| `/excluir_sobra` | POST | `excluir_sobra` | redirect | Apaga fisicamente uma sobra |
| `/excluir_item` | POST | `excluir_item` | `inventario.html` | Marca item para exclusão do staging |
| `/alternar_status` | POST | `alternar_status` | `inventario.html` | Alterna Localizado/Não localizado manualmente |
| `/marcar_lote` / `/excluir_lote` | POST | idem | `inventario.html` | Ação em massa sobre selecionados |
| `/excluir_foto` | POST | `excluir_foto` | redirect | Remove foto (R2 + planilha) |
| `/movimentacao` | GET/POST | `movimentacao` | `movimentacao.html` | Busca um bem por número; "confirmar" **não implementado** ("Funcionalidade em manutenção", `inventory.py:970`) |
| `/termos_ccustos` | GET/POST | `termos_ccustos` | `termos_ccustos.html` | **Placeholder** — "Funcionalidade em migração" (`inventory.py:974-978`), sem lógica |
| `/dashboard` | GET | `index` | `dashboard.html` | KPIs e progresso |
| `/relatorios` | GET | `reports` | `relatorios.html` | Grid filtrável |
| `/relatorios/exportar` | POST | `export_reports` | — (download) | Gera xlsx/pdf |
| `/admin` | GET | `index` | `admin.html` | CRUD de usuários + toggle de fonte de dados |
| `/admin/datasource/toggle` | POST | `alterar_datasource` | — | Troca sheets↔excel (aplica após restart) |
| `/admin/usuarios/*` | POST | criar/editar/desativar/reativar/remover/resetar_senha | — | CRUD de usuários |
| `/health` | GET | `health` | — | Health check |

**JS relevante**:
- `app/templates/inventario.html` (script embutido, ~960 linhas de JS): scanner de câmera
  (`html5-qrcode`, CDN), confirmação de sobrescrita, modal de "ficha completa" do bem
  (renderiza `campos_fonte` dinamicamente), modal de cadastro de sobra, auto-save on
  blur/change (`_autoSalvarCampo`, chama `/atualizar_info_bem` via fetch+JSON), filtro/ordenação
  client-side da tabela carregada, toggle de painel colapsável (persistido em `localStorage`).
- `app/static/js/theme.js` — dark/light mode.
- `app/static/js/pwa.js`, `app/static/js/service-worker.js` — PWA instalável (ver
  `docs/pwa.md`), cache offline básico.
- Sem framework JS (React/Vue) — tudo vanilla JS + Jinja2 server-rendered.

## 7. Regras de negócio e validações (arquivo:linha)

- `valida_numero_bem`: obrigatório, ≤50 chars, regex `^[A-Z0-9.\-]+$` —
  `app/services/validation.py:14-22`.
- `valida_local`: obrigatório e deve estar na lista de locais conhecidos —
  `app/services/validation.py:46-52`.
- `valida_conservacao`: opcional, mas se preenchido deve estar em
  `['Bom','Regular','Ruim','Inservivel']` — `app/services/validation.py:55-61`,
  lista canônica em `app/blueprints/inventory.py:95`.
- `valida_arquivo_upload`: extensão + tamanho + Pillow.verify() + formato —
  `app/services/validation.py:64-104`.
- Normalização de número de bem (remove zeros à esquerda) para comparação —
  `inventory.py:39-41`, `excel.py:132-133`.
- Só itens com `situacao=='ATIVO'` entram no cálculo de status do dashboard/relatório —
  `inventory_rows.py:180-185`.
- Linha sem `numero_bem` e sem `local_verificado` é descartada na conversão para
  dashboard/relatório (registro legado incompleto) — `inventory_rows.py:330-333`.
- "Local Divergente" é sempre **calculado**, nunca um valor gravado — `reports.py:112-132`.
- Auto-flush a cada 20 itens alterados — `inventory.py:19,356-359`.
- Sobra exige foto + descrição + observação (server-side, autoritativo mesmo que o front também
  valide) — `inventory.py:688-703`.
- Política "somente sobra pode ser excluída fisicamente" (nunca um bem cadastrado) — checada no
  servidor antes de `delete_row`, nunca confia no cliente — `inventory.py:803-817`, comentário
  explícito citando teste `T-04.7-03-02`.
- Rollback de sobra: se o upload da foto falhar depois de criada a linha, a linha é apagada —
  `inventory.py:717-746`.
- `update_row_fields` só grava colunas de **cadastro** (não mexe em `local_verificado`/
  `usuario_inventario`/`data_inventario`/`status`) quando o alvo é uma sobra sendo editada —
  `inventory.py:554-570`.
- Rate limits: login 5/15min (`auth.py:17`), upload de foto 20/h (`inventory.py:598`),
  cadastro de sobra 20/h (`inventory.py:674`).
- Senha: verificação híbrida hash/texto puro; troca obrigatória se legado ou senha temporária
  `12345678` — `app/models/user.py:47-71`, `app/services/user_source.py:35-49`.
- `role` nunca vira `admin` por ausência de coluna/valor — só `ADMIN_EMAIL` explícito ou valor
  literal `admin` na planilha — comentário no topo de `user_source.py`.
- `MAX_CONTENT_LENGTH` padrão 5MB por upload — `config.py:41`.
- CSP, CSRF (Flask-WTF), Talisman, rate limiting via Flask-Limiter — `app/__init__.py:44-55`.

## 8. O que NÃO precisa migrar × o que é essencial

**Não precisa migrar** (específico deste sistema/infra, não do processo de negócio):
- Autenticação própria com Flask-Login, hash híbrido, troca obrigatória de senha — o sistema
  destino provavelmente já tem seu próprio esquema de acesso.
- Multi-datasource plugável Sheets↔Excel e a tela `/admin` de toggle — o destino já decidiu
  SQLite fixo.
- Sessão em filesystem / `flask_session`, cache em filesystem (`flask_cache`) — decisões de
  infra do sga, não do domínio.
- Armazenamento de fotos em Cloudflare R2 via boto3 — a decidir no destino (pode ser R2 também,
  ou disco local; ver dúvidas).
- PWA (manifest, service worker, instalável) — feature própria da UI do sga, não do inventário.
- CRUD de usuários (`/admin/usuarios/*`) — duplica o que o sistema destino já tem
  (`pessoas`/auth do termos-responsabilidade).
- `/movimentacao` — não implementado de fato ("Funcionalidade em manutenção"), não vale
  migrar como está.
- `/termos_ccustos` e `app/services/terms.py` — **este último é código morto**: importa um
  módulo `conexao` que não existe em `app/` (não é importado por nenhum blueprint, `grep`
  confirmou), parece um resquício de versão anterior de single-arquivo (provavelmente do
  tenant `crcrr`, dado o texto fixo "CFC" no termo). A rota `/termos_ccustos` atual é só um
  placeholder sem lógica. **Não copiar `terms.py` como base** — é órfão e desatualizado; o
  sistema destino já tem sua própria geração de termos (`Termo_de_Responsabilidade.py`,
  `termos_html.py`) que deve continuar sendo a fonte de verdade para termos.
- Outros tenants no mesmo host (`sga-crcrr`, `sgadb-web`, `sgadb-ng-web` etc.) — fora do
  escopo, cada um é um deploy separado.
- Multi-fonte de dados (Google Sheets) — o destino já decidiu SQLite.

**É essencial** (a lógica de domínio que o usuário quer preservar):
- Fluxo sala → carregar bens da sala → ler código de barras → confirmar/divergente/faltante
  (seção 3) — é o núcleo do pedido.
- Regra de divergência: bem lido pertence a outra sala (`local_sistema != sala atual`) é aceito
  mas sinalizado, não bloqueado.
- Conceito de "sobra" (bem sem cadastro, com foto+descrição+observação obrigatórias e política
  de exclusão restrita) — mapeável para o destino como "bem não localizado no catálogo,
  registrado avulso".
- Regra de fotos: captura via `capture="environment"`, compressão/redimensionamento antes de
  guardar, foto obrigatória para sobra.
- Exportação com filtros/colunas selecionáveis e resumo de filtros no cabeçalho.
- Auto-save silencioso ao editar conservação/observação/quem-usa (sem esperar "Salvar tudo").
- Confirmação ao reler bem já confirmado (evita sobrescrita acidental sem aviso).
- A **ausência** de um "evento de inventário" formal no sistema atual é uma lacuna a preencher
  no design novo, não algo a copiar (ver seção 9).

## 9. Riscos/dúvidas para perguntar ao usuário

1. **"Evento de inventário" não existe no sistema atual** (seção 3.7) — cada bem guarda só o
   último estado, sobrescrito a cada nova conferência, sem histórico. O sistema novo precisa
   decidir: um evento por sala (abre ao carregar, fecha ao enviar) ou uma campanha guarda-chuva
   com várias salas e data de início/fim? Isso muda o esquema de tabelas novo.
2. **Bem "divergente" nunca é bloqueado nem move a localização automaticamente** — só sinaliza.
   No sistema novo, ao confirmar um bem "trazido de outra sala", o campo `bens.localizacao`
   deve ser atualizado de fato (como o `local_verificado` faz aqui) ou fica só registrado como
   achado sem alterar o cadastro mestre?
3. **"Sala" não tem cadastro próprio** neste sistema — é só a união dos valores livres em
   `local_sistema`/`local_verificado`. No destino, `bens.localizacao` já é texto livre também,
   mas há `localizacoes` cadastradas (100 ativas) — o inventário deve trabalhar só com as
   localizações já cadastradas em `localizacoes`, ou aceitar leitura de qualquer valor livre
   como o sga faz?
4. **Fotos**: o sga guarda em Cloudflare R2 (bucket externo). O sistema destino roda em
   container único com SQLite local — fotos vão para disco (bind mount, como `dados/`) ou para
   algum bucket também? Isso decide se replicamos a etapa de compressão/redimensionamento
   (recomendável de qualquer forma, pelo tamanho) e o nome de arquivo.
5. **"Faltante" não é persistido** no sga — é status calculado (bem da sala, sem leitura nesta
   sessão). Isso é aceitável no destino, ou o usuário quer um relatório histórico de "o que
   faltou no evento X" que sobreviva ao encerramento (o que exige persistir o vínculo
   bem↔evento↔status, não só o último estado no bem)?
6. Bens sem cadastro (sobras) no sga viram **linha nova na planilha mestre**. No destino, a
   tabela `bens` tem `numero INTEGER PRIMARY KEY` (não aceita string vazia) — será preciso
   decidir se sobras entram como linha em `bens` com um número sintético, ou como tabela
   separada (ex.: `inventario_sobras`) fora de `bens`.
7. Código de barras: o sga não valida prefixo nem dígito verificador, só formato alfanumérico
   genérico — confirmar se os leitores/códigos do CFC têm algum padrão que valha reforçar no
   sistema novo (ex.: sempre numérico, sempre = `bens.numero`).
8. `app/services/terms.py` (legado, código morto) sugere que já existiu uma tentativa de gerar
   "Termo de Responsabilidade por Centro de Custo" dentro do próprio sga — mas está órfã e
   provavelmente pertence a outro tenant (texto fixo menciona "CFC" mas importa de `conexao`,
   módulo inexistente aqui). Não deve ser usada como referência; o destino já tem sua própria
   geração de termos — confirmar que o inventário novo não deve tentar gerar termos, só
   alimentar dados para o fluxo de termos já existente.

## 10. Mapa de arquivos (sga/cfc) relevantes ao inventário

- `README.MD` — visão geral, stack, setup.
- `AGENTS.md` — convenções do projeto (GSD), branch `googlesheets` é a prioridade.
- `config.py` — toda a configuração (sessão, cache, R2, datasource, admin).
- `run.py` — entrypoint (`flask run` equivalente).
- `docker-compose.yml`, `Dockerfile` — deploy do container `sga-cfc`, porta 3002 interna.
- `config/datasource.example.yaml` — mapeamento de colunas (o real, `config/datasource.yaml`, é
  gitignored/nao existe nesse checkout; a aplicação lê de `secrets/datasource.yaml` em
  produção conforme `docker-compose.yml`).
- `app/__init__.py` — application factory, registro de blueprints, guarda de troca de senha.
- `app/blueprints/inventory.py` — **coração do módulo**: todas as rotas de bipagem/sobra/foto.
- `app/blueprints/dashboard.py` — dashboard e relatórios/exportação.
- `app/blueprints/admin.py` — CRUD de usuários + toggle de fonte de dados (fora de escopo).
- `app/blueprints/auth.py` — login/logout/troca de senha (fora de escopo).
- `app/blueprints/pwa.py` — manifest/service worker (fora de escopo).
- `app/services/datasources/base.py` — contrato `DataSource` (Protocol).
- `app/services/datasources/config.py` — parsing/validação do descritor YAML, resolução de
  colunas por cabeçalho.
- `app/services/datasources/excel.py` — implementação Excel local (a mais próxima do que o
  destino vai fazer com SQLite — mesma ideia de backup atômico antes de escrever).
- `app/services/datasources/google_sheets.py`, `app/services/sheets.py` — implementação Sheets
  (fora de escopo, não lidos em detalhe).
- `app/services/datasources/factory.py` — escolhe a implementação por config/estado.
- `app/services/datasource_state.py` — persiste a escolha sheets/excel em arquivo.
- `app/services/inventory_rows.py` — conversão linha bruta ↔ dict lógico; **referência central
  para nomes de campo e labels em português**.
- `app/services/validation.py` — todas as validações server-side.
- `app/services/pending_store.py` — staging JSON por usuário (durável).
- `app/services/storage.py` — upload/delete no Cloudflare R2.
- `app/services/images.py` — compressão/redimensionamento WebP.
- `app/services/reports.py` — schema de colunas, filtros, ordenação, "Local Divergente" (puro,
  sem Flask — bom modelo a seguir no destino).
- `app/services/report_exports.py` — geração de `.xlsx`/PDF.
- `app/services/cached_datasource.py` — cache com invalidação em cima de qualquer `DataSource`.
- `app/services/user_source.py`, `app/models/user.py` — auth (fora de escopo).
- `app/services/datahora.py` — timezone Brasil para timestamps.
- `app/services/terms.py` — **código morto/órfão**, não usar como referência (seção 8).
- `app/templates/inventario.html` — tela principal (1600 linhas: HTML + CSS + JS embutidos).
- `app/templates/relatorios.html`, `dashboard.html` — telas de relatório/dashboard.
- `app/templates/movimentacao.html`, `termos_ccustos.html` — placeholders não funcionais.
- `app/templates/admin.html`, `login.html`, `alterar_senha.html`, `home.html`, `holding.html`,
  `offline.html` — fora do escopo do inventário em si.
- `app/templates/components/*.html` — layout compartilhado (sidebar, topbar, footer, scripts).
- `app/static/js/pwa.js`, `service-worker.js`, `theme.js` — fora de escopo.
- `app/static/css/*.css` — estilo próprio do sga (não é DSGov — o destino já usa DSGov visual).
- `data/inventario.xlsx`, `data/backups/*.xlsx` — dados reais (não lidos além dos cabeçalhos
  via YAML), dono root, só leitura de nomes/tamanhos aqui.
- `data/pending/*.json` — staging real de usuários do CFC (não lido, exceto 1 arquivo de teste
  com 1 item sem dado sensível).
- `tests/` — ~30 arquivos de teste (nomes lidos, conteúdo não inspecionado em detalhe); cobrem
  datasource (excel/sheets/factory/config), rotas de admin/dashboard/relatórios, segurança,
  upload, PWA, mobile responsivo, sobras (fase 04.7) — bom indício de quais comportamentos são
  considerados "contrato" pelo time atual.
- `docs/deploy-vps.md`, `docs/pwa.md` — fora de escopo do inventário.

## Mapeamento para as tabelas já existentes no destino (`termos-responsabilidade`, `db.py`)

- `bens` (destino: `numero INTEGER PK, situacao, descricao, complemento, classificacao,
  localizacao, data_entrada, valor_compra, valor_atual`) ≈ quase 1:1 com as colunas cadastrais
  do sga (`numero_bem, situacao, descricao, complemento, classificacao_contabil, local_sistema,
  data_entrada, valor_compra, valor_atual`). **Diferença chave**: o sga tem DOIS campos de
  local (`local_sistema` = cadastral/"deveria estar" e `local_verificado` = achado no último
  inventário); o destino tem só `bens.localizacao` (um valor). Ou seja, `bens.localizacao` no
  destino corresponde a `local_sistema` do sga — o resultado do inventário (`local_verificado`,
  `estado_conservacao`, `usuario_bem`, `foto`, `usuario_inventario`, `data_inventario`) **não
  tem onde morar em `bens`** hoje; precisa de tabela(s) nova(s), o que já é a expectativa do
  usuário ("novas tabelas no termos.db, só acréscimos").
- `localizacoes` (mapeadas a `responsaveis`/centro de custo) ≈ o papel de "sala" no sga, mas com
  cadastro formal (o sga não tem isso — ver dúvida 3). O inventário novo deveria iterar sobre
  `localizacoes` já existentes como a lista de "salas" a inventariar.
- `pessoas`/`atribuicoes` — não têm equivalente direto no módulo de inventário do sga (que só
  tem `usuario_bem`, texto livre, não uma FK para uma tabela de pessoas). Se o inventário novo
  quiser usar "quem usa" apontando para `pessoas`, é uma melhoria em relação ao sga, não uma
  migração direta.
- Não há no destino, hoje, equivalente a: evento/campanha de inventário, leitura individual,
  divergência, sobra, foto de bem — todas precisam ser desenhadas do zero (aproveitando os
  *nomes de campo e regras* do sga como referência, não a estrutura de dados, que era
  "tudo em uma linha de planilha").
