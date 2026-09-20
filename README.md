# Termos de Responsabilidade — CFC

Programa local (Windows) da Gerência de Serviços Administrativos (Gersev) para emitir Termos de Responsabilidade (por centro de
custo e individuais) e Termos de Devolução, com botão **Copiar para o SEI** e download em `.docx`.
Os dados ficam num SQLite (`dados/termos.db`) mantido pelo próprio programa.

## Uso

1. Abra `TermosCFC.exe`. A janela abre em `http://127.0.0.1:12345`.
   **Pesquisa** (lupa no cabeçalho, para quem tem acesso ao acervo): número do bem abre a ficha; texto procura em bens
   (descrição, complemento, localização, centro, pessoa), pessoas e centros de custo. Várias palavras: todas
   têm que bater (`computador GEX-LIC` = computadores do GEX-LIC; o "e" solto é ignorado). Texto simples
   busca "contém"; com `*` o padrão é literal (`GEX*` começa com GEX, `*ITEC` termina com ITEC). Nos
   resultados, centro ou pessoa têm *Ver bens* (lista filtrada) e *Termo*.
   **Processos SEI** (Cadastros → Processos SEI): um vigente por tipo de termo (centro de custo,
   individual, devolução). Sem processo vigente o termo não pode ser copiado nem baixado.
   **Termos emitidos**: cada cópia ou download registra data, processo e a lista de bens daquele
   momento (foto). Na tela do termo aparece o último registro e se entraram/saíram bens desde então
   (termo *desatualizado*). O número do documento SEI pode ser anotado depois, no registro.
   **Atualizar base** guarda o que mudou a cada importação (novos, removidos, movidos, situação) e a
   ficha do bem mostra o histórico dele.
   **Início** é enxuto: até seis atalhos (termo por centro de custo, termo individual, termo de devolução,
   termos emitidos, realizar inventário e Análise), os indicadores permitidos (ativos, valor sem imóveis,
   imóveis, sem centro nem pessoa, termos a emitir, última importação, andamento do inventário) e o card
   *Sobre o patrimônio*. Não tem mais gráficos nem busca no meio da tela — os gráficos ficam na Análise,
   e a busca, na lupa do cabeçalho. Cada usuário vê só os atalhos e indicadores cujas telas ele pode abrir.
   **Análise** filtra bens por qualquer combinação, com gráficos por dimensão, os cards de valor não
   informado × valor zero (contados à parte um do outro), *Abrir termo completo* do centro/pessoa quando
   couber — com o aviso *Os filtros desta análise não limitam o termo*, porque o documento sai inteiro —
   e *Exportar .xlsx* (`analise.xlsx`, com todos os bens do filtro, sem o limite de 1000 linhas da tabela
   na tela). A antiga tela **Recorte** passou a se chamar **Análise**; `/recorte` e `/recorte/xlsx`
   continuam funcionando e redirecionam para os novos endereços.
   **Menu** em árvore, recortado pelas funções de quem entrou: Início, Termos de Responsabilidade (grupo),
   Análise, Inventário (grupo), Cadastros (grupo), Textos, Atualizar base, Usuários e Ajuda. Grupo sem
   nenhum item permitido não aparece; o grupo da tela aberta já vem expandido e só o item da tela fica
   marcado (o relatório de um evento encerrado não acende o relatório do evento aberto).
   **Ajuda** (`/ajuda`) é o guia do sistema: sumário e seções apenas das funções do usuário, incluindo
   perguntas frequentes. O ícone de interrogação ao lado do título das telas de trabalho abre o guia direto
   na seção correspondente; a própria Ajuda, o login, as telas de erro e os documentos não têm esse botão.
   **Inventário** (menu próprio): abra um evento (nome, portaria, comissão; todas as salas com bens
   ativos ou uma amostra), escolha o integrante e leia as plaquetas por sala com leitor de código de
   barras, câmera do celular ou digitação. Bem lido na sala cadastrada = localizado; em outra sala =
   divergente (o cadastro do SPW não muda); não lido = pendente; sem cadastro = sobra (com foto).
   Bem baixado lido fica registrado (continua baixado). Conservação, quem usa, observação e foto por bem.
   Relatório e `.xlsx` por evento. O evento fica aberto até ser encerrado; encerrar congela tudo.
   Cada evento tem **Painel** (KPIs e gráficos por situação, integrante, conservação, andar e sala —
   o andar é o texto antes do primeiro `-` no nome da sala), relatório com filtros/ordenação e `.xlsx`
   com opção *Incluir fotos* (`=IMAGEM`; exige Microsoft 365 — em Excel antigo aparece `#NOME?`). Na sala, é possível marcar bens como localizados em lote
   (plaqueta ilegível) e desmarcar. Ao encerrar, os bens do evento são congelados
   (`inventario_bens_encerrados`): o relatório de um evento encerrado não muda quando a base do SPW
   é atualizada.
   Fotos vão para o bucket R2 configurado em `secrets/.env` (variáveis `R2_*`); sem ele, fotos ficam
   desativadas. Cada bem aceita várias fotos por evento (a sala mostra todas; relatório e `.xlsx` só a
   primeira, com "+N"); a chave no bucket é `<pasta>/<n>-<número do bem>.webp`, com a pasta sendo o nome
   do evento em minúsculas sem acento (ex.: `inventario2026/1-12334.webp`), por isso dois eventos não podem
   ter nomes que gerem a mesma pasta. O cadastro do bem (`/bem`) lista as fotos agrupadas por evento.
   A planilha de cadastros ganha abas `inv_*` para exportar/importar inventários inteiros
   (migração de outros sistemas). No menu, *Inventário* é um grupo com *Eventos* e, quando há
   evento aberto, o próprio evento, *Painel* e *Relatório*.
   **Usuários e funções** (site): entrar com login (ou e-mail, se cadastrado) e senha. O e-mail é opcional e
   único. Cada usuário soma quantas funções quiser — não são perfis excludentes: *Administrador* (tudo:
   usuários, abrir/encerrar/excluir inventário, exclusões e importação de cadastros), *Operador* (termos,
   cadastros, textos, atualizar base), *Consulta* (só vê o acervo; não emite termo), *Inventário* (lê bens,
   tira foto e registra sobra nos eventos em que está na comissão) e *Consulta de inventários* (acompanha
   telas, painel, relatório e `.xlsx` de qualquer evento, aberto ou encerrado, sem poder ler bens). As funções
   se somam: por exemplo, Inventário + Consulta de inventários enxerga todos os eventos, mas só grava leitura,
   foto e sobra no evento da própria comissão. A comissão de cada evento é escolhida pelo administrador entre
   os usuários com Administrador ou Inventário, na tela do próprio evento; criar um usuário com a função
   Inventário não o inclui automaticamente em nenhuma comissão — é preciso adicioná-lo depois, evento a
   evento. A leitura grava o nome de quem está logado.
   A entrada depende do acesso: quem enxerga o acervo cai no **Início**; quem só tem função de inventário é
   levado direto a **Inventário** (`/inventario`), sem lupa de pesquisa, Análise nem termos — nem por URL
   digitada (403). Dentro do módulo as duas funções se separam: quem tem só *Inventário* também não alcança
   painel, relatório nem exportação (403), e enxerga apenas os eventos de cujas comissões participa; quem tem
   *Consulta de inventários* alcança painel, relatório e `.xlsx` de todos os eventos, mas nenhuma ação de
   conferência. Quem tem Inventário e ainda não está em nenhuma comissão vê *Nenhum inventário atribuído a
   você*, sem nome nem total de evento alheio.
   Usuário não é excluído, só inativado. *Nova senha* gera uma senha temporária mostrada uma vez, com troca
   obrigatória no primeiro acesso. Cinco senhas erradas seguidas bloqueiam o login por 15 minutos. O programa
   Windows não pede senha: entra como "Administrador local". Renomear um usuário mantém o nome atualizado na
   comissão do evento aberto (se ele estiver nela); leituras já feitas continuam com o nome antigo.
   Quem vem de uma instalação com os quatro perfis antigos recebe, na migração, exatamente a função do mesmo
   código (*admin*→Administrador, *operador*→Operador, *consulta*→Consulta, *inventariante*→Inventário);
   isso é restrito por identidade, não por nome — usuários homônimos na comissão de um evento não são
   religados automaticamente, e o administrador precisa selecionar as contas certas na tela Comissão de cada
   evento que ainda estiver aberto.
2. **Atualizar base**: envie o export do sistema de patrimônio (`.xlsx`). Só a tabela de bens muda.
   *Exportar bens (formato SPW)* devolve a mesma tabela em `.xlsx`, nas 9 colunas do export — backup reimportável.
3. **Cadastros**: quatro áreas com busca visível, filtros, ordenação e paginação de 10/20/50 registros.
   Em **Centros de custo**, encontre a sigla ou o responsável e use *Editar*; alterar a sigla mantém
   as localizações e referências dos termos. Em **Localizações**, filtre as que estão sem centro,
   vincule-as ou altere o centro de uma ou várias localizações com revisão de origem e destino.
   A seleção em lote vale para a página atual e é limpa ao mudar busca, filtros ou página.
   Em **Pessoas**, use *Editar* para corrigir o nome ou *Ver bens* para consultar um patrimônio antes
   de atribuí-lo. Bem atribuído a pessoa não entra no termo do setor. Em **Processos SEI**, cadastre
   processos e revise o impacto antes de substituir o vigente, encerrar ou excluir.
   Inclusão e edição têm formulários próprios; erros mantêm os dados preenchidos. Salvar e cancelar
   preservam o contexto da lista. Exclusões e remoções de vínculo exigem confirmação; centros com
   bens ativos sob sua guarda e processos com termos registrados mantêm seus bloqueios de exclusão.
4. **Termos**: escolha o centro/pessoa → página do termo → *Copiar para o SEI* ou *Baixar .docx*.
5. **Textos**: os dizeres dos termos (abertura, compromissos, parágrafos, quem recebe a devolução,
   cidade, sigla do órgão) são editáveis no menu Textos, com marcadores como `{nome}` e `{ccustos}`;
   "Restaurar padrão" volta ao texto original.
6. **Planilha de cadastros**: Cadastros → *Exportar cadastros* gera `cadastros.xlsx` (4 abas; quando já
   existe inventário, mais 7 abas `inv_*`). Edite no Excel e importe em *Atualizar base → Importar
   cadastros* — substitui as 4 tabelas inteiras; as abas `inv_*` só são aceitas todas juntas (e então
   substituem o inventário inteiro) ou nenhuma (inventário preservado). A 6ª aba, `inv_bens_encerrados`,
   é o retrato dos bens de cada evento encerrado (gravado no encerramento); é opcional na importação —
   ausente, a tabela é mantida. A 7ª aba, `inv_fotos`, tem as fotos dos bens (evento, número, nfoto, url);
   também é opcional — ausente, uma coluna `foto_url` em `inv_leituras` (planilhas anteriores) vira a
   foto 1 de cada bem. Atenção: diferente de `inv_bens_encerrados`, quando as abas `inv_*` são importadas
   sem `inv_fotos` e sem a coluna `foto_url`, a tabela de fotos fica vazia (as imagens continuam no bucket,
   mas sem referência). A coluna `fotos_seq` de `inv_leituras` é o contador interno das fotos (maior número
   já usado por bem); é opcional na importação. Desmarcar um bem apaga a leitura e esse contador; se o bem
   for lido e fotografado de novo, a numeração recomeça em 1.

Backup = copiar a pasta `dados/`. `dados/segredo.txt` — chave que assina a sessão do navegador, criada
na primeira execução; na web pode vir da variável `TERMOS_SEGREDO`. Na VPS o container cria esse arquivo como root; para rodar o sistema fora do Docker na mesma pasta, defina `TERMOS_SEGREDO` no ambiente ou ajuste o dono do arquivo.

## Proposta comercial

`proposta/proposta.html` é uma página avulsa, sem servidor, para montar a proposta de preço a uma prefeitura:
abre com duplo clique, calcula a mensalidade pela quantidade de bens do acervo em faixas decrescentes
(tudo editável na própria página, com as edições guardadas no navegador) e gera o PDF pelo botão
*Salvar em PDF*. Não faz parte do sistema hospedado.

## Desenvolvimento

    python -m venv .venv && .venv/bin/pip install -r requirements.txt
    .venv/bin/pytest
    .venv/bin/python app.py        # http://127.0.0.1:12345 (debug)
    .venv/bin/python main.py       # como o programa: janela (ou navegador, se não houver WebView)

A porta pode ser trocada com a variável `TERMOS_PORTA` (padrão 12345), útil para testar sem conflitar
com outra instância já rodando.

Migração inicial a partir das planilhas antigas: `python importar_planilhas.py acervo.xlsx geral.xlsx`.

## Gerar o executável (Windows)

    python -m venv .venv && .venv\Scripts\activate && pip install -r requirements.txt
    build.bat

Sai em `dist\TermosCFC\`. Distribua a pasta inteira (zip). Requer o WebView2 Runtime (já vem no
Windows 10/11 atualizados); sem ele o programa abre no navegador padrão.

Copie o `dados\termos.db` já migrado (por exemplo, da máquina onde rodou o `importar_planilhas.py`)
para `dist\TermosCFC\dados\` antes de distribuir; sem isso o programa abre com a base vazia.

### Observações

Os parágrafos do termo por centro de custo não carregam mais 4 espaços em branco no início (era um
resíduo de indentação do código antigo) — confira o `.docx` gerado.

## Servir na web (Docker)

O mesmo código roda em `https://patrimonio.sistemascfc.org`, num container nesta máquina, atrás do nginx do host
(container em `127.0.0.1:12012`, Certbot). O login é do próprio sistema (`TERMOS_LOGIN=1` no `compose.yml`).

    docker compose up -d --build   # (re)constrói e sobe; dados em ./dados (termos.db, timbrado.docx)
    docker compose logs -f         # acompanhar
    docker compose exec web python usuarios.py criar-admin antonio "Antônio Sousa" antonio@cfc.org.br   # primeiro administrador (ou redefinir a senha de um admin)

Publicação da Fase 4 (uma vez): subir o container; criar o administrador pelo comando acima; entrar e criar os
usuários da comissão do evento aberto com **exatamente** os nomes já gravados nas leituras (Inventário → evento →
*Comissão* mostra quem já tem leituras); remover `auth_basic` e `auth_basic_user_file` do vhost
`/etc/nginx/conf.d/patrimonio.sistemascfc.org.conf` e recarregar (`nginx -t && systemctl reload nginx`).
Enquanto o `auth_basic` ficar, o site pede as duas senhas, sem prejuízo.

`Dockerfile` (alvos `web` e `robo`), `compose.yml` e `.dockerignore` são só do site; o programa de desktop não os
usa. O vhost fica em `/etc/nginx/conf.d/patrimonio.sistemascfc.org.conf`. Backup continua sendo copiar a pasta
`dados/`.

### Emissão no SEI

O botão **Emitir Termo no SEI** (página do termo) e **Atualizar com SPW** (Atualizar base) enfileiram pedidos em
`robo_pedidos`; quem atende é `atender_pedidos.py`, rodando **dentro do container `robo`** (`compose.yml`,
alvo `robo` do `Dockerfile`), que tem Playwright e xlrd por cima da imagem do site:

    docker compose up -d --build   # constrói as duas imagens (web e robo) e sobe os dois containers
    docker compose logs -f robo    # acompanhar o trabalhador (pedidos atendidos, erros)

O container `robo` é permanente (`restart: unless-stopped`), sem porta exposta; monta `./dados` (banco, logs,
capturas de erro) e `./secrets` (somente leitura). Sem ele de pé, o site continua funcionando: o pedido fica
"aguardando a vez" e a tela avisa depois de 2 minutos.

Segredos em `secrets/sei.env` (`SEI_LOGIN_URL`, `SEI_ORGAO`; chmod 600); usuário, senha e unidade são os cadastrados
por quem emitiu em *Meus acessos* (ver "Acessos por usuário (SEI e SPW)" abaixo). No SEI, crie à mão
um bloco de assinatura por unidade, com o nome exato `Termos {UNIDADE}` (ex.: `Termos GECONT`); o sistema só inclui o
documento — disponibilizar o bloco continua sendo feito por vocês. Em Cadastros, informe a "Unidade no SEI" das pessoas
que recebem termo individual (e a exceção no centro de custo cuja sigla difere da unidade do SEI). Em Textos, o nome do
tipo de documento por tipo de termo.

Diagnóstico: `dados/robo_pedidos.log` (exceções), `dados/sei/erro.png` (tela do SEI no erro), tabela `robo_pedidos`.

### Acessos por usuário (SEI e SPW)

As senhas de cada operador ficam cifradas (Fernet, `cofre.py`) no banco, nunca em texto puro. Uma vez, gere a
chave e guarde-a:

    .venv/bin/python -c "import cofre; print(cofre.gerar_chave())"   # cole o resultado em secrets/chaves.env
    # secrets/chaves.env:
    # CHAVE_SENHAS=...
    chmod 600 secrets/chaves.env

As duas imagens (`web` e `robo`) montam `./secrets` (somente leitura). **`backup.sh` não copia `chaves.env`** —
ele só leva `termos.db` para o bucket R2; guarde a chave à parte, à mão, em outro lugar (por exemplo, no
gerenciador de senhas do administrador). Sem ela as senhas cifradas em `termos.db` são inúteis; quem restaura um
backup sem a chave junto precisa pedir que cada operador cadastre o acesso de novo. Com isso, `secrets/sei.env`
fica só com `SEI_LOGIN_URL` e `SEI_ORGAO`.

Cada operador cadastra o próprio acesso em **Meus acessos** (link no cabeçalho, `/meus-acessos`): ao SEI (usuário,
senha e sigla da unidade — é nessa unidade e em nome dessa pessoa que os termos nascem no SEI, e é lá que o bloco
`Termos {SIGLA}` precisa existir) e ao SPW (usuário e senha — usado por *Atualizar com SPW* pelo site). Sem acesso
cadastrado, o site recusa antes de enfileirar ("Cadastre seu acesso ao SEI em Meus acessos para emitir."/"...ao SPW
... para atualizar."). Ao trocar a senha no SEI ou no SPW, atualize na mesma tela (deixar a senha em branco mantém
a atual). O administrador, na lista de usuários, só vê se há acesso cadastrado (coluna "Acessos") e pode apagá-lo
(*Apagar acessos*) — nunca vê a senha.

O cron da madrugada (`./atualizar_base.sh`) continua usando `secrets/spw.env` completo, sem depender de acesso de
ninguém (ver "Robô do SPW" abaixo).

### Robô do SPW

`importar_spw.py` entra no SPW, exporta a relação de bens (Excel/Detalhado), compara com a última execução e,
se mudou, importa como o upload de Atualizar base faria. Roda dentro do container `robo` (acima), em dias úteis
às 3h, e registra cada execução em `robo_execucoes`: o card "última importação" do Início mostra `robô ok`/`robô
falhou` e Atualizar base lista as últimas execuções. Segredos em `secrets/spw.env` (`SPW_USUARIO`, `SPW_SENHA`,
`SPW_LOGIN_URL`, `SPW_CONSULTA_URL`, chmod 600). Pelo site, *Atualizar com SPW* usa a credencial de quem clicou
(ver "Acessos por usuário" acima); o cron continua com `spw.env`.

`atualizar_base.sh` não roda mais o robô diretamente: ele entra no container já de pé com `docker compose exec`
(e sai com um aviso claro se o `robo` não estiver rodando):

    ./atualizar_base.sh            # docker compose exec -T robo python importar_spw.py; sai 0/1
    ./atualizar_base.sh --teste    # mesma coisa numa cópia em dados/robo-teste, sem tocar na base real

Crontab (`crontab -e`, usuário dono de `dados/termos.db`); o script já grava em `dados/robo_spw.log` (o container
`robo` precisa estar de pé — `docker compose up -d` — para o cron funcionar):

    0 3 * * 1-5 /opt/web/termos-responsabilidade/atualizar_base.sh >/dev/null 2>>/opt/web/termos-responsabilidade/dados/robo_spw.log

O robô não importa se o export vier com menos de 90% dos bens da base (protege contra export vazio ou truncado);
nesse caso registra erro e a baixa em massa, se for real, passa por Atualizar base. Um upload manual entre
execuções fica valendo até o SPW mudar: o robô compara o export com a própria execução anterior, não com a base.

Diagnóstico: `dados/robo_spw.log` (uma linha por execução) e `dados/spw/erro.png` (tela do SPW no momento do erro).
Se o SPW mudar o layout, os seletores ficam todos em `baixar_export`.

`requirements-robo.txt` (Playwright, xlrd) continua no repositório: a imagem `robo` não o usa (as dependências
estão fixas no `Dockerfile`), mas os scripts de spike em `docs/superpowers/notes/` ainda dependem dele.

## Arquivos

| Arquivo | Função |
|---|---|
| `app.py` | rotas Flask |
| `app_cadastros.py` | navegação, formulários e revisão das alterações de cadastros |
| `db.py` | esquema, importação, consultas, cadastros |
| `termos_html.py` | corpo HTML dos termos (padrão gelic; tabelas 90 %) |
| `textos.py` | textos padrão dos termos e marcadores |
| `Script_Termo_Individual.py`, `Termo_de_Responsabilidade.py`, `termo_devolucao.py` | geradores `.docx` |
| `config.py` | pasta de dados (`TERMOS_DADOS` sobrepõe) |
| `main.py`, `build.bat` | programa de desktop e build |
| `Dockerfile` | dois alvos: `web` (site) e `robo` (`FROM web`, + Playwright e xlrd, atende `robo_pedidos`) |
| `compose.yml` | serviços `web` (site, porta 12012) e `robo` (trabalhador da fila, sem porta) |
| `templates/`, `static/dsgov/` | telas DSGov 3.7.0 (offline) |
| `painel.py`, `graficos.py` | cards de gráfico (ECharts embutido, tema DSGov) |
| `inventario.py`, `fotos.py`, `app_inventario.py` | módulo de inventário (dados, fotos no R2, rotas) |
| `usuarios.py`, `app_usuarios.py` | usuários, senhas, matriz de permissões e telas de login/usuários |
| `cofre.py` | cifra (Fernet) das senhas do SEI e do SPW guardadas por usuário; chave em `secrets/chaves.env` |
| `atualizar_base.sh` | Chama o robô do SPW dentro do container `robo` via `docker compose exec` (`--teste` usa uma cópia da base); o cron chama o mesmo script |
| `importar_spw.py` | Robô do SPW: exporta, converte e importa os bens (roda no container `robo`, por cron) |
| `requirements-robo.txt` | Não é usado pela imagem `robo` (dependências fixas no `Dockerfile`); ainda serve os scripts de spike em `docs/superpowers/notes/` |
| `robo_sei.py` | Robô do SEI: login, cria o documento no processo e inclui no bloco de assinatura |
| `atender_pedidos.py` | Trabalhador do container `robo`: atende a fila `robo_pedidos` (emissão no SEI e atualização com o SPW) |
| `segredos.py` | Lê os arquivos `secrets/*.env` (SPW, SEI) |
