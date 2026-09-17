# Termos de Responsabilidade — CFC

Programa local (Windows) da Gerência de Serviços Administrativos (Gersev) para emitir Termos de Responsabilidade (por centro de
custo e individuais) e Termos de Devolução, com botão **Copiar para o SEI** e download em `.docx`.
Os dados ficam num SQLite (`dados/termos.db`) mantido pelo próprio programa.

## Uso

1. Abra `TermosCFC.exe`. A janela abre em `http://127.0.0.1:12345`.
   **Pesquisa** (lupa no cabeçalho, em todas as telas): número do bem abre a ficha; texto procura em bens
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
   **Início** é o painel: cards e gráficos por situação, centro, classificação, localização, idade, ano e
   faixa de valor, todos clicáveis. **Recorte** filtra bens por qualquer combinação, com os mesmos
   gráficos, o termo do centro/pessoa quando couber e *Exportar .xlsx*.
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
   desativadas. A planilha de cadastros ganha abas `inv_*` para exportar/importar inventários inteiros
   (migração de outros sistemas).
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
   existe inventário, mais 6 abas `inv_*`). Edite no Excel e importe em *Atualizar base → Importar
   cadastros* — substitui as 4 tabelas inteiras; as abas `inv_*` só são aceitas todas juntas (e então
   substituem o inventário inteiro) ou nenhuma (inventário preservado). A 6ª aba, `inv_bens_encerrados`,
   é o retrato dos bens de cada evento encerrado (gravado no encerramento); é opcional na importação —
   ausente, a tabela é mantida.

Backup = copiar a pasta `dados/`. `dados/segredo.txt` — chave que assina a sessão do navegador, criada
na primeira execução; na web pode vir da variável `TERMOS_SEGREDO`. Na VPS o container cria esse arquivo como root; para rodar o sistema fora do Docker na mesma pasta, defina `TERMOS_SEGREDO` no ambiente ou ajuste o dono do arquivo.

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
(padrão dos demais sites: container em `127.0.0.1:12012`, Certbot, e senha `auth_basic` em
`/etc/nginx/.htpasswd_patrimonio`, porque o sistema não tem login próprio).

    docker compose up -d --build   # (re)constrói e sobe; dados em ./dados (termos.db, timbrado.docx)
    docker compose logs -f         # acompanhar
    sudo htpasswd /etc/nginx/.htpasswd_patrimonio patrimonio   # trocar a senha do site

`Dockerfile`, `compose.yml` e `.dockerignore` são só do site; o programa de desktop não os usa. O vhost fica em
`/etc/nginx/conf.d/patrimonio.sistemascfc.org.conf`. Backup continua sendo copiar a pasta `dados/`.

## Arquivos

| Arquivo | Função |
|---|---|
| `app.py` | rotas Flask |
| `app_cadastros.py` | navegação, formulários e revisão das alterações de cadastros |
| `db.py` | esquema, importação, consultas, cadastros |
| `termos_html.py` | corpo HTML dos termos (padrão gelic; tabelas 80 % / 100 %) |
| `textos.py` | textos padrão dos termos e marcadores |
| `Script_Termo_Individual.py`, `Termo_de_Responsabilidade.py`, `termo_devolucao.py` | geradores `.docx` |
| `config.py` | pasta de dados (`TERMOS_DADOS` sobrepõe) |
| `main.py`, `build.bat` | programa de desktop e build |
| `Dockerfile`, `compose.yml` | site em patrimonio.sistemascfc.org |
| `templates/`, `static/dsgov/` | telas DSGov 3.7.0 (offline) |
| `painel.py`, `graficos.py` | cards de gráfico (ECharts embutido, tema DSGov) |
| `inventario.py`, `fotos.py`, `app_inventario.py` | módulo de inventário (dados, fotos no R2, rotas) |
