# Fase 5: menu em árvore, Início enxuto, Análise no lugar de Recorte e página Ajuda

**Data:** 2026-09-17
**Estado:** referência histórica, substituída em 2026-09-17 pelas specs
`2026-09-17-fase5-acessos-design.md` e `2026-09-17-fase5-experiencia-design.md`,
conforme decisões registradas em `../notes/2026-09-17-fase5-grill.md`.
**Atenção ao executor:** não executar as regras antigas de perfil único, inventariante
com consulta geral, indicador único "sem valor" ou Ajuda igual para todas as funções.
**Base:** `main` após a Fase 4 (`2026-09-17-fase4-usuarios-design.md`). Esta fase depende dela: o menu e a
Ajuda são filtrados por perfil, e "Usuários" é item do menu.

## 1. Objetivo e decisões

Pedido do usuário: um menu mais organizado, com "Termos de Responsabilidade" expandindo como o Inventário
faz hoje; Recorte renomeado e reunindo "todas as informações, filtros e emissão de relatórios sobre o
acervo" que hoje se dividem entre Início e Recorte; Início só com acesso rápido, KPIs e um texto
informativo sobre a atividade de patrimônio, "uma versão web do README para usuários"; a barra de pesquisa
sai do meio da tela e fica só no cabeçalho.

Decisões dele (2026-09-17):

- A tela de filtros e gráficos chama-se **Análise**, rota `/analise`; `/recorte` redireciona.
- A tabela "Termos por centro de custo" **sai do Início** e não vai para lugar nenhum: a tela Termo por
  centro de custo já a tem, com mais colunas (último termo, e-mail).
- Texto informativo: **resumo no Início + página Ajuda** no menu, com o guia completo. Texto fixo no
  template (não editável em Textos), escrito a partir do README em linguagem de usuário; o texto integral
  está na seção 6 desta spec para revisão.
- **Cadastros também vira grupo** no menu (Centros de custo, Localizações, Pessoas, Processos SEI).
  Textos, Atualizar base e Usuários continuam itens soltos.

Fora: editar textos da Ajuda pela interface, busca dentro da Ajuda, tour guiado, mudar o conteúdo dos
gráficos ou filtros da Análise, novos KPIs além dos já calculados, PDF.

---

## 2. Menu (`menu.py`, `app.py`, `templates/base.html`)

A estrutura sai do `context_processor` de `app.py` (onde a Fase 4 já a filtra por perfil) para uma função
em módulo próprio:

```python
def montar(perfil: str, evento_aberto: dict | None, endpoint_atual: str) -> list[dict]
# cada item: {"rotulo", "icone", "url", "endpoint", "filhos": [ {rotulo, url, endpoint} ], "aberto": bool}
```

Árvore completa (antes do filtro por perfil):

| # | Item | Ícone | Filhos |
|---|---|---|---|
| 1 | Início | `fa-home` | |
| 2 | Termos de Responsabilidade | `fa-file-signature` | Termo por centro de custo · Termo individual · Termo de devolução · Termos emitidos |
| 3 | Análise | `fa-chart-bar` | |
| 4 | Inventário | `fa-clipboard-check` | Eventos · *nome do evento aberto* · Painel · Relatório (os três últimos só com evento aberto) |
| 5 | Cadastros | `fa-address-book` | Centros de custo · Localizações · Pessoas · Processos SEI |
| 6 | Textos | `fa-pen-nib` | |
| 7 | Atualizar base | `fa-upload` | |
| 8 | Usuários | `fa-users` | (só admin, Fase 4) |
| 9 | Ajuda | `fa-question-circle` | |

Regras:

- Um grupo só aparece se ao menos um filho for permitido ao perfil (`usuarios.permitido`, Fase 4); um item
  solto só se o endpoint dele for permitido. Ajuda é permitida a todos.
- **Grupo aberto ao carregar** quando um dos filhos é a tela atual (`endpoint_atual`), ou quando a tela atual
  pertence ao grupo sem estar listada (ex.: `termo`, `termo_emitido` → grupo Termos; `inventario.sala_tela`
  → grupo Inventário; `cadastros_novo`, `cadastros_editar` → grupo Cadastros). Mapeamento explícito em
  `menu.py` (`PERTENCE = {"termo": "termos", ...}`). Abrir = classe `active` no `.menu-folder` e
  `aria-expanded="true"` no título, que é o que o core do DSGov usa ao clicar.
- Título do grupo continua sendo `<a href="javascript:void(0)">` (decisão de 2026-09-17, commit `aa1f826`):
  clicar abre e fecha, não navega.
- Item da tela atual recebe `active` (o DSGov o destaca). Vale para filhos e para itens soltos.
- O ícone de Textos muda de `fa-file-signature` para `fa-pen-nib` porque `fa-file-signature` passa a ser o
  grupo Termos.

`templates/base.html` renderiza a lista com a mesma marcação de hoje mais as classes acima. Nada de
JavaScript novo.

---

## 3. Início (`app.py:home`, `templates/index.html`)

Três blocos, nesta ordem:

1. **Acesso rápido**: o carrossel de hoje com seis cards: Termo por centro de custo, Termo individual,
   Termo de devolução, Termos emitidos, Realizar inventário e **Análise** (`fa-chart-bar`, "Filtros,
   gráficos e exportação do acervo"). Filtrados por perfil com a mesma regra do menu (card só aparece se a
   tela é permitida).
2. **KPIs**: os 7 cards de hoje, sem mudança de conteúdo. Os que hoje apontam para `/recorte` passam a
   apontar para `/analise` (é o mesmo `painel.url_recorte`, que passa a gerar `/analise`). O aviso "N bens
   ativos sem valor" continua acima deles.
3. **Sobre o patrimônio**: card com o texto da seção 6.1 (2 ou 3 parágrafos) e botão secundário "Ver guia"
   para `/ajuda`.

Saem do Início: o card "Pesquisar" (a lupa do cabeçalho já abre a busca em todas as telas), os cards de
gráfico e a tabela "Termos por centro de custo". A rota `home` deixa de chamar `painel.cards_graficos`;
`db.painel` deixa de calcular `dimensoes` (passa a ser chamado só com o que o Início usa: contadores,
última importação, centros e pessoas para "a emitir"). O template deixa de carregar ECharts.

Título da página continua "Início"; trilha vazia como hoje.

---

## 4. Análise (`app.py`, `templates/analise.html`, `painel.py`)

O Recorte de hoje, renomeado e com KPIs no topo. Nada do que existe sai.

- Rotas `GET /analise` e `GET /analise/xlsx` (endpoints `analise` e `analise_xlsx`). `GET /recorte` e
  `GET /recorte/xlsx` devolvem 301 para as novas com a query string intacta. `painel.url_recorte` e
  `url_recorte_xlsx` passam a gerar `/analise...` (nomes de função ficam; só o caminho muda). O nome
  `recorte` em `db.py` (`db.recorte`, `db.exportar_recorte`) fica: é o conceito, não a tela.
- Cabeçalho: título "Análise", frase descritiva do recorte (`painel.descrever`), botões "Termo de X" quando
  couber e "Exportar .xlsx", como hoje.
- Filtros: o formulário de hoje, sem mudança.
- **KPIs sob filtro** (linha nova, substitui os dois contadores atuais): bens no recorte · valor atual ·
  imóveis (quantidade e valor) · sem centro nem pessoa · sem valor. Todos calculados sobre o recorte
  (`db.recorte` passa a devolver `imoveis`, `valor_imoveis`, `sem_centro`, `sem_valor` além de
  `quantidade` e `valor_total`; uma consulta agregada a mais, com o mesmo `where`). Clicáveis: imóveis
  acrescenta `classificacao=imoveis` ao filtro atual, sem centro acrescenta `ccusto=-&pessoa=-`, sem valor
  acrescenta `valor_ate=0`. Quando o filtro já fixa a dimensão, o card mostra o número sem link.
- Gráficos por dimensão e tabela dos bens: como hoje. Aviso de truncamento em 1.000 bens: como hoje.
- Trilha: `[("Análise", None)]`. Nome do arquivo exportado: `analise.xlsx`.
- Texto da página Textos, README e demais telas que citam "Recorte" passam a dizer "Análise" (busca por
  "Recorte"/"recorte" nos templates e no README; `test_app.py` e `test_painel.py` atualizados).

---

## 5. Ajuda (`app.py:ajuda`, `templates/ajuda.html`)

- `GET /ajuda`, permitida a todo perfil (entra em `PERMISSOES` da Fase 4 para os quatro).
- Página com sumário no topo (lista de âncoras) e uma seção `<section id="...">` por tema, na ordem do
  menu: `inicio`, `pesquisa`, `termos`, `analise`, `inventario`, `cadastros`, `textos`, `atualizar-base`,
  `usuarios`, `perguntas`. Cada seção tem `<h2>` e o texto da seção 6.2.
- Seções de telas que o perfil não pode abrir não são renderizadas (mesma regra do menu: a seção aparece
  se ao menos um endpoint dela é permitido; `perguntas` sempre aparece; `usuarios` só para admin). A seção
  "Sua conta" (trocar senha, sair) aparece só quando há login (`g.usuario["id"]` não nulo).
- **Link "Ajuda desta tela"**: em `base.html`, ao lado do `<h1>` de cada tela, um `<a class="br-button
  circle small" aria-label="Ajuda desta tela">` com `fa-question-circle` para `/ajuda#<âncora>`. A âncora
  vem de `menu.AJUDA[endpoint]` (mapa endpoint → âncora, com `PERTENCE` cobrindo as telas filhas). Tela sem
  mapeamento não mostra o botão. Como os `<h1>` hoje estão dentro de cada template, o botão é inserido
  pelo bloco `conteudo` via macro `titulo(texto)` em `_macros.html`, que os templates passam a usar no
  lugar do `<div class="d-flex ..."><h1>` repetido (cerca de 20 templates; mudança mecânica).
- Marcação DSGov: `br-card` por seção, `br-list` para passos, `br-message info` para avisos. Sem imagens.

---

## 6. Textos

Escritos a partir do README, para quem usa e não para quem programa: sem nomes de arquivo, tabela ou
variável. O usuário revisa aqui; ajustes de redação depois da implementação são bem-vindos e não exigem
nova spec.

### 6.1 Início: "Sobre o patrimônio"

> **Patrimônio do CFC**
>
> Todo bem permanente do Conselho (móveis, equipamentos, veículos, imóveis) tem um número de patrimônio e
> um responsável pela guarda: um centro de custo, representado pelo seu titular, ou uma pessoa, quando o
> bem é de uso individual. O Termo de Responsabilidade formaliza essa guarda e é assinado no SEI; o Termo
> de Devolução registra quando o bem volta ao Setor de Patrimônio.
>
> Este sistema reúne o cadastro dos bens (importado do sistema de patrimônio), os responsáveis, a emissão
> dos termos, a análise do acervo e o inventário anual, feito sala a sala com leitura das plaquetas.
>
> Use os atalhos acima para as tarefas do dia a dia, a lupa no cabeçalho para achar um bem, uma pessoa ou
> um centro, e o guia para saber como cada tela funciona.

Botão: **Ver guia** → `/ajuda`.

### 6.2 Ajuda: guia do usuário

**Sumário**: Início · Pesquisa · Termos de Responsabilidade · Análise · Inventário · Cadastros · Textos ·
Atualizar base · Usuários · Sua conta · Perguntas frequentes.

#### Início (`#inicio`)

A página inicial tem atalhos para as tarefas mais comuns e um resumo do acervo em cards: bens ativos,
valor dos ativos sem imóveis, imóveis, bens sem centro nem pessoa, termos a emitir ou reemitir, data da
última atualização da base e andamento do inventário aberto. Clique em qualquer card para ver a lista
correspondente na Análise. Um aviso aparece quando há bens ativos sem valor cadastrado: eles não entram
nas somas.

#### Pesquisa (`#pesquisa`)

A lupa no cabeçalho funciona em todas as telas.

- Digite o **número do patrimônio** para abrir a ficha do bem.
- Digite um **texto** para procurar em bens (descrição, complemento, localização, centro, pessoa), pessoas
  e centros de custo. Com várias palavras, todas precisam bater: `computador GEX-LIC` traz os computadores
  do GEX-LIC.
- Use `*` para começar ou terminar com: `GEX*` acha o que começa com GEX; `*ITEC` o que termina com ITEC.
- Nos resultados, um centro ou uma pessoa têm **Ver bens** (a lista do que está sob a guarda deles) e
  **Termo**.

A **ficha do bem** mostra os dados atuais, o responsável, o histórico de mudanças registrado a cada
atualização da base (entrada, saída, mudança de sala ou de situação) e as fotos tiradas em inventários.

#### Termos de Responsabilidade (`#termos`)

Há três termos, cada um com um processo SEI vigente cadastrado em Cadastros → Processos SEI. **Sem processo
vigente, o termo não pode ser copiado nem baixado.**

- **Termo por centro de custo**: todos os bens ativos nas localizações do centro, exceto os atribuídos a
  uma pessoa. A tabela da tela mostra, por centro, quantos bens, o valor, quando saiu o último termo e a
  situação: *vigente*, *desatualizado* (entraram ou saíram bens desde a última emissão) ou *sem termo*.
- **Termo individual**: os bens atribuídos a uma pessoa em Cadastros → Pessoas. Bem atribuído a pessoa não
  entra no termo do setor.
- **Termo de devolução**: escolha a pessoa, marque os bens devolvidos (ou todos) e gere o termo.

Na página do termo: **Copiar para o SEI** copia o documento formatado para colar num documento do SEI;
**Baixar .docx** gera o arquivo Word com o timbrado. Qualquer das duas ações **registra a emissão**: data,
processo e a lista de bens daquele momento. Depois, em **Termos emitidos**, anote o número do documento
SEI e o bloco de assinatura, e registre o envio do e-mail de assinatura (o botão abre seu programa de
e-mail com o texto pronto). Emissões não podem ser apagadas: são o histórico.

#### Análise (`#analise`)

Filtre o acervo por situação, centro de custo, pessoa, localização, classificação, idade, valor e data de
entrada. A tela mostra os totais do recorte (bens, valor, imóveis, sem responsável, sem valor), um gráfico
por dimensão e a lista dos bens. Tudo é clicável: clicar numa fatia ou coluna acrescenta aquele filtro.
Quando o recorte é um centro ou uma pessoa, aparece o botão do termo correspondente. **Exportar .xlsx**
baixa a lista completa do recorte, mesmo quando a tela mostra só os primeiros 1.000 bens.

#### Inventário (`#inventario`)

O inventário é feito por **eventos**. Só há um evento aberto por vez.

1. **Abrir evento** (administrador): nome, portaria, comissão (usuários que poderão ler) e salas (todas com
   bens ativos ou uma amostra).
2. **Ler por sala**: abra a sala e leia as plaquetas com leitor de código de barras, câmera do celular ou
   digitando. Você precisa estar na comissão do evento.
   - Bem lido na sala em que está cadastrado: **localizado**.
   - Lido em outra sala: **divergente** (o cadastro do sistema de patrimônio não muda; a divergência fica
     registrada).
   - Não lido: **pendente**.
   - Plaqueta sem cadastro: registre como **sobra**, com descrição e foto.
   - Bem baixado lido fica registrado e continua baixado.
   Para cada bem lido você pode informar conservação, quem usa, observação e tirar fotos (várias por bem).
   Quando a plaqueta está ilegível, marque bens em lote como localizados; **Desmarcar** apaga a leitura.
3. **Acompanhar**: o **Painel** do evento tem indicadores e gráficos por situação, integrante,
   conservação, andar e sala; o **Relatório** lista tudo com filtros e ordenação e exporta `.xlsx`, com a
   opção de incluir as fotos (exige Excel do Microsoft 365).
4. **Encerrar** (administrador) congela o evento: o relatório de um evento encerrado não muda quando a base
   é atualizada.

O administrador pode alterar a comissão de um evento aberto e excluir um evento inteiro (com confirmação
pelo nome; não há como desfazer).

#### Cadastros (`#cadastros`)

Quatro áreas, todas com busca, filtros, ordenação e paginação.

- **Centros de custo**: sigla, responsável, e-mail, matrícula e função. Alterar a sigla mantém as
  localizações e os termos ligados a ela.
- **Localizações**: cada sala pertence a um centro de custo. Filtre as que estão **sem centro** e vincule
  uma ou várias de uma vez, revisando origem e destino antes de confirmar.
- **Pessoas**: quem tem bens de uso individual. **Ver bens** mostra o que está com a pessoa; **Atribuir**
  e **Desatribuir** movem bens entre o setor e a pessoa.
- **Processos SEI**: um processo vigente por tipo de termo. Substituir o vigente, encerrar ou excluir mostra
  antes o impacto nos termos.

Exclusões pedem confirmação e só o administrador as faz. Um centro com bens ativos sob sua guarda e um
processo com termos registrados não podem ser excluídos.

**Planilha de cadastros**: *Exportar cadastros* gera um `.xlsx` com as quatro áreas (e, quando há
inventário, as abas do inventário). Edite no Excel e importe em Atualizar base → *Importar cadastros*: a
importação **substitui as tabelas inteiras**, por isso só o administrador a faz. As abas de inventário
são aceitas todas juntas (substituem o inventário inteiro) ou nenhuma (inventário preservado).

#### Textos (`#textos`)

Os dizeres dos termos (abertura, compromissos, parágrafos, quem recebe a devolução, cidade, sigla do
órgão, texto do e-mail de assinatura) são editáveis, com marcadores como `{nome}` e `{ccustos}` que o
sistema preenche. **Restaurar padrão** volta ao texto original.

#### Atualizar base (`#atualizar-base`)

Envie o arquivo `.xlsx` exportado do sistema de patrimônio. Só a tabela de bens muda: responsáveis,
localizações, pessoas e atribuições ficam como estão. Cada atualização registra o que mudou (novos,
removidos, movidos, situação) e alimenta o histórico da ficha do bem. *Exportar bens* devolve a base atual
no mesmo formato, como cópia de segurança.

#### Usuários (`#usuarios`, só administrador)

Crie um usuário com login, nome e perfil: **administrador** (tudo, inclusive usuários, exclusões e
inventário), **operador** (termos, cadastros, textos, atualizar base), **inventariante** (lê bens nos
inventários em que estiver na comissão) ou **consulta** (só vê). Usuário não é excluído, é **inativado**.
**Nova senha** gera uma senha temporária, mostrada uma única vez: passe-a à pessoa, que terá de trocá-la
no primeiro acesso. Após cinco tentativas erradas o login fica bloqueado por 15 minutos.

#### Sua conta (`#conta`, só com login)

Seu nome e perfil aparecem no cabeçalho. **Sair** encerra a sessão. Para trocar a senha, abra
**Trocar senha** no cabeçalho: informe a atual e a nova (mínimo de 8 caracteres). A sessão expira após
12 horas sem uso.

#### Perguntas frequentes (`#perguntas`)

- **O termo está "desatualizado". O que faço?** Entraram ou saíram bens desde a última emissão. Abra o
  termo, confira a lista e emita de novo; a emissão anterior continua no histórico.
- **Um bem aparece no termo do setor, mas é de uso pessoal.** Atribua-o à pessoa em Cadastros → Pessoas.
  Ele sai do termo do setor e entra no termo individual.
- **Divergente ou pendente?** Divergente foi lido, mas em outra sala; pendente não foi lido. Ambos
  aparecem no relatório do inventário para providências.
- **Li o bem na sala errada.** Leia de novo na sala certa: a leitura mais recente vale.
- **A foto não sobe.** Verifique a conexão; a foto é enviada na hora. Se o problema persistir, avise o
  administrador.
- **Onde está o Recorte?** Virou **Análise**, com os mesmos filtros e mais indicadores.
- **A base do sistema de patrimônio mudou e o inventário encerrado não refletiu.** Correto: o evento
  encerrado é congelado. Só o evento aberto acompanha a base.

---

## 7. Arquivos

Novos: `menu.py`, `templates/analise.html` (renomeado de `recorte.html`), `templates/ajuda.html`,
`tests/test_menu.py`, `tests/test_ajuda.py`.

Alterados: `app.py` (rotas `analise`, `analise_xlsx`, redirecionamentos, `ajuda`, `home` enxuta,
`context_processor` chamando `menu.montar`), `painel.py` (`url_recorte` → `/analise`), `db.py`
(`painel` sem `dimensoes`; `recorte` com os agregados novos), `templates/index.html`,
`templates/base.html` (árvore com `active`/`aria-expanded`), `templates/_macros.html` (macro `titulo`),
todos os templates com `<h1>` (macro `titulo`), `README.md` (Recorte → Análise; menu; Ajuda),
`tests/test_app.py`, `tests/test_painel.py`.

---

## 8. Testes

- **Menu** (`test_menu.py`): árvore completa para admin; filtro por perfil (inventariante não vê Termos nem
  Cadastros; consulta não vê Textos, Atualizar base, Usuários; todos veem Ajuda); grupo da tela atual vem
  `aberto` (inclusive telas filhas via `PERTENCE`); item atual `active`; Inventário mostra as telas do
  evento só com evento aberto.
- **Início**: sem `<form` de pesquisa, sem `echarts`, sem tabela "Termos por centro"; 6 atalhos para admin
  e só os permitidos para os outros perfis; KPIs apontam para `/analise`; texto "Sobre o patrimônio" e
  botão "Ver guia" presentes.
- **Análise**: `/recorte?ccusto=CCI` → 301 para `/analise?ccusto=CCI`; `/recorte/xlsx?...` → 301;
  `/analise` renderiza os KPIs sob filtro com os valores certos no cenário semeado (imóveis, sem centro,
  sem valor) e os links de drill-down; `/analise/xlsx` baixa `analise.xlsx`; nenhum template nem o README
  contém a palavra "Recorte" fora do histórico das specs.
- **Ajuda**: `/ajuda` abre para os quatro perfis; seções omitidas conforme perfil; "Sua conta" só com
  login; todas as âncoras do sumário existem na página; toda tela com `menu.AJUDA` mapeado exibe o botão
  "Ajuda desta tela" com âncora existente (teste percorre `app.url_map` com GETs simples e confere o
  `href`).
- Os testes existentes que chamam `/recorte` passam a chamar `/analise` (ou seguem o redirect).
