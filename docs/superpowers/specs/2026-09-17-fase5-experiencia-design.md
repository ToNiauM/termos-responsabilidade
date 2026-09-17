# Fase 5B–5C — navegação, Ajuda e Análise

**Data:** 2026-09-17.
**Estado:** comportamento consolidado após o grill; implementação não iniciada.
**Depende de:** `2026-09-17-fase5-acessos-design.md`.
**Substitui:** `2026-09-17-fase5-menu-inicio-analise-ajuda-design.md`.
**Planos, na ordem de execução:** `../plans/2026-09-17-fase5b-analise.md` e
`../plans/2026-09-17-fase5c-navegacao-ajuda.md`.

## 1. Menu e entrada

| Item | Ícone | Filhos |
|---|---|---|
| Início | fa-home | nenhum |
| Termos de Responsabilidade | fa-file-signature | Por centro, individual, devolução, emitidos |
| Análise | fa-chart-bar | nenhum |
| Inventário | fa-clipboard-check | Eventos, evento aberto acessível, Painel e Relatório permitidos |
| Cadastros | fa-address-book | Centros, Localizações, Pessoas, Processos SEI |
| Textos | fa-pen-nib | nenhum |
| Atualizar base | fa-upload | nenhum |
| Usuários | fa-users | nenhum; apenas admin, com login ligado |
| Ajuda | fa-question-circle | nenhum |

Filtrar filhos e itens pela matriz efetiva; grupo vazio desaparece. Inventário no
menu nunca divulga evento aberto de outra comissão ao usuário que só pode conferir.
Os links de eventos encerrados estão na lista de eventos, sob a mesma autorização.
Início e Análise ficam ausentes para quem tem apenas funções de inventário.

`menu.montar(funcoes, evento_aberto, endpoint_atual, argumentos, login_ativo)` recebe
um evento já autorizado ou `None`. Usa endpoints e argumentos, não prefixos de URL:
quatro cadastros compartilham um endpoint e diferentes eventos usam o mesmo endpoint.
Grupo da tela atual abre; detalhe/edição abrem o grupo correspondente. Só o destino
correto é marcado com `active` e `aria-current="page"`. Um relatório de evento
encerrado não marca como atual o link do relatório do evento aberto.

Manter título expansível como `<a href="javascript:void(0)">`, marcação `br-menu`,
ícones e estilos existentes. O core local DSGov em `_setDropMenu` força
`aria-expanded=false`; `dsgov.js` também marca por prefixo. Um pequeno script próprio
`static/js/menu-estado.js`, carregado após esses arquivos, restaura a seleção do
servidor e sincroniza `aria-expanded` com `.menu-folder.active` na inicialização.
Não duplicar o mecanismo de abrir/fechar, modificar vendor ou criar CSS novo.

## 2. Início

O Início geral só é calculado e renderizado para quem pode consultar o acervo.
Inventário/Consulta de inventários sozinhos entram na lista de eventos.

Três blocos: acesso rápido, indicadores permitidos, “Sobre o patrimônio”.

- Até seis atalhos: termo por centro, individual, devolução, emitidos, realizar
  inventário e Análise. Filtrar pela ação: “Realizar inventário” exige permissão de
  conferência; quem só consulta inventários usa o menu de consulta.
- Manter contagens de ativos, valor sem imóveis, imóveis, sem centro nem pessoa,
  termos pendentes, última importação e andamento do inventário quando autorizado.
  Não exibir indicador administrativo sem acesso ao destino; consulta comum não vê
  histórico de importação nem andamento de inventário sem função correspondente.
- Remover busca do conteúdo, gráficos e tabela “Termos por centro”. Busca do
  cabeçalho aparece somente para quem tem acesso à pesquisa geral.
- `db.painel` não calcula dimensões; `home` não monta gráficos e não carrega ECharts.
- Aviso de **valor não informado** refere-se somente a NULL. Mostrar separadamente
  a informação de valor zero, em texto neutro com seu filtro, quando existir.

Texto do card:

> **Patrimônio do CFC**
>
> Este sistema reúne o cadastro dos bens do Conselho Federal de Contabilidade, os
> responsáveis por sua guarda, os termos de responsabilidade e o inventário.
>
> O Termo de Responsabilidade registra a guarda dos bens por um centro de custo ou
> por uma pessoa. O Termo de Devolução registra a devolução dos bens. Os documentos
> são preparados aqui para uso no SEI.
>
> Use os atalhos para suas atividades e a pesquisa do cabeçalho para localizar bens,
> pessoas ou centros de custo. O guia explica as funções disponíveis para você.

Botão secundário **Ver guia**, destino `/ajuda`.

## 3. Análise, URLs e documentos

`GET /analise` e `/analise/xlsx` substituem a tela Recorte. As rotas antigas ficam
como aliases 301 com query string integral, inclusive parâmetros vazios e repetidos.
Aplicar as mesmas permissões nos aliases, sem caminho alternativo de acesso.
Helpers `painel.url_recorte*` mantêm os nomes internos e geram apenas URLs novas.
`db.recorte` e `db.exportar_recorte` permanecem conceitos internos.

Manter filtros, gráficos por dimensão, tabelas dos gráficos e aviso de limite de
1.000 bens. A exportação contém todos os bens filtrados, arquivo `analise.xlsx` e
aba `analise`. Sem filtros, situação é ATIVO; `situacao=` significa todas e deve
permanecer nos links. Contagens/valores nunca usam apenas as 1.000 linhas da tela.

Seis indicadores: bens no recorte; valor atual; imóveis (quantidade e valor); sem
centro nem pessoa; **Valor não informado**; **Valor zero**. Uma agregação sobre o
mesmo WHERE alimenta tudo, com zeros em resultado vazio.

- Não informado: `valor_atual IS NULL`.
- Zero: `valor_atual = 0`, sem NULL ou negativos.
- Adicionar filtro `valor_status` com opções vazio, `nao_informado`, `zero`. Nenhuma
  limpeza/conversão dos valores importados faz parte desta fase.
- Filtros numéricos passam a comparar valor real: NULL não satisfaz “até R$ 0”.
  Combinar “não informado” com intervalo numérico produz resultado vazio, sem
  alterar silenciosamente o intervalo. Os filtros são visíveis e removíveis.
- Tabela mostra **Não informado** para NULL e moeda para valores numéricos, inclusive
  R$ 0,00. Excel mantém célula vazia para NULL e número 0 para zero.

Clique em indicador acrescenta uma restrição sem trocar filtros conflitantes.
Imóveis só oferece clique se ainda não há classificação selecionada. Sem centro
nem pessoa só oferece clique se ambos não estão fixados, ou um já está em `-` e
o outro está livre. Valor zero/ausente só oferece clique se `valor_status` estiver
livre e a contagem for positiva. Cartão sem resultado ou sem refinamento possível
é texto, sem link; demais filtros são preservados. Assim o clique não aumenta o
universo nem transforma um zero do indicador em uma lista não vazia.

Para centro/pessoa existente, manter acesso **Abrir termo completo**; exibir ao lado
o contexto do titular e “Os filtros desta análise não limitam o termo”. Não emitir
termo parcial. Consulta pode abrir o documento; somente Operador/Admin podem emitir.
A prévia não registra emissão; copiar/baixar conserva as regras atuais.

## 4. Ajuda e títulos

`GET /ajuda` disponível às cinco funções. Sumário e seções resultam das mesmas
capacidades; esconder também parágrafos de ação não concedida dentro de uma seção.
Exemplo: Consulta vê como abrir termos, mas não recebe instruções de emitir; quem
tem Inventário vê conferência, sem instruções de baixar relatórios; admin vê tudo.

Seções: `inicio`, `pesquisa`, `termos`, `analise`, `inventario`,
`consulta-inventarios`, `cadastros`, `textos`, `atualizar-base`, `usuarios`, `conta`,
`perguntas`. `conta` exige login ligado e usuário real. Perguntas frequentes são
filtradas por tema; nenhuma âncora aponta para seção ausente. Não há busca, editor,
tour nem imagens na Ajuda.

Conteúdo factual a cobrir, adaptado à permissão:

| Seção | Conteúdo |
|---|---|
| Início | Atalhos e indicadores autorizados, diferença entre NULL e zero, guia |
| Pesquisa | Número de patrimônio, palavras, curingas, ficha e histórico autorizado |
| Termos | Centro/pessoa/devolução; processo vigente; prévia; emissão e histórico somente para quem pode executar |
| Análise | Filtros, indicadores, gráficos, Excel completo, termo completo independente dos filtros |
| Inventário | Comissão; escolher sala; ler plaqueta; localizado/divergente/pendente; fotos, conservação, sobras; encerrado só leitura |
| Consulta de inventários | Todos os eventos; painel, relatório e filtros; Excel com/sem fotos; acesso de consulta não permite conferir |
| Cadastros | Centros, localizações, pessoas, processos; atribuições; exclusões/importação integral só admin |
| Textos | Editar modelos e marcadores; restaurar padrão |
| Atualizar base | Importar export do patrimônio; bens substituídos, demais cadastros mantidos; histórico e exportação |
| Usuários | Cinco funções combináveis; menor privilégio; ativação, senha temporária; login ou e-mail opcional |
| Conta | Trocar senha, sair, troca obrigatória e duração da sessão |
| Perguntas | Termo desatualizado; bem individual; divergente/pendente; leitura na sala errada; foto; evento encerrado |

Reaproveitar o texto da spec histórica somente após corrigir frases falsas:
“todo card vai para Análise”, “tudo é clicável”, “zero é sem cadastro”, “quatro
perfis” e instruções que a função não pode executar. A referência ao antigo nome
Recorte é permitida exclusivamente na FAQ “Onde está o Recorte?” e nos redirects.

Criar macro **ajuda_titulo(ancora)** junto ao `<h1>` existente, preservando estrutura,
ações e subtítulos. A substituição em massa do cabeçalho por uma macro única da spec
histórica fica descartada: os templates têm cabeçalhos diferentes. Rotas de login,
erros e documentos incorporados não ganham ajuda contextual. A página Ajuda não
aponta para si mesma. Cada tela mapeada recebe uma âncora realmente renderizada.

## 5. Aceitação

Validar cada função isolada, combinações e modo local. Menu ativo correto inclusive
cadastro específico e evento encerrado; sem links negados ou nomes de evento fora
do escopo. Início sem gráficos/busca central e sem conteúdo de módulo não concedido.
Ajuda com âncoras válidas e instruções de ação filtradas. Separação NULL/zero/negativo,
recorte vazio, intervalo numérico, situações todas/ATIVO, preservação de filtros,
contagens além de 1.000 e Excel coerente. Termo completo claramente identificado.
Verificação de navegador: desktop e celular, abertura/fechamento e acessibilidade
do menu após inicialização do core, navegação por teclado e leitura do contexto.
