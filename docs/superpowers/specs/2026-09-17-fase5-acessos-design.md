# Fase 5A — funções cumulativas e menor privilégio

**Data:** 2026-09-17.
**Estado:** comportamento definido com o usuário no grill; desenho técnico consolidado para planejamento.
**Base:** `main`, commit `c0dd564`, após a fase 4.
**Decisões:** `../notes/2026-09-17-fase5-grill.md`, seções 1–10.
**Plano:** `../plans/2026-09-17-fase5a-acessos.md`.

## 1. Resultado e limites

O administrador escolhe várias funções na criação e edição do usuário. Os acessos
se somam; função ausente não é concedida por conveniência de navegação. Todas as
rotas aplicam a mesma política que menus, botões, atalhos e Ajuda.

São cinco funções fixas. Não entra editor de funções personalizadas, autorização
por sala, concessão de consulta por evento, login externo ou alteração de senhas.
Manter Flask, SQLite, Jinja e os componentes DSGov existentes.

## 2. Matriz normativa

| Área/ação | Administrador | Operador | Consulta | Inventário | Consulta de inventários |
|---|---|---|---|---|---|
| Início geral, pesquisa, ficha, Análise e exportação da Análise | sim | sim | sim | não | não |
| Ver termos e histórico de termos | sim | sim | sim | não | não |
| Emitir termos, registrar SEI e e-mail | sim | sim | não | não | não |
| Manter cadastros, atribuições, textos, atualizar bens | sim | sim | não | não | não |
| Excluir cadastros/importar cadastros completos | sim | não | não | não | não |
| Lista de eventos, salas e pendências | todos | não | não | sua comissão | todos, só leitura |
| Conferir, registrar/editar/desmarcar leitura, fotos e sobras | sua comissão | não | não | sua comissão | não |
| Painel, relatório e exportação de inventário | todos | não | não | não | todos |
| Abrir/encerrar/excluir eventos e definir comissão | sim | não | não | não | não |
| Administrar usuários e funções | sim | não | não | não | não |
| Própria senha, sair e Ajuda pertinente | sim | sim | sim | sim | sim |

Operador é o pacote mantido pelo usuário. As consultas do acervo necessárias à
operação pertencem a esse pacote; ele não recebe Inventário automaticamente.
Consulta de inventários cobre eventos anteriores, atuais e futuros, sem exigir
participação na comissão. Inventário + Consulta não libera relatórios de inventário.
Inventário + Consulta de inventários libera consulta global e conferência somente
nos eventos da própria comissão. Eventos encerrados não aceitam alterações por ninguém.
Administrador mantém acesso global às rotas conhecidas, mas também precisa integrar
a comissão para conferir bens. Rotas não cadastradas são negadas inclusive a ele.

## 3. Dados e migração

`usuarios_funcoes(usuario_id, funcao)` tem chave composta, FK para `usuarios.id` e
CHECK para `admin`, `operador`, `consulta`, `inventariante`, `consulta_inventarios`.
O identificador interno `inventariante` permanece para reduzir renomeações; o rótulo
da função é **Inventário**. Não há função pré-marcada na criação: exigir ao menos uma.

Migrar o campo antigo `usuarios.perfil` para exatamente uma função de mesmo código,
sem ampliar a concessão; remover esse campo após a migração. Guardar um marcador de
migração e executar a conversão em transação. Reinicializar o sistema não recria
funções que tenham sido removidas. Manter IDs, hashes de senha, e-mail, atividade e
demais dados das contas. A instalação nova já nasce no esquema normalizado.

`usuarios.criar` e `usuarios.editar` passam a receber uma coleção `funcoes`;
`por_id`, `por_login`, `por_email` e `listar` devolvem `funcoes` em ordem estável.
Listagem filtra pela presença de uma função, sem duplicar usuários com várias.
Criar/editar conta e substituir suas funções ocorre em uma única transação.
Remover `admin` conta como rebaixamento mesmo se outras funções permanecerem.
Preservar a proteção contra auto-rebaixamento e contra perda do último admin ativo;
a CLI `criar-admin` continua funcionando. Não guardar funções no cookie: buscar no
banco a cada requisição, para revogação valer já na próxima chamada.

## 4. Identidade e comissão

Criar `inventario_comissao_usuarios(evento_id, usuario_id, nome_na_comissao)`.
O par evento/usuário é a chave; o nome é uma fotografia textual, não uma credencial.
Manter `inventario_integrantes` e os nomes nas leituras/sobras como registros legados
e de intercâmbio. A autorização web exige ID vinculado e função atual de conferência.

Na migração, vincular nomes antigos apenas quando existir exatamente uma conta com
esse nome, ativa e com `admin` ou `inventariante`. Homônimos, contas inelegíveis e
nomes sem conta ficam sem vínculo de autorização. Não inferir permissão pelo nome
em acessos posteriores. Mostrar ao administrador os nomes sem vínculo na comissão
do evento aberto para que selecione a conta correta. Em eventos encerrados, não
alterar o registro histórico; a função Consulta de inventários permite a consulta.

Os formulários de comissão enviam IDs; mostrar nome + login para distinguir pessoas.
Somente usuários ativos com Inventário ou Administrador são elegíveis. Operador
precisa receber Inventário para atuar na comissão. Inclusão/remoção e checagem de
eventos usam o vínculo por ID. Renomear uma conta não muda quem tem acesso nem a
autoria antiga. O programa local continua com o Administrador local e a regra
de comissão existente; esse caso é explícito, sem criar conta fictícia no banco.

Importar uma planilha de cadastros com inventário não importa funções de usuários.
Ao substituir eventos, preservar somente vínculos locais cujo ID de evento, nome
e data de abertura identificam o mesmo evento e cujo nome de comissão persistiu.
Evento novo/importado não autoriza contas pelo nome: o admin define a comissão
do evento aberto. Planilha sem abas de inventário preserva todos os vínculos.
Excluir evento remove os vínculos por FK. Exportação histórica mantém os nomes.

## 5. Política e integração

Criar `permissoes.py` para a matriz pura por `(endpoint, método)`, sem banco nem Flask;
`comissoes.py` concentra autorização por evento e concessão por ID. `usuarios.py`
mantém autenticação e funções do usuário. Rotas continuam nos módulos atuais.

O servidor verifica nesta ordem: sessão válida, troca obrigatória de senha, função
da rota e, para inventário, escopo do evento. Requisição recusada não lê dados do
evento nem chama storage, gera documento ou modifica banco. Para conferência, exigir
evento aberto no serviço, como hoje; falta de função/vínculo devolve 403 e evento
encerrado conserva o erro de negócio. JSON recebe o envelope `{"erro": ...}`.

Registrar GET/POST explicitamente. HEAD usa a permissão do GET e conserva as
proteções existentes contra registrar emissão. OPTIONS usa os métodos declarados
da rota sem autorizar mutações. Não usar fallback permissivo do GET para POST.

`/` permanece entrada compatível: admin, operador e consulta veem o Início geral;
usuário somente de inventário é redirecionado para `/inventario` antes de qualquer
consulta do painel. Login, troca de senha e breadcrumb usam o mesmo destino inicial.
`proximo` só é respeitado para GET interno autorizado, incluindo o escopo de evento;
destino inacessível cai na entrada permitida, sem loop. URL direta proibida retorna 403.

O inventariante sem evento atribuído recebe “Nenhum inventário atribuído a você”,
sem nome, contagem ou existência de eventos alheios. Retirar pesquisa geral e links
do acervo desse usuário. No inventário, controles de escrita exigem função, comissão
e evento aberto; apenas estar na comissão não exibe ações para quem perdeu a função.

A ficha geral do bem hoje expõe fotos e nomes dos inventários. Filtrar esses grupos
pela mesma autorização de evento: Consulta/Operador sozinhos não recebem conteúdo
de inventário; a combinação com Inventário recebe apenas os eventos vinculados;
Consulta de inventários/Admin recebem todos. Esta fase controla a entrega dos links
pelo aplicativo; não altera o serviço de armazenamento nem invalida URLs de fotos
que já tenham sido compartilhadas.

## 6. Critérios de aceitação

- Cinco funções isoladas e suas combinações produzem exatamente a matriz.
- Inventário sozinho não acessa acervo, termos, gráficos, relatórios ou exportações,
  inclusive por GET/POST/HEAD direto e `proximo`; `/` redireciona sem consultar o painel.
- Conferência exige função atual, vínculo por ID e evento aberto.
- Homônimo não vinculado não pode ler, alterar, fotografar, desmarcar ou excluir sobra.
- Remover função ou comissão revoga o acesso com sessão já aberta na próxima chamada.
- Consulta de inventários vê todos os eventos, inclusive criados depois da concessão,
  sem poder conferir; Consulta comum não alcança o módulo por conta própria.
- Migração é repetível sem ampliar funções nem religar comissões removidas.
- Importação, exclusão e renomeação mantêm os limites de autorização e o histórico.
- Menus e formulários não apresentam ações que as rotas recusariam.
- CLI de admin, modo local, CSRF, login por e-mail e proteção do último admin continuam.
