# Administração de inventários — chave aberto/fechado e painel do administrador

**Data:** 2026-09-18.
**Estado:** implementado e publicado em 2026-09-20 (noite); ver §10.
**Base:** `main`, commit `724bee1`, depois do robô do SPW.
**Plano:** `../plans/2026-09-18-administracao-inventarios.md`.

## 1. Resultado e limites

O administrador ganha uma tela **Administração**, só dele, com duas abas: **Inventários** (criar, definir comissão,
abrir/fechar, finalizar, excluir) e **Usuários** (a tela atual, sem mudanças). Cada inventário não finalizado tem
uma única chave **Abrir/Fechar**: ligar a chave em um fecha o que estava aberto. A tela Inventário passa a servir
só à conferência pela comissão.

Não entra: vários inventários abertos ao mesmo tempo, salas por integrante, mudança no fluxo de leitura, fotos ou
relatório, editor de funções fora do cadastro de usuários. Manter Flask, SQLite, Jinja e os componentes DSGov.
O modo desktop (sem login, comissão por nomes) continua funcionando com as mesmas telas.

## 2. Decisões do usuário (2026-09-18)

| Pergunta | Decisão |
|---|---|
| "Definir atribuições" no painel | Funções dos usuários: quem entra na comissão sem a função Inventário recebe a função na hora |
| Quantos abertos por vez | Um. A chave de um fecha os demais automaticamente, sem mensagem de erro |
| Onde fica o painel | Item "Administração" no menu, só admin, com abas Inventários e Usuários |
| Estados | Aberto, fechado (suspenso, reabrível) e finalizado (o encerrar de hoje, permanente) |

## 3. Estados

Coluna nova `suspenso_em TEXT` em `inventario_eventos`; `aberto_em` e `encerrado_em` ficam como estão.

| Estado | No banco | Quem faz o quê |
|---|---|---|
| **aberto** | `encerrado_em IS NULL AND suspenso_em IS NULL` | comissão lê, fotografa, registra sobras; no máximo um |
| **fechado** | `encerrado_em IS NULL AND suspenso_em NOT NULL` | ninguém lê; comissão e consulta veem salas, painel e relatório; admin pode reabrir |
| **finalizado** | `encerrado_em NOT NULL` | permanente; congela os bens (`inventario_bens_encerrados`) como hoje |

Transições, todas do administrador, cada uma numa transação:

- **Criar**: nasce fechado (`suspenso_em = agora`). O formulário tem a caixa "abrir agora", marcada por padrão
  quando não há nenhum aberto; marcada, a criação já liga a chave.
- **Ligar a chave (abrir)**: `UPDATE ... SET suspenso_em = NULL WHERE id = ?` e, na mesma transação,
  `UPDATE ... SET suspenso_em = agora WHERE encerrado_em IS NULL AND id <> ?`. Ligar num finalizado → erro.
- **Desligar a chave (fechar)**: `suspenso_em = agora`. Fechar um já fechado não faz nada.
- **Finalizar**: de aberto ou fechado; é o `encerrar_evento` de hoje, que também zera `suspenso_em`.
  Finalizado não tem chave nem volta.
- **Excluir**: como hoje, em qualquer estado, com a confirmação pelo nome e apagando as fotos.

`_evento_aberto_ou_erro` passa a recusar também o fechado, com "Evento fechado: não aceita leituras até ser
reaberto." Todas as rotas de leitura, foto e sobra já passam por ela. A edição de comissão usa uma checagem mais
branda, `_evento_nao_finalizado_ou_erro`, porque a comissão pode mudar com o evento fechado.

`evento_aberto(conn)` continua devolvendo só o aberto (nenhum ou um). Função nova `evento_corrente(conn)`:
o aberto ou, se não houver, o fechado mais recente por `aberto_em`; é o que o card do Início e o menu mostram.

## 4. Tela Administração (`/administracao` e `/usuarios`, com abas)

Item de menu "Administração" (ícone `fa-cogs`), no lugar de "Usuários", visível só ao admin e só com login ativo
(no desktop o item aparece sem a aba Usuários). Abas no padrão de `cadastros.html`.

**Aba Inventários**

- Formulário "Novo inventário": nome, descrição, comissão (lista de usuários ativos; no desktop, nomes elegíveis
  como hoje), escopo de salas (todas ou escolher) e a caixa "abrir agora".
- Tabela "Inventários" com todos: nome, estado como tag (`bg-success` aberto, `bg-warning` fechado,
  `bg-gray-20` finalizado), aberto em, finalizado em, comissão, leituras/salas conferidas (`resumo`), e ações:
  - aberto: **Fechar**, **Finalizar**, **Comissão**, **Excluir**;
  - fechado: **Abrir**, **Finalizar**, **Comissão**, **Excluir**;
  - finalizado: **Relatório**, **Excluir**.
  Finalizar e Excluir mantêm as confirmações de hoje (Excluir pede o nome). Abrir/Fechar são um POST cada, sem
  confirmação, com flash "Inventário X aberto; Y foi fechado." quando houver troca.
- **Comissão** abre a tela `inventario_comissao.html` atual, agora sob Administração, e funciona também com o
  evento fechado (só finalizado recusa).

**Aba Usuários**: a própria tela `/usuarios` de hoje, com a barra de abas no topo e o botão "Novo usuário". As rotas
`usuarios.*` não mudam de nome nem de URL; a aba Inventários é `/administracao`.

## 5. Tela Inventário, Início e menu

- `inventario_eventos.html` perde o formulário "Abrir evento" e a tabela passa a listar os eventos visíveis com o
  estado. O admin vê o botão "Administrar inventários" apontando para a aba. `inventario_evento.html` perde os
  botões Encerrar, Comissão e Excluir (o admin vê "Administrar" no lugar) e, no evento fechado, mostra a mensagem
  "Inventário fechado: leitura suspensa" no topo, sem os controles de leitura.
- Card do Início: usa `evento_corrente`; título "Inventário em andamento" (aberto) ou "Inventário fechado".
- Menu lateral: o grupo Inventário mostra o evento corrente (aberto ou fechado) como hoje mostra o aberto.
- Ajuda: a seção Usuários vira "Administração" (mesmo id) e descreve a chave, finalizar e comissão.

## 6. Comissão e funções

No modo web, a lista de comissão mostra todos os usuários **ativos**, não só os elegíveis. Ao salvar (criar ou
editar comissão), quem não tem `inventariante` nem `admin` recebe `inventariante` na mesma transação
(`usuarios_funcoes`), e o flash informa "Função Inventário concedida a: Fulano, Beltrana." Remover da comissão não
tira a função. Usuário inativo nunca entra. No desktop nada muda (nomes, sem funções).

## 7. Código

- `db.py`: `suspenso_em TEXT` no esquema e `ALTER TABLE ... ADD COLUMN` em `criar_esquema` quando faltar.
- `inventario.py`: `estado(e) -> 'aberto'|'fechado'|'finalizado'`; `criar_evento(...)` (o `abrir_evento` de hoje
  sem a recusa por evento aberto, gravando `suspenso_em`; `abrir_evento` vira alias que cria e liga a chave, para
  os testes atuais); `ligar_chave(conn, id)`, `desligar_chave(conn, id)`; `encerrar_evento` zera `suspenso_em`;
  `_evento_aberto_ou_erro` recusa fechado; `evento_corrente`; `eventos()` ordena aberto, fechados, finalizados.
- `comissoes.py`: `abrir` vira `criar(conn, nome, descricao, ids, salas, abrir_agora)`; `definir` e `criar` chamam
  `usuarios.conceder_funcao(conn, ids, "inventariante")` novo, que devolve os nomes que ganharam a função;
  `_usuarios_selecionados` aceita qualquer ativo.
- `app_admin.py` novo, blueprint `admin`: só `tela` (GET `/administracao`, aba Inventários). As rotas administrativas
  existentes ficam em `app_inventario.py` com nome e URL de hoje (`inventario.abrir` = criar, com o campo
  `abrir_agora`; `inventario.encerrar` = finalizar; `inventario.comissao`; `inventario.excluir`) e ganham
  `inventario.abrir_chave` (POST `/inventario/<id>/abrir`) e `inventario.fechar` (POST `/inventario/<id>/fechar`);
  todas redirecionam para `admin.tela`, e `inventario_comissao.html`/`inventario_excluir.html` ganham trilha
  "Administração › ..." e Cancelar voltando à Administração. Isso evita reescrever os testes que usam essas URLs.
- `app_usuarios.py`: `lista` continua renderizando `usuarios/lista.html`, que ganha a barra de abas
  (`_abas_admin.html`) e a trilha "Administração › Usuários".
- `permissoes.py`: `admin.tela`, `inventario.abrir_chave` e `inventario.fechar` em ADMIN; as linhas antigas ficam.
- `menu.py`: item "Administração" → `admin.tela` (nos dois modos); `destino_atual` mapeia `usuarios.*`,
  `inventario.comissao` e `inventario.excluir` para `admin.tela`; a seção da Ajuda mantém o id `usuarios` com o
  título "Administração" e passa a existir também no desktop.
- Exportar/importar cadastros (`inventario.exportar_abas`/`validar_abas`/`substituir_tabelas`): a aba `inv_eventos`
  ganha a coluna `suspenso_em` (opcional; planilhas antigas sem a coluna importam com `NULL`). A validação "mais
  de um evento aberto" passa a contar só os sem `encerrado_em` e sem `suspenso_em`.

## 8. Testes

- Estados: criar nasce fechado; criar com "abrir agora" fecha o que estava aberto; ligar a chave em B fecha A na
  mesma transação; ligar em finalizado → erro; fechar e reabrir mantêm leituras; finalizar fechado congela os
  bens; `evento_corrente` prefere o aberto e cai no fechado mais recente.
- Leitura bloqueada em evento fechado (`ler`, `lote`, `foto`, `sobra`, `atualizar_leitura`) com a mensagem nova.
- Permissões: `/administracao` e cada ação para as cinco funções (só admin passa); `/usuarios` redireciona;
  inventariante não vê "Administrar"; consulta de inventários vê o fechado no relatório.
- Comissão: usuário sem função recebe `inventariante` e o flash lista o nome; inativo é recusado; desktop continua
  por nomes.
- Telas: aba Inventários lista os três estados com as ações certas; card do Início com "Inventário fechado"; menu
  com o evento corrente; Ajuda com a seção nova.
- Cadastros: exportar traz `suspenso_em`; importar planilha antiga (sem a coluna) funciona; dois abertos →
  rejeitado; um aberto e um fechado → aceito.
- Migração: banco antigo sem `suspenso_em` ganha a coluna; evento aberto existente continua aberto.

## 9. Fora do escopo

Salas por integrante, vários abertos, histórico de quem abriu/fechou (fica nas datas), notificação à comissão.

## 10. Evidências

Publicado em 2026-09-20 ~21:10, main=`e806441` (merge ff do ramo `admin-inventarios`, 8 commits, apagado depois). Base do
ramo: `76e87ac` (o plano foi escrito sobre `724bee1`; as diferenças — SEI por usuário, Meus acessos, permissões — foram
reconciliadas nos despachos).

- Commits: `1b89896` estados e chave · `a1ae180` comissão aceita qualquer ativo e concede a função · `e84fd6c` planilha com
  `suspenso_em` · `850f750` + `6261f97` tela Administração, rotas da chave, permissões e menu · `30c10ce` tela Inventário só
  para conferência, Início e Ajuda · `4e1e038` README · `e806441` correções da revisão final (teste de amostragem por tela,
  estado na tela de excluir, testes desktop e de atomicidade).
- Suíte: `.venv/bin/pytest -q` → 3496 passed (antes: 3385).
- Publicação: `docker compose up -d --build`; `db._colunas(conn, 'inventario_eventos')` no container inclui `suspenso_em`;
  "Inventário 2026" (id 1) continua **aberto** (`suspenso_em` e `encerrado_em` nulos) e "Inventário Setorial 2026"
  finalizado; `/login` 200; `/administracao` sem sessão → 302; trabalhador reiniciado 21:10.
- Desvios conscientes em relação ao texto da spec (implementação prevalece): mensagem do evento fechado é "Inventário
  fechado. A leitura está suspensa até o administrador reabrir; consulta liberada." (§5 dizia "Inventário fechado: leitura
  suspensa"); colunas "Criado em" e "Localizados" (§4 dizia "aberto em" e "leituras/salas conferidas" — o evento agora
  nasce fechado, então "Criado em" é o correto); §8 "`/usuarios` redireciona" contradizia §4/§7 — `/usuarios` continua
  respondendo 200 com a barra de abas. O card e o menu usam a versão de `evento_corrente` recortada pela visibilidade
  (`app._evento_corrente_visivel`), não `inventario.evento_corrente` diretamente.
- Deferidos (revisão final): `ligar_chave` sem `BEGIN IMMEDIATE` (os dois UPDATEs caem na mesma transação; só o flash
  pode citar um "fechado" defasado sob dois admins simultâneos); helpers `_elegiveis/_local/_conn` duplicados entre
  `app_admin.py` e `app_inventario.py`; N+1 de `resumo` na lista de inventários; card do Início diz "Nenhum inventário
  aberto" quando só há finalizados; `consulta_inventarios` fora da matriz de `test_permissoes` (pré-existente).

Pendências do usuário: conferir no site o menu **Administração** (abas Inventários e Usuários) com o "Inventário 2026"
aberto e a chave; a comissão pode agora incluir qualquer usuário ativo.
