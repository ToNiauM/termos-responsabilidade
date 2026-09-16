# Plano de UX/UI — Cadastros

Data: 16/09/2026  
Estado: aprovado em conversa, implementado e publicado em produção em 16/09/2026.

## 1. Objetivo e escopo

Facilitar encontrar, cadastrar e alterar informações na área **Cadastros**, mantendo a identidade DSGov 3.7.0 e a aplicação Flask/SQLite existente.

Premissa da proposta: “alterar o sistema” significa manter os cadastros e seus vínculos pela interface. Abrange centros de custo/responsáveis, localizações, pessoas e processos SEI. A prioridade sugerida é a manutenção dos registros existentes.

O plano foi aprovado em conversa. A implementação mantém o esquema de dados existente; a publicação foi autorizada posteriormente e concluída. Evidências em `../notes/2026-09-16-ux-cadastros-validacao.md`.

## 2. Diagnóstico do código atual

Análise dos templates, rotas e regras de dados; ainda sem validação visual em navegador ou observação de usuários.

| Situação observada | Efeito provável na experiência |
| --- | --- |
| Formulários de inclusão ocupam o início das abas | Quem quer alterar um registro precisa passar primeiro pela inclusão |
| “Responsáveis” representa também os centros de custo | O nome da aba não explica todas as informações que ela mantém |
| Busca das tabelas fica atrás do ícone de lupa | Encontrar registros exige descobrir esse controle |
| Pessoas exige selecionar um nome e clicar em “Ver bens” para acessar sua manutenção | Editar uma pessoa não é uma ação imediatamente visível |
| Transferência de localizações fica abaixo da tabela | Seleção e ação podem ficar distantes em listas extensas |
| Erros de negócio redirecionam à página anterior | Os valores enviados não são reapresentados pelo servidor no formulário |
| Ações de exclusão, remoção de vínculo e mudança de vigência têm apresentações diferentes | As consequências nem sempre aparecem antes da ação |
| As quatro abas são renderizadas na mesma resposta | A manutenção da tela e do estado de navegação fica concentrada em um template |

Referências: `templates/cadastros.html`, `templates/editar_responsavel.html`, `templates/editar_pessoa.html`, `templates/_macros.html`, seção de cadastros de `app.py` e funções de manutenção em `db.py`.

## 3. Fluxo proposto

**Cadastros → escolher a área → buscar o registro → Editar → Salvar alterações → voltar ao registro na lista.**

Para inclusão: **Novo cadastro → preencher → Salvar → consultar o registro criado**, com acesso claro aos próximos vínculos necessários.

Para ações que alteram vínculos: **escolher registros → informar destino → revisar impacto → confirmar → ver resultado**.

Organização da área:

- **Centros de custo** — descrição curta: “Centros, responsáveis e localizações vinculadas”.
- **Localizações** — descrição: “Vincule as localizações da base aos centros de custo”.
- **Pessoas** — descrição: “Pessoas cadastradas e bens sob sua responsabilidade”.
- **Processos SEI** — descrição: “Processos usados na emissão dos termos”.

Manter a troca direta entre as quatro áreas, cada uma com URL própria. O link abre somente o conteúdo da área escolhida, com navegação compatível com Voltar/Avançar do navegador. Cadastros continua entrando diretamente em Centros de custo, sem uma página intermediária obrigatória.

### Estrutura das listagens

1. Título, explicação curta e ação principal contextual: “Novo centro de custo”, “Vincular localização”, “Nova pessoa” ou “Novo processo”.
2. Busca sempre visível, com rótulo específico, e filtros pertinentes à área.
3. Quantidade de resultados, tabela, ações identificáveis e mensagem quando não houver registros.
4. Paginação de 10/20/50 itens e ordenação nos cabeçalhos relevantes.

Busca e filtros atuam no conjunto completo, antes da paginação. Exportar cadastros continua sendo uma ação secundária de exportação completa, identificada como tal.

### Estrutura dos formulários

- Inclusão e edição em páginas próprias, reaproveitando os mesmos campos de cada cadastro.
- Uma ação principal: “Salvar” na inclusão e “Salvar alterações” na edição; “Cancelar” ao lado.
- Campos obrigatórios identificados, ajuda curta abaixo do campo e tipo de entrada adequado, como e-mail.
- Erro junto ao campo, resumo quando necessário e preservação de todos os valores preenchidos.
- Após salvar ou cancelar, retorno ao contexto de origem, preservando busca, filtros e página. O registro alterado deve ficar identificável; se deixar de corresponder ao filtro, a mensagem explica isso e oferece acesso ao registro.
- Navegação por teclado, foco visível, rótulos acessíveis e composição responsiva com os componentes existentes do DSGov.

## 4. Melhorias por área

### Centros de custo

- Lista focada em sigla, responsável, função e ações; demais dados disponíveis no formulário e na consulta do registro.
- Busca por sigla ou responsável e acesso direto a “Editar”.
- Formulário com identificação do centro e dados do responsável; ajuda específica para a mudança de sigla.
- Resumo de localizações vinculadas e atalho “Gerenciar localizações”, já filtrado pelo centro.
- Explicitar que a troca de sigla mantém os vínculos existentes.
- Exclusão em página de confirmação, informando as localizações afetadas. Quando houver bens ativos sob guarda do centro, mostrar o bloqueio e o caminho para reorganizar os vínculos.

### Localizações

- Lista com filtro de situação (“Todas”, “Sem centro de custo”, “Vinculadas”) e centro de custo, além da busca por localização.
- Mostrar as pendências com um atalho direto para resolvê-las; considerar pendentes as localizações com bens ativos sem centro, conforme a regra atual.
- Ação por linha “Vincular centro” ou “Alterar centro”, conforme a situação.
- Manter transferência em lote das localizações já vinculadas: selecionar, escolher destino e revisar uma relação de origem → destino antes da confirmação.
- Exibir a quantidade selecionada junto da ação. “Selecionar todas” afeta somente a página visível; ao trocar busca, filtros ou página, limpar a seleção e comunicar isso. A revisão identifica exatamente as localizações enviadas.
- Usar “Remover vínculo” para desfazer o mapeamento, com explicação de que essa operação não exclui o bem nem modifica sua localização na base importada.
- Localizações continuam vindo da base patrimonial; o fluxo cuida do vínculo ao centro.

### Pessoas

- Substituir a seleção inicial por uma lista pesquisável com nome, quantidade de bens e ações “Editar” e “Ver bens”.
- Cadastro e alteração de nome em formulário simples; informar que os bens atribuídos acompanham a mudança de nome.
- “Ver bens” abre a manutenção dos vínculos da pessoa: bens atuais, atribuição por patrimônio e acesso ao termo individual.
- Na atribuição, consultar o patrimônio e mostrar descrição e responsável atual antes de confirmar o novo vínculo. Se já estiver com outra pessoa, identificar origem e destino na confirmação.
- “Remover vínculo” informa que o bem deixa de estar atribuído à pessoa e passa a seguir a responsabilidade definida pela localização, quando houver centro vinculado.
- Exclusão explica a remoção das atribuições, preservando os bens da base.

### Processos SEI

- Busca por número ou descrição; filtros por tipo de termo e situação.
- Destacar quais processos estão vigentes e os tipos de termo sem processo vigente.
- Inclusão em formulário próprio, com explicação da opção de vigência.
- Antes de substituir o vigente, mostrar qual processo deixará de vigorar; antes de encerrar, explicar o efeito na emissão daquele tipo de termo.
- Confirmação de exclusão e manutenção do bloqueio para processos com termos registrados.
- O escopo cobre inclusão, consulta, vigência, encerramento e exclusão existentes. Edição do número ou tipo de processo já registrado exige uma definição própria sobre o histórico e não integra esta proposta.

## 5. Sequência de execução após aprovação

| Etapa | Entrega | Verificação |
| --- | --- | --- |
| 1. Estrutura e Centros de custo | Navegação, listagem com busca, formulário compartilhado e retorno ao contexto | Encontrar, editar, cancelar, salvar e corrigir erro sem perder valores |
| 2. Localizações | Filtros de pendências, alteração individual e transferência em lote com revisão | Conferir origem/destino, seleção paginada e efeitos nos vínculos |
| 3. Pessoas | Lista pesquisável, edição direta e atribuição com consulta do patrimônio | Renomear mantendo bens, transferir e remover atribuições |
| 4. Processos SEI | Listagem, inclusão e apresentação do impacto das mudanças de vigência | Um vigente por tipo, bloqueio de exclusão e preservação dos termos emitidos |
| 5. Validação e entrega | Ajustes de acessibilidade, responsividade, regressão e documentação | Percorrer os cenários abaixo e revisar visualmente as quatro áreas |

Cada etapa aplica o mesmo padrão visual e de navegação estabelecido na primeira. Não requer uma nova aprovação entre etapas, desde que o plano aprovado seja mantido.

## 6. Implementação prevista

- `templates/cadastros.html`: reduzir à estrutura da área e separar as telas por domínio em `templates/cadastros/`.
- `templates/editar_responsavel.html` e `templates/editar_pessoa.html`: integrar os formulários compartilhados de inclusão/edição.
- `templates/_macros.html`: estender componentes somente quando necessário, preservando os consumidores de outras telas.
- `app.py`: rotas de consulta/formulário, parâmetros de busca e paginação, respostas de validação com valores preenchidos e retornos ao contexto original. Preservar os endereços já usados por outras partes do sistema.
- `db.py`: consultas de listagem e contagem; reutilizar as regras existentes. Validar todas as entradas antes de qualquer gravação e tornar atômicas as operações compostas necessárias ao fluxo.
- JavaScript específico de Cadastros somente quando necessário para seleção, consulta de patrimônio e gerenciamento do foco; manter os arquivos de fornecedor e o CSS DSGov existentes.
- `tests/test_app.py` e `tests/test_db.py`: ampliar casos de comportamento e regressão. Usar base temporária de testes.

Não há migração de framework ou mudança de esquema de dados prevista. Painel, inventário, textos, importação em planilha e geração dos documentos permanecem fora do redesenho. Links para consulta e emissão podem ser reutilizados dentro dos novos fluxos.

## 7. Critérios de aceite

- A partir da lista de centros ou pessoas, é possível abrir diretamente a edição do registro encontrado.
- Salvar, cancelar e Voltar mantêm uma navegação previsível e o contexto da busca.
- Formulários inválidos preservam os dados digitados e indicam o que corrigir; nada é parcialmente gravado.
- Trocar a sigla do centro ou o nome da pessoa mantém os vínculos e as referências do histórico previstas pelas regras atuais.
- Pendências de localização podem ser encontradas e resolvidas pela própria área, sem exportar planilha.
- Transferência em lote mostra todos os itens selecionados, a origem e o destino antes de gravar; o servidor revalida os vínculos antes da confirmação.
- Atribuir um patrimônio permite identificar o bem e eventual responsável anterior.
- Exclusão, remoção de vínculo e encerramento usam textos distintos e explicam seus efeitos antes da confirmação.
- As regras de exclusão de centros e processos continuam sendo aplicadas no servidor.
- Busca e paginação não escondem resultados correspondentes fora da página atual.
- As quatro áreas funcionam por teclado e em telas estreitas, com rolagem da tabela contida em seu componente.
- Testes de regressão de cadastros e termos passam; revisão visual cobre listas vazias, erros, confirmações, nomes longos e listas com várias páginas.

Na execução, registrar o resultado de `.venv/bin/python -m pytest -q` e do verificador DSGov (`python3 /home/ToNiauM/.codex/skills/dsgov/scripts/verificar.py .`). Por ser uma aplicação Flask existente, verificar a compatibilidade do verificador e distinguir problemas preexistentes de violações introduzidas, sem declarar conformidade com verificações não realizadas.

## 8. Aprovação solicitada

Aprovar o redesenho das quatro áreas com prioridade em **encontrar e editar**, conforme os fluxos e limites descritos. A aprovação autoriza iniciar a implementação local e sua validação; publicação será tratada separadamente.
