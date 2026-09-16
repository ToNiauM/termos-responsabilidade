# Validação — UX/UI dos Cadastros

Data: 16/09/2026. Implementação local do plano aprovado em conversa.

## Entrega

- Quatro áreas com busca visível, filtros específicos, ordenação e paginação no servidor.
- Criação e edição em formulários próprios, com erros por campo e preservação dos valores.
- Retorno à busca e à página de origem; foco no registro após cancelar/editar. Registro salvo fora do filtro tem acesso explícito.
- Lista pesquisável de pessoas e manutenção dos bens em uma tela própria.
- Revisão antes de transferir localizações, atribuir patrimônios, remover vínculos, excluir registros ou substituir/encerrar processos vigentes.
- Revisões vinculadas aos dados apresentados, válidas por 30 minutos; mudanças no estado exigem nova conferência. Revalidação e gravação da confirmação ocorrem sob trava de escrita.
- Alteração da sigla do centro e das referências dos termos em uma única transação.
- Endpoints existentes preservados, sem alteração de esquema, dependências ou dados de produção. Rotas de cadastro separadas em `app_cadastros.py`.

## Testes automatizados

Comando: `.venv/bin/python -m pytest -q`.

Resultado: **164 testes passaram**. Inclui os testes existentes de cadastro, termos, documentos, importação, painel e inventário, mais 18 cenários específicos de UX e integridade.

Os novos cenários cobrem busca antes da paginação, acentos, parâmetros inválidos, dados preservados em erros, atomicidade da renomeação, manutenção dos filtros, destino de retorno restrito às rotas locais, duplicidade de pessoas, pendências conforme a guarda individual, confirmação sem gravação antecipada, estado alterado entre revisão e confirmação, troca de patrimônio na requisição, preservação dos bens e do histórico, seleção em lote mantida após erro e acesso a registro que saiu do filtro.

`git diff --check`: sem erros.

## Navegador

Chromium headless, servidor local Flask e SQLite temporário com dados fictícios. Nenhuma operação foi executada sobre a base de produção.

- Inspeção das quatro áreas em 1366 px e 390 px: sem transbordamento horizontal da página, um botão principal por tela e ausência de IDs duplicados.
- Capturas de tela inspecionadas para listagem de centros, localizações em celular e formulário de edição.
- Formulário inválido manteve a sigla e direcionou o foco ao resumo do erro.
- Fluxo completo: página 2 da lista de pessoas → editar → cancelar → mesma página e foco na linha.
- Fluxo completo: selecionar localizações → escolher destino no componente DSGov → revisar → confirmar → mensagem de sucesso.
- Fluxo completo: consultar patrimônio → conferir descrição e origem → confirmar atribuição → pessoa com os bens atualizados.
- Fluxo completo: cadastrar processo vigente → encerrar → conferir impacto → confirmar.

A passagem final desses fluxos terminou sem exceções ou avisos de console. Na primeira inspeção das listagens em celular apareceram avisos capturados pelo inicializador compartilhado de dropdown durante a compactação do breadcrumb; esse código compartilhado não foi alterado. Não é uma certificação completa de acessibilidade por leitor de tela.

O teste de navegador encontrou e permitiu corrigir uma integração específica: o componente `BRCard` substitui o ID do elemento em que é inicializado. O formulário de transferência passou a ficar dentro do cartão, preservando seu ID e a associação dos checkboxes da tabela. A transferência foi repetida e confirmada com sucesso após a correção.

## DSGov

Comando: `python3 /home/ToNiauM/.codex/skills/dsgov/scripts/verificar.py . --json`.

| Resultado | Antes | Depois |
| --- | --- | --- |
| Erros | 41 | 40 |
| Avisos | 4 | 3 |
| Erros novos | — | 0 |

Os templates novos não apresentam erros no verificador. O resultado global continua com saída 1 por apontamentos já existentes, incluindo a estrutura Flask do layout compartilhado, arquivos estáticos e outras telas. Isso não equivale a conformidade global com a skill DSGov. A comparação utilizou arquivo, regra e mensagem, ignorando mudanças de número de linha.

## Situação de publicação

Código e documentação preparados localmente. Não houve rebuild/restart do container de produção, publicação nem alteração da base real. A publicação permanece uma etapa separada, conforme o plano aprovado.
