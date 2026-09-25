# Números da pesquisa — item a item, com a origem de cada um

Registro de conferência dos números usados no TCC *"Automação de inventários no setor público com Data
Science e infraestrutura open source replicável"* (MBA Data Science e Analytics, USP/Esalq, 2026).
Posição dos dados: 22 set. 2026 (campanha), 24 set. 2026 (ensaio de carga) e 25 set. 2026 (carga efetiva no banco). Escrito para que qualquer
pessoa — banca, auditoria ou o próprio autor daqui a anos — consiga refazer a conta.

Origem de cada número:

| Sigla | Fonte | Como reproduzir |
|---|---|---|
| **R** | Relatório exportado do sistema de conferência em campo (`relatorio_20260922_211049.xlsx`, 7.428 linhas, uma por bem; não versionado porque traz nomes de servidores) | `recalcular_metricas.py` gera `metricas_campanha_2026.json` (versionado aqui) e as figuras |
| **P** | Inventário-piloto de 26 maio 2026 (setor de TI) e simulação de Monte Carlo, etapa anterior do TCC | Esqueleto 10 do TCC; não recalculado nesta etapa |
| **E** | Ensaio de carga das leituras neste sistema, contra o cadastro do SPW de 24 set. 2026; a carga efetiva, em 25 set. 2026 com `--gravar`, deu os mesmos números | `dados/migrar_evento_2026.py` + `dados/depara_localizacoes_2026.csv` (saída impressa pelo script); entrada = `relatorio_20260922_211049.xlsx` convertido por `dados/relatorio_para_base.py` |
| **L** | Literatura (Paul et al., 2024; Madeira Junior e Silveira, 2024): redução de 70–80% do tempo com automação | Cálculo: 37,6 h ÷ 0,30 e ÷ 0,20 |
| **C** | Cálculo derivado de números acima | Indicado na linha |

## Base patrimonial (Tabela 2, Figuras 3–5)

| Número no TCC | Valor | Fonte |
|---|---|---|
| Registros na base | 7.428 | R `registros` |
| Ativos / baixados / doados / inservíveis | 3.519 (47,4%) / 2.051 (27,6%) / 1.854 (25,0%) / 4 | R `situacao` |
| Registros já desfeitos | 52,6% | C: (2.051 + 1.854 + 4) ÷ 7.428 |
| Localizações com bens ativos | 97 (+ 1 registro sem tombamento e sem localização) | R `localizacoes` |
| Classes: móveis / proc. dados / máquinas | 1.906 (54,2%) / 1.189 (33,8%) / 342 (9,7%) | R `classes_ativos` |
| Valor atual dos ativos | R$ 102,8 mi | R `valor_ativos_M` |
| Valor dos bens móveis (sem sede, terrenos, instalações) | R$ 15,58 mi; 3.513 bens | R `valor_moveis_M`, `moveis_n` |
| Cinco maiores localizações | 30% do acervo ativo | R (top 5 de `Local Sistema`: 29,9%) |
| Idade média / mediana dos móveis | 14,1 / 13,7 anos (referência 22 set. 2026) | R `idade_media`, `idade_mediana` |
| Bens com mais de 10 anos | 60,0% | R `pct_mais_10_anos` |
| Entrados desde 2020 / antes de 2000 | 1.124 / 521 | R `n_desde_2020`, `n_antes_2000` |
| Curva ABC (A / B / C) | 352 (10,0%) / 1.068 (30,4%) / 2.092 (59,6%) | R `abc` |

## Inventário-piloto (Tabela 4, Figuras 6–7)

| Número no TCC | Valor | Fonte |
|---|---|---|
| Bens verificados; lotes analisados | 266; 14 lotes (222 bens) | P |
| Produtividade; tempo mediano por bem | ≈ 94 bens/h; 37 s | P |
| Acurácia locacional; divergências | 94,4%; 15 bens (5,6%) | P |
| Monte Carlo (10.000 iterações, bootstrap dos lotes) | 37,6 h; IC 95% 30,6–46,6 h; universo 3.518 | P |

## Campanha censitária (Tabelas 5–6, Figuras 8–10)

| Número no TCC | Valor | Fonte |
|---|---|---|
| Bens conferidos | 3.457 = 3.428 ativos + 27 baixados + 2 doados | R `conferidos`, `conferidos_sit` |
| Cobertura do acervo ativo; pendentes | 97,4%; 91 | R `cobertura`, `pendentes` |
| Período; dias de campo; servidores | 18 ago. a 22 set. 2026; 15; 4 (dois com 98,7% das leituras) | R `dias_campo`, `servidores`, `pct_dois_servidores` |
| Lotes de sincronização | 345 | R `lotes` |
| Tempo efetivo; produtividade; tempo por bem | 28,2 h; 113 bens/h; 31,8 s (mediana por lote 26,4 s) | R `tempo_efetivo_h`, `produtividade`, `seg_por_bem`, `seg_por_bem_mediana_lote` |
| Regra do tempo efetivo | intervalos entre lotes consecutivos do mesmo dia, sem pausas > 1 h e sem o 1º lote do dia | `recalcular_metricas.py`, função `prod_rows` |
| Etapa 1 (18–21 ago., pavimentos administrativos) | 4 dias; 1.546 bens; 137 lotes; 7,8 h; 185 bens/h; 19,5 s | R `etapas[0]` |
| Etapa 2 (24–28 ago., depósitos, subsolos, áreas técnicas) | 5 dias; 1.593; 139; 13,2 h; 114 bens/h; 31,7 s | R `etapas[1]` |
| Etapa 3 (1–22 set., varredura residual) | 6 dias; 318; 69; 7,2 h; 35 bens/h; 102,7 s | R `etapas[2]` |
| Volume diário (ativos) | etapa 1: 325–450; etapa 2: 115–540; etapa 3: 1–111 | R `diario` |
| Cobertura acumulada ao fim das etapas 1 / 2 / atual | 43,8% / 88,4% / 97,4% | R (soma de `diario`) |
| Horas para concluir os 91 pendentes; total projetado | 2,6 h ao ritmo da etapa 3; ≈ 31 h | R `horas_para_concluir_ritmo_e3`, `horas_total_projetado` |
| Comparação com o modelo | 31 h dentro do IC (30,6–46,6); 18% abaixo de 37,6 h | C: 1 − 30,8 ÷ 37,6 |
| Extrapolação otimista após a etapa 1 | ≈ 19 h; erro de quase 40% | C: 3.518 ÷ 185 = 19,0; 1 − 19 ÷ 30,8 |
| Horas-pessoa (soma por servidor e dia) | 28,8 h | R (análise exploratória; ≈ horas de equipe porque quase sempre um servidor por dia) |
| Produtividade individual dos dois principais servidores | 134 e 94 bens/h | R (análise exploratória, não versionada por conter nomes) |
| Ganho sobre o piloto | ≈ 20% | C: 113 ÷ 94 |

## Acurácia locacional (Tabelas 7–9)

| Número no TCC | Valor | Fonte |
|---|---|---|
| Concordantes / divergentes (bens ativos) | 2.859 (83,4%) / 569 (16,6%) | R `concordantes`, `divergentes`, `concordancia_pct` |
| Parcial ao fim da etapa 1 (esqueleto 10) | 81,2% | Esqueleto 10 (1.580 ativos conferidos até 21 ago.) |
| Pares origem-destino; dez mais frequentes; pares com ≤ 5 casos | 187; 179 casos; 167 pares (320 casos) | R `div_pares_n`, `div_top10`, `div_top10_sum`, `div_pares_ate5` |
| Categorias: entre pavimentos / recolhimento a depósito / mesmo pavimento / granularidade / origem lógica | 267 (46,9%) / 150 (26,4%) / 103 (18,1%) / 40 (7,0%) / 9 (1,6%) | R `div_categorias` |
| Taxa por pavimento: maior / menor | 7º pav. 28,1% / 11º pav. 1,9% | R `div_taxa_pavimento` |
| Taxa por classe: móveis / proc. dados / máquinas | 17,6% / 17,4% / 10,3% | R `div_taxa_classe` |
| **Reconciliação com o cadastro reestruturado** | 3.456 leituras; 280 divergentes (8,1%) | E (saída do `migrar_evento_2026.py` com o de-para) |
| Localizações lidas renomeadas ou extintas | 38 de 91; 1.970 leituras (57,0% de 3.457) | E (cruzamento `Local Inventariado` × localizações ativas do SPW em 24/09) |
| Salas resultantes no evento | 98 = localizações ativas do cadastro | E |

## Conservação, baixados, cobertura e pendentes (Tabela 10)

| Número no TCC | Valor | Fonte |
|---|---|---|
| Estado de conservação | 3.457 × "Bom" (100%) | R `conservacao` |
| Bens baixados/doados localizados fisicamente; em depósitos | 29 (27 + 2); 12 | R `nao_ativos_encontrados` |
| Pavimentos com 100%; menor cobertura | 7; 92,3% (2º pavimento) | R `cobertura_pavimento` |
| Localizações em 100% / ≥ 90% / total | 74 / 91 / 97 | R `loc_100`, `loc_90`, `loc_n` |
| Pendentes por localização | Depósito Rampa 28; comunicação 14; Depósito SEPAT 5 | R `pendentes_loc` |
| Pendentes em localizações lógicas; físicos | 12 (4 licenças amortizadas, 4 apuração de responsabilidade, 4 bens novos); 79 | R `pendentes_logicos`; C: 91 − 12 |
| Pendentes por classe | 56 móveis; 26 informática | R `pendentes_classe` |
| Valor pendente; parcela em software | R$ 715,7 mil; R$ 610,6 mil | R `pendentes_valor_k`, `pendentes_valor_softwares_k` |
| Cobertura por valor (móveis) | 95,4% (R$ 14,87 de 15,58 mi) | R `cobertura_valor_moveis_pct`, `valor_moveis_conferido_M` |
| Cobertura por classe ABC | A 97,4% / B 97,7% / C 97,3% | R `abc` |
| Cobertura por valor na posição parcial (esqueleto 10) | 13,8% | Esqueleto 10 |

## Comparação com o processo manual (Tabela 11, Figura 11)

| Número no TCC | Valor | Fonte |
|---|---|---|
| Tempo manual estimado | 125 a 188 h | L: 37,6 ÷ 0,30 = 125; 37,6 ÷ 0,20 = 188 |
| Redução de tempo | > 75% | C: 1 − 28,2 ÷ 125 = 77% (piso) |
| Automatizado | 28,2 h medidas; ≈ 31 h projetadas; 37,6 h simuladas | R, C, P |

## Pipeline e carga (Tabela 12)

| Número no TCC | Valor | Fonte |
|---|---|---|
| Leituras aceitas / rejeitadas na carga | 3.456 / 0 | E |
| Registro tratado como sobra | 1 (bem sem número de tombamento) | E |
| Bens fora da base atual | 0 | E |
| Carga efetiva no banco (25 set. 2026): leituras / divergentes / salas / sobras | 3.456 / 280 / 98 / 1 (iguais ao ensaio); as 11 leituras feitas neste sistema em 24–25 set. eram testes e foram descartadas (`--descartar-feitas-aqui`): a última leitura no banco é de 22 set. 2026 11:34. Evento encerrado em 25 set. 2026 13:33 com o nome "Inventário Eventual 2026"; o snapshot de encerramento congelou 3.549 bens (3.520 ativos das salas + 29 baixados/doados lidos), e os 280 divergentes se mantêm contra ele | E (`--gravar`; `inventario.encerrar_evento`) |
| Painel do sistema após a carga: lidos / divergentes / cobertura | 3.167 / 260 / 90,0% — definição diferente do TCC: *lido* = bem ativo lido na própria sala do SPW; *divergente* só conta ativos (os 280 incluem 20 baixados/doados); 97,4% do TCC = conferidos ÷ ativos | E (`inventario.resumo`) |

## O que não está em nenhum arquivo versionado (e por quê)

- O relatório `relatorio_20260922_211049.xlsx` e a planilha `INVENTARIO_SEPAT` trazem nomes de servidores
  e do usuário de cada bem; ficam fora do repositório. Os números agregados estão em
  `metricas_campanha_2026.json`.
- As produtividades individuais (134 e 94 bens/h) e as horas-pessoa (28,8 h) vieram de uma análise
  exploratória com os nomes; o TCC só reporta os agregados e não identifica ninguém.
- O de-para de localizações (`dados/depara_localizacoes_2026.csv`) foi derivado por regra de maioria e
  contém quatro casos ambíguos (setores desmembrados) e um sem correspondência clara (DEPEV → Sala
  Diretoria, 16 dos 27 bens); qualquer mudança nele altera o 280 (8,1%).

## Conferência sistema × TCC (25 set. 2026, após a carga e o encerramento)

Feita com `conferir_sistema.py` (só lê o banco; usa o snapshot congelado no encerramento do evento).

| Número | TCC | Sistema | Situação |
|---|---|---|---|
| Leituras / sobras (= 3.457 conferidos) | 3.456 / 1 | 3.456 / 1 | igual |
| Divergentes após o de-para | 280 (8,1%) | 280 | igual |
| Salas do evento | 98 | 98 | igual |
| Abertura (primeira leitura) | 18 ago. 2026 | 18 ago. 2026 14:16 | igual |
| Período; dias de campo; servidores | 18 ago. a 22 set.; 15; 4 | 18 ago. 14:17 a 22 set. 11:34; 15; 4 | igual |
| Baixados/doados localizados | 29 | 29 | igual |
| Conservação | 3.457 × "Bom" | 3.456 × "Bom" (+ a sobra) | igual |
| Cobertura do acervo ativo | 97,4% | 3.427 ÷ 3.520 = 97,4% | igual |
| Base patrimonial: registros / ativos / localizações | 7.428 / 3.519 / 97 (posição 22 set.) | 7.429 / 3.520 / 98 (SPW de 25 set.) | um bem entrou depois; as 98 salas são as do cadastro reestruturado (E) |
| Pendentes | 91 | 93 | o bem novo + a sobra sem tombamento, que o relatório antigo contava como ativo conferido |
| Divergentes com os nomes antigos das salas | 569 (16,6%) | não existe | o sistema só conhece os nomes novos; o TCC apresenta 569 e 280 e explica a reconciliação |
| Painel do sistema: lidos / divergentes / cobertura | — | 3.167 / 260 / 90,0% | definição da tela: "lido" = ativo encontrado na própria sala; divergentes só entre ativos |
| Tempo efetivo, produtividade, etapas | 28,2 h; 113 bens/h; 3 etapas | não conferível | dependem dos lotes de sincronização do sistema antigo (R), não migrados |
