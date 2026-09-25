# Análise do Esqueleto 11 contra as regras da banca (MBA USP/Esalq, turmas 251/252)

Data: 25 set. 2026. Objeto: `11_TCC_Esqueleto_para_revisao_ajustado.docx` (30 páginas). Fontes das regras: Manual de
Instruções e Normas (67 p.), Templates 07 e 08, Manual de Metodologias, Termo de Anuência. "p. N" = página do Manual.
Verificações objetivas feitas por script (contagem de palavras, DOIs no Crossref, doi.org, LexML, Drive, docx).

## 1. Riscos altos (podem derrubar nota ou gerar rejeição)

| # | Problema | Evidência | Regra | Correção |
|---|---|---|---|---|
| 1 | **Jain et al. (2016) não existe como citado.** O DOI 10.1109/CCEM.2016.012 resolve para "Investigating factors affecting cloud computing adoption by SMEs" (Kumar e Samalia); o título "Docker and container cloud: challenges and opportunities" não consta no Crossref | Crossref + doi.org | p. 45-47: obra citada deve existir e constar corretamente | Remover a citação (Metodologia, "Arquitetura progressiva") e a referência; Boettiger (2015) e Nüst et al. (2020) já sustentam a frase |
| 2 | **Zelbst et al. com ano, volume, páginas e DOI errados.** Real: 2012, IJOPM 32(3): 329-350, DOI 10.1108/01443571211212600 (texto diz 2010, 30(12): 1318-1337, DOI 404) | Crossref | idem | Corrigir referência e as duas citações "(Zelbst et al., 2010)" → 2012 |
| 3 | **Brito et al. (2019) com DOI errado** (404). Real: 10.29069/forscience.2019v7n2.e661 | doi.org | idem | Corrigir DOI |
| 4 | **Kar et al. (2022) e Ton et al. (2024) com autores errados.** Reais: Kar, S.; Bhimrajka, S.; Kumar, A.; Mukherjee, S. 2022. In: IEEE CONECCT, 2022. p. 1-6, DOI 10.1109/CONECCT55679.2022.9865739. Ton, N.T.N.; Le, M.T.; Lam, T.T.; Do, T.D. 2024. In: IEEE GTSD, 2024. p. 362-367, DOI 10.1109/GTSD62346.2024.10675053 | Crossref | idem; modelo de Anais p. 56 exige evento, ano, cidade, país e páginas | Reescrever as duas entradas |
| 5 | **Acórdão TCU nº 1.886/2007 não trata de segregação de funções**: ementa é "Auditoria de conformidade. Pessoal. Pensão civil. Falta de recadastramento anual". Citado na Introdução e em "Estado de conservação e desdobramentos administrativos". O guia de 24 ago. já havia retirado essa citação; o Esqueleto 11 a reintroduziu | LexML | Critério da banca: consistência e domínio do conteúdo (p. 28) | Retirar as duas menções ou substituir por norma que de fato trate do tema, verificada |
| 6 | **Termo de Anuência assinado não localizado.** O texto afirma "a pesquisa dispõe de Termo de Anuência institucional"; só existe o modelo em branco (local e Drive, 27 mar. 2026). O CFC é nomeado no título de rosto, Resumo, Abstract e corpo | Drive + pasta | p. 16-17: obrigatório para dados privados de organização e para citar o nome; permanece exigido mesmo anonimizando | Obter assinatura (modelo 1, com autorização de divulgar o nome) e entregar com o TCC; sem ela, anonimizar integralmente ("autarquia federal em Brasília/DF") |
| 7 | **Resumo com 309 palavras** (limite 250) e com "Concluiu-se que". Abstract idem (opcional, mas mesmas regras) | contagem | p. 25 e p. 57 item 7 | Cortar ~60 palavras; trocar a frase final por afirmação direta |
| 8 | **Metodologia sem a justificativa de não submissão ao CEP.** Há apenas "sem envolver seres humanos como sujeitos de pesquisa" | texto | p. 15: justificativa e fundamentação no art. 1º da Res. CNS 510/2016, obrigatória na Metodologia | Acrescentar parágrafo com a base legal (incisos aplicáveis) e o FDE entregue |
| 9 | **Fase 2 (PostgreSQL, Grafana, Streamlit, ETL) declarada "objeto central deste estudo" sem evidência no texto**: nenhuma figura de painel (nota de revisão pendente), repositório da Fase 2 não publicado (nota no Apêndice A), indicadores calculados por script Matplotlib. A banca pode pedir para ver | texto + notas | p. 19: reprodução do estudo; p. 28: consistência entre objetivo, metodologia e resultados | Ou evidenciar (capturas do painel citadas no texto + repositório público) ou reposicionar: objeto central = sistema de conferência + análise em Python; painéis como desdobramento |

## 2. Riscos médios (redução de nota ou pedido de correção)

- **Sete notas de revisão** ainda no texto ("[NOTA DE REVISÃO ...]"). Remover todas (p. 21; templates).
- **Gráficos inseridos como imagem PNG** (11 figuras do Matplotlib). Manual p. 35: gerados no Excel e "em hipótese alguma inseridos como imagem", salvo software que o Excel não reproduz. Justificável para Monte Carlo e curva ABC, mas as barras simples (Figuras 3, 4, 8, 9, 11) são reproduzíveis no Excel. Decidir e, se mantiver, garantir sem grade, sem borda, eixos pretos 1,5 pt (Tabela 8 do Manual).
- **Subtítulos de Resultados e Discussão não espelham os da Metodologia** (p. 43). Metodologia: caracterização, local e base, arquitetura, coleta, análise, síntese. Resultados: 11 subtítulos temáticos. Alinhar ao menos a ordem e a nomenclatura (ex.: "Base patrimonial", "Inventário-piloto e simulação", "Campanha censitária", "Acurácia e reconciliação", "Pipeline e painéis").
- **Tempo verbal**: trechos no presente descrevendo o sistema ("o operador seleciona", "o processamento é idempotente", "a importação é atômica", "A Fase 3 prevê"). Manual p. 42 e 56: pretérito perfeito em todo o texto.
- **Conclusões repetem os números dos Resultados** (94, 113, 97,4%, 28,2 h, 125-188 h, 569, 280, 29, 345). p. 25-26: "não podem ser meras reproduções dos resultados". Reduzir a afirmações e implicações.
- **Normas citadas e ausentes das Referências**: CF/1988 art. 70, IN SEDAP 205/1988, NBC TSP 07, ISO 55000:2014 (e o acórdão, se ficar). p. 45 e 47: tudo citado consta nas Referências; há modelo para lei (p. 56).
- **Inconsistência interna**: nota da Tabela 9 diz "28 bens baixados ou doados"; Tabelas 5 e 12, o texto e o sistema dizem 29 (27 + 2). Corrigir para 29.
- **Ordem em citações múltiplas** (p. 46): autor único, depois dois autores, depois et al. Corrigir "(Broomfield e Reutter, 2021; Merhi, 2021)" → "(Merhi, 2021; Broomfield e Reutter, 2021)" e "(Hanis et al., 2011; Kaganova e Amoils, 2020)" → "(Kaganova e Amoils, 2020; Hanis et al., 2011)".
- **"et al." em itálico** em uma ocorrência (p. 46: sem itálico).
- **Tabela 3 (em Resultados) com "Fonte: Dados originais da pesquisa"**; em Resultados a fonte é "Resultados originais da pesquisa" (p. 34).
- **Referências de periódico com "Disponível em / Acesso em"**: o modelo do Manual (p. 51) para artigo é sem URL; a forma com URL é para documentos online. Não é erro grave, mas destoa do modelo. Madeira Junior e Silveira (2024) é TCC: fonte a evitar (p. 55); manter só se indispensável.
- **Folha de rosto sem titulação do autor** (modelo p. 41: "Titulação. Instituição. E-mail"). Hoje: "Supervisor do Setor de Patrimônio do CFC". Acrescentar a formação (ex.: "Bacharel em ...").
- **Limite de 30 páginas já no teto**: ao remover as notas sobra pouco, mas as inserções previstas (capturas do painel, custo/ROI) estouram. Cortar antes de inserir; Apêndice conta no limite.
- **Hipótese não enunciada na Introdução**: pesquisa-ação pede problema e hipóteses na Introdução (Met. p. 31-33); a hipótese de "redução superior a 75%" só aparece nos Resultados.
- **Expressões a evitar** (p. 57): "Cabe registrar" (2), "Destaca-se" (1), "Trata-se" (3). Não estão na lista literal, mas são da mesma família de muletas.
- **Título com expressões em outro idioma** ("Data Science", "open source"): p. 24 pede evitar. Tolerado na área, mas é regra escrita; alternativa: "ciência de dados" e "código aberto".

## 3. O que está em conformidade (conferido)

Template 07 correto (pesquisa-ação sem algoritmo de ML; limite de 30 p.); 30 páginas; Arial 11 (9 nos endereços), margens 2,5 cm, justificado, recuo 1,25 cm, espaçamento 1,5 no texto e 1,0 em resumo/tabelas/referências; cabeçalho e paginação à direita em Arial 9; título com 14 palavras, negrito, sem "estudo de/análise de"; título repetido antes do Resumo; 5 palavras-chave diferentes do título; Introdução em 2 páginas (p. 3-4), sem figuras, com objetivo no último parágrafo; seções na ordem exigida, títulos em negrito sem numeração; "Metodologia" (não "Material e Métodos"), com 5W2H e fluxograma pedidos pelo orientador; 12 tabelas sem negrito, números à direita, título acima e fonte abaixo; 11 figuras com legenda abaixo e fonte, todas citadas no parágrafo anterior; siglas entre colchetes; datas no formato "18 ago. 2026"; citações indiretas, sem apud; Conclusões sem citações nem figuras; Agradecimento com 2 linhas e sem nomes; referências em ordem alfabética, sem negrito, à esquerda; 17 das 21 referências com DOI válido e dados corretos; Turnitin dos Resultados Preliminares em 6%; nenhum nome de empregado no texto; números do sistema batem com o texto (ver `docs/tcc/numeros-da-pesquisa.md` no repositório).

## 4. Ordem sugerida de correção (para o Esqueleto 12, sem sobrescrever o 11)

1. Referências: remover Jain; corrigir Zelbst, Brito, Kar, Ton; acrescentar normas legais.
2. Retirar o Acórdão 1.886/2007 (2 menções).
3. Resumo e Abstract a 250 palavras; sem "Concluiu-se que".
4. Parágrafo do CEP (Res. CNS 510/2016) na Metodologia.
5. Decidir a Fase 2: evidenciar ou reposicionar; publicar repositório ou retirar a promessa do Apêndice A.
6. Termo de Anuência assinado ou anonimização.
7. Remover notas, alinhar subtítulos, pretérito perfeito, Conclusões sem números, nota da Tabela 9 (29), ordem das citações, fonte da Tabela 3, titulação na folha de rosto.
8. Recontar páginas (≤ 30) e regenerar PDF.

## 5. Aplicado no Esqueleto 12 (25 set. 2026, `12_build_esqueleto.py` → `12_TCC_Esqueleto_corrigido.docx/.pdf`)

Feito: Jain removido; Zelbst (2012, 32(3): 329-350, DOI correto), Brito (DOI correto), Kar e Ton (autores, evento,
páginas e DOI reais); acórdão 1.886/2007 retirado das duas menções (a segunda virou descrição da segregação sem
citar norma); Resumo 248 e Abstract 249 palavras, sem "Concluiu-se que"; parágrafo do CEP (Res. CNS 510/2016,
art. 1º, parágrafo único) na Metodologia; CF/1988, IN SEDAP 205/1988, Res. 510/2016, NBC TSP 07 e ISO 55000 citadas
no texto e nas Referências (25 entradas, todas citadas); hipótese enunciada no último parágrafo da Introdução;
"objeto central" retirado da Fase 2; Apêndice A aponta os scripts de recálculo em docs/tcc do repositório; pretérito
perfeito nas descrições do sistema; "Cabe registrar", "Destaca-se" e "Trata-se" eliminados; Conclusões reescritas sem
repetir números; nota da Tabela 9 = 29; fonte da Tabela 3 = Resultados originais; ordem das citações múltiplas;
sete notas de revisão removidas. Resultado: 30 páginas, Introdução nas p. 3-4.

Pendente, porque depende do autor: Termo de Anuência assinado (ou anonimizar); titulação na folha de rosto;
decisão sobre gráficos em Excel vs. imagem; capturas do painel e custo/ROI (só se couberem nas 30 páginas).

## 6. Esqueleto 13 (25 set. 2026, `13_build_esqueleto.py` → `13_TCC_Esqueleto_corrigido.docx/.pdf`)

Subtítulos de Resultados e Discussão espelhando os da Metodologia (Manual p. 43). Metodologia: Caracterização da
pesquisa; Local do estudo e base patrimonial; Diagnóstico do processo manual; Arquitetura progressiva de automação;
Inventário-piloto e projeção por simulação; Campanha censitária; Acurácia locacional e reconciliação cadastral;
Cobertura do inventário, conservação e bens baixados; Comparação com o cenário manual e implicações econômicas;
Síntese metodológica. Resultados: os mesmos nomes, na mesma ordem (a arquitetura fica por último, como "pipeline e
carga no sistema de gestão", porque depende dos números da campanha), mais Limitações do estudo. As antigas
"Procedimentos de coleta" e "Análise dos dados" foram redistribuídas; as duas seções de acurácia e as duas de
cobertura/conservação foram fundidas. 30 páginas; Introdução nas p. 3-4.
