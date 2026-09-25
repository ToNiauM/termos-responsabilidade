# docs/tcc — TCC "Automação de inventários no setor público com Data Science e infraestrutura open source replicável"

MBA Data Science e Analytics, USP/Esalq, 2026. O texto do TCC vive no repositório privado `ToNiauM/tcc`;
aqui ficam os arquivos que ligam o TCC a este sistema e permitem refazer os números.

| Arquivo | O que é |
|---|---|
| `numeros-da-pesquisa.md` | Cada número do TCC com a origem e como reproduzir; inclui a conferência sistema × TCC |
| `recalcular_metricas.py` | Recalcula os indicadores da campanha a partir do relatório do sistema antigo (não versionado: traz nomes) e gera as figuras |
| `metricas_campanha_2026.json` | Saída agregada do script acima (sem nomes) |
| `conferir_sistema.py` | Lê o banco deste sistema e imprime os números comparáveis com o TCC |
| `analise_regras_banca_esqueleto11.md` | Análise do Esqueleto 11 contra o Manual de Normas do programa, com o que foi corrigido nos Esqueletos 12 e 13 |
| `12_build_esqueleto.py`, `13_build_esqueleto.py` | Scripts que geram o docx do TCC (versões 12 e 13) a partir do template oficial, das figuras e das métricas; rodam na pasta do TCC |
