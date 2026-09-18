# Robô de importação do SPW — handoff para a próxima sessão

**Data:** 2026-09-17. **Estado:** acesso e download PROVADOS; robô ainda não desenhado nem implementado.
**Próximo passo:** `superpowers:brainstorming` curto → spec → plano pequeno → implementação. Não há decisão tomada
sobre agendamento, nomes ou tratamento de erro além do que está em "Ideia acordada".

## O que o usuário faz hoje à mão
1. Entra no SPW (sistema de patrimônio do CFC, ASP.NET WebForms + DevExpress) com usuário e senha.
2. Abre "Consulta de Bens Patrimoniais" (Tipo RELAÇÃO DOS BENS, Sem Agrupamento, Situação TODOS) e clica Exportar → Excel → Detalhado.
3. Abre o `.xls`, copia as linhas e cola na planilha no formato de importação (a que sai de "Exportar bens" do nosso sistema), salva `.xlsx`.
4. Faz upload em `/upload` (Atualizar base). O robô deve substituir os passos 1–4 sem mudar a importação.

## O que foi provado (scripts nesta pasta, rodados com Playwright; ver "Ambiente")
- `testar_acesso.py`: login + abre a consulta + lista controles. `baixar_export.py`: login + Exportar Excel/Detalhado + salva o arquivo.
- Login: `https://www3.cfc.org.br/spw/sistemasgerenciaispadronizado/chamador/login.aspx`
  - usuário: `#ContentPlaceHolder1_ASPxRoundPanel1_txtUsuario_I` (fill)
  - senha: clicar o placeholder `#..._txtSenha_I_CLND`, esperar ~300 ms, `fill(force=True)` em `#..._txtSenha_I`
  - combo Conselho já vem "CFC"; botão `#ContentPlaceHolder1_ASPxRoundPanel1_btnEntrar` (div DevExpress; usar click + expect_navigation)
  - após login cai em `MenuChamador.aspx`; a URL da consulta funciona direto com a sessão.
- Consulta: URL completa com P1..P6 está em `secrets/spw.env` (SPW_CONSULTA_URL). Grid com 7.429 itens/149 páginas.
  - Exportar: `#ContentPlaceHolder1_ASPxButton1` abre popup `PCExportacao`: selects `#ContentPlaceHolder1_PCExportacao_cboArquivo`
    (PDF/Excel — o padrão é PDF!) e `#..._cboModeloExportacao` (Detalhado/Resumido); botão `#..._imgExportar` (input image) dispara o download.
  - Excel+Detalhado → `gvSemAgrupamento.xls`, **BIFF binário** (~1,4 MB), aba `Sheet`, 2 linhas de título, cabeçalho na 3ª linha com
    exatamente as 9 colunas de `db.COLUNAS_EXPORT` (há uma coluna vazia entre "Data Entrada" e "Valor Compra"; datas vêm como texto dd/mm/aaaa).
- `db.importar_bens` só aceita `.xlsx` com cabeçalho na linha 1 (`db._aba_do_export`). Conversão que funcionou:
  `xlrd` lê o `.xls`, acha a linha com "Número Bem", escreve dali em diante com `openpyxl` (células vazias → None, datas → datetime).
  Resultado num banco temporário: 7.429 bens importados, 3.520 ativos — igual ao upload manual.

## Ideia acordada com o usuário (ainda sem brainstorming)
- Um único script `importar_spw.py`: entra, baixa, compara hash com a última importação (não repete), converte, chama `db.importar_bens`
  (mesma validação/rollback/histórico do upload), registra log; na 2ª falha seguida avisa por e-mail.
- Rodar por cron **dentro do container** (`docker compose exec web ...` ou serviço de cron no compose) por volta das 4h. **Sem rebuild nem restart**:
  o app lê o SQLite a cada requisição e o indicador "última importação" do Início já mostra o resultado.
- Dependências novas no container: `playwright` + chromium (ou tentar `requests` reproduzindo os postbacks — não testado) e `xlrd`.
  Avaliar no brainstorming: Playwright no container pesa ~400 MB; alternativa é rodar o robô no host e importar via `docker compose exec`.
- Perguntas abertas: usuário de serviço no SPW em vez da senha pessoal; horário; o que fazer se o SPW mudar a tela; onde logar.

## Credenciais e segurança
- `secrets/spw.env` (chmod 600, pasta ignorada pelo git): SPW_USUARIO, SPW_SENHA, SPW_LOGIN_URL, SPW_CONSULTA_URL.
- A senha do SPW foi colada no chat em 2026-09-17: recomendar troca depois que o robô estiver no ar. Nunca colocar em memória, docs ou git.

## Ambiente usado na prova (fora do repo, pode ter sumido)
- Playwright Python em venv separado: `python3 -m venv pwenv && pwenv/bin/pip install playwright xlrd openpyxl && pwenv/bin/playwright install chromium`.
- Os scripts desta pasta leem `secrets/spw.env` por caminho absoluto `/opt/web/termos-responsabilidade/secrets/spw.env`.

## Contexto do sistema (para não reler tudo)
- Fase 5 mesclada na main (17e7594) e publicada em 2026-09-17; ver `docs/superpowers/specs/2026-09-17-fase5-*.md` §evidências e README.
- Preferências do usuário: simplicidade acima de tudo; sem Django/Postgres/servidor novo; DSGov só visual; decisões passam por ele.
