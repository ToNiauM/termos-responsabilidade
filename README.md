# Termos de Responsabilidade — CFC

Programa local (Windows) do Setor de Patrimônio para emitir Termos de Responsabilidade (por centro de
custo e individuais) e Termos de Devolução, com botão **Copiar para o SEI** e download em `.docx`.
Os dados ficam num SQLite (`dados/termos.db`) mantido pelo próprio programa.

## Uso

1. Abra `TermosCFC.exe`. A janela abre em `http://127.0.0.1:5000`.
2. **Atualizar base**: envie o export do sistema de patrimônio (`.xlsx`). Só a tabela de bens muda.
3. **Cadastros**: responsáveis por centro de custo (com *renomear*), localização → centro de custo,
   pessoas e bens atribuídos. Bem atribuído a pessoa não entra no termo do setor.
4. **Termos**: escolha o centro/pessoa → página do termo → *Copiar para o SEI* ou *Baixar .docx*.
5. **Textos**: os dizeres dos termos (abertura, compromissos, parágrafos, quem recebe a devolução,
   cidade, sigla do órgão) são editáveis no menu Textos, com marcadores como `{nome}` e `{ccustos}`;
   "Restaurar padrão" volta ao texto original.
6. **Planilha de cadastros**: Cadastros → *Exportar cadastros* gera `cadastros.xlsx` (4 abas). Edite no
   Excel e importe em *Atualizar base → Importar cadastros* — substitui as 4 tabelas inteiras.

Backup = copiar a pasta `dados/`.

## Desenvolvimento

    python -m venv .venv && .venv/bin/pip install -r requirements.txt
    .venv/bin/pytest
    .venv/bin/python app.py        # http://127.0.0.1:5000 (debug)
    .venv/bin/python main.py       # como o programa: janela (ou navegador, se não houver WebView)

A porta pode ser trocada com a variável `TERMOS_PORTA` (padrão 5000), útil para testar sem conflitar
com outra instância já rodando.

Migração inicial a partir das planilhas antigas: `python importar_planilhas.py acervo.xlsx geral.xlsx`.

## Gerar o executável (Windows)

    python -m venv .venv && .venv\Scripts\activate && pip install -r requirements.txt
    build.bat

Sai em `dist\TermosCFC\`. Distribua a pasta inteira (zip). Requer o WebView2 Runtime (já vem no
Windows 10/11 atualizados); sem ele o programa abre no navegador padrão.

Copie o `dados\termos.db` já migrado (por exemplo, da máquina onde rodou o `importar_planilhas.py`)
para `dist\TermosCFC\dados\` antes de distribuir; sem isso o programa abre com a base vazia.

### Observações

Os parágrafos do termo por centro de custo não carregam mais 4 espaços em branco no início (era um
resíduo de indentação do código antigo) — confira o `.docx` gerado.

## Arquivos

| Arquivo | Função |
|---|---|
| `app.py` | rotas Flask |
| `db.py` | esquema, importação, consultas, cadastros |
| `termos_html.py` | corpo HTML dos termos (padrão gelic; tabelas 80 % / 100 %) |
| `textos.py` | textos padrão dos termos e marcadores |
| `Script_Termo_Individual.py`, `Termo_de_Responsabilidade.py`, `termo_devolucao.py` | geradores `.docx` |
| `config.py` | pasta de dados (`TERMOS_DADOS` sobrepõe) |
| `main.py`, `build.bat` | programa de desktop e build |
| `templates/`, `static/dsgov/` | telas DSGov 3.7.0 (offline) |
