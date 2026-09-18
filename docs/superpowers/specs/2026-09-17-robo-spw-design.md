# Robô de importação do SPW

**Data:** 2026-09-17.
**Estado:** desenho aprovado; ainda não implementado.
**Base:** `main`, commit `33e975f`, após a fase 5.
**Plano:** `../plans/2026-09-17-robo-spw.md`.
**Prova de acesso:** `../notes/2026-09-17-robo-spw/README.md` (login, exportação e conversão provados com Playwright).

## 1. Resultado e limites

O robô substitui os passos manuais de entrar no SPW, exportar a relação de bens,
colar na planilha de importação e enviar em "Atualizar base". Roda no host, em
dias úteis às 4h, e chama `db.importar_bens` com a mesma validação, rollback e
histórico do upload manual. O upload manual continua existindo e não muda.

Não entra: e-mail de alerta, retry automático, usuário de serviço no SPW,
Playwright dentro da imagem Docker, alteração do formato de importação.
Manter Flask, SQLite, Jinja e os componentes DSGov existentes.

## 2. Decisões

| Pergunta | Decisão | Motivo |
|---|---|---|
| Onde roda | Host, venv próprio `.venv-robo/` | Chromium (+400 MB) fica fora da imagem; sem rebuild para o robô |
| Alerta de falha | Só no site: card do Início em alerta | Não há SMTP nem `mail` no host; quem usa o sistema vê |
| Horário | `0 4 * * 1-5` | O SPW quase não muda no fim de semana; antes do backup das 10h |
| Repetição | Não importa se o export for igual ao anterior | Evita histórico de importações vazio todo dia |
| Comparação | SHA-256 das linhas de dados, não do arquivo | As linhas de título do `.xls` podem trazer data de emissão |

## 3. Componentes

### 3.1 `importar_spw.py` (raiz, ao lado de `importar_planilhas.py`)

Roda com `.venv-robo/bin/python`, que tem `playwright`, `xlrd` e `openpyxl`.
Usa `db.py` e `config.py` do checkout diretamente; `TERMOS_DADOS` não é definido,
então `config.pasta_dados()` resolve para `./dados`, a mesma pasta montada no container.

Funções:

- `ler_env(caminho) -> dict`: lê `secrets/spw.env` (SPW_USUARIO, SPW_SENHA, SPW_LOGIN_URL, SPW_CONSULTA_URL).
- `baixar_export(env, destino: Path) -> Path`: Playwright com Chromium headless; login, consulta,
  Exportar → Excel → Detalhado, salva em `destino`. Seletores e esperas iguais ao
  `baixar_export.py` das notas. Em exceção, salva `dados/spw/erro.png` antes de relançar.
- `ler_xls(caminho) -> list[list]`: xlrd; acha a linha cujo primeiro texto é "Número Bem" e devolve
  dessa linha em diante. Células vazias → `None`; datas → `datetime`.
- `hash_linhas(linhas) -> str`: SHA-256 das linhas com valores normalizados como `db._texto` faz
  (`None` → vazio, `str` sem espaços duplicados, `datetime` como dd/mm/aaaa, números como `float`).
- `linhas_para_xlsx(linhas) -> BytesIO`: openpyxl, cabeçalho na linha 1, uma aba.
- `executar(conn, baixar=baixar_export, agora=db._agora) -> dict`: orquestra o fluxo da §4 e devolve
  `{"resultado", "mensagem", "importacao_id"}`. `baixar` é parâmetro para os testes injetarem linhas falsas.
- `main()`: abre `db.conectar()`, chama `executar`, imprime uma linha de log
  (`2026-09-18 04:00:12 importado 7429 bens, 3 novos, 1 removido`) e sai com 0 em
  `importado`/`sem_mudanca` e 1 em `erro`.

### 3.2 `db.py`

Tabela nova no esquema:

```sql
CREATE TABLE IF NOT EXISTS robo_execucoes (
  id            INTEGER PRIMARY KEY,
  iniciado_em   TEXT NOT NULL,
  terminado_em  TEXT NOT NULL,
  resultado     TEXT NOT NULL CHECK (resultado IN ('importado','sem_mudanca','erro')),
  hash          TEXT,
  importacao_id INTEGER REFERENCES importacoes(id) ON DELETE SET NULL,
  mensagem      TEXT
);
```

Funções: `registrar_execucao_robo(conn, iniciado_em, resultado, hash=None, importacao_id=None, mensagem=None)`,
`execucoes_robo(conn, limite=10) -> list[dict]` e `ultimo_hash_robo(conn) -> str | None`
(hash da última execução com resultado `importado` ou `sem_mudanca`).

`painel()` devolve também `"robo": (execucoes_robo(conn, 1) or [None])[0]` e
`"importacao_desatualizada": bool` (última importação há mais de 4 dias, ou nenhuma).

### 3.3 Início (`templates/index.html`)

O card "última importação" ganha uma segunda linha com o estado do robô:

- `robô ok · dd/mm` quando a última execução foi `importado` ou `sem_mudanca`;
- `robô falhou dd/mm: <mensagem curta>` quando foi `erro`;
- nada quando o robô nunca rodou.

O card fica em alerta (borda e ícone de aviso DSGov, classe `dsgov-kpi-alerta`) quando
a última execução do robô foi `erro` ou quando `importacao_desatualizada` é verdadeiro.
O link continua indo para a tela da importação.

### 3.4 Atualizar base (`templates/upload.html`)

Abaixo do formulário, tabela DSGov "Execuções do robô" com as últimas 10 linhas:
data/hora, resultado (tag verde/cinza/vermelha), mensagem, e link para a importação quando houver.
Visível para quem já vê a tela (`pode('upload')`), sem permissão nova.

## 4. Fluxo de uma execução

1. `iniciado_em = agora()`. Cria `dados/spw/` se não existir.
2. `baixar_export` salva `dados/spw/ultimo.xls` (sobrescreve).
3. `ler_xls` → linhas; valida que o cabeçalho contém as chaves de `db.COLUNAS_EXPORT`
   (senão erro "cabeçalho do SPW mudou: faltam ...").
4. `hash_linhas` → se igual a `ultimo_hash_robo`, registra `sem_mudanca` e encerra.
5. `linhas_para_xlsx` → `db.importar_bens(conn, buffer, nome_arquivo="SPW automático")`.
6. Registra `importado` com `importacao_id` e o hash.

Tempo estimado de 1 a 3 minutos, ainda não medido: login e carga da consulta (~20 s), geração do export
pelo SPW (7.429 itens, provavelmente 30–90 s), conversão e importação (~10 s). O download tem timeout de 5 min;
o tempo real da primeira execução vai para a §9.
Não há trava contra execuções simultâneas: o cron roda uma vez por dia e o SQLite serializa a escrita.

## 5. Erros

Qualquer exceção em `executar` vira registro `erro` com `str(exc)[:500]` como mensagem, log e
código de saída 1. Casos previstos:

| Situação | Mensagem registrada | O que o usuário vê |
|---|---|---|
| SPW fora do ar, senha errada, tela mudou | texto do Playwright (timeout, seletor não achado) + `erro.png` | card em alerta; `dados/spw/erro.png` para diagnóstico |
| Cabeçalho do export mudou | "cabeçalho do SPW mudou: faltam Valor Atual" | card em alerta |
| Bem atribuído sumiu do export | texto de `ImportacaoInvalida`, o mesmo do upload manual | card em alerta com a instrução de remover a atribuição |
| Segredo ausente | "secrets/spw.env não encontrado ou incompleto" | card em alerta |

Se `registrar_execucao_robo` falhar (banco travado), a mensagem vai só para o log. O card
"desatualizada" cobre esse caso depois de 4 dias.

## 6. Instalação

- `requirements-robo.txt`: `playwright`, `xlrd`, `openpyxl`.
- `.gitignore`: acrescentar `.venv-robo/`.
- Comandos, documentados no README:

```
python3 -m venv .venv-robo
.venv-robo/bin/pip install -r requirements-robo.txt
.venv-robo/bin/playwright install --with-deps chromium
```

- Crontab do usuário `ToNiauM` (dono de `dados/termos.db`, o mesmo do `backup.sh`):

```
0 4 * * 1-5 cd /opt/web/termos-responsabilidade && .venv-robo/bin/python importar_spw.py >> dados/robo_spw.log 2>&1
```

- Rebuild da imagem uma vez, porque `db.py` e os templates mudam. O robô em si não precisa do container.

## 7. Testes

Sem rede. Em `tests/test_importar_spw.py`:

- `hash_linhas` é estável para linhas iguais e muda com uma célula diferente.
- `linhas_para_xlsx` gera um `.xlsx` que `db.importar_bens` aceita.
- `executar` com `baixar` falso: primeira execução → `importado` e uma linha em `importacoes`;
  segunda com as mesmas linhas → `sem_mudanca` e nada novo em `importacoes`; linhas com bem atribuído
  ausente → `erro` com a mensagem de `ImportacaoInvalida` e `bens` intacta; cabeçalho sem coluna → `erro`.
- `db.registrar_execucao_robo` / `execucoes_robo` / `ultimo_hash_robo`.

Em `tests/test_painel.py` e `tests/test_app.py`: `painel()["robo"]` e `importacao_desatualizada`;
Início mostra "robô falhou" e a classe de alerta; `/upload` lista as execuções.

`baixar_export` e `ler_xls` não têm teste automático: são provados à mão contra o SPW real e a evidência
(data, tamanho do arquivo, total importado) fica na seção 9 desta spec.

## 8. Recomendações fora do escopo

- Pedir ao SPW um usuário de serviço para o robô, em vez da senha pessoal.
- Trocar a senha do SPW depois que o robô estiver no ar (ela foi colada no chat em 2026-09-17).
- Se o SPW trocar de DevExpress ou de layout, os seletores em `baixar_export` são o único ponto a ajustar.

## 9. Evidências

Preenchido na implementação: data da primeira execução real, total de bens, tempo, e a saída do
`pytest`.
