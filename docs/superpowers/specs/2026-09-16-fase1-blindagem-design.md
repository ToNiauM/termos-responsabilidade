# Fase 1 — Blindagem barata: design

Data: 16/09/2026. Estado: aprovado em conversa (triagem em `../notes/2026-09-16-revisao-system-improvement-plan.md`).
Base: `main` em `9950dc1`.

## 1. Objetivo

Fechar os defeitos confirmados do `SYSTEM_IMPROVEMENT_PLAN.md` que custam pouco e protegem dados ou o
perímetro, sem mudar fluxo, telas, esquema de banco ou dependências. Estimativa: ~80 linhas de código,
~11 commits, um branch `fase1-blindagem`.

Fora do escopo (decidido): CSRF, emissão via POST, marcação de e-mail, refatorações, testes de
concorrência/browser, CI, migrações versionadas, compensação de fotos no R2, otimização de consultas.

## 2. Mudanças

### 2.1 Segredo de sessão por instalação
- `config.chave_secreta()`: devolve `TERMOS_SEGREDO` se a variável existir; senão lê `pasta_dados()/segredo.txt`;
  se não existir, cria com `secrets.token_hex(32)` usando `os.open(O_CREAT|O_EXCL)` (dois processos não
  brigam: quem perder lê o arquivo). Cria a pasta de dados se preciso (`main.py` importa `app` antes de
  `db.inicializar()`).
- `app.py`: `app.secret_key = config.chave_secreta()` no import, no lugar do literal.
- Efeito: cookies e tokens de revisão assinados só valem para a instalação. Trocar o arquivo invalida a
  sessão (seleção da devolução em andamento), o que é aceitável. Na web, `dados/` é bind mount, então o
  arquivo persiste entre rebuilds. No desktop fica ao lado do `termos.db`. Backup não precisa incluir.
- `tests/test_cadastros_ux.py` passa a importar `app` dentro da fixture (como `test_app.py`), para o
  segredo de teste nascer na pasta temporária.

### 2.2 Contexto do Docker
`.dockerignore` ganha: `secrets`, `*.env`, `*.db`, `*.db-*`, `*.xlsx`, `*.gz`, `*.md`, `backup.sh`.
Verificação: construir a imagem com tag de teste, listar `/app` e remover a imagem.

### 2.3 Operações atômicas
- `db._renomear_pessoa(conn, antigo, novo)` faz os dois UPDATEs sem commit; `renomear_pessoa` chama e commita
  (contrato mantido); `salvar_pessoa` chama o helper, faz o UPDATE de e-mail/matrícula e commita uma vez;
  qualquer exceção faz `rollback`.
- `textos._gravar(conn, chave, valor)` sem commit (INSERT OR REPLACE, ou DELETE quando igual ao padrão);
  `salvar` e `restaurar` continuam commitando (contrato mantido); novo `textos.salvar_todos(conn, valores)`
  valida tudo, grava tudo e commita uma vez, com `rollback` em exceção. `app.textos_salvar` usa `salvar_todos`.

### 2.4 Importação de cadastros: abas `inv_*` tudo ou nada
Em `db.importar_cadastros`, depois de ler as abas opcionais: se houver alguma aba `inv_*` mas não todas as
cinco, `ImportacaoInvalida("Planilha de cadastros: abas de inventário incompletas (faltam: …). Envie as 5
abas inv_* ou nenhuma.")`. Nada é gravado. Planilha só com as 4 abas continua preservando o inventário.

### 2.5 Deduplicação de emissão considera o processo
`db.registrar_emissao`: reaproveita o último termo do dia apenas se `ultimo["processo_id"] == proc["id"]`
além da mesma lista de números. Trocar o processo vigente e reemitir no mesmo dia gera termo novo.

### 2.6 Entradas inválidas com resposta previsível
- `app._exigir_processo`: `tipo not in db.TIPOS_TERMO` → `abort(404)` (antes: KeyError 500).
- `app.termo_docx`: `request.method == "HEAD"` → responde 200 vazio sem gerar nem registrar.
- `app_inventario.ler` e `atualizar_leitura`: corpo JSON que não seja objeto → `{"erro": "Envie um objeto
  JSON."}`, 400.
- `app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024` (mesmo limite do nginx) e
  `@app.errorhandler(413)` com flash "Arquivo muito grande: o limite é 20 MB." e redirect ao referrer.

### 2.7 Texto literal no XLSX
`db.acrescentar_linha(ws, valores)`: `ws.append(valores)` e, na linha recém-criada, toda célula `str`
com `data_type == "f"` passa a `"s"`. Números e datas não mudam. Usado em `exportar_cadastros`,
`exportar_recorte`, `exportar_bens`, `inventario.exportar_abas`, `inventario.exportar_xlsx` e
`Termo_de_Responsabilidade.gerar_planilha_centro`. Verificado com openpyxl 3.1.5: roundtrip devolve
a string `=1+1` com `data_only=True` e `False`.

### 2.8 Frontend
- `inventario_sala.html`: o `<form id="form-sobra">` sai de dentro da classe `br-card`; passa a existir
  `<div class="br-card mt-3" id="card-sobra" hidden>` envolvendo o form. O JS alterna `hidden` no card.
  Motivo: `core.min.js` do DSGov faz `setAttribute("id", "card…")` em todo `.br-card`; o mesmo mecanismo
  quebrou o formulário de transferência dos cadastros em 16/09.
- `termo.html`: o botão Copiar só chama `/registrar` se a cópia deu certo (Clipboard API sem exceção, ou
  `execCommand("copy")` devolvendo `true`); caso contrário mostra "Não foi possível copiar. Selecione o
  texto do documento e copie com Ctrl+C." e não registra.
- Verificação manual no navegador (sem teste automatizado de JS).

### 2.9 Testes e documentação
- `tests/test_app.py:71`: ano fixo `"2026"` vira o ano corrente.
- `README.md:51-52`: "4 abas" → "4 abas; quando há inventário, mais 5 abas `inv_*`, que só são
  substituídas quando as 5 vêm juntas".

## 3. Testes
Um teste por mudança de comportamento, no arquivo que já cobre a função (`test_db.py`, `test_textos.py`,
`test_inventario.py`, `test_app.py`, `test_config.py`). Falhas de banco no meio da operação são
provocadas com `CREATE TRIGGER … RAISE(ABORT, …)` na própria conexão da fixture, e o estado é lido por
uma conexão nova (`db.conectar()`).

## 4. Entrega
Branch `fase1-blindagem`; commits atômicos; suíte completa verde; merge em `main`, push e
`docker compose up -d --build` na VPS só com autorização do usuário. Primeira subida cria
`dados/segredo.txt`; o usuário deve fechar e reabrir a sessão do navegador se tiver devolução em andamento.
