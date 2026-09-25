# ASTRA — Auditoria Técnica do Sistema de Patrimônio

Data: 21/09/2026. Referência do código: `f28cd8fd087d8be41803d934121782655328da59`.

**Escopo: análise e planejamento, sem implementação.** Este é o único arquivo criado/alterado no repositório pela auditoria. Não foram executados robôs reais, migrations na instalação, scripts de limpeza, deploy, commits ou alterações em configuração. Os bancos usados nas reproduções são sintéticos e temporários. Nenhum conteúdo de `secrets/` foi copiado para os testes ou transcrito aqui.

## 1. Executive Summary

O sistema possui uma base funcional coerente: monólito Flask, regras parcialmente separadas das rotas, SQLite com foreign keys, autorização centralizada, proteção CSRF nos POSTs, revisões assinadas de alterações cadastrais, criptografia de credenciais e testes numerosos. Não foi demonstrada execução remota de código, SQL injection ou acesso anônimo ao aplicativo na configuração Docker declarada.

Os riscos mais importantes estão na **integridade documental e na concorrência**, e não na ausência completa de controles de segurança:

1. Uma planilha manual vazia pode apagar todos os bens sem atribuições individuais. Números fracionários são truncados silenciosamente.
2. Emissões diferentes podem receber o mesmo número. Fotografias simultâneas usam a mesma chave externa e uma operação falha após ambas terem enviado conteúdo.
3. O histórico dos termos não é integralmente imutável: nomes e siglas são reescritos, responsáveis/textos são consultados no cadastro atual e a deduplicação considera somente números de bens.
4. Uma leitura pode ser gravada depois que outra conexão finalizou o inventário.
5. A troca de senha não revoga cookies anteriores. O download DOCX grava emissão por GET, fora da proteção CSRF.
6. Uma validação inválida do número do termo ainda grava documento e bloco SEI por causa de um `finally`.
7. O esquema permite duas pessoas para um patrimônio; a migração inicial consegue produzir esse estado e também pode deixar cadastros apagados quando falha.

**Prioridade imediata:** proteger importações destrutivas, alocação de números e fotos, finalização concorrente e migração inicial. Cada intervenção deve começar por um teste que reproduza o defeito. Depois, estabilizar a identidade e o conteúdo histórico dos termos, sessões e estados da integração.

A suíte existente terminou com código de saída 0; uma coleta independente identificou **3.500 casos parametrizados**. Isso não significa 3.500 cenários de negócio independentes: parte expressiva vem da matriz de combinações de funções. As reproduções desta auditoria encontraram defeitos fora dessa cobertura.

### Método, confiança e limites

- **CONFIRMADO:** comportamento demonstrado em ambiente temporário ou diretamente estabelecido por um caminho completo de código; quando só estático, isso é indicado.
- **PROVÁVEL:** caminho de falha consistente, mas sem execução completa do cenário externo/concorrente.
- **HIPÓTESE A VALIDAR:** depende de política, volume, configuração efetivamente ativa ou comportamento de navegador/serviço que não foi medido.
- Severidade de segurança: CRÍTICO, ALTO, MÉDIO, BAIXO, INFORMATIVO. **Não há vulnerabilidade classificada CRÍTICO nesta auditoria.** P0 também abrange corrupção/perda de dados, mesmo sem invasão.
- Prioridades: P0 imediato; P1 alta; P2 média; P3 evolução; P4 opcional. Complexidade no backlog: pequena, média ou grande, relativa a este projeto. Risco no backlog é o risco da alteração futura.
- Houve leitura dos módulos de aplicação, esquema, importações, integrações, documentos, permissões, templates principais, JS/CSS próprios, testes, README, configurações Docker, scripts e documentação de decisões em `docs/superpowers/`. Bibliotecas minificadas não receberam revisão manual linha a linha.
- Não foram encontrados `AGENTS.md` nos diretórios ancestrais consultados nem `CLAUDE.md` no inventário do projeto. O catálogo de skills foi examinado; `gsd-code-review/SKILL.md` e seu workflow foram lidos para avaliar aplicabilidade. O workflow é orientado a fases GSD e artefatos próprios, por isso não foi executado. A auditoria seguiu o escopo específico deste pedido, sem gerar `REVIEW.md`, `.planning/` ou agentes auxiliares. A pasta local `.superpowers/` não apresentou instruções de skill no inventário consultado.
- Não houve teste visual em navegador, medição de contraste, pentest na URL pública nem acesso ao banco operacional. O vhost Nginx indicado no README foi lido, mas não se comprovou que representa a configuração carregada pelo processo. Não foram consultados valores de segredos.
- As linhas abaixo são aproximadas e referem-se à revisão auditada. O nome da função é a referência mais estável.

## 2. Arquitetura identificada

```text
Navegador → Nginx/TLS → Waitress, 4 threads → Flask/Jinja
                                           ├─ app.py / blueprints / app_cadastros.py
                                           ├─ permissoes.py + usuarios.py + comissoes.py
                                           ├─ db.py / inventario.py → SQLite (dados/termos.db)
                                           ├─ fotos.py → R2/S3
                                           └─ HTML / DOCX / XLSX em memória

SQLite.robo_pedidos → atender_pedidos.py → Playwright → SEI / SPW
cron → scripts/atualizar_base.sh → importar_spw.py ────────┘

Desktop: main.py → Flask em 127.0.0.1 → pywebview/navegador
Backup: scripts/backup.sh → SQLite .backup → gzip → rclone/R2
```

Stack: Python, Flask, Werkzeug/Jinja, sqlite3, openpyxl, python-docx, Pillow, boto3, cryptography/Fernet; Playwright e xlrd no trabalhador. UI DSGov com fontes e bibliotecas locais, ECharts e html5-qrcode. Não é Django, não usa ORM nem possui API pública versionada. Há endpoints JSON internos de inventário e registro de emissão.

O compose define web e robô compartilhando `dados/`, ambos montando `secrets/` somente para leitura. Web publica apenas `127.0.0.1:12012`; Nginx faz a exposição externa. O Dockerfile exige login, enquanto o desktop usa administrador local sem conta. `db.criar_esquema()` aplica DDL e ajustes legados no startup dos dois serviços.

Entidades principais: bens; centros/responsáveis; localizações; pessoas; atribuições; processos SEI; emissões e seus bens congelados; importações e diferenças; usuários/funções; inventários, salas, comissão por usuário, leituras, sobras, fotos e snapshot de encerramento; pedidos e execuções de robôs; textos configuráveis.

**Regra central confirmada:** bem atribuído a pessoa sai do termo do centro; sem atribuição, segue o centro mapeado pela localização (`db.bens_do_centro`, linhas 595–602). A conferência física registra divergências, mas não movimenta o cadastro oficial do SPW. Pessoas e usuários de login são entidades distintas. Emissão de devolução não remove a atribuição automaticamente.

## 3. Mapa funcional

| Fluxo | Implementação e comportamento encontrado | Criticidade |
|---|---|---|
| Entrada/atualização de bens | Upload XLSX ou exportação automatizada SPW substitui `bens`; não há CRUD manual geral de bens | Crítica |
| Cadastros | Centros, responsáveis, pessoas, mapeamentos e processos SEI; revisão assinada para várias ações | Alta |
| Mudança de responsabilidade | Atribuir/desatribuir pessoa; mover localizações entre centros; editar responsável | Crítica |
| Localização física/situação/baixa | Vem da importação SPW; não há workflow próprio de baixa | Crítica |
| Termos | Centro, individual, devolução; prévia HTML, copiar, DOCX e planilha de centro | Crítica |
| Histórico de termos | Snapshot parcial de bens, processo, documento/bloco e registro de e-mail | Crítica |
| SEI | Fila, credencial por solicitante, criação/reconciliação por rótulo e inclusão em bloco | Crítica |
| Assinatura | Realizada fora do aplicativo; não há validação criptográfica nem confirmação de assinatura interna | Alta |
| E-mail | `mailto:` abre cliente e registra clique como envio; não há servidor de e-mail | Alta |
| Inventário | Evento fechado/aberto/finalizado, comissão, salas, leitura individual/lote, divergências, sobras e fotos | Crítica |
| Consulta/relatórios | Busca global, ficha, análise por dimensões, painéis, filtros e XLSX | Alta |
| Importação/exportação de cadastros | Substituição integral de quatro tabelas e, opcionalmente, inventários | Crítica |
| Usuários | Funções cumulativas, ativação, senha temporária, bloqueio e acessos SEI/SPW cifrados | Crítica |
| Anexos | Fotos de inventário; não há gestão genérica de anexos documentais | Alta |
| PDF | Não foi encontrado gerador PDF patrimonial; DOCX/HTML são os documentos do sistema. `proposta/proposta.html` imprime proposta comercial à parte | Informativa |
| Operação | Backup SQLite, cron SPW, trabalhador, scripts destrutivos de preparação | Crítica |

## 4. Pontos fortes encontrados

- `permissoes.permitido()` nega endpoint/método não cadastrado, inclusive normalização de HEAD. A proteção é no servidor, não só no menu.
- `app.resolver_usuario()` consulta usuário ativo e funções a cada request. Inativação e mudanças de função afetam requests seguintes.
- `comissoes.py` autoriza por ID de usuário e evento; homônimos não recebem acesso por nome na web. Há testes dedicados de escopo.
- CSRF em todos os POSTs roteados, token aleatório e comparação constante; login também protegido; cookies HttpOnly, SameSite=Lax e Secure na configuração web.
- Senhas com hash Werkzeug/scrypt; credenciais externas com Fernet; chave de criptografia fora do banco; redefinição temporária com `Cache-Control: no-store`.
- SQLite habilita foreign keys por conexão e timeout de 30 segundos; índices parciais garantem processo vigente por tipo e pedido ativo por termo/SPW.
- `app_cadastros.confirmar()` assina dados/estado e revalida sob `BEGIN IMMEDIATE`. `usuarios.editar()` protege o último administrador sob trava de escrita.
- Importações principais validam diversas relações e têm rollback. Bens atribuídos ausentes impedem a substituição da base.
- HTML dos termos escapa conteúdo; SQL de valores é parametrizado; ordenações/filtros dinâmicos consultados usam listas permitidas.
- XLSX de saída neutraliza fórmulas em textos por `db.acrescentar_linha()`; a fórmula IMAGE intencional escapa aspas.
- Inventários finalizados possuem snapshot dos bens. Robô SEI tenta localizar documento existente antes de criar outro; falha de bloco preserva documento.
- Backup existente usa a API `.backup` do SQLite, faz `integrity_check` e envia cópia externa. Downloads de documentos são gerados em memória.

## 5. Vulnerabilidades de segurança

| ID | Severidade | Problema | Local | Impacto | Recomendação |
|---|---|---|---|---|---|
| S01 | ALTO | Cookie antigo continua válido após troca de senha | `app.py:115`, `usuarios.py:230–255`, `app_usuarios.py:84–102` | Persistência de acesso após comprometimento de sessão | Versão de sessão por usuário e revogação na troca/reset; política explícita de logout |
| S02 | MÉDIO | GET do DOCX modifica histórico sem CSRF | `app.py:402–427 → termo_docx` | Emissão induzida e timestamp alterado por navegação | POST autenticado/CSRF para emitir; GET exclusivamente de artefato existente |
| S03 | ALTO | Restrições de estado SEI ficam só na tela | `app.py:568–577`, `db.py:953–958`, `templates/termo_emitido.html` | Adulteração ou corrida com registros de documentos já emitidos | Transições verificadas no servidor; correção excepcional auditada |
| S04 | MÉDIO | Modo sem login é o padrão fora da imagem | `config.py:69`, `app.py:126` | Publicação acidental como administrador local | Modo desktop explicitamente opt-in; validação de configuração no startup web |
| S05 | MÉDIO | URLs de fotos importadas sem validação de esquema/origem | `inventario.py:720–875`, templates de fotos | Links ativos maliciosos/rastreamento e associação indevida de objetos | Validar HTTPS, origem, chave e vínculo; tratar links legados explicitamente |
| S06 | MÉDIO | Recursos de upload não limitados após descompressão | `db.py:338`, `db.py:1206`, `fotos.py:29–61` | Esgotamento de memória/CPU por usuário com upload | Limites de ZIP, linhas, células, dimensões de imagem e tempo |
| S07 | MÉDIO | Dependências sem lock e pip local com avisos conhecidos | `requirements*.txt`, `Dockerfile:15–30`, `.venv` | Build não reproduzível e risco na instalação de pacotes | Inventário por alvo, pins/hashes e atualização validada da ferramenta de build |
| S08 | BAIXO | Redirecionamento de erro confia em Referer | `app.py:197–206` | Redirecionamento externo induzido | Reutilizar validação de retorno local |
| S09 | MÉDIO | Trabalhador não revalida ativo/função do solicitante | `usuarios.py:291–338`, `atender_pedidos.py:65–113` | Pedido pendente executado com credencial de conta revogada | Definir política de revogação e checá-la antes da operação externa |

### S01 — Revogação insuficiente de sessões — P1 — CONFIRMADO

Evidência: a sessão contém `usuario_id`; o resolvedor só verifica existência, ativo e funções. `trocar_senha()` troca o hash, mas não atualiza um identificador de sessão. Reprodução: autenticar usuário de teste, guardar cookie, trocar senha pela função de domínio, usar o cookie guardado em outro cliente; `/` respondeu **200**. Pré-condição de exploração: atacante já obteve cookie válido e consegue alcançar a aplicação, inclusive passando eventual Basic Auth do proxy.

Reset temporário restringe inicialmente o usuário à troca de senha; depois da troca legítima, o cookie anterior volta a alcançar as rotas. Logout limpa somente o cookie do cliente que saiu. A duração configurada é 12 horas com sessão permanente; renovação automática exige distinguir inatividade de prazo absoluto. Recomendação: `session_version` no banco e cookie, incrementada na troca/reset/inativação conforme política; reautenticação; decidir logout global versus local. Testar dois clientes, reset seguido de troca e expiração. Não é necessário armazenar todas as sessões em Redis.

### S02 — Emissão por GET — P1 — CONFIRMADO

`termo_docx()` chama `registrar_emissao()` antes de responder ao download. Uma chamada GET isolada gerou DOCX com status 200 e incrementou o número de emissões em 1; HEAD já é protegido contra esse efeito. Um link externo aberto pelo operador em navegação de topo pode levar cookie SameSite=Lax e gerar emissão sem intenção/CSRF. Também há risco de ferramentas de pré-visualização. Impacto limitado aos privilégios do usuário que abriu o link, sem escalada para perfil Consulta. Emitir via POST e tornar o download histórico um GET sem mutações.

### S03 — Integridade do estado SEI — P1 — CONFIRMADO por código

O template oculta o formulário em `andamento`/`concluido` e torna documento readonly em `erro_bloco`. O endpoint aceita POST direto com CSRF de qualquer operador autorizado, sem checar esses estados. `salvar_documento_sei()` permite substituir/apagar documento e bloco, mantendo `id_documento` antigo via COALESCE. Um operador pode alterar o vínculo de documento concluído ou concorrer com o robô; número visível e hiperlink podem apontar para documentos distintos. Não é IDOR entre setores: o modelo atual concede acesso global ao acervo para operadores. É ausência de invariantes de estado e auditoria.

Definir correção manual explícita, com motivo, antes/depois, autor e validação de compatibilidade do ID. Bloquear edição durante execução e impedir que apagar documento seja usado para contornar o congelamento da numeração. As especificações antigas autorizavam edição ampla; reconciliar esse contrato com a automação nova, preservando um caminho legítimo de correção.

### S04 — Configuração permissiva fora do Docker — P1 — CONFIRMADO condicional

Qualquer valor de `TERMOS_LOGIN` diferente de `1`, inclusive ausência ou erro de digitação, ativa administrador local. Dockerfile/compose declaram `1` e `main.py` escuta loopback: **não foi comprovada exposição anônima em produção**. O risco aparece ao servir `app:app` fora desse caminho. `app.py` executado diretamente usa debug, apenas em loopback; não confundir com debug no Waitress. Fazer modo web falhar sem configuração válida e manter modo desktop explícito. Verificar Host/entrada desktop contra exposição involuntária; DNS rebinding não foi testado.

### S05 — URLs não confiáveis — P1 — PROVÁVEL

`validar_abas()` apenas exige URL não vazia e importa `foto_url` de sobras/legado sem allowlist. Templates fazem `href="{{ url }}"` e `src="{{ url }}"`; escape HTML não valida protocolos. Cenário: administrador importa planilha externa com link `javascript:`/origem de rastreamento e outra pessoa clica na foto. Execução do JavaScript com `target=_blank` depende do navegador e não foi demonstrada: **não classificar como XSS executado**. URLs públicas previsíveis produzidas por `fotos._url()` também não passam pelas permissões do Flask; exposição real depende da configuração do bucket, não inspecionada.

Rejeitar protocolos ativos, restringir host/chave, validar associação ao evento e separar identificador de objeto de URL de exibição. `fotos.apagar()` aceita trecho legado `inventario/` em qualquer URL: validar que o objeto pertence ao registro antes de excluir. Se as fotos forem restritas, usar URLs temporárias ou entrega autenticada; não presumir que o bloqueio de telas protege o bucket.

### S06 — Upload e consumo de recursos — P2 — PROVÁVEL

Há limites úteis: request de 20 MB e foto de 5 MB. Porém XLSX é ZIP/XML; `read_only=True` não limita tamanho descompactado e o código acumula linhas em listas. Foto pequena em bytes pode expandir antes do resize. Pillow possui proteções próprias, mas não há orçamento explícito de pixels e `comprimir()` pode falhar fora do tratamento de validação. Não foram enviados ZIP bombs nem imagens de estresse.

A documentação do [openpyxl](https://openpyxl.readthedocs.io/en/3.0/) recomenda proteção XML adicional; `defusedxml` não apareceu no inventário local. Isso, isoladamente, **não comprova XXE**, pois parser, lxml e Python têm comportamentos próprios. Validar a configuração efetiva e adicionar casos pequenos, seguros, de DTD/ZIP desproporcional. Limitar comprimento/quantidade também na criação de documentos e campos livres.

### S07 — Supply chain e inventário de CVEs — P1 — CONFIRMADO no ambiente local

Nenhuma dependência de `requirements.txt`, `requirements-robo.txt` ou instalação do Dockerfile tem versão/hash fixado. Dockerfile não consome os requirements; `waitress` existe na imagem, mas não no requirements principal. `python:3.12-slim` é tag móvel. Não se pode reproduzir uma release apenas pelo commit. Desktop traz pywebview/pyinstaller; mantê-los separados do runtime web é adequado.

Consulta somente de nomes/versões de **36 distribuições da `.venv`** à API OSV em 21/09/2026 retornou avisos apenas para `pip 25.1.1`: seis GHSA e seis aliases PYSEC, correspondentes aos mesmos seis problemas, não doze vulnerabilidades independentes:

| CVE | GHSA | Versão corrigida indicada pelo OSV |
|---|---|---|
| CVE-2025-8869 | GHSA-4xh5-x5gv-qwph | 25.3 |
| CVE-2026-1703 | GHSA-6vgw-5pg2-w6jp | 26.0 |
| CVE-2026-3219 | GHSA-58qw-9mgm-455v | 26.1 |
| CVE-2026-6357 | GHSA-jp4c-xjxw-mgf9 | 26.1 |
| CVE-2026-8643 | GHSA-wf93-45jw-7689 | 26.1.2 |
| CVE-2026-13346 | GHSA-qwm4-qh6w-59xr | 26.2.0 |

São riscos de instalação/extração/resolução de pacotes, dependentes de artefato/índice controlado pelo atacante; não são endpoints Flask exploráveis demonstrados. O [changelog oficial do pip](https://pip.pypa.io/en/stable/news/) confirma correções de traversal, nomes de scripts e dupla decodificação. Validar atualização para versão que cubra todos os avisos, registrando ambiente de build e testes. A API de evidência é `https://api.osv.dev/v1/querybatch`; detalhes por `https://api.osv.dev/v1/vulns/<GHSA>`.

Versões locais relevantes: Flask 3.1.3, Werkzeug 3.1.8, Jinja2 3.1.6, openpyxl 3.1.5, python-docx 1.2.0, Pillow 12.3.0, cryptography 50.0.1. Avisos públicos de Werkzeug para versões anteriores não devem ser atribuídos automaticamente à 3.1.8. Não foram auditados todos os pacotes da `.venv-robo`, imagens em execução, SO/Chromium ou JS vendorizado; ausência de resposta OSV para os demais 35 pacotes não prova segurança completa nem abandono de bibliotecas.

### S08 — Retorno externo por erro — P2 — CONFIRMADO estático

Os handlers de negócio e 413 redirecionam para `request.referrer` sem validação. Um GET com filtro inválido, vindo de site externo, pode retornar a esse site; clientes também podem fornecer Referer arbitrário. Não há execução de comandos nem vazamento de cookie demonstrado. Aplicar a mesma abordagem de reconstrução local usada em `app_cadastros.retorno()`.

### S09 — Revogação não alcança a fila — P1 — CONFIRMADO estático; política a validar

`credencial_sei()`/`credencial_spw()` consultam login e credencial sem `ativo`/função. Um usuário enfileira, é inativado ou perde permissão e o trabalhador ainda usa sua credencial. Pré-condição: pedido anterior autorizado e credencial ainda válida no sistema externo. Confirmar se o negócio quer executar pedidos previamente aprovados ou cancelar na revogação; registrar identidade estável e decisão. Para contenção de conta comprometida, a revogação deve impedir novos efeitos externos pendentes.

### Cobertura negativa e riscos de configuração

- Não identificado SQL injection explorável nos caminhos examinados: interpolação de identificadores vem de constantes/allowlists; valores são parametrizados. Não há ORM.
- Não identificado caminho HTTP de command injection, path traversal de upload, SSRF servidor a partir de URL de usuário, mass assignment de privilégios ou execução de templates fornecidos pelo usuário. Fotos/IMAGE podem disparar requisições do navegador/Excel, o que é diferente de SSRF servidor.
- `textos.validar()` restringe marcadores, conversões e format specs; `termos_html.esc()` escapa conteúdo. Isso reduz template injection e XSS no texto documental.
- Nenhum `.env`, banco ou arquivo de `secrets/` apareceu na consulta de caminhos versionados feita. Não houve varredura de todos os commits históricos nem prova de inexistência de segredo em todo o histórico Git.
- Nota do spike SPW (`docs/superpowers/notes/2026-09-17-robo-spw/README.md`, seção Credenciais) registra senha compartilhada em conversa e recomenda troca. **HIPÓTESE A VALIDAR:** confirmar rotação sem recuperar/publicar a senha. Capturas/JSON de evidências versionadas merecem revisão de dados pessoais antes de compartilhar o repo.
- Nginx lido ainda declara Basic Auth, divergindo da etapa do README que orienta sua retirada após login próprio. Isso adiciona uma barreira, mas pode causar dupla autenticação e contas compartilhadas no proxy. Não foi alterado.
- Ausência de CSP/frame-ancestors e política de cache explícita no app é dívida de hardening, não prova de exploração. Cabeçalhos globais de Nginx não foram integralmente auditados. Definir CSP compatível com scripts inline/iframe antes de habilitá-la.
- Login bloqueia a conta por cinco falhas, mas contador usa read-modify-write sem serialização e a mensagem de bloqueio diferencia conta existente. **PROVÁVEL:** concorrência pode subcontar falhas; ataques podem bloquear contas conhecidas. Testar paralelismo e limites no proxy sem eliminar defesa por conta.

## 6. Bugs identificados

| ID | Prioridade | Bug / confiança | Como reproduzir com dados sintéticos | Impacto | Local |
|---|---|---|---|---|---|
| B01 | P0 | Importação vazia substitui acervo — CONFIRMADO | Remover atribuições da fixture; enviar XLSX só com cabeçalho válido: total vira 0 | Perda de cadastro; leituras abertas ficam sem base | `db.importar_bens:329–403`, `app.upload:599` |
| B02 | P1 | Identificador fracionário truncado — CONFIRMADO | Importar número 1001.9: bem gravado como 1001 | Associação de bem errado; pode atingir atribuição preexistente | `db.importar_bens:351–357` |
| B03 | P0 | Número de termo duplicado — CONFIRMADO | Dois termos, mesma unidade/tipo/ano; barreira depois do cálculo, antes do UPDATE: ambos 01/2026 | Ambiguidade documental e reconciliação SEI incorreta | `db.proximo_numero_termo:997`, `preparar_envio_sei:1011` |
| B04 | P1 | Snapshot reutilizado com conteúdo diferente — CONFIRMADO | Emitir; alterar descrição/valor; emitir mesmos números no dia: mesmo ID e snapshot antigo | DOCX novo diverge do histórico; indicador diz vigente | `db.registrar_emissao:907`, `situacao_termo:1106` |
| B05 | P1 | Documento histórico reconstruído com dados atuais — CONFIRMADO | Alterar responsável após emissão e chamar `_html_do_registro`: HTML muda | Reenvio deixa de representar o termo original | `app._html_do_registro:444`, `db.salvar_centro:574`, `_renomear_pessoa:812` |
| B06 | P1 | Formulário inválido grava documento/bloco — CONFIRMADO | POST número INVALIDO e documento 999: resposta 302 de erro, documento 999 persistido | Alteração parcial silenciosa | `app.termo_emitido_documento:568–577` |
| B07 | P0 | Uploads concorrentes usam mesma chave — CONFIRMADO | Duas conexões em `adicionar_foto`; barreira no callback enviar: ambas `probe/1-1001.webp`; uma IntegrityError | Uma foto pode sobrescrever outra no R2 antes de erro SQL | `inventario.adicionar_foto:409–426` |
| B08 | P0 | Leitura após finalizar inventário — CONFIRMADO | Pausar após `_evento_aberto_ou_erro`; finalizar por outra conexão; retomar leitura: ela é inserida | Evento finalizado alterado e snapshot incompleto | `inventario.ler:333`, `encerrar_evento:199` |
| B09 | P1 | Exclusão externa parcial mantém referências inválidas — CONFIRMADO por código | Callback apaga primeira foto e falha na segunda de `excluir_evento`/`desfazer_leituras` | Banco preservado referencia objeto já apagado; mensagem “nada alterado” incorreta | `inventario.py:241–256,474–491`, `app_inventario._apagar_no_bucket` |
| B10 | P1 | Responsabilidade individual não é única no banco — CONFIRMADO | Inserir pessoa B e atribuí-la ao número já de A: duas linhas aceitas; migração permite o mesmo | Dois termos individuais para o mesmo bem; joins duplicam totais | `db.ESQUEMA:43–47`, `importar_planilhas.migrar:61–67` |
| B11 | P0 | Migração inicial destrutiva falha pela metade — CONFIRMADO | Fixture populada; `migrar()` com export inválido: pessoas/atribuições já apagadas | Perda de cadastros antes da validação | `importar_planilhas.py:24–32` |
| B12 | P1 | Prévia copiada não corresponde necessariamente ao registro — CONFIRMADO por código | Abrir iframe; mudar base em outra sessão; copiar: copia HTML antigo, POST lê bens atuais | Histórico não prova conteúdo entregue | `templates/termo.html:41–66`, `app.termo_registrar:431` |
| B13 | P2 | Clique em mailto marcado como envio — CONFIRMADO por código | Clicar e cancelar cliente de e-mail: POST ocorre após 500 ms | Falso registro de comunicação | `templates/termo_emitido.html:73–78`, `db.registrar_email:968` |
| B14 | P2 | Histórico limita a 200 sem navegação completa — CONFIRMADO por código | Criar 201 emissões elegíveis: uma fica fora e não há página seguinte | Registros antigos invisíveis na listagem usual | `db.termos_emitidos:929`, `templates/termos_emitidos.html` |
| B15 | P1 | Importação remove bem conferido de evento aberto — CONFIRMADO por código | Ler bem sem atribuição; importar conjunto sem esse número; relatório faz JOIN com bens atuais | Leitura existe mas desaparece do relatório; encerramento não captura o bem ausente | `db.importar_bens`, `inventario.relatorio:576`, `encerrar_evento:209` |
| B16 | P2 | Exportação histórica pode não ser reimportável — CONFIRMADO por código | Finalizar evento; remover bem do acervo; exportar cadastros; reimportar: validação da leitura exige bem atual | Snapshot histórico válido não restaura por XLSX | `inventario.validar_abas:781`, `db.importar_cadastros` |

### Orientação de correção dos bugs

- **B01/B02:** preparar e validar todas as linhas antes de qualquer DELETE; rejeitar vazio, número não inteiro/não finito e linha inválida com localização do erro. Prévia de removidos/atribuídos/conferidos e confirmação assinada de impacto; para grande redução, aprovação explícita. O robô já exige pelo menos 90% da contagem, mas o upload manual não. Preservar opção legítima de baixa em massa com procedimento revisável.
- **B03:** reservar número sob transação que adquira escrita antes de ler contador; constraint de unicidade por unidade/tipo/ano/sequencial. Unificar número manual e automático. Não deduplicar registros antigos silenciosamente: reconciliar documentos SEI e registrar decisão. Colisão é especialmente perigosa porque o robô procura documento por rótulo e, para pessoas, usa apenas primeiro nome (`robo_sei:431–447`).
- **B04/B05/B12:** snapshot completo e imutável com identidade do responsável, textos, data, bens, valores, processo e representação/hash do documento. Deduplicação por intenção/conteúdo, não só números/dia. Preview, download, registro e fila devem referenciar a mesma versão. Não inventar snapshots antigos ausentes: marcar legado como incompleto.
- **B06:** validar todo o formulário antes de persistir; gravar número/documento/bloco/ID coerentemente em uma transação. Remover efeitos do caminho `finally` no futuro. Erro deve preservar todas as colunas.
- **B07:** identificador de objeto único por upload, não MAX+1 sem reserva. Não manter transação SQLite aberta durante rede; reservar operação curta, enviar, finalizar e reconciliar pendências. Preserve a ordem de exibição por sequência independente da chave do bucket.
- **B08:** checar aberto e escrever na mesma seção transacional; testar leitura, edição, foto, sobra e comissão contra fechar/finalizar/excluir. Permissão por comissão também precisa política contra revogação concorrente.
- **B09:** usar estados de exclusão/limpeza externa reexecutável e log por objeto; só prometer atomicidade do banco, não do R2. Soft delete/retensão de evidência reduz dano; não adotar transação distribuída.
- **B10:** reconciliar duplicados, adicionar UNIQUE(numero), usar a mesma regra na migração. Na rota atual, revisão com BEGIN IMMEDIATE e DELETE/INSERT já impede inferir automaticamente duplicação pelo fluxo web normal; o problema comprovado é no esquema/caminho legado.
- **B11:** script inicial deve recusar banco preenchido por padrão, validar todas as fontes e permitir rollback integral. Não executar novamente em produção para testar.
- **B13/B14:** chamar registro de “solicitação de envio” ou pedir confirmação humana; paginação/filtros por data e ID preservados ao abrir/voltar.
- **B15/B16:** preservar evidência mínima do bem desde a leitura, ou bloquear remoção na carga enquanto afetar inventário; no restore usar snapshot para validar leitura histórica, não só o acervo atual.

Outros riscos **PROVÁVEIS**: `termo_enviar_sei` faz emissão, numeração e enqueue em commits separados; falha pode deixar registro sem pedido. `atender_um()` seleciona pedido antes do flock e não o relê depois; duas instâncias podem executar o mesmo pedido sequencialmente. Requeue de órfãos usa tempo desde início, sem heartbeat; uma segunda instância pode reenfileirar trabalho ainda ativo. A configuração declara um trabalhador; não há evidência de duplicação operacional atual. Testar antes de permitir múltiplas instâncias.

## 7. Integridade dos dados patrimoniais

### 7.1 O que está protegido e o que falta

| Invariante | Situação atual | Ação |
|---|---|---|
| Número patrimonial único | PK em bens; import duplicado faz rollback | Preservar; validar inteiro antes da conversão (B02) |
| Uma pessoa por bem | Só PK composta nome/número; B10 comprovado | UNIQUE(numero), reconciliação prévia |
| Centro ou pessoa | Consulta exclui bens atribuídos do centro | Preservar regra; testar mudança em ambos os sentidos |
| Responsável/localização com histórico | UPDATE/DELETE sem evento de auditoria | Registro transacional de autor, antes/depois, motivo e bens afetados |
| Termo representa emissão original | Snapshot de bens parcial; B04/B05/B12 | Versão documental completa e identidade estável |
| Troca de responsável encerra termo anterior | Não existe lifecycle de validade/supersessão | Definir substituição explícita; não confundir rótulo “vigente” com assinatura |
| Baixado sem carga ativa | `bens_da_pessoa()` inclui BAIXADO e `atribuir()` aceita qualquer situação | **HIPÓTESE A VALIDAR:** decidir se deve impedir atribuição/emissão ativa; manter leitura de baixado como evidência legítima |
| Excluir pessoa com bens | Permitido com revisão; CASCADE remove atribuições | Não chamar isso bypass: é regra atual. Preferir inativação, transferência explícita e auditoria |
| Excluir centro com bens ativos | Bloqueado quando sob guarda; atribuições individuais excluídas da contagem | Testar futuras remoções de atribuição e centros já excluídos |
| Excluir processo com termos | Bloqueado e protegido por FK | Preservar |
| Inventário finalizado imutável | Snapshot presente, mas B08 e importação administrativa o alteram | Serialização e via explícita de retificação histórica |
| Lote completo ou resultado parcial explícito | `ler_lote()` chama `ler()` com commit por bem | Se falha no meio, informar quais foram gravados; decidir contrato de atomicidade |

A devolução aceita pessoa/nome e qualquer bem existente a partir da sessão; a URL não vincula a seleção à pessoa. **HIPÓTESE A VALIDAR:** confirmar se devolução por terceiro é permitida. Se for, exigir origem/justificativa; se não, validar pertencimento no servidor. Testar duas abas com seleções diferentes: ambas compartilham cookie e podem gerar termo com seleção alterada pela outra. Não remover responsabilidade automaticamente só por copiar um documento, pois isso não comprova entrega/assinatura.

`_mudancas()` registra apenas novo/removido/localização/situação. Valor, descrição, classificação e complemento mudam sem antes/depois; mapeamento de sala para centro e atribuições individuais nem entram nesse histórico. A ficha não permite responder “quem era responsável ontem e por que mudou”. Leituras sobrescrevem local, data e integrante; alterações de observação/conservação não registram autor. Nomes de integrantes sem ID na leitura impedem distinguir homônimos historicamente, apesar de a autorização por ID estar correta.

### 7.2 Contrato recomendado para o executor

Manter cadastro atual e histórico separados. Introduzir identidade interna estável para pessoa/centro e snapshot textual na emissão; preservar sigla/nome anteriores como dados históricos. Operações críticas devem receber ator e motivo e gravar evento na mesma transação. Termos novos podem referenciar termo substituído, com motivo e estado; termos legados continuam consultáveis com qualidade histórica indicada. Nenhuma etapa deve reescrever história para parecer completa.

## 8. Banco de dados

SQLite é proporcional ao sistema observado. Notas do SPW mencionam aproximadamente 7,4 mil bens em uma execução passada; isso não é medição da base atual. Não há justificativa demonstrada para trocar de banco.

- **Concorrência:** `timeout=30` equivale a busy_timeout de 30.000 ms; confirmado em banco temporário. `journal_mode` padrão observado foi `delete`; código não liga WAL. Não foi lido o modo da base real, que pode diferir. WAL deve ser avaliado com teste de carga e backup, não ativado como solução para B03/B07/B08: essas são corridas lógicas.
- **Transações:** SELECTs anteriores ao primeiro DML não adquirem reserva de escrita no comportamento atual do sqlite3. Usar `BEGIN IMMEDIATE` onde a decisão precisa ser serializada; evitar rede dentro da transação. `with conn` sozinho não torna a leitura inicial protegida.
- **Dados monetários:** REAL e somas float são usados em modelos/termos. Potencial de diferenças de centavos; não foi medido erro financeiro concreto. Fixar política de arredondamento e avaliar centavos inteiros/Decimal nas bordas; reconciliar antes de conversão em massa.
- **Datas:** bens guardam datas DD/MM/AAAA e filtros usam substrings; datas impossíveis/texto arbitrário não são rejeitadas pela importação. `datetime.now()` é naive e depende do TZ do processo; compose declara São Paulo, desktop depende do SO. Normalizar entrada e definir timezone/instante de negócio com testes de virada de dia/ano antes de mudar armazenamento.
- **FKs:** bens.localizacao não aponta a localizacoes porque salas não mapeadas são aceitas; isso é funcional. Snapshots não têm FK para bem atual, adequadamente preservando histórico. Leituras abertas também não têm FK para bem: B15 demonstra o custo. Chave textual de termo não referencia pessoa/centro, permitindo órfãos lógicos.
- **Constraints:** manter UNIQUE parcial de processo vigente e pedidos ativos. Acrescentar unicidade de atribuição e número documental após saneamento; avaliar unique parcial de evento aberto e chave normalizada de pasta após testar o fluxo que fecha evento anterior. CHECKs de booleanos/quantidade/estados devem corresponder a regras confirmadas.
- **Índices candidatos:** `atribuicoes(numero)`, `bens(localizacao,situacao)`, `importacoes_mudancas(numero,importacao_id)`, `termos_emitidos(tipo,chave,emitido_em,id)`, `termos_emitidos_bens(numero,termo_id)`, `inventario_leituras(evento_id,localizacao)`. Usar EXPLAIN QUERY PLAN e medidas antes/depois; não criar todos por reflexo.
- **Migrations:** `db.criar_esquema()` mistura criação, evolução por inspeção e limpeza de dados; há marcador específico só para acesso. Web e robô chamam a rotina no startup, e `depends_on` não espera migração concluída. **PROVÁVEL:** disputa em upgrade e schema parcial se falhar. Eleger um migrador, ledger de versão, transação por etapa e teste de bancos legados. Proibir limpeza histórica silenciosa sem relatório/backup.
- **Filesystem:** não foi comprovado uso de NFS. Exigir disco local compatível com locking para SQLite e lockfile; não colocar banco compartilhado em filesystem de rede sem validação específica.

## 9. UI/UX

Avaliação estática de templates/JS e respostas Flask; sem certificação visual ou de acessibilidade.

| ID | Tipo / prioridade | Evidência | Proposta e verificação |
|---|---|---|---|
| U01 | Bug de contexto, P2 | `termo_emitido.html:5` e `upload.html:29` fazem meta refresh a cada 5 s | Poll somente do status; arquivo selecionado, foco e rolagem devem sobreviver ao andamento |
| U02 | Acessibilidade, P2 | `inventario_sala.html` campos `.campo` sem label individual; upload em label com input hidden; `focar()` força foco após requests | Nome acessível com campo e patrimônio; operação por teclado; evitar roubar foco durante edição |
| U03 | Usabilidade/funcional, P2 | Tabela de trazidos não oferece conservação/foto; miniaturas mostram uma foto e +N; inventariante puro não pode abrir `/bem` | Exibir detalhe/fotos dentro do evento autorizado, sem conceder acervo global |
| U04 | Prevenção de erro, P2 | Formulários de excluir foto/sobra não pedem confirmação; lote desmarcar pede | Padronizar confirmação e alternativa de desfazer/retensão conforme B09 |
| U05 | Feedback, P2 | Autosave mantém valor digitado após erro; sucesso sem indicador por campo; requests podem chegar fora de ordem | Estado pendente/salvo/erro, retry e proteção contra sobrescrever edição mais nova |
| U06 | Permissões visuais, P2 | `upload.html` sempre mostra importar cadastros, embora operador receba 403 | Condicionar à mesma `pode('importar_cadastros','POST')`; manter proteção backend |
| U07 | Contexto de navegação, P2 | Ficha e termo emitido não carregam retorno da análise/lista; breadcrumb usa rota sem filtros | Retorno local preservando query/page; teste abrir→voltar |
| U08 | Clareza, P2 | “E-mail enviado” após clique e “Bens iguais aos de hoje” por números somente | Corrigir semântica junto a B04/B13; distinguir emissão, envio e assinatura |

Pontos positivos: idioma pt-BR, skiplinks, cabeçalhos de tabela com scope, breadcrumbs, macros compartilhadas, ícones com nomes acessíveis em diversas ações, vazio e erros em formulários, grids responsivos e tabelas com overflow. Cadastros já preservam filtros, página e âncora (`app_cadastros.retorno/voltar`); não reimplementar indiscriminadamente. JS restaura scroll de forms GET no mesmo pathname; não cobre todo fluxo de detalhe/retorno. O refresh integral está confirmado; “piscar” e perda efetiva de posição precisam teste de navegador. Contraste, zoom, tamanho de alvos e quebra mobile permanecem **HIPÓTESE A VALIDAR**.

## 10. Testes

### Existentes

`tests/` cobre dados, importações, DOCX/HTML/XLSX, textos, configurações, CSRF, login, funções, escopo de inventário, UI renderizada, cadastros, cofre, comissão/migrações, SPW e SEI com doubles. Não foi identificada suíte E2E de navegador nos testes versionados. Scripts Playwright em `docs/superpowers/notes/` são experimentos com serviços reais e não devem ser rodados como testes automatizados seguros.

### Executados

1. Copiados módulos Python, tests, templates, static e timbrado para `/tmp/astra-audit-f3b6p_jp`. Não foram copiados banco real, `dados/`, `secrets/`, `.git` ou ambiente virtual. Interpretador da `.venv` original usado somente como runtime; `PYTHONDONTWRITEBYTECODE=1` impediu novos caches Python.
2. Na cópia: `TERMOS_DADOS=<pasta temporária> PYTHONDONTWRITEBYTECODE=1 <repo>/.venv/bin/python -m pytest -q -p no:cacheprovider`. **Saída 0, sem falhas reportadas.** Coleta adicional: **3500 tests collected**. Não foi obtida medida de cobertura de linhas/branches e não se deve inventar percentual.
3. `ast.parse` em **61 arquivos** Python de raiz/tests: sem erro de sintaxe, sem geração de bytecode.
4. Inventário de pacotes e consulta OSV: 36 distribuições; resultado descrito em S07. Não houve instalação/atualização de pacotes.
5. Provas dirigidas em bancos sintéticos e test client, abaixo. Callbacks de upload eram falsos e concorrência foi controlada por barreiras, não por esperas probabilísticas.

| Prova | Resultado observado |
|---|---|
| XLSX vazio, sem atribuições | `EMPTY_IMPORT 0` |
| Número 1001.9 | `FRACTIONAL_ID 1001` |
| Mesmos números, descrição/valor alterados | `STALE_SNAPSHOT True CADEIRA vigente` |
| Responsável alterado após emissão | `HISTORY_RENDER_CHANGED True` |
| Número de termo inválido com documento 999 | `INVALID_FORM_WRITES 302 999` |
| GET de DOCX | `GET_MUTATION 200 1` |
| Cookie anterior à troca de senha | `OLD_COOKIE_AFTER_PASSWORD_CHANGE 200` |
| Dois uploads simultâneos | mesma chave duas vezes; `ok` e `IntegrityError` |
| Duas atribuições para 1002 | `DUPLICATE_ASSIGNMENT_SCHEMA 2` |
| Defaults SQLite sintético | `delete`, `30000` ms |
| Dois números simultâneos | `NUMBER_RACE ['01/2026', '01/2026']` |
| Finalização intercalada com leitura | `WRITE_AFTER_FINALIZED True 1` |
| Falha na migração inicial | `MIGRATION_FAILURE ImportacaoInvalida 0` pessoas restantes |

### Falhando e fragilidades

Não foram alterados testes. A suíte existente não apresentou falhas; as provas demonstram requisitos ausentes, não falhas já capturadas por ela. `ClienteComCSRF` insere token automaticamente, útil para testes funcionais mas incapaz de provar sozinho que cada formulário real envia token. `test_csrf.py` cobre a proteção com cliente cru e inspeção de fontes. Muitos testes verificam presença de texto/HTML; não provam comportamento de JS, câmera, foco ou clipboard. Hash rápido em fixture testa fluxo, não custo de scrypt. `SEIFalso` prova orquestração, não compatibilidade atual dos seletores reais.

### Matriz mínima recomendada

| Área | Antes de alterar | Resultado exigido após |
|---|---|---|
| Termo centro/individual/devolução | Fixtures com snapshot e processo, conteúdo conhecido e datas fixas | Mesmo conteúdo/hash em preview, DOCX/HTML, registro e fila; versão nova quando conteúdo muda |
| Numeração/emissão | Duas conexões com barreira; duplo clique; manual+automático | Número único por escopo; idempotência da mesma intenção; estados sem registros parciais |
| Responsabilidade/localização | Pessoa→pessoa, pessoa→centro, centro→centro, baixado e cadastro removido | Uma carga por bem; antes/depois/ator completos; termos anteriores preservados |
| Exclusão | Pessoa com bens, centro com bens, processo usado, evento/fotos | Bloqueios/revisões corretos; nenhuma exclusão externa não rastreável |
| Permissões | Todos os endpoints/métodos, funções cumulativas, evento alheio, anônimo, inativo | 403/redirect sem efeitos; frontend consistente; cookies revogados |
| Importação | Vazio, parcial, linhas inválidas, duplicado, fracionário, datas/valores, duas cargas concorrentes | Rejeição explicada sem alterações; confirmação vinculada ao conteúdo; histórico real |
| Exportação | Textos começando `=`, Unicode, filtros, 201+ termos, snapshot de bem removido | Sem fórmulas injetadas; não truncar export; restore de histórico válido |
| Inventário | Releitura, divergente, baixado, fechar durante ler/editar, lote com falha no meio | Nenhuma gravação após finalização; resultado parcial explícito ou rollback conforme contrato |
| Fotos | Upload paralelo, falha no envio/DB, exclusão parcial, chaves legadas, URL maliciosa | Chaves únicas, compensação observável, autorização e validação de URL |
| SEI/SPW | Falha antes/depois de criação, timeout, repetição, revogação, dois trabalhadores | Reconciliação sem duplicar/associar documento errado; fila consistente |
| UI | Browser real com rede lenta/falha, teclado, viewport estreita | Sem perda de seleção/filtro/foco; feedback de erro não finge salvamento |
| Recovery | Backup criado durante escrita sintética + objetos/chave fictícios | Restore íntegro, FK ok, fotos/chaves disponíveis, sem reenviar pedidos sem reconciliação |

## 11. Arquitetura e código

Separação de módulos de domínio sem Flask é uma boa base; não exige uma nova arquitetura. Hotspots: `db.py` 1.489 linhas (DDL, imports, regras, consultas e relatórios), `inventario.py` 905 (estado, validação, fotos, export), `app.py` 729 (segurança, rotas, documentos e fila), `robo_sei.py` 470 e `app_cadastros.py` 410 (closures de rotas e revisão). Tamanho é indicador de concentração, não defeito por si só.

**D01 — P2 — Dívida técnica:** propriedade das transações inconsistente. Helpers dão commit e podem ser chamados dentro de outra operação, enquanto outros deliberadamente não dão commit. Ex.: criação de emissão→preparo→enqueue; edição de usuário→atualização de nome na comissão. Definir uma transação por caso de uso e separar helpers internos sem commit, somente após testes de comportamento. Não extrair camadas genéricas que só repassem argumentos.

**D02 — P2 — Dívida técnica operacional:** evoluções de schema por inspeção/ALTER e limpeza automática sem ledger completo; scripts de inicialização/migração têm contratos diferentes. Introduzir um mecanismo simples versionado, sem exigir framework novo.

**D03 — P2 — DX e documentação:** não foram encontrados CI/GitHub Actions, lint/formatter/type-check configurados nem lock de dependências. `docs/SYSTEM_IMPROVEMENT_PLAN.md` contém apenas texto de teste de Git, não um plano válido. README tem boas instruções, mas chama dependências do Dockerfile de “fixas” apesar de não fixar versões e repete “backup = copiar dados/”. Falta template de ambiente sem segredos e procedimento verificável de restore. `scripts/backup.sh` existe localmente e está ignorado pelo Git, portanto não acompanha um clone novo.

**D04 — P3 — Legado:** JS/CSS reutilizados mencionam Django, HTMX e `/tabela`, embora o sistema seja Flask. `dsgov.js` tem handler HTMX que aponta para `/login/?next=`, distinto de `/login?proximo=`; HTMX não está carregado na base inspecionada, portanto isso é código residual, não bug de login atualmente comprovado. Remover apenas após inventário de usos. Duplicação de formatação monetária e construção DOCX é pequena; centralizar só o que os testes mostrarem divergir.

Contratos em dicionários extensos dificultam evoluir snapshots/pedidos. Tipagem pontual de estruturas críticas ajuda; não é necessário tipar retroativamente todo o app para corrigir P0. Exceções de negócio possuem mensagem amigável, mas IntegrityError/OperationalError de concorrência podem virar 500. `textos.obter()` descarta configuração inválida silenciosamente: registrar diagnóstico sem expor dados.

## 12. Performance

**Não foram medidos tempos/capacidade de produção.** A classificação abaixo evita confundir otimização preventiva com gargalo observado.

| ID | Evidência | Classificação | Ação proporcional |
|---|---|---|---|
| P01 | `situacoes_centros/pessoas` fazem consultas de bens e termo por item; `painel()` chama ambas | N+1 evidente pelo código | Medir queries/latência por volume; agregar carga e últimos termos em consultas por conjunto |
| P02 | `bens_da_sala()` busca fotos por cada bem lido; `eventos_visiveis()` faz vínculo por evento; `usuarios.listar()` busca funções por usuário | N+1 evidente | Buscar fotos/vínculos/funções em lote e preservar escopo |
| P03 | Export recorte usa `recorte(limite=None)`, calculando também dimensões que não exporta | Trabalho redundante evidente | Separar consulta de linhas do cálculo de gráficos |
| P04 | XLSX/docx e relatórios acumulam resultados e rodam sincronamente; quatro threads Waitress | Risco preventivo, sem saturação medida | Benchmark sintético compatível com volume real; limites; export streaming quando demonstrado |
| P05 | Relatório filtra/ordena em Python; busca normaliza campos por item/palavra | Custo evidente, impacto não medido | Manter para poucos milhares se satisfaz objetivo; medir antes de reescrever SQL |

Estabelecer orçamento acordado de latência e memória; registrar tamanho da fixture, hardware e contagem de queries. Aumentar threads não resolve bloqueio SQLite. Fila já existe para integração demorada; não adicionar fila de relatórios sem evidência de necessidade.

## 13. Observabilidade e diagnóstico

**O01 — P1:** trilha de auditoria insuficiente para patrimônio: `historico_do_bem()` junta importações e termos, mas não atribuições, edição de responsável, alterações de textos ou ator de importação. Pedido tem `criado_por` textual, leitura tem integrante textual, usuário tem último acesso; isso não basta para atribuir todas as ações inequivocamente. Acrescentar evento transacional com ID estável do ator, origem, entidade, antes/depois, motivo e correlação de termo/importação/request; nunca senha, cookie ou token.

**O02 — P2:** há arquivo de traceback do trabalhador, screenshots de erro e mensagens em tabelas; não há request ID nem esquema uniforme de logs, health/readiness, métricas ou alerta de fila. `_registrar()` ignora OSError. Logs locais crescem sem rotação declarada. Incluir logging estruturado simples, rotação, latência, resultado, fila atrasada e idade do último backup; logs de exceção não devem depender da transação que falhou. Um endpoint de saúde não deve expor segredos ou inventário.

Não seria possível responder de forma completa hoje quem mudou determinada responsabilidade, valor anterior e termo afetado. É possível rastrear parcialmente localização/situação importada, emissão, passo do robô e último leitor. Capturas de SEI/SPW e mensagens externas podem conter dados pessoais: acesso restrito, retenção e sanitização; não foi comprovado vazamento público delas.

## 14. Backup, Recovery e continuidade

**R01 — P1 — CONFIRMADO quanto ao escopo do script:** `scripts/backup.sh` copia somente banco, valida integridade, compacta, envia por rclone, remove remotos acima de 30 dias e locais acima de 7. Não copia timbrado customizado, segredo de sessão, chave Fernet, configurações/credenciais nem fotos do bucket. README reconhece que `chaves.env` exige guarda separada. Não foram confirmados agendamento efetivo, sucesso recente, criptografia/permissões do destino, versionamento de fotos, restore testado ou RPO/RTO. Retenção remota usa credencial com permissão de deletar: cópia externa não equivale a backup imutável.

**R02 — P1 — CONFIRMADO estático:** README/compose recomendam copiar `dados/`; `atualizar_base.sh --teste` usa `cp dados/termos.db` com aplicação potencialmente ativa. Cópia crua pode não representar snapshot consistente e, sob WAL, omitir mudanças. Usar `.backup`/API SQLite ou parar todos os escritores, sem executar isso na produção durante a auditoria. O modo `--teste` ainda faz login/download real do SPW; não é teste hermético.

**R03 — P1 — PROVÁVEL:** scripts `zerar_banco.sh` e `apagar_termos_emitidos.sh` verificam pedidos antes da limpeza, mas não adquirem exclusão coordenada contra web/robô. Outra requisição pode enfileirar após a verificação. Há cópia `.backup`, porém apagar histórico reinicia numeração e não apaga documentos externos; isto pode reutilizar número existente no SEI. Restringir esses scripts a ambiente explicitamente descartável, com pré-condições verificadas e identidade da base, preservando capacidade de manutenção autorizada.

### Procedimento futuro verificável

1. Definir RPO/RTO com responsável operacional e frequência compatível; registrar evidência de execução e alertar atraso/falha.
2. Backup consistente de banco + manifesto de versão/schema/hash; cópia protegida de timbrado e chave Fernet em custódia separada; inventário/retensão dos objetos R2. Segredo de sessão pode ser rotacionado no restore para invalidar sessões, em vez de restaurado automaticamente.
3. Restaurar em diretório/host temporário sem saída para SEI/SPW/R2 e sem trabalhador automático; `integrity_check` e `foreign_key_check`; conferir contagens, snapshots, amostras e descriptografia de credencial fictícia.
4. Restaurar fotos/chaves e validar referências. Reconciliação de pedidos/documentos externos antes de retomar fila: banco antigo pode esquecer documento já criado.
5. Medir tempo real de restauração e idade dos dados; documentar resultado, responsável e periodicidade. Não declarar continuidade assegurada somente por `integrity_check`.
6. Rollback de release com imagem identificável e backup de schema anterior. Após migrations incompatíveis, rollback de código sozinho não basta; testar restore/forward fix e reconciliação de escritas externas.

## 15. Melhorias funcionais propostas

| ID | Problema resolvido e benefício | Complexidade | Risco | Dependências | Prioridade |
|---|---|---|---|---|---|
| F01 | Timeline de responsabilidade/valor e versões documentais: comprovar cadeia de guarda | Grande | Alto em legado | B04/B05/B10, O01 | P1 para base histórica; P3 para UI expandida |
| F02 | Prévia de importação com diferenças e conflitos: impedir remoção acidental e orientar saneamento | Média | Médio | B01/B02/B15 | P1 |
| F03 | Relatório de inconsistências: duplicados, carga baixada, sala sem centro, termo divergente, foto ausente, pedido parado | Média | Baixo se somente leitura | Regras explícitas e snapshots | P2 |
| F04 | Retificação/substituição de termo com motivo, versão e referência SEI | Média | Alto | B03/B05/S03, O01 | P2 |
| F05 | Alertas operacionais de backup/fila/importação: evitar falhas silenciosas | Pequena | Baixo | O02/R01 | P2 |
| F06 | Detalhe de bem dentro do inventário autorizado, inclusive fotos/divergentes | Média | Médio | Testes de escopo, U03 | P2 |
| F07 | Confirmação real de assinatura por integração, se houver demanda e mecanismo oficial | Grande | Alto | Snapshot imutável, acordo com SEI | P3, condicionada |

Busca global, gráficos, leitura por câmera e operações em lote **já existem**. Não propor como se fossem ausentes. QR com link de consulta pode ser avaliado depois, sem liberar acervo a quem só tem inventário. API externa, anexos genéricos e notificações de negócio precisam caso de uso concreto; não entram no caminho crítico.

## 16. Melhorias que NÃO devem ser feitas

- Não migrar para Django, microsserviços ou SPA para corrigir as falhas: Flask/Jinja atende aos fluxos observados.
- Não introduzir Kubernetes, service mesh, event sourcing completo ou transações distribuídas.
- Não trocar SQLite por preferência. Primeiro corrigir transações, constraints e medir concorrência. PostgreSQL só com requisito real de escala/alta disponibilidade/contenção não resolvida.
- Não adicionar Redis/Celery/Broker apenas porque existe fila: uma fila SQLite e um trabalhador podem bastar com claim/idempotência corretos.
- Não prometer prova de assinatura a partir de download/cópia/clique em e-mail ou inclusão em bloco.
- Não reescrever todo DSGov/CSS; corrigir problemas concretos de foco, refresh e controles e preservar recursos acessíveis existentes.
- Não eliminar o modo desktop; torná-lo explícito sem enfraquecer web.
- Não reconstruir histórico passado com dados atuais como se fossem verdade histórica; não sanear duplicados automaticamente sem reconciliação.
- Não executar scripts de spike/limpeza contra SEI ou banco real como parte da suíte; não remover dependências desktop como “inúteis” ao produto inteiro.

## 17. Backlog consolidado

IDs agrupados compartilham uma entrega, mas os critérios individuais permanecem obrigatórios. Nenhum item desta tabela foi implementado.

| ID | Categoria | Prioridade | Complexidade | Risco | Dependências |
|---|---|---|---|---|---|
| B01/B02/F02 | Importação segura | P0 contenção; P1 prévia completa | Média | Alto | Regressões de import/rollback |
| B03 | Numeração documental | P0 | Média | Alto | Diagnóstico de duplicados; testes concorrentes |
| B07 | Identidade/concorrência de fotos | P0 | Média | Alto | Double de storage; testes de falha |
| B08 | Finalização e escrita | P0 | Média | Alto | Testes com duas conexões |
| B11 | Migração inicial | P0 | Pequena | Alto | Fixture de banco preenchido/erro |
| S01/S02 | Sessão e método HTTP | P1 | Média | Médio | Test client cru e dois clientes |
| S03/B06 | Estado/edição de documento SEI | P1 | Média | Alto | Transições definidas; SEIFalso |
| S04 | Separação web/desktop | P1 | Pequena | Médio | Testes de startup e empacotamento |
| S05 | URLs/escopo de fotos | P1 | Média | Médio | Inventário de URLs legadas; testes browser |
| S06 | Limites de recursos | P2 | Média | Médio | Fixtures benignas de limites |
| S07/D03 | Build, CVEs e CI | P1 ferramenta; P2 processo | Média | Médio | Inventário de imagens/desktop/robô |
| S08 | Retorno de erro | P2 | Pequena | Baixo | Teste Referer interno/externo |
| S09 | Revogação e fila | P1 | Média | Médio | Política de pedidos já autorizados |
| B04/B05/B12/F01 | Snapshot e histórico | P1 | Grande | Alto | B03, testes de conteúdo; preservação legado |
| B09 | Consistência banco/objetos | P1 | Média | Alto | B07, estados e retry |
| B10 | Unicidade de carga | P1 | Média | Alto | Reconciliação e backup validado |
| B15/B16 | Inventário histórico/importação | P1 segurança; P2 intercâmbio | Média | Alto | B01/B08, fixtures históricas |
| B13/U08 | Semântica de e-mail/estado | P2 | Pequena | Baixo | Contrato de negócio |
| B14/U07 | Paginação e retorno | P2 | Média | Baixo | Testes browser de contexto |
| U01–U06/F06 | UX operacional/acessibilidade | P2 | Média | Médio | Permissões e estados estabilizados |
| D01/D02 | Transações e schema versionado | P2 | Média | Alto | Correções P0/P1 e testes de legado |
| D04 | Resíduos/duplicação | P3 | Pequena | Baixo | Inventário de usos |
| P01–P05 | Consultas/performance | P2 se medido; P3 preventivo | Média | Médio | Baseline e equivalência de resultados |
| O01 | Auditoria transacional | P1 | Média | Médio | Identidade estável e casos de uso |
| O02/F05 | Diagnóstico e alertas | P2 | Média | Baixo | Logging e responsáveis definidos |
| R01/R02/R03 | Recovery e scripts | P1 | Média | Alto | Inventário de backup; ambiente isolado |
| F03/F04 | Reconciliação/retificação | P2 | Média | Médio | Histórico e constraints |
| F07 | Assinatura integrada | P3 condicionada | Grande | Alto | Fonte oficial e demanda validada |
| Auditoria visual ampliada | Conveniência visual | P4 | Pequena | Baixo | Acessibilidade/UX principal resolvidas |

## 18. Plano de execução por fases e testes

**Regra para todas as fases:** implementar apenas em trabalho futuro autorizado. Antes de mudar schema/fluxo, escrever a regressão correspondente e comprová-la falhando no código anterior em base sintética. Nenhuma fase libera alteração com teste obrigatório falhando. Antes de aplicar em dados reais, backup restaurável e reconciliação de legado; nunca “corrigir” duplicado apagando arbitrariamente. Manter escopo Flask/SQLite/DSGov e testar desktop e web quando afetados.

### Fase 0 — Proteção imediata, em cinco entregas pequenas

Pré-condições: branch/ambiente de execução futuros próprios; cópia descartável e sem integrações externas; baseline da suíte preservada. **Não esperar uma refatoração para proteger P0.**

1. B01/B02/B11: recusar import vazio/IDs inválidos; conter migração em banco preenchido; validar antes de gravar.
2. B03: diagnóstico somente leitura, reserva transacional e unicidade após reconciliação planejada.
3. B07: chave única e recuperação de falha, preservando URLs antigas.
4. B08: estado e gravação atômicos para todas as mutações de inventário.
5. R02/R03: corrigir procedimento de cópia e impedir limpeza acidental na operação ativa; ensaio mínimo de restore antecede qualquer aplicação real.

Módulos futuros: `db.py`, `inventario.py`, `fotos.py`, `importar_planilhas.py`, scripts, testes desses módulos; eventual migration versionada mínima, sem refatoração ampla.

Testes antes: reproduções B01/B02/B03/B07/B08/B11; duas conexões com barreiras; rollback de imports já existentes; esquema novo e legado.

Testes depois: mesma planilha inválida deixa banco idêntico; N emissões concorrentes têm N números distintos; duas fotos têm objetos distintos sem overwrite; escrita após encerramento falha sem efeitos; migração rejeitada mantém cadastros. Testes de backup em base sintética ativa e de scripts bloqueados na configuração de produção simulada.

Regressões: import legítimo de baixa em massa, numeração manual, URLs antigas, fotos ordenadas, lote e fluxo de abertura/fechamento. **DoD:** seis provas principais não reproduzem o dano, constraints verificadas, suíte passa e plano de saneamento/rollback existe; nenhuma alteração de legado sem conciliação.

### Fase 1 — Rede de segurança e build reproduzível

Pré-condições: contenções F0 com suas regressões passando. Não adiar esses testes locais para esta fase.

Módulos futuros: tests, configuração de CI, requirements/lock por alvo, Dockerfile, documentação de ambiente. Atualizar ferramenta de instalação S07 com validação; registrar versões efetivas de web/robô/desktop e bibliotecas JS.

Testes antes: baseline existente e regressões F0. Depois: pipeline isolado sem rede de integração/segredos, rotas/métodos fora da matriz negados, testes de schema legado→novo, smoke de geração de documentos e inicialização de cada alvo. Adicionar infraestrutura E2E local para foco, clipboard simulado e navegação; manter teste real de SEI fora do CI.

Regressões: PyInstaller, Python da imagem, Chromium/xlrd e dependências de fontes. **DoD:** clone limpo tem procedimento reproduzível, CI bloqueia falha, nenhum teste usa dados reais, avisos de dependência triados por alvo; número de testes e resultados registrados sem apresentar parametrização como cobertura total.

### Fase 2 — Segurança de sessão e correções funcionais delimitadas

Pré-condições: F1; política de sessões e correção manual SEI definida.

Módulos futuros: `app.py`, `app_usuarios.py`, `usuarios.py`, `config.py`, `permissoes.py`, templates de termo/e-mail, trabalhador. Entregar separadamente S01/S02/S04; S03/B06/S09; B13/B14/S08; validação de URLs S05 e limites S06.

Testes antes: dois clientes com cookie antigo, GET/HEAD sem efeito, POST CSRF, conta temporária, perfis; S03 via POST direto; B06 com todas as colunas comparadas; worker após revogação; URLs perigosas e legado permitido.

Testes depois: cookies antigos negados conforme política; emissão só por intenção POST; número inválido não grava nada; documento concluído protegido com correção auditável; usuário inativo não produz novo efeito externo indevido; mailto não afirma envio; navegação alcança 201º termo; retorno externo rejeitado.

Regressões: download em desktop, back/refresh do navegador, acesso SEI manual legítimo, cookies Secure e recuperação de admin. **DoD:** controles do servidor independem do template; nenhuma elevação de permissão; suite/E2E relevantes passam; estados e mensagens correspondem ao que foi persistido.

### Fase 3 — Integridade histórica e regras patrimoniais

Pré-condições: F0–F2; decisões documentadas para baixados, devolução por terceiro, validade/substituição de termo, import de inventário encerrado. Backup/restore mínimo validado antes de saneamento.

Módulos futuros: `db.py`, `inventario.py`, `app.py`, `termos_html.py`, geradores DOCX, importação/exportação, `robo_sei.py`; schema de snapshots/identidades/auditoria. Entregar B10 primeiro; depois B04/B05/B12/O01; depois B15/B16/B09 e fila atômica.

Testes antes: conteúdo exato de documento com data fixa; renomeação sem perda de identidade; mudança de textos/valores; edição entre preview e POST; restauração de histórico sem bem atual; falha parcial de bucket; duplo enqueue e retry SEI.

Testes depois: nenhuma dupla atribuição; termo emitido imutável; nova versão para conteúdo novo; HTML/copiar/DOCX/fila referenciam mesmo snapshot; autor e antes/depois completos; histórico continua disponível após baixa/import; exclusão externa tem estado rastreável; falhas não geram documento fantasma silencioso.

Regressões: referências de termos antigos, PDFs não existentes não devem ser inventados como dependência, fórmulas XLSX neutralizadas, reimport de planilhas antigas, atribuição centro/pessoa. **DoD:** constraints e testes de concorrência passam; cada operação crítica tem evento/auditoria; legado incompleto é explicitamente rotulado, sem reconstrução falsa; conciliação aprovada pelos responsáveis pelos dados antes de aplicação real.

### Fase 4 — Arquitetura, schema e performance proporcional

Pré-condições: invariantes F3 estabilizadas e medidas de queries/latência com dados sintéticos representativos.

Módulos futuros: `db.py`, módulos de casos de uso que vierem a ser extraídos, `inventario.py`, `comissoes.py`, migrations e startup Docker; D01/D02/P01–P05.

Testes antes: contratos de transação, falha após cada gravação, startup de dois serviços, versões legadas, contagens/totais/filtros. Depois: único migrador determinístico; funções internas não encerram transação do caso de uso; mesma saída com menos consultas; benchmark atende orçamento acordado sem aumentar memória indiscriminadamente.

Regressões: import/geração sob escrita concorrente, cursor/fechamento de conexão, timeouts, desktop Windows sem `fcntl`. **DoD:** separação reduz concentração demonstrada, não muda regras; EXPLAIN/contagem de queries sustentam índices; rollback/forward recovery testados; WAL só adotado se validar workload e backup.

### Fase 5 — UI/UX e acessibilidade

Pré-condições: endpoints/estados estabilizados, harness E2E e fixtures de permissões F1–F3.

Módulos futuros: templates de inventário/termos/upload, macros, `static/dsgov/js/dsgov.js` e CSS próprio, rotas de retorno/detalhe; U01–U08/F06.

Testes antes: capturar fluxo atual com 201 termos, filtros/página, três fotos, bem divergente, rede lenta e inventariante sem acervo. Depois: polling não perde arquivo/foco/scroll, voltar mantém query, operação inteira por teclado, nome acessível de campos/ícones, confirmação de exclusão, fotos acessíveis dentro do evento, erro de autosave visível sem sucesso falso.

Regressões: inicialização BRCard/BRTable, câmera, clipboard, comportamento mobile e menus cumulativos. **DoD:** fluxos críticos passam em desktop e viewport móvel, teclado e zoom; contraste medido; ausência de acesso adicional a eventos/acervo; revisão visual com evidências, sem redesenho geral.

### Fase 6 — Observabilidade e recuperação completa

Pré-condições: auditoria transacional de F3; cópia consistente e contenção de scripts de F0. Esta fase completa o recovery, não autoriza adiar backup básico.

Módulos futuros: logging de app/robôs, health/readiness, scripts/documentação de backup/restore, configuração operacional proposta; O02/R01/F05.

Testes antes: simular disco indisponível, falha de upload de backup, chave ausente, objeto ausente, fila sem worker. Depois: alerta útil com correlação, sem segredo; rotação de logs; restore em host limpo com integridade/FKs/snapshots/fotos; reconciliação de pedidos sem repetir emissão externa; rollback compatível com schema.

Regressões: volume de logs, informação sensível em screenshots, permissões de arquivos/container, falso health positivo. **DoD:** RPO/RTO medidos e aprovados operacionalmente; evidência de restore registrada; custodiante de chave e retenção definidos; procedimentos versionados sem credenciais e execução de alertas comprovada.

### Fase 7 — Evolução funcional justificada

Pré-condições: todos os bloqueadores P0/P1 pertinentes encerrados; integridade e observabilidade verificadas. Validar demanda antes de F07/API/QR adicional.

Módulos futuros: consultas e UI de timeline/reconciliação, retificação, alertas de negócio e adaptador oficial de assinatura apenas se aprovado. Entregar primeiro F03/F04, ampliando F01 sem reescrever histórico.

Testes antes: fixtures de inconsistências reais simuladas e regras de retificação; todos os controles de escopo. Depois: relatório detecta divergências esperadas sem falsos “saneamentos”; retificação mantém versão original e motivo; eventual assinatura só muda estado com evidência oficial validada.

Regressões: escopo de acesso, mistura entre informação atual e histórica, custo de painéis. **DoD:** cada funcionalidade responde a problema desta auditoria ou demanda documentada, possui critério de aceitação mensurável e testes; nenhuma depende de vulnerabilidade pendente ou presume que uma ação externa ocorreu sem evidência.

## 19. Receita das reproduções para implementação futura

Usar `tests.conftest.semear(conn)` sobre banco temporário criado com `db.criar_esquema()`. Não usar `db.conectar()` sem caminho/TERMOS_DADOS isolado. Desativar chamadas externas com doubles. Abaixo estão os pontos de sincronização que evitaram testes intermitentes:

- **B03:** cadastrar dois termos distintos cuja `unidade_sei` resolva a mesma sigla. Decorar temporariamente `proximo_numero_termo` no teste: calcular resultado original, aguardar `threading.Barrier(2)`, retornar. Cada thread tem conexão própria. Verificar unicidade no banco e número de efeitos externos simulados, não apenas ausência de exceção.
- **B07:** criar evento aberto e leitura. Callback `enviar(chave)` registra a chave numa lista, espera barreira e retorna URL fictícia. Duas conexões chamam `adicionar_foto`. Antes da correção: mesma chave e um IntegrityError. Depois: chaves distintas e duas operações recuperáveis/confirmadas.
- **B08:** interceptar a verificação de evento aberto apenas no teste. Após a verificação, finalizar pela segunda conexão, então retomar `ler`. O problema atual permite inserir; depois a operação deve ser serializada antes do encerramento ou recusada após ele, sem timeout artificial como critério de sucesso.
- **B04/B05:** emitir centro CCI, guardar ID/HTML; alterar descrição/valor mantendo números e responsável em etapa separada; repetir. Testar imutabilidade do registro antigo e criação explícita da versão nova.
- **B06:** cliente autorizado com token CSRF, POST `numero_termo=INVALIDO`, `documento_sei=999`, `bloco_sei=888`. Comparar todas as colunas antes/depois; status de redirect sozinho não é prova de rollback.
- **S01:** guardar cookie antes da troca; trocar com função/rota normal; instalar cookie guardado em segundo cliente. Repetir com reset temporário→troca legítima e com logout/inativação conforme política definida.
- **B01/B02:** Workbook em BytesIO com cabeçalho `list(db.COLUNAS_EXPORT)`; vazio ou linha de número 1001.9. Comparar hash lógico de todas as tabelas relevantes antes/depois da rejeição, incluindo histórico.
- **B11:** chamar `importar_planilhas.migrar` com fonte inválida em banco sintético preenchido. Verificar preservação das quatro tabelas e bens; testar falha também depois de ler parcialmente os cadastros.
- **B10:** testar UNIQUE(numero) por inserção direta para provar garantia de banco e por migração para provar o caminho real. Não usar só a rota web, que já contém uma proteção transacional adicional.

As provas temporárias não foram incorporadas aos testes do repositório nesta etapa, conforme a regra de análise exclusiva. Este documento especifica o comportamento a capturar, as pré-condições, os módulos e os critérios para que o executor não precise redescobrir os defeitos.
