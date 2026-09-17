# Fase 5B — Análise Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reunir análise do acervo com indicadores corretos, filtros consistentes e termo completo explicitamente identificado.

**Architecture:** Estender a consulta agregada de `db.recorte`, introduzir `valor_status` e manter os geradores de gráficos existentes. Renomear rotas de apresentação para Análise com aliases autorizados. Preparar indicadores no Python e renderizar com os cards DSGov existentes.

**Tech Stack:** Flask, SQLite, Jinja, openpyxl, ECharts e pytest.

---

**Depende de:** 5A concluída.
**Spec:** `../specs/2026-09-17-fase5-experiencia-design.md`, seção 3.

| Arquivo | Mudança |
|---|---|
| `db.py` | filtros de valor, agregação e nome da aba exportada |
| `painel.py` | URLs novas, descrição e indicadores com links seguros |
| `app.py` | rotas novas, redirects e contexto da Análise |
| `templates/recorte.html` → `templates/analise.html` | título, seis indicadores, estado de valor e termo completo |
| `templates/index.html` | corrigir aviso de NULL e link; simplificação vem em 5C |
| `tests/test_analise.py` (novo), `tests/test_db.py`, `tests/test_painel.py`, `tests/test_app.py` | dados, filtros e fluxo |

## Tarefa 1 — separar NULL, zero e intervalos numéricos

**Modify:** `db.py:FILTROS`, `_where`, `recorte`, `painel`; `tests/test_db.py`.
**Create:** `tests/test_analise.py`.

- [ ] Escrever o teste com NULL, zero e negativo:

```python
import io
import pytest
from openpyxl import load_workbook
import db
from tests.conftest import semear

def test_zero_e_ausente_sao_conjuntos_distintos(dados):
    semear(dados)
    dados.execute('UPDATE bens SET valor_atual=NULL WHERE numero=1001')
    dados.execute('UPDATE bens SET valor_atual=0 WHERE numero=1002')
    dados.execute('UPDATE bens SET valor_atual=-10 WHERE numero=1004')
    r = db.recorte(dados,{'situacao':'ATIVO'})
    assert (r['quantidade'],r['valor_nao_informado'],r['valor_zero']) == (3,1,1)
    n = db.recorte(dados,{'situacao':'ATIVO','valor_status':'nao_informado'})
    z = db.recorte(dados,{'situacao':'ATIVO','valor_status':'zero'})
    assert [b['numero'] for b in n['bens']] == [1001]
    assert [b['numero'] for b in z['bens']] == [1002]
    assert db.recorte(dados,{'valor_status':'nao_informado','valor_de':'0'})['quantidade']==0
    assert {b['numero'] for b in db.recorte(dados,{'valor_ate':'0'})['bens']} == {1002,1004}
    with pytest.raises(db.ErroDeNegocio):
        db.recorte(dados,{'valor_status':'outro'})
```

- [ ] Rodar `.venv/bin/python -m pytest tests/test_analise.py -q`; esperado: falha nas novas chaves.
- [ ] Acrescentar `valor_status` ao final de `FILTROS`. Em `_where`, inserir antes dos limites de valor e trocar `_V` por `b.valor_atual` somente nas comparações de intervalo:

```python
status_valor = f.get('valor_status')
if status_valor == 'nao_informado':
    cl.append('b.valor_atual IS NULL')
elif status_valor == 'zero':
    cl.append('b.valor_atual = 0')
elif status_valor:
    raise ErroDeNegocio('Situação do valor inválida.')
if f.get('valor_de'):
    cl.append('b.valor_atual >= ?'); p.append(float(f['valor_de']))
if f.get('valor_ate'):
    cl.append('b.valor_atual <= ?'); p.append(float(f['valor_ate']))
```

Não alterar a definição das faixas dos gráficos nesta fase. A divisão NULL/zero é
uma dimensão explícita adicional; as faixas históricas continuam conforme a base.

- [ ] Substituir a consulta de totais de `recorte` e devolver as novas chaves:

```python
def recorte(conn, f, limite=1000):
    where, p = _where(f)
    sql = f'SELECT b.*, l.ccustos AS ccustos, a.nome AS pessoa {_DE} WHERE {where} ORDER BY b.numero'
    bens = _todos(conn,sql+(f' LIMIT {limite+1}' if limite else ''),*p)
    totais = dict(conn.execute(f'''SELECT
      count(*) AS quantidade,
      coalesce(sum(b.valor_atual),0) AS valor_total,
      coalesce(sum(CASE WHEN {_IMOVEIS_SQL} THEN 1 ELSE 0 END),0) AS imoveis,
      coalesce(sum(CASE WHEN {_IMOVEIS_SQL} THEN b.valor_atual ELSE 0 END),0) AS valor_imoveis,
      coalesce(sum(CASE WHEN l.ccustos IS NULL AND a.nome IS NULL THEN 1 ELSE 0 END),0) AS sem_centro,
      coalesce(sum(CASE WHEN b.valor_atual IS NULL THEN 1 ELSE 0 END),0) AS valor_nao_informado,
      coalesce(sum(CASE WHEN b.valor_atual = 0 THEN 1 ELSE 0 END),0) AS valor_zero
      {_DE} WHERE {where}''',p).fetchone())
    return dict(totais,bens=bens[:limite] if limite else bens,
                truncado=bool(limite) and len(bens)>limite,dimensoes=dimensoes(conn,f))
```

- [ ] Em `db.painel`, retirar `sem_valor` do dicionário inicial e acrescentar as duas contagens antes de calcular `a_emitir_centros`:

```python
d.update({
    'valor_nao_informado': um(f'SELECT count(*) FROM bens b WHERE {ativo} AND b.valor_atual IS NULL'),
    'valor_zero': um(f'SELECT count(*) FROM bens b WHERE {ativo} AND b.valor_atual = 0'),
})
```

- [ ] Atualizar `templates/index.html`: `p.sem_valor` → `p.valor_nao_informado`, rótulo “bem(ns) ativo(s) com valor não informado”, link `url_recorte(f, valor_status='nao_informado')`. Adicionar informação neutra de zero:

```jinja
{% if p.valor_zero %}<p class="text-gray-70">{{ p.valor_zero }} bem(ns) ativo(s) com valor zero.
  <a href="{{ url_recorte(f, valor_status='zero') }}">Ver bens com valor zero</a>.</p>{% endif %}
```

- [ ] Atualizar `test_painel_cards` para as duas chaves novas. Ainda manter `dimensoes` até a simplificação de 5C.
- [ ] Acrescentar teste com recorte vazio: todas as seis contagens/somas são 0; teste com imóvel e pessoa atribuída para comprovar que “sem centro nem pessoa” exige ambas as ausências.
- [ ] Rodar `.venv/bin/python -m pytest tests/test_analise.py tests/test_db.py -q`; esperado: passa.
- [ ] Commit `feat: distinguish missing and zero asset values` com os arquivos desta tarefa.

## Tarefa 2 — rotas Análise e compatibilidade das URLs

**Modify:** `app.py:recorte/recorte_xlsx`, `painel.py:url_recorte*`, `db.py:exportar_recorte`, testes de rotas.
**Rename:** `templates/recorte.html` para `templates/analise.html`.

- [ ] Escrever teste do alias e da query preservada:

```python
@pytest.mark.parametrize('sufixo',['','/xlsx'])
def test_alias_preserva_query(cliente,sufixo):
    query='situacao=&ccusto=GEX%2BLIC&ccusto=CCI&valor_status=zero'
    r=cliente.get('/recorte'+sufixo+'?'+query)
    assert r.status_code==301
    assert r.headers['Location']=='/analise'+sufixo+'?'+query
```

- [ ] Rodar esse teste; esperado: falha (rota antiga ainda renderiza).
- [ ] Renomear funções e decorators para `analise`/`analise_xlsx`. Renderizar `analise.html`, trilha Análise e arquivo `analise.xlsx`. Acrescentar aliases:

```python
def _alias_analise(endpoint):
    destino = url_for(endpoint)
    if request.query_string:
        destino += '?' + request.query_string.decode('latin-1')
    return redirect(destino,code=301)

@app.route('/recorte')
def recorte():
    return _alias_analise('analise')

@app.route('/recorte/xlsx')
def recorte_xlsx():
    return _alias_analise('analise_xlsx')
```

As quatro permissões já foram declaradas em 5A. Não redirecionar por login para
uma rota antiga sem aplicar a mesma matriz do destino.

- [ ] Helpers mantêm filtros e `situacao=`:

```python
def url_recorte(f, **extra):
    args = {k:v for k,v in {**f,**extra}.items() if v}
    args.setdefault('situacao',f.get('situacao',''))
    return url_for('analise',**args)

def url_recorte_xlsx(f):
    args = {k:v for k,v in f.items() if v}
    args.setdefault('situacao',f.get('situacao',''))
    return url_for('analise_xlsx',**args)
```

- [ ] Trocar `url_for('recorte')` para `url_for('analise')` em templates e no item do menu atual; trocar título/trilha “Recorte” por “Análise”. Não renomear `db.recorte`, `db.exportar_recorte`, `_filtros_recorte` ou helpers internos. `ws.title` em `exportar_recorte` passa a `analise`.
- [ ] Testes normais que verificam conteúdo usam `/analise`; somente testes do alias usam `/recorte`. Atualizar URLs esperadas em `tests/test_app.py` e `tests/test_painel.py`, e download esperado para `analise.xlsx`.
- [ ] Rodar `.venv/bin/python -m pytest tests/test_analise.py tests/test_app.py tests/test_painel.py tests/test_permissoes.py -q`; esperado: passa.
- [ ] Commit `feat: expose analysis routes with legacy redirects` com os arquivos renomeados e alterados.

## Tarefa 3 — indicadores e refinamentos sem trocar filtros

**Modify:** `painel.py`, `app.py:analise`, `templates/analise.html`; `tests/test_painel.py`.

- [ ] Testar que filtros incompatíveis não são sobrescritos:

```python
def test_indicadores_nao_trocam_filtro(dados):
    from app import app
    import painel
    r=dict(quantidade=1,valor_total=10,imoveis=0,valor_imoveis=0,
           sem_centro=0,valor_nao_informado=0,valor_zero=0)
    with app.test_request_context():
        cards=painel.indicadores_analise(r,{'ccusto':'CCI','classificacao':'MÓVEIS','valor_de':'1'})
    assert all(c['url'] is None for c in cards)
```

- [ ] Rodar o teste; esperado: função ausente.
- [ ] Criar a preparação de indicadores:

```python
def indicadores_analise(r, f):
    def refinar(contagem, **novos):
        if not contagem or any(f.get(k) and f[k] != v for k,v in novos.items()):
            return None
        if all(f.get(k)==v for k,v in novos.items()):
            return None
        return url_recorte(f,**novos)
    return [
      dict(rotulo='Bens no recorte',valor=r['quantidade'],detalhe=None,url=None),
      dict(rotulo='Valor atual',valor=moeda(r['valor_total']),detalhe=None,url=None),
      dict(rotulo='Imóveis',valor=r['imoveis'],detalhe=moeda(r['valor_imoveis']),
           url=refinar(r['imoveis'],classificacao='imoveis')),
      dict(rotulo='Sem centro nem pessoa',valor=r['sem_centro'],detalhe=None,
           url=refinar(r['sem_centro'],ccusto='-',pessoa='-')),
      dict(rotulo='Valor não informado',valor=r['valor_nao_informado'],detalhe=None,
           url=refinar(r['valor_nao_informado'],valor_status='nao_informado')),
      dict(rotulo='Valor zero',valor=r['valor_zero'],detalhe=None,
           url=refinar(r['valor_zero'],valor_status='zero')),
    ]
```

O teste de contagem positiva evita link contraditório com intervalo numérico. O
teste de filtros conflitantes evita mudar centro, pessoa, classe ou valor_status.

- [ ] Passar `indicadores=painel.indicadores_analise(r,f)` na renderização. Acrescentar ao formulário, usando a macro existente:

```jinja
<div class="col-md-3 mb-3">{{ select('valor_status', 'Situação do valor',
  [('', 'Todos'), ('nao_informado', 'Valor não informado'), ('zero', 'Valor zero')],
  selecionado=f.get('valor_status', ''), obrigatorio=False) }}</div>
```

- [ ] Substituir os dois contadores antigos por seis cards:

```jinja
<div class="row mb-2">
{% for k in indicadores %}
  <div class="col-sm-6 col-lg-4 mb-3">
    {% if k.url %}<a class="br-card h-100 dsgov-kpi" href="{{ k.url }}">
    {% else %}<div class="br-card h-100 dsgov-kpi">{% endif %}
      <div class="card-content"><div class="valor">{{ k.valor }}</div>
        {% if k.detalhe %}<div class="detalhe">{{ k.detalhe }}</div>{% endif %}
        <div class="text-down-01 text-gray-70">{{ k.rotulo }}</div>
      </div>
    {% if k.url %}</a>{% else %}</div>{% endif %}
  </div>
{% endfor %}
</div>
```

- [ ] Na célula de valor da tabela, usar `{{ 'Não informado' if b.valor_atual is none else moeda(b.valor_atual) }}`. Não usar `or 0`.
- [ ] Em `painel.descrever`, antes do retorno, acrescentar:

```python
if f.get('valor_status'):
    partes.append({'nao_informado':'valor não informado','zero':'valor zero'}[f['valor_status']])
```

A entrada já foi validada por `_where`; teste direto da função usa valores válidos.

- [ ] Testar o clique completo: extrair a URL de um card positivo, converter sua query com `urllib.parse.parse_qs`, chamar `db.recorte` e comparar quantidade ao card. Incluir `situacao=`, centro, pessoa, imóvel, intervalo, NULL, zero e filtros já fixados.
- [ ] Rodar `.venv/bin/python -m pytest tests/test_analise.py tests/test_painel.py tests/test_db.py -q`; esperado: passa.
- [ ] Commit `feat: add exact analysis indicators and drill-down`.

## Tarefa 4 — termo completo e exportação integral

**Modify:** `templates/analise.html`, `tests/test_analise.py`.

- [ ] Escrever teste da comunicação, sem emissão na visita:

```python
def test_termo_completo_nao_e_limitado_pelo_filtro(cliente,dados):
    antes=dados.execute('SELECT count(*) FROM termos_emitidos').fetchone()[0]
    html=cliente.get('/analise?ccusto=CCI&classificacao=MÓVEIS').get_data(as_text=True)
    assert 'Abrir termo completo' in html
    assert 'Os filtros desta análise não limitam o termo' in html
    assert '/termo/ccusto/CCI' in html
    assert dados.execute('SELECT count(*) FROM termos_emitidos').fetchone()[0]==antes
```

- [ ] Rodar; esperado: falha pelo rótulo antigo.
- [ ] Substituir o bloco de ação de termo:

```jinja
{% if termo_de and pode('termo') %}
<a class="br-button primary" href="{{ url_for('termo', tipo=termo_de[0], chave=termo_de[1]) }}">
  <i class="fas fa-file-alt mr-1" aria-hidden="true"></i>Abrir termo completo
</a>
<span class="text-down-01">{{ termo_de[1] }}</span> {{ tag_termo(termo_de[2]) }}
{% endif %}
```

Abaixo da frase descritiva, sob a mesma condição:

```jinja
<p class="text-down-01 text-gray-70">Os filtros desta análise não limitam o termo.</p>
```

- [ ] Acrescentar teste de quantidade além do limite e valores exportados:

```python
def test_exportacao_integral_e_valores_preservados(dados):
    semear(dados)
    dados.executemany('''INSERT INTO bens VALUES (?,'ATIVO','BEM','','MÓVEIS',
      '99 - SEM MAPA','01/01/2020',1,?)''',
      [(n,None if n%2 else 0) for n in range(2000,3105)])
    f={'situacao':'ATIVO'}
    r=db.recorte(dados,f)
    assert r['truncado'] and len(r['bens'])==1000 and r['quantidade']==1108
    arq=io.BytesIO(); db.exportar_recorte(dados,f,arq)
    ws=load_workbook(io.BytesIO(arq.getvalue())).active
    assert ws.title=='analise' and ws.max_row-1==1108
    valores={row[0]:row[-1] for row in ws.iter_rows(min_row=2,values_only=True)}
    assert valores[2000]==0 and valores[2001] is None
```

- [ ] Testar Consulta abrindo a Análise e o termo, sem botões de copiar/baixar; Inventário sozinho recebe 403 no novo endpoint, Excel e aliases. Não alterar regras de geração dos termos.
- [ ] Rodar `.venv/bin/python -m pytest tests/test_analise.py tests/test_app.py tests/test_permissoes.py -q`; esperado: passa.
- [ ] Atualizar referências visíveis no README para Análise, arquivo exportado e distinção de valores. Manter referência ao nome antigo apenas na explicação da migração/FAQ.
- [ ] Conferir `git diff --check` e commit `feat: clarify full terms and verify complete analysis exports`.

**Saída da etapa:** Análise funcional no layout existente. Prosseguir para 5C para
retirar gráficos do Início e entregar a árvore e o guia de uso.
