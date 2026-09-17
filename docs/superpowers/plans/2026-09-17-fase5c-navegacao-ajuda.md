# Fase 5C — navegação e Ajuda Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Organizar o menu e o Início e oferecer Ajuda correspondente às funções efetivamente concedidas.

**Architecture:** `menu.py` prepara árvore, seleção e seções de Ajuda; recebe eventos já autorizados e não consulta banco. Templates usam as mesmas permissões das rotas. Um script pequeno sincroniza o estado inicial do menu depois do core DSGov, sem substituir seus controles.

**Tech Stack:** Flask, Jinja, DSGov local, JavaScript nativo e pytest.

---

**Depende de:** 5A e 5B.
**Spec:** `../specs/2026-09-17-fase5-experiencia-design.md`, seções 1, 2 e 4.
**Referências locais:** skill DSGov `references/componentes/menu.md`, `card.md`,
`checkbox.md`, `message.md` e `references/acessibilidade.md`. O app continua Flask;
não executar scaffold Django nem sobrescrever CSS/vendor.

| Arquivo | Responsabilidade |
|---|---|
| `menu.py` (novo) | itens, grupo atual, destinos de ajuda e seções permitidas |
| `static/js/menu-estado.js` (novo) | restaurar seleção/ARIA após inicialização do core |
| `templates/ajuda.html` (novo) | guia por capacidades |
| `app.py` | contexto de menu/ajuda, rota de ajuda, Início enxuto |
| `templates/base.html`, `_macros.html` e templates interativos | árvore e botão contextual |
| `templates/index.html`, `db.py:painel` | remoção de trabalho desnecessário no Início |
| `tests/test_menu.py`, `tests/test_ajuda.py` (novos), testes existentes | seleção, acesso, âncoras e regressões |

## Tarefa 1 — árvore e seleção por endpoint e argumentos

**Create:** `menu.py`, `tests/test_menu.py`.

- [ ] Escrever teste com o endpoint compartilhado de Cadastros:

```python
from app import app
import menu

def test_somente_cadastro_correto_fica_ativo(dados):
    with app.test_request_context():
        arvore=menu.montar({'admin'},None,'cadastros',{'aba':'pessoas'},True)
    grupo=next(i for i in arvore if i['id']=='cadastros')
    assert grupo['aberto']
    assert [f['rotulo'] for f in grupo['filhos'] if f['ativo']]==['Pessoas']

def test_inventariante_nao_tem_painel_nem_acervo(dados):
    with app.test_request_context():
        arvore=menu.montar({'inventariante'},{'id':7,'nome':'Meu evento'},
                          'inventario.sala_tela',{'id':7,'localizacao':'Sala'},True)
    assert [i['rotulo'] for i in arvore]==['Inventário','Ajuda']
    assert [f['rotulo'] for f in arvore[0]['filhos']]==['Eventos','Meu evento']
    assert arvore[0]['aberto']
```

- [ ] Rodar `.venv/bin/python -m pytest tests/test_menu.py -q`; esperado: módulo ausente.
- [ ] Implementar árvore, mapa de detalhes e seleção:

```python
from flask import url_for
from permissoes import permitido

CADASTROS = [('Centros de custo','responsaveis'),('Localizações','localizacoes'),
             ('Pessoas','pessoas'),('Processos SEI','processos')]
TERMOS = [('Termo por centro de custo','centro_custos'),('Termo individual','termos_individuais'),
          ('Termo de devolução','termo_devolucao'),('Termos emitidos','termos_emitidos_tela')]
CADASTRO_ENDPOINTS = {
  'responsaveis_editar':'responsaveis','responsaveis_incluir':'responsaveis','responsaveis_excluir':'responsaveis',
  'pessoas_editar':'pessoas','pessoas_incluir':'pessoas','pessoas_excluir':'pessoas',
  'pessoas_atribuir':'pessoas','pessoas_desatribuir':'pessoas',
  'localizacoes_alterar':'localizacoes','localizacoes_incluir':'localizacoes',
  'localizacoes_mover':'localizacoes','localizacoes_excluir':'localizacoes',
  'processos_incluir':'processos','processos_vigente':'processos','processos_encerrar':'processos','processos_excluir':'processos',
}

def destino_atual(endpoint, args):
    args=dict(args or {})
    if endpoint in CADASTRO_ENDPOINTS:
        return 'cadastros',{'aba':CADASTRO_ENDPOINTS[endpoint]}
    if endpoint=='cadastro_novo':
        return 'cadastros',{'aba':args.get('aba')}
    if endpoint=='termo':
        return {'ccusto':'centro_custos','individual':'termos_individuais','devolucao':'termo_devolucao'}.get(args.get('tipo'),'centro_custos'),{}
    if endpoint in {'termo_emitido_tela','termo_emitido_documento','termo_emitido_email'}:
        return 'termos_emitidos_tela',{}
    if endpoint in {'gerar','gerar_individual'}:
        return ('centro_custos' if endpoint=='gerar' else 'termos_individuais'),{}
    if endpoint in {'usuarios.novo','usuarios.incluir','usuarios.editar','usuarios.nova_senha'}:
        return 'usuarios.lista',{}
    if endpoint in {'inventario.sala_tela','inventario.comissao','inventario.excluir'}:
        return 'inventario.evento_tela',{'id':args.get('id')}
    if endpoint=='importacao_tela':
        return 'upload',{}
    return endpoint,args

def montar(funcoes, evento_aberto, endpoint_atual, argumentos=None, login_ativo=True):
    ep, args=destino_atual(endpoint_atual,argumentos)
    def item(rotulo,icone,endpoint,kw=None):
        kw=kw or {}
        return dict(rotulo=rotulo,icone=icone,endpoint=endpoint,url=url_for(endpoint,**kw),
                    ativo=ep==endpoint and all(args.get(k)==v for k,v in kw.items()),filhos=[],aberto=False)
    def grupo(id,rotulo,icone,filhos,aberto):
        filhos=[f for f in filhos if permitido(funcoes,f['endpoint'])]
        return dict(id=id,rotulo=rotulo,icone=icone,filhos=filhos,aberto=aberto,ativo=False,url=None)
    itens=[]
    if permitido(funcoes,'analise'):
        itens.append(item('Início','fa-home','home'))
    g=grupo('termos','Termos de Responsabilidade','fa-file-signature',
            [item(r,'',e) for r,e in TERMOS],ep in {e for _,e in TERMOS})
    if g['filhos']: itens.append(g)
    if permitido(funcoes,'analise'):
        itens.append(item('Análise','fa-chart-bar','analise'))
    filhos=[item('Eventos','','inventario.eventos_tela')]
    if evento_aberto:
        for r,e in [(evento_aberto['nome'],'inventario.evento_tela'),('Painel','inventario.painel_tela'),('Relatório','inventario.relatorio_tela')]:
            filhos.append(item(r,'',e,{'id':evento_aberto['id']}))
    g=grupo('inventario','Inventário','fa-clipboard-check',filhos,(endpoint_atual or '').startswith('inventario.'))
    if g['filhos']: itens.append(g)
    g=grupo('cadastros','Cadastros','fa-address-book',
            [item(r,'','cadastros',{'aba':aba}) for r,aba in CADASTROS],ep=='cadastros')
    if g['filhos']: itens.append(g)
    for r,i,e in [('Textos','fa-pen-nib','textos_tela'),('Atualizar base','fa-upload','upload'),
                  ('Usuários','fa-users','usuarios.lista'),('Ajuda','fa-question-circle','ajuda')]:
        if permitido(funcoes,e) and (e!='usuarios.lista' or login_ativo):
            itens.append(item(r,i,e))
    return itens
```

- [ ] Testar relatório de evento fechado com ID diferente do aberto: grupo Inventário aberto, nenhum filho do evento aberto marcado como atual. Testar modo local sem Usuários e união de funções sem itens duplicados.
- [ ] Integrar a rota real de Ajuda da tarefa 2 antes da verificação final desta tarefa: a árvore usa `url_for('ajuda')`. Não criar rota fictícia nos testes. O teste inicialmente vermelho passa após as duas tarefas serem integradas.
- [ ] Commit com a tarefa 2, mantendo a árvore ligada somente quando seus destinos existem.

## Tarefa 2 — seções de Ajuda, rota e título contextual

**Modify:** `menu.py`, `app.py`, `templates/_macros.html`.
**Create:** `templates/ajuda.html`, `tests/test_ajuda.py`.

- [ ] Definir capacidades de cada seção e resolver ajuda da tela:

```python
SECOES = [
 ('inicio','Início','analise','GET'),('pesquisa','Pesquisa','pesquisa','GET'),
 ('termos','Termos de Responsabilidade','termo','GET'),('analise','Análise','analise','GET'),
 ('inventario','Conferência do inventário','inventario.ler','POST'),
 ('consulta-inventarios','Consulta de inventários','inventario.relatorio_tela','GET'),
 ('cadastros','Cadastros','cadastros','GET'),('textos','Textos','textos_tela','GET'),
 ('atualizar-base','Atualizar base','upload','GET'),('usuarios','Usuários','usuarios.lista','GET'),
 ('conta','Sua conta','usuarios.senha','GET'),('perguntas','Perguntas frequentes','ajuda','GET'),
]
AJUDA = {
 'home':'inicio','pesquisa':'pesquisa','bem':'pesquisa','analise':'analise',
 'centro_custos':'termos','termos_individuais':'termos','termo_devolucao':'termos','termos_emitidos_tela':'termos',
 'cadastros':'cadastros','textos_tela':'textos','textos_salvar':'textos','upload':'atualizar-base',
 'usuarios.lista':'usuarios','usuarios.senha':'conta',
 'inventario.painel_tela':'consulta-inventarios','inventario.relatorio_tela':'consulta-inventarios',
}

def secoes_ajuda(funcoes, login_ativo):
    return [dict(id=id,titulo=titulo) for id,titulo,ep,m in SECOES
            if permitido(funcoes,ep,m) and (id not in {'conta','usuarios'} or login_ativo)]

def ancora_ajuda(funcoes, endpoint, argumentos, login_ativo):
    if endpoint in {None,'ajuda','usuarios.login'}:
        return None
    ep,_=destino_atual(endpoint,argumentos)
    if (ep or '').startswith('inventario.') and ep not in AJUDA:
        ancora='inventario' if permitido(funcoes,'inventario.ler','POST') else 'consulta-inventarios'
    else:
        ancora=AJUDA.get(ep)
    return ancora if ancora in {s['id'] for s in secoes_ajuda(funcoes,login_ativo)} else None
```

- [ ] Adicionar a rota e, no ramo autenticado do contexto, as duas variáveis. No ramo sem usuário devolver lista vazia e âncora `None`:

Adicionar `import menu` aos imports de `app.py` antes de usar o módulo no contexto.

```python
@app.route('/ajuda')
def ajuda():
    return render_template('ajuda.html',trilha=[('Ajuda',None)])
```

```python
contexto['SECOES_AJUDA']=menu.secoes_ajuda(usuario['funcoes'],config.exigir_login())
contexto['AJUDA_ANCORA']=menu.ancora_ajuda(usuario['funcoes'],request.endpoint,
                                       request.view_args,config.exigir_login())
```
- [ ] Criar a macro contextual, sem contexto implícito:

```jinja
{% macro ajuda_titulo(ancora) %}
{% if ancora %}<a class="br-button circle small ml-2" href="{{ url_for('ajuda') }}#{{ ancora }}" aria-label="Ajuda desta tela">
  <i class="fas fa-question-circle" aria-hidden="true"></i>
</a>{% endif %}
{% endmacro %}
```

- [ ] Criar o guia abaixo. Cada `<section>` é gerada apenas quando seu ID pertence a `SECOES_AJUDA`; ações internas possuem suas próprias condições:

```jinja
{% extends 'base.html' %}
{% block titulo %}Ajuda{% endblock %}
{% block conteudo %}
<h1>Ajuda</h1>
<p>Consulte as orientações para as funções atribuídas ao seu usuário.</p>
<nav class="br-card mb-4" aria-label="Sumário da ajuda"><div class="card-content br-list">
{% for s in SECOES_AJUDA %}<a class="br-item" href="#{{ s.id }}">{{ s.titulo }}</a>{% endfor %}
</div></nav>
{% for s in SECOES_AJUDA %}
<section class="br-card mb-4" id="{{ s.id }}" aria-labelledby="titulo-{{ s.id }}">
<div class="card-header"><h2 class="text-up-02" id="titulo-{{ s.id }}">{{ s.titulo }}</h2></div>
<div class="card-content">
{% if s.id == 'inicio' %}
  <p>O Início reúne atalhos e indicadores das áreas que você pode acessar. Clique nos indicadores que têm link para abrir a lista correspondente.</p>
  <p>Valor não informado significa que o cadastro não traz esse dado. Valor zero é um valor cadastrado, igual a R$ 0,00. Os dois casos têm filtros separados na Análise.</p>
{% elif s.id == 'pesquisa' %}
  <p>Use a lupa do cabeçalho em qualquer tela em que ela estiver disponível. Digite o número de patrimônio para abrir a ficha ou palavras para procurar bens, pessoas e centros de custo.</p>
  <p>Com várias palavras, todas precisam aparecer no resultado. Use <code>GEX*</code> para começar com GEX e <code>*ITEC</code> para terminar com ITEC. A ficha mostra dados do bem e histórico disponível para suas funções.</p>
{% elif s.id == 'termos' %}
  <p>O termo por centro de custo reúne os bens ativos sob a guarda do setor, excluindo os atribuídos individualmente. O termo individual reúne os bens da pessoa. O termo de devolução registra os bens que retornam ao patrimônio.</p>
  <p>Abra as telas de termos para consultar os documentos e o histórico. Um termo pode estar vigente, desatualizado ou ainda não ter sido emitido.</p>
  {% if pode('termo_docx') %}
  <p>Para emitir, é necessário um processo SEI vigente do tipo correspondente. Confira o documento e use <strong>Copiar para o SEI</strong> ou <strong>Baixar .docx</strong>. As duas ações registram a emissão; abrir a prévia não registra.</p>
  <p>Em Termos emitidos, registre o documento SEI, o bloco de assinatura e o envio do pedido de assinatura. O botão de e-mail abre seu programa de e-mail com o texto pronto. A emissão anterior permanece no histórico.</p>
  {% endif %}
{% elif s.id == 'analise' %}
  <p>Filtre por situação, centro, pessoa, localização, classificação, idade, data de entrada, faixa de valor e situação do valor. Os indicadores representam o conjunto filtrado inteiro.</p>
  <p>Os gráficos e indicadores com link refinam a seleção. Valor não informado e valor zero são filtros diferentes. Combinar valor não informado com um intervalo numérico não retorna bens.</p>
  <p><strong>Exportar .xlsx</strong> baixa todos os bens filtrados, mesmo quando a tela limita a lista a 1.000. Célula vazia de valor indica ausência de informação; zero permanece um número.</p>
  <p><strong>Abrir termo completo</strong> leva ao documento integral do centro ou da pessoa. Os filtros desta análise não limitam o termo.</p>
{% elif s.id == 'inventario' %}
  <p>A função Inventário permite conferir bens nos eventos em que você integra a comissão. Sem evento atribuído, solicite ao administrador sua inclusão. Evento encerrado aceita somente consulta.</p>
  <ol class="br-list">
    <li class="br-item">Abra o evento e escolha a sala que está conferindo.</li>
    <li class="br-item">Leia a plaqueta pelo leitor de código de barras, pela câmera ou digitando o número.</li>
    <li class="br-item">Confira conservação, quem usa o bem e observações; acrescente as fotos necessárias.</li>
    <li class="br-item">Acompanhe os bens pendentes e as divergências da sala.</li>
  </ol>
  <p>Localizado foi lido na sala cadastrada; divergente foi lido em outra sala; pendente ainda não foi lido. A conferência não muda o cadastro de localização do sistema de patrimônio. Bens baixados continuam baixados.</p>
  <p>Se a plaqueta não existir no cadastro, registre uma sobra com descrição, observação e foto quando o recurso estiver disponível. Para plaquetas ilegíveis, use a marcação em lote após conferir os bens.</p>
  <p>Desmarcar remove a leitura e as fotos associadas. Se leu na sala errada, leia novamente na sala correta; vale a leitura mais recente.</p>
  {% if pode('inventario.abrir','POST') %}
  <p>O administrador abre o evento, define a comissão e o escopo das salas. Só pode haver um evento aberto. Encerrar congela os registros; excluir pede confirmação pelo nome e não pode ser desfeito.</p>
  {% endif %}
{% elif s.id == 'consulta-inventarios' %}
  <p>A função Consulta de inventários permite acompanhar todos os eventos, anteriores e futuros, pelo Painel, Relatório e exportação. Ela não permite conferir bens por si só.</p>
  <p>O Painel mostra o andamento e as distribuições do evento. No Relatório, filtre sala, situação, integrante, conservação, fotos e texto e use a ordenação das colunas.</p>
  <p>Exporte a planilha com ou sem fotos. A opção com fotos usa recursos do Excel do Microsoft 365. Eventos encerrados preservam os registros congelados, mesmo depois de atualizar a base.</p>
{% elif s.id == 'cadastros' %}
  <p>Centros de custo mantêm responsáveis e contatos. Localizações associam salas aos centros. Pessoas recebem bens de uso individual. Processos SEI definem o processo vigente de cada tipo de termo.</p>
  <p>Use busca, filtros e paginação para encontrar registros. Revise origem e destino ao mover localizações ou atribuir bens. Bem atribuído a uma pessoa deixa de entrar no termo do setor.</p>
  {% if pode('responsaveis_excluir','POST') %}<p>Exclusões pedem confirmação. Centro com bens ativos e processo com termos registrados não podem ser excluídos.</p>{% endif %}
  {% if pode('importar_cadastros','POST') %}<p>A planilha de cadastros substitui as tabelas inteiras. As abas de inventário devem vir todas juntas ou ser omitidas. Importar nomes não concede acesso a usuários; revise a comissão do evento aberto.</p>{% endif %}
{% elif s.id == 'textos' %}
  <p>Edite os dizeres dos termos e o modelo de e-mail. Preserve marcadores como <code>{nome}</code> e <code>{ccustos}</code>, preenchidos pelo sistema. Restaurar padrão recupera o texto original.</p>
{% elif s.id == 'atualizar-base' %}
  <p>Envie o arquivo .xlsx exportado do sistema de patrimônio. A atualização substitui a base de bens e mantém os cadastros de responsáveis, pessoas e atribuições, conforme as validações apresentadas.</p>
  <p>O histórico registra entradas, saídas e mudanças dos bens. Exportar bens gera uma planilha no formato de importação para cópia de segurança.</p>
{% elif s.id == 'usuarios' %}
  <p>Crie usuários e selecione somente as funções necessárias. Administrador gerencia o sistema; Operador emite termos, mantém cadastros e textos e atualiza a base; Consulta lê acervo e termos; Inventário faz conferência na própria comissão; Consulta de inventários acompanha painéis, relatórios e exportações de todos os eventos.</p>
  <p>As funções selecionadas se somam. Criar um usuário com Inventário não o inclui automaticamente em uma comissão. Login é único; o e-mail é opcional, único e também pode ser usado para entrar.</p>
  <p>Usuários são inativados, não excluídos. Nova senha gera uma senha temporária mostrada uma única vez, com troca obrigatória no acesso seguinte. Cinco tentativas erradas bloqueiam o login por 15 minutos.</p>
{% elif s.id == 'conta' %}
  <p>Seu nome e suas funções aparecem no cabeçalho. Para trocar a senha, informe a atual e a nova, com pelo menos oito caracteres. Sair encerra a sessão. A sessão expira após 12 horas sem uso.</p>
{% elif s.id == 'perguntas' %}
  {% if pode('termo') %}<h3 class="text-up-01">Por que o termo está desatualizado?</h3><p>Entraram ou saíram bens desde a última emissão. A equipe com função Operador pode conferir e emitir outro; o anterior permanece no histórico.</p>{% endif %}
  {% if pode('cadastros') %}<h3 class="text-up-01">Um bem de uso pessoal aparece no termo do setor.</h3><p>Atribua o bem à pessoa em Cadastros → Pessoas para que passe ao termo individual.</p>{% endif %}
  {% if pode('inventario.ler','POST') %}<h3 class="text-up-01">A foto não foi enviada.</h3><p>Confira a conexão e tente novamente. A foto é enviada no momento do registro. Se continuar falhando, avise o administrador.</p>{% endif %}
  {% if pode('analise') %}<h3 class="text-up-01">Onde está o Recorte?</h3><p>A tela passou a se chamar Análise. Ela reúne filtros, gráficos, indicadores e exportação do acervo.</p>{% endif %}
  {% if pode('inventario.relatorio_tela') %}<h3 class="text-up-01">Por que a atualização da base não mudou o evento encerrado?</h3><p>O encerramento congela os dados do evento. Essa preservação permite consultar o que foi registrado naquele inventário.</p>{% endif %}
{% endif %}
</div></section>
{% endfor %}
{% endblock %}
```

- [ ] Acrescentar parser de teste para IDs e links, sem depender de texto bruto para inferir âncoras:

```python
from html.parser import HTMLParser

class Estrutura(HTMLParser):
    def __init__(self,html):
        super().__init__(); self.ids=set(); self.links=[]; self.h1=0
        self.feed(html)
    def handle_starttag(self,tag,attrs):
        a=dict(attrs)
        if a.get('id'): self.ids.add(a['id'])
        if tag=='a' and a.get('href'): self.links.append(a['href'])
        if tag=='h1': self.h1+=1

def test_sumario_aponta_para_secoes_existentes(cliente):
    r=cliente.get('/ajuda')
    assert r.status_code==200
    estrutura=Estrutura(r.get_data(as_text=True))
    assert estrutura.h1==1
    assert all(link[1:] in estrutura.ids for link in estrutura.links if link.startswith('#'))
```

- [ ] Rodar `.venv/bin/python -m pytest tests/test_menu.py tests/test_ajuda.py -q`; esperado: passa depois de ligar o contexto.
- [ ] Commit `feat: define accessible navigation and scoped help` com arquivos explícitos.

## Tarefa 3 — renderizar árvore e corrigir estado após DSGov

**Modify:** `app.py:contexto_dsgov`, `templates/base.html`.
**Create:** `static/js/menu-estado.js`.

- [ ] No contexto, substituir a montagem inline de MENU por:

```python
visiveis=comissoes.eventos_visiveis(obter_conn(),usuario) if pode('inventario.eventos_tela') else []
aberto=next((e for e in visiveis if not e['encerrado_em']),None)
contexto['MENU']=menu.montar(usuario['funcoes'],aberto,request.endpoint,request.view_args,
                            config.exigir_login())
```

Não passar `inventario.evento_aberto` sem autorização. O menu precisa só de ID e nome,
não de resumo ou leituras. Preservar CSRF, usuário, URL inicial e auxiliares de 5A.

- [ ] Substituir somente o conteúdo de `<nav class="menu-body">` em `base.html`:

```jinja
{% for i in MENU %}
{% if i.filhos %}
<div class="menu-folder{% if i.aberto %} active{% endif %}">
  <a class="menu-item" href="javascript:void(0)" role="treeitem" aria-expanded="{{ 'true' if i.aberto else 'false' }}">
    <span class="icon"><i class="fas {{ i.icone }}" aria-hidden="true"></i></span><span class="content">{{ i.rotulo }}</span>
  </a>
  <ul role="group">{% for f in i.filhos %}<li>
    <a class="menu-item{% if f.ativo %} active{% endif %}" href="{{ f.url }}" role="treeitem"{% if f.ativo %} data-atual="true" aria-current="page"{% endif %}>
      <span class="content">{{ f.rotulo }}</span>
    </a>
  </li>{% endfor %}</ul>
</div>
{% else %}
<a class="menu-item{% if i.ativo %} active{% endif %}" href="{{ i.url }}" role="treeitem"{% if i.ativo %} data-atual="true" aria-current="page"{% endif %}>
  <span class="icon"><i class="fas {{ i.icone }}" aria-hidden="true"></i></span><span class="content">{{ i.rotulo }}</span>
</a>
{% endif %}
{% endfor %}
```

- [ ] Adicionar script após `dsgov.js` e antes do bloco de scripts da página:

```jinja
<script src="{{ url_for('static', filename='js/menu-estado.js') }}"></script>
```

Conteúdo integral de `static/js/menu-estado.js`:

```javascript
(function () {
  function sincronizar() {
    var menu = document.getElementById('main-navigation');
    if (!menu) return;
    menu.querySelectorAll('a.menu-item').forEach(function (a) {
      var atual = a.dataset.atual === 'true';
      a.classList.toggle('active', atual);
      if (atual) a.setAttribute('aria-current', 'page');
      else a.removeAttribute('aria-current');
    });
    menu.querySelectorAll('.menu-folder > a.menu-item').forEach(function (a) {
      a.setAttribute('aria-expanded', String(a.parentElement.classList.contains('active')));
    });
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', sincronizar);
  } else {
    sincronizar();
  }
})();
```

O core inicia antes; seu clique continua controlando expansão e ARIA. Esse script
só corrige o estado inicial e a marcação indevida por prefixo feita por `dsgov.js`.

- [ ] Testar HTML da árvore para grupos vazios ausentes, seleção única e links de eventos autorizados. Verificar no navegador após inicialização: expandido true, clique fecha/abre, teclado funciona, sem erro no console.
- [ ] Rodar `.venv/bin/python -m pytest tests/test_menu.py tests/test_permissoes.py -q`; esperado: passa com novas asserções de árvore.
- [ ] Commit `feat: render function-aware navigation tree`.

## Tarefa 4 — Início enxuto sem consultas de gráficos

**Modify:** `db.py:painel`, `app.py:home`, `templates/index.html`; `tests/test_db.py`, `tests/test_app.py`.

- [ ] Escrever teste que detecta trabalho desnecessário:

```python
def test_inicio_nao_calcula_graficos(cliente,monkeypatch):
    import db
    def proibido(*args,**kwargs): raise AssertionError('Início não usa dimensões')
    monkeypatch.setattr(db,'dimensoes',proibido)
    html=cliente.get('/').get_data(as_text=True)
    conteudo=html.split('id="main-content"',1)[1].split('</main>',1)[0]
    assert '<form' not in conteudo
    assert 'echarts.min.js' not in html
    assert 'Termos por centro de custo' not in conteudo
    assert 'Sobre o patrimônio' in conteudo and 'Ver guia' in conteudo
```

- [ ] Rodar; esperado: falha em `db.dimensoes`.
- [ ] Retirar a chave `dimensoes` de `db.painel` e o argumento `cards=...` de `home`. Preservar a primeira verificação de destino inicial da etapa 5A. Atualizar o teste de `db.painel` para afirmar `'dimensoes' not in p`.
- [ ] Em `index.html`, remover importações `grafico`, `tag_termo`, `cabecalho_tabela`, o formulário Pesquisar, a linha de gráficos, a tabela de termos e o bloco ECharts. Manter cabeçalho e indicadores com as permissões de 5A e a distinção de valor de 5B.
- [ ] Renderizar os seis atalhos condicionais, com destino Análise já existente:

```jinja
<div class="dsgov-atalhos mb-4" role="list" aria-label="Acesso rápido">
{% for rotulo,icone,descricao,ep,permissao,metodo in [
 ('Termo por centro de custo','fa-building','Bens sob guarda de um setor','centro_custos','centro_custos','GET'),
 ('Termo individual','fa-user-check','Bens atribuídos a uma pessoa','termos_individuais','termos_individuais','GET'),
 ('Termo de devolução','fa-box-open','Devolução de bens ao patrimônio','termo_devolucao','termo_devolucao','GET'),
 ('Termos emitidos','fa-history','Histórico e documentos SEI','termos_emitidos_tela','termos_emitidos_tela','GET'),
 ('Realizar inventário','fa-clipboard-check','Conferência de bens por sala','inventario.eventos_tela','inventario.ler','POST'),
 ('Análise','fa-chart-bar','Filtros, gráficos e exportação do acervo','analise','analise','GET')
] %}
{% if pode(permissao,metodo) %}
<a class="br-card dsgov-atalho" href="{{ url_for(ep) }}" role="listitem"><div class="card-content">
  <div class="dsgov-atalho-icone"><i class="fas {{ icone }}" aria-hidden="true"></i></div>
  <div class="text-weight-semi-bold text-up-01">{{ rotulo }}</div><div class="text-down-01 text-gray-70">{{ descricao }}</div>
</div></a>
{% endif %}{% endfor %}
</div>
```

- [ ] Acrescentar card “Sobre o patrimônio” com os três parágrafos integrais da seção 2 da spec de experiência. Estrutura:

```jinja
<section class="br-card mt-4" aria-labelledby="sobre-patrimonio">
  <div class="card-header"><h2 class="text-up-02" id="sobre-patrimonio">Sobre o patrimônio</h2></div>
  <div class="card-content">
    <p>Este sistema reúne o cadastro dos bens do Conselho Federal de Contabilidade, os responsáveis por sua guarda, os termos de responsabilidade e o inventário.</p>
    <p>O Termo de Responsabilidade registra a guarda dos bens por um centro de custo ou por uma pessoa. O Termo de Devolução registra a devolução dos bens. Os documentos são preparados aqui para uso no SEI.</p>
    <p>Use os atalhos para suas atividades e a pesquisa do cabeçalho para localizar bens, pessoas ou centros de custo. O guia explica as funções disponíveis para você.</p>
    <a class="br-button secondary" href="{{ url_for('ajuda') }}">Ver guia</a>
  </div>
</section>
```

- [ ] Confirmar admin com seis atalhos; Operador/Consulta sem Inventário não recebem atalho de conferência. Consulta sem acesso a importações não recebe esse KPI. Inventário sozinho não renderiza o Início.
- [ ] Rodar `.venv/bin/python -m pytest tests/test_app.py tests/test_db.py tests/test_menu.py tests/test_analise.py -q`; esperado: passa.
- [ ] Commit `feat: simplify home to permitted shortcuts and indicators`.

## Tarefa 5 — integrar Ajuda contextual nas telas existentes

**Modify:** templates interativos e `tests/test_ajuda.py`.

- [ ] Aplicar a edição mecânica abaixo aos títulos existentes. O script exige exatamente um título em cada template e preserva seus textos, ações e subtítulos:

```python
from pathlib import Path
import re

nomes='''index.html analise.html bem.html pesquisa.html centro_custos.html
termos_individuais.html termo_devolucao.html termos_emitidos.html termo_emitido.html
termo.html upload.html importacao.html textos.html cadastros.html
cadastros/formulario.html cadastros/atribuir.html cadastros/mover.html
cadastros/confirmar.html inventario_eventos.html inventario_evento.html
inventario_sala.html inventario_relatorio.html inventario_painel.html
inventario_comissao.html inventario_excluir.html usuarios/lista.html
usuarios/formulario.html senha.html'''.split()
for nome in nomes:
    p=Path('templates')/nome
    texto=p.read_text()
    if '{{ ajuda_titulo(AJUDA_ANCORA) }}' in texto:
        continue
    assert texto.count('</h1>')==1, nome
    texto=texto.replace('</h1>','</h1>{{ ajuda_titulo(AJUDA_ANCORA) }}',1)
    texto,n=re.subn(r'(\{% extends [^\n]+%\})',
                   r"\1\n{% from '_macros.html' import ajuda_titulo %}",texto,count=1)
    assert n==1,nome
    p.write_text(texto)
p=Path('templates/cadastros/_pessoa.html')
texto=p.read_text().replace('<h1>','<h2 class="text-up-02">').replace('</h1>','</h2>')
p.write_text(texto)
```

Não modificar login, erro ou documento incorporado. O parcial de pessoa passa a
título de seção: o título principal e a Ajuda pertencem à página de cadastros.
Revisar a posição do botão nos cabeçalhos complexos da sala e do usuário; mantê-lo
no contêiner do título, antes do subtítulo e sem mover as ações existentes.
- [ ] Criar casos de tela sem efeito colateral: Início, Análise, Pesquisa, bem, listas de termos, cadastros, Textos, Upload, Usuários, Senha e telas de evento semeado. Para cada GET permitido, extrair o link `/ajuda#...`, buscar `/ajuda` com o mesmo usuário e afirmar que a âncora existe.
- [ ] Não percorrer indiscriminadamente todos os GETs de `url_map`: download de termo registra emissão. Usar lista explícita de páginas interativas, com fixtures de evento e processo. Incluir telas filhas e POSTs que retornam formulário com erro, sem disparar emissão.
- [ ] Testar seções e ações por funções:

```python
def test_ajuda_inventariante_sem_relatorios(cliente):
    from tests.conftest import logar,SENHA_PADRAO
    cliente.post('/sair'); logar(cliente,'beltrana',SENHA_PADRAO)
    html=cliente.get('/ajuda').get_data(as_text=True)
    estrutura=Estrutura(html)
    assert {'inventario','conta','perguntas'} <= estrutura.ids
    assert not {'inicio','pesquisa','termos','analise','consulta-inventarios','usuarios'} & estrutura.ids
    assert 'Copiar para o SEI' not in html and 'Exporte a planilha' not in html
```

Consulta comum tem seção Termos sem instrução de emissão; Operador tem emissão sem
instruções de exclusão/admin; Consulta de inventários tem relatórios globais sem
seção de conferência; combinação soma seções. Desktop não mostra Conta nem Usuários.

- [ ] Rodar `.venv/bin/python -m pytest tests/test_ajuda.py tests/test_menu.py tests/test_permissoes.py -q`; esperado: passa.
- [ ] Commit `feat: connect page help to effective capabilities`.

## Tarefa 6 — integração, documentação e validação visual

**Modify:** `README.md`, specs da fase 5 (estado e evidências).

- [ ] Substituir a documentação de quatro perfis por cinco funções cumulativas, entrada por acesso, nome Análise, NULL/zero, termo completo e Ajuda. Documentar que os usuários antigos recebem apenas a função correspondente e que nomes de comissão ambíguos exigem vínculo pelo administrador.
- [ ] Rodar `.venv/bin/python -m pytest -q`; esperado: suíte completa passa. Corrigir falhas de regressão e rodar novamente somente os testes afetados antes de repetir a suíte se necessário.
- [ ] Rodar `git diff --check`; esperado: sem erro. Buscar `Recorte` em templates/README: permitido apenas na FAQ/migração, não em títulos, menu, arquivo exportado ou texto de botões. Conferir ausência de `USUARIO.perfil` e autorização por nome no caminho web.
- [ ] Iniciar o app em uma pasta temporária de dados com login habilitado e contas de teste; usar apenas uma porta local livre. Nunca usar o banco de produção para semear usuários de validação.
- [ ] No navegador, verificar os cenários:

| Cenário | Resultado obrigatório |
|---|---|
| Admin, 1280px | Nove itens/grupos, até seis atalhos, Início sem gráficos |
| Inventário sozinho | Entrada em eventos próprios; sem lupa, Análise, termos ou relatórios |
| Inventário sem comissão | Estado vazio; nenhum nome ou total de evento alheio |
| Consulta de inventários | Todos os eventos; painel/relatório/exportação; nenhuma ação de conferência |
| Combinação das duas funções de inventário | Consulta global; escrita só na comissão |
| Cadastros → Pessoas/edição | Só Pessoas ativa; Cadastros aberto |
| Relatório de evento encerrado | Inventário aberto; não marca o relatório do evento aberto como atual |
| Menu após core inicializar | ARIA coerente; clicar abre/fecha uma vez; teclado e foco funcionam |
| Celular, 390px | Sem perda de controles; menu abre/fecha; ajuda contextual alcançável |
| Análise com NULL/zero/filtro | Números, lista e Excel concordam; aviso do termo completo visível |
| Ajuda por função | Sumário sem links quebrados; instruções de ação correspondem às permissões |

- [ ] Registrar evidências de testes e navegador nas specs, com data e commit real.
- [ ] Commit `docs: document phase 5 access and user workflows` apenas dos arquivos planejados.

**Entrega:** os três planos concluídos, suíte e cenários visuais verificados.
Publicação e execução de migração no ambiente real são etapas posteriores à
implementação; esta fase pode ser inteiramente revisada localmente.
