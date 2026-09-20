"""Ajuda: seções conforme as funções do usuário, sumário coerente e âncora contextual de cada tela."""
import pathlib
import re
from html.parser import HTMLParser

from flask import g

import comissoes
import menu
import usuarios
from tests.conftest import SENHA_PADRAO, logar


class Estrutura(HTMLParser):
    """IDs, seções e links da página: confere âncoras sem depender do texto exibido."""

    def __init__(self, html):
        super().__init__()
        self.ids = set()
        self.links = []
        self.secoes = []
        self.h1 = 0
        self.feed(html)

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if a.get('id'):
            self.ids.add(a['id'])
        if tag == 'section' and a.get('id'):
            self.secoes.append(a['id'])
        if tag == 'a' and a.get('href'):
            self.links.append(a['href'])
        if tag == 'h1':
            self.h1 += 1


def ajuda(cliente):
    r = cliente.get('/ajuda')
    assert r.status_code == 200
    return Estrutura(r.get_data(as_text=True))


def test_sumario_aponta_para_secoes_existentes(cliente):
    estrutura = ajuda(cliente)
    assert estrutura.h1 == 1
    assert all(link[1:] in estrutura.ids for link in estrutura.links if link.startswith('#'))


def test_sumario_lista_exatamente_as_secoes_renderizadas(cliente):
    estrutura = ajuda(cliente)
    ancoras = [link[1:] for link in estrutura.links if link.startswith('#')]
    assert [a for a in ancoras if a in set(estrutura.secoes)] == estrutura.secoes


def test_administrador_ve_todas_as_secoes(cliente):
    assert ajuda(cliente).secoes == [id for id, *_ in menu.SECOES]
    assert "Administra".encode() in cliente.get("/ajuda").data


def test_inventariante_ve_a_conferencia_e_nao_a_consulta(cliente):
    assert logar(cliente, 'beltrana', SENHA_PADRAO).status_code == 302
    assert ajuda(cliente).secoes == ['inventario', 'conta', 'perguntas']


def test_consulta_de_inventarios_ve_a_consulta_e_nao_a_conferencia(cliente, dados):
    usuarios.criar(dados, 'ci', 'Consulta Inv', SENHA_PADRAO, ['consulta_inventarios'], trocar_senha=False)
    assert logar(cliente, 'ci', SENHA_PADRAO).status_code == 302
    assert ajuda(cliente).secoes == ['consulta-inventarios', 'conta', 'perguntas']


def test_modo_local_nao_traz_conta(cliente_local):
    secoes = ajuda(cliente_local).secoes
    assert 'conta' not in secoes
    assert secoes == [id for id, *_ in menu.SECOES if id != 'conta']


def test_ancora_da_tela_atual():
    assert menu.ancora_ajuda({'admin'}, 'bem', {'numero': '1'}, True) == 'pesquisa'
    assert menu.ancora_ajuda({'admin'}, 'responsaveis_editar', {'ccustos': 'CCI'}, True) == 'cadastros'
    assert menu.ancora_ajuda({'admin'}, 'termo', {'tipo': 'individual', 'chave': 'X'}, True) == 'termos'
    assert menu.ancora_ajuda({'admin'}, 'usuarios.senha', {}, True) == 'conta'
    assert menu.ancora_ajuda({'admin'}, 'usuarios.acessos', {}, True) == 'conta'
    assert menu.ancora_ajuda({'inventariante'}, 'inventario.sala_tela', {'id': 1, 'localizacao': 'S'}, True) == 'inventario'
    assert menu.ancora_ajuda({'consulta_inventarios'}, 'inventario.eventos_tela', {}, True) == 'consulta-inventarios'
    assert menu.ancora_ajuda({'consulta_inventarios'}, 'inventario.painel_tela', {'id': 1}, True) == 'consulta-inventarios'


def test_telas_sem_ajuda_contextual():
    """A própria Ajuda, o login, o documento embutido e o que o usuário não enxerga não ganham atalho."""
    assert menu.ancora_ajuda({'admin'}, 'ajuda', {}, True) is None
    assert menu.ancora_ajuda({'admin'}, 'usuarios.login', {}, True) is None
    assert menu.ancora_ajuda({'admin'}, None, None, True) is None
    assert menu.ancora_ajuda({'admin'}, 'termo_documento', {'tipo': 'ccusto', 'chave': 'CCI'}, True) is None
    assert menu.ancora_ajuda({'inventariante'}, 'home', {}, True) is None
    assert menu.ancora_ajuda({'admin'}, 'usuarios.senha', {}, False) is None


def test_contexto_leva_secoes_e_ancora_para_as_telas(dados):
    from app import app
    with app.test_request_context('/pesquisa'):
        g.usuario = usuarios.USUARIO_LOCAL
        contexto = {}
        app.update_template_context(contexto)
    assert contexto['AJUDA_ANCORA'] == 'pesquisa'
    assert [s['id'] for s in contexto['SECOES_AJUDA']] == [id for id, *_ in menu.SECOES if id != 'conta']


def test_contexto_sem_usuario_nao_oferece_ajuda(dados):
    from app import app
    with app.test_request_context('/login'):
        contexto = {}
        app.update_template_context(contexto)
    assert contexto['SECOES_AJUDA'] == [] and contexto['AJUDA_ANCORA'] is None


def test_dsgov_devolve_o_id_do_servidor_depois_do_brcard():
    """Guarda de regressão do navegador (invisível ao cliente de teste, que não roda JS): o BRCard do core
    reescreve o id de TODO .br-card, o que apagaria as âncoras das seções de /ajuda depois de a página
    carregar. O bridge `dsgov.js` precisa guardar o id do servidor e devolvê-lo após construir o BRCard.

    A primeira asserção lê o texto minificado do core 3.7.0. **Se ela falhar depois de atualizar o
    `core.min.js`, não remende o texto procurado**: confira antes, no core novo, se o BRCard ainda
    reescreve o id. Se não reescrever mais, o remendo do bridge virou desnecessário e este teste deve
    ser reescrito ou apagado junto com ele; se reescrever de outro jeito, ajuste o trecho procurado.
    A verificação de verdade é o roteiro de navegador da fase 5 (âncoras do sumário de /ajuda).
    """
    raiz = pathlib.Path(__file__).resolve().parent.parent
    core = (raiz / 'static/dsgov/vendor/govbr-ds/core.min.js').read_text(encoding='utf-8')
    assert 'setAttribute("id",`card${n}`)' in core, 'core atualizado: o BRCard ainda reescreve o id?'
    bloco = (raiz / 'static/dsgov/js/dsgov.js').read_text(encoding='utf-8')
    bloco = bloco.split('querySelectorAll(".br-card")')[1].split('});')[0]
    codigo = re.sub(r'/\*.*?\*/', '', bloco, flags=re.S)              # só o código: comentário não vale
    assert 'var idDoServidor = el.getAttribute("id");' in codigo
    assert codigo.index('new window.core.BRCard') < codigo.index('el.setAttribute("id", idDoServidor);')


def test_macro_de_ajuda_da_tela(dados):
    from app import app
    with app.test_request_context():
        macros = app.jinja_env.get_template('_macros.html').module
        assert 'href="/ajuda#pesquisa"' in str(macros.ajuda_titulo('pesquisa'))
        assert str(macros.ajuda_titulo(None)).strip() == ''


# ---------------------------------------------------------- botão de ajuda nas telas interativas

def _links_ajuda(html):
    """Links /ajuda#... presentes na página: o botão de ajuda contextual do título, quando existe."""
    return [l for l in Estrutura(html).links if l.startswith('/ajuda#')]


def _checar_pagina(cliente, resposta, funcoes, endpoint, args, login_ativo=True):
    """Cada página varrida tem exatamente um <h1>; o botão de ajuda (se houver) bate com
    menu.ancora_ajuda e leva a uma seção de fato renderizada em /ajuda para o mesmo usuário;
    sem âncora, a página não mostra nenhum botão."""
    assert resposta.status_code == 200, endpoint
    html = resposta.get_data(as_text=True)
    estrutura = Estrutura(html)
    assert estrutura.h1 == 1, (endpoint, estrutura.h1)
    links = [l for l in estrutura.links if l.startswith('/ajuda#')]
    esperado = menu.ancora_ajuda(funcoes, endpoint, args, login_ativo)
    if esperado is None:
        assert links == [], (endpoint, links)
        return
    assert links == [f'/ajuda#{esperado}'], (endpoint, links, esperado)
    assert esperado in ajuda(cliente).ids, (endpoint, esperado)


def test_telas_interativas_oferecem_ajuda_coerente(cliente, dados):
    """Roteiro sem efeito colateral pelas telas interativas listadas na tarefa: só GETs que não emitem
    termo e um POST que devolve formulário com erro. A contagem de termos emitidos não pode mudar."""
    antes = dados.execute("SELECT COUNT(*) FROM termos_emitidos").fetchone()[0]
    beltrana_id = usuarios.por_login(dados, 'beltrana')['id']
    eid = comissoes.abrir(dados, 'Evento', '', [beltrana_id], None)
    funcoes = {'admin'}
    telas = [
        ('/', 'home', {}),
        ('/analise', 'analise', {}),
        ('/pesquisa', 'pesquisa', {}),
        ('/bem?numero=1001', 'bem', {}),
        ('/centro-custos', 'centro_custos', {}),
        ('/termos-individuais', 'termos_individuais', {}),
        ('/termo_devolucao', 'termo_devolucao', {}),
        ('/termos-emitidos', 'termos_emitidos_tela', {}),
        ('/cadastros/responsaveis', 'cadastros', {'aba': 'responsaveis'}),
        ('/cadastros/localizacoes', 'cadastros', {'aba': 'localizacoes'}),
        ('/cadastros/pessoas', 'cadastros', {'aba': 'pessoas'}),
        ('/cadastros/pessoas?nome=ANA SILVA', 'cadastros', {'aba': 'pessoas'}),
        ('/cadastros/processos', 'cadastros', {'aba': 'processos'}),
        ('/textos', 'textos_tela', {}),
        ('/upload', 'upload', {}),
        ('/usuarios', 'usuarios.lista', {}),
        ('/usuarios/novo', 'usuarios.novo', {}),
        (f'/usuarios/{beltrana_id}/editar', 'usuarios.editar', {'id': beltrana_id}),
        ('/senha', 'usuarios.senha', {}),
        ('/meus-acessos', 'usuarios.acessos', {}),
        ('/inventario', 'inventario.eventos_tela', {}),
        (f'/inventario/{eid}', 'inventario.evento_tela', {'id': eid}),
        (f'/inventario/{eid}/sala/01 - SALA CCI', 'inventario.sala_tela', {'id': eid, 'localizacao': '01 - SALA CCI'}),
        (f'/inventario/{eid}/relatorio', 'inventario.relatorio_tela', {'id': eid}),
        (f'/inventario/{eid}/painel', 'inventario.painel_tela', {'id': eid}),
        (f'/inventario/{eid}/comissao', 'inventario.comissao', {'id': eid}),
        (f'/inventario/{eid}/excluir', 'inventario.excluir', {'id': eid}),
    ]
    for url, endpoint, args in telas:
        _checar_pagina(cliente, cliente.get(url), funcoes, endpoint, args)
    _checar_pagina(cliente, cliente.get('/cadastros/localizacoes/alterar', query_string={'localizacao': '01 - SALA CCI'}),
                   funcoes, 'localizacoes_alterar', {})
    antes_atribuicoes = dados.execute("SELECT COUNT(*) FROM atribuicoes").fetchone()[0]
    # Tela filha via POST com erro: cadastros/atribuir.html com número inválido, sem gravar nada.
    _checar_pagina(cliente, cliente.post('/cadastros/pessoas/atribuir', data={'nome': 'ANA SILVA', 'numero': 'abc'}),
                   funcoes, 'pessoas_atribuir', {})
    # Tela filha de revisão: cadastros/confirmar.html só é exibida; sem o token assinado, nada é gravado.
    _checar_pagina(cliente, cliente.post('/cadastros/pessoas/atribuir', data={'nome': 'ANA SILVA', 'numero': '1001'}),
                   funcoes, 'pessoas_atribuir', {})
    assert dados.execute("SELECT COUNT(*) FROM atribuicoes").fetchone()[0] == antes_atribuicoes
    # Tela filha via POST com erro: não cria usuário, não emite termo.
    r = cliente.post('/usuarios/incluir', data={
        'login': 'novato', 'nome': 'Fulano', 'senha': 'Senha!234', 'confirmacao': 'outra-senha', 'funcoes': ['consulta'],
    })
    assert usuarios.por_login(dados, 'novato') is None
    _checar_pagina(cliente, r, funcoes, 'usuarios.incluir', {})
    assert dados.execute("SELECT COUNT(*) FROM termos_emitidos").fetchone()[0] == antes


def test_paginas_sem_ajuda_nao_mostram_botao(cliente):
    """Login e a negação de acesso (403) não pertencem a nenhuma seção do guia: nenhum botão de ajuda."""
    cliente.post('/sair')
    assert _links_ajuda(cliente.get('/login').get_data(as_text=True)) == []
    assert logar(cliente, 'beltrana', SENHA_PADRAO).status_code == 302
    r = cliente.get('/usuarios')          # só ADMIN; a inventariante é negada
    assert r.status_code == 403
    assert _links_ajuda(r.get_data(as_text=True)) == []


# ---------------------------------------------------------- conteúdo da ajuda por combinação de funções

def test_ajuda_inventariante_sem_relatorios(cliente):
    from tests.conftest import logar,SENHA_PADRAO
    cliente.post('/sair'); logar(cliente,'beltrana',SENHA_PADRAO)
    html=cliente.get('/ajuda').get_data(as_text=True)
    estrutura=Estrutura(html)
    assert {'inventario','conta','perguntas'} <= estrutura.ids
    assert not {'inicio','pesquisa','termos','analise','consulta-inventarios','usuarios'} & estrutura.ids
    assert 'Copiar para o SEI' not in html and 'Exporte a planilha' not in html


def test_ajuda_consulta_sem_instrucao_de_emissao(cliente, usuarios_exemplo):
    cliente.post('/sair'); logar(cliente, *usuarios_exemplo['consulta'])
    html = cliente.get('/ajuda').get_data(as_text=True)
    estrutura = Estrutura(html)
    assert 'termos' in estrutura.ids
    assert 'Copiar para o SEI' not in html and 'Baixar .docx' not in html


def test_ajuda_operador_emissao_sem_instrucoes_de_exclusao_ou_admin(cliente, usuarios_exemplo):
    cliente.post('/sair'); logar(cliente, *usuarios_exemplo['operador'])
    html = cliente.get('/ajuda').get_data(as_text=True)
    estrutura = Estrutura(html)
    assert {'termos', 'cadastros'} <= estrutura.ids
    assert 'usuarios' not in estrutura.ids
    assert 'Copiar para o SEI' in html
    assert 'Exclusões pedem confirmação' not in html and 'A planilha de cadastros substitui' not in html


def test_ajuda_consulta_de_inventarios_relatorios_sem_conferencia(cliente, dados):
    usuarios.criar(dados, 'ci', 'Consulta Inv', SENHA_PADRAO, ['consulta_inventarios'], trocar_senha=False)
    cliente.post('/sair'); logar(cliente, 'ci', SENHA_PADRAO)
    html = cliente.get('/ajuda').get_data(as_text=True)
    estrutura = Estrutura(html)
    assert 'consulta-inventarios' in estrutura.ids and 'inventario' not in estrutura.ids
    assert 'Exporte a planilha' in html
    assert 'Conferência do inventário' not in html


def test_ajuda_combinacao_de_funcoes_soma_secoes(cliente, dados):
    usuarios.criar(dados, 'multi', 'Multi Função', SENHA_PADRAO, ['operador', 'consulta_inventarios'], trocar_senha=False)
    cliente.post('/sair'); logar(cliente, 'multi', SENHA_PADRAO)
    ids_combinado = ajuda(cliente).secoes
    ids_operador = {s['id'] for s in menu.secoes_ajuda(['operador'], True)}
    ids_consulta_inventarios = {s['id'] for s in menu.secoes_ajuda(['consulta_inventarios'], True)}
    esperado = [id for id, *_ in menu.SECOES if id in ids_operador | ids_consulta_inventarios]
    assert ids_combinado == esperado
    assert ids_operador < set(ids_combinado) and ids_consulta_inventarios < set(ids_combinado)
