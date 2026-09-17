"""Ajuda: seções conforme as funções do usuário, sumário coerente e âncora contextual de cada tela."""
from html.parser import HTMLParser

from flask import g

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


def test_inventariante_ve_a_conferencia_e_nao_a_consulta(cliente):
    assert logar(cliente, 'beltrana', SENHA_PADRAO).status_code == 302
    assert ajuda(cliente).secoes == ['inventario', 'conta', 'perguntas']


def test_consulta_de_inventarios_ve_a_consulta_e_nao_a_conferencia(cliente, dados):
    usuarios.criar(dados, 'ci', 'Consulta Inv', SENHA_PADRAO, ['consulta_inventarios'], trocar_senha=False)
    assert logar(cliente, 'ci', SENHA_PADRAO).status_code == 302
    assert ajuda(cliente).secoes == ['consulta-inventarios', 'conta', 'perguntas']


def test_modo_local_nao_traz_usuarios_nem_conta(cliente_local):
    secoes = ajuda(cliente_local).secoes
    assert 'usuarios' not in secoes and 'conta' not in secoes
    assert secoes == [id for id, *_ in menu.SECOES if id not in {'usuarios', 'conta'}]


def test_ancora_da_tela_atual():
    assert menu.ancora_ajuda({'admin'}, 'bem', {'numero': '1'}, True) == 'pesquisa'
    assert menu.ancora_ajuda({'admin'}, 'responsaveis_editar', {'ccustos': 'CCI'}, True) == 'cadastros'
    assert menu.ancora_ajuda({'admin'}, 'termo', {'tipo': 'individual', 'chave': 'X'}, True) == 'termos'
    assert menu.ancora_ajuda({'admin'}, 'usuarios.senha', {}, True) == 'conta'
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
    assert [s['id'] for s in contexto['SECOES_AJUDA']] == [id for id, *_ in menu.SECOES
                                                          if id not in {'usuarios', 'conta'}]


def test_contexto_sem_usuario_nao_oferece_ajuda(dados):
    from app import app
    with app.test_request_context('/login'):
        contexto = {}
        app.update_template_context(contexto)
    assert contexto['SECOES_AJUDA'] == [] and contexto['AJUDA_ANCORA'] is None


def test_macro_de_ajuda_da_tela(dados):
    from app import app
    with app.test_request_context():
        macros = app.jinja_env.get_template('_macros.html').module
        assert 'href="/ajuda#pesquisa"' in str(macros.ajuda_titulo('pesquisa'))
        assert str(macros.ajuda_titulo(None)).strip() == ''
