"""Árvore de navegação: quem vê cada item e qual item fica marcado como o atual."""
import pytest

import menu


@pytest.fixture
def requisicao(dados):
    """`url_for` exige contexto de requisição; `dados` isola a pasta da instalação."""
    from app import app
    with app.test_request_context():
        yield


def test_somente_cadastro_correto_fica_ativo(requisicao):
    arvore = menu.montar({'admin'}, None, 'cadastros', {'aba': 'pessoas'}, True)
    grupo = next(i for i in arvore if i['id'] == 'cadastros')
    assert grupo['aberto']
    assert [f['rotulo'] for f in grupo['filhos'] if f['ativo']] == ['Pessoas']


def test_inventariante_nao_tem_painel_nem_acervo(requisicao):
    arvore = menu.montar({'inventariante'}, {'id': 7, 'nome': 'Meu evento'},
                         'inventario.sala_tela', {'id': 7, 'localizacao': 'Sala'}, True)
    assert [i['rotulo'] for i in arvore] == ['Inventário', 'Ajuda']
    assert [f['rotulo'] for f in arvore[0]['filhos']] == ['Eventos', 'Meu evento']
    assert arvore[0]['aberto']


def test_tela_derivada_de_cadastro_marca_a_aba_de_origem(requisicao):
    """Editar um centro de custo é uma tela sem item próprio: acende Centros de custo."""
    arvore = menu.montar({'admin'}, None, 'responsaveis_editar', {'ccustos': 'CCI'}, True)
    grupo = next(i for i in arvore if i['id'] == 'cadastros')
    assert grupo['aberto']
    assert [f['rotulo'] for f in grupo['filhos'] if f['ativo']] == ['Centros de custo']


def test_relatorio_de_evento_encerrado_nao_acende_o_evento_aberto(requisicao):
    """Consultar um evento antigo abre o grupo Inventário sem marcar nada do evento aberto."""
    arvore = menu.montar({'admin'}, {'id': 7, 'nome': 'Meu evento'},
                         'inventario.relatorio_tela', {'id': 9}, True)
    grupo = next(i for i in arvore if i['id'] == 'inventario')
    assert grupo['aberto']
    assert [f['rotulo'] for f in grupo['filhos']] == ['Eventos', 'Meu evento', 'Painel', 'Relatório']
    assert [f['rotulo'] for f in grupo['filhos'] if f['ativo']] == []


def test_relatorio_do_evento_aberto_acende_o_item_do_evento(requisicao):
    arvore = menu.montar({'admin'}, {'id': 7, 'nome': 'Meu evento'},
                         'inventario.relatorio_tela', {'id': 7}, True)
    grupo = next(i for i in arvore if i['id'] == 'inventario')
    assert [f['rotulo'] for f in grupo['filhos'] if f['ativo']] == ['Relatório']


def test_modo_local_nao_mostra_usuarios(requisicao):
    local = menu.montar({'admin'}, None, 'home', {}, False)
    web = menu.montar({'admin'}, None, 'home', {}, True)
    assert 'Usuários' in [i['rotulo'] for i in web]
    assert [i['rotulo'] for i in local] == [r['rotulo'] for r in web if r['rotulo'] != 'Usuários']


def test_uniao_de_funcoes_nao_duplica_itens(requisicao):
    """As funções se somam: cada destino aparece uma vez só, e o Início é o item atual."""
    todas = {'admin', 'operador', 'consulta', 'inventariante', 'consulta_inventarios'}
    arvore = menu.montar(todas, {'id': 3, 'nome': 'Inventário 2026'}, 'home', {}, True)
    rotulos = [i['rotulo'] for i in arvore]
    assert rotulos == ['Início', 'Termos de Responsabilidade', 'Análise', 'Inventário', 'Cadastros',
                       'Textos', 'Atualizar base', 'Usuários', 'Ajuda']
    tudo = rotulos + [f['rotulo'] for i in arvore for f in i['filhos']]
    assert len(tudo) == len(set(tudo))
    assert [i['rotulo'] for i in arvore if i['ativo']] == ['Início']


def test_menu_nao_consulta_o_banco():
    """A árvore recebe o evento já autorizado: o módulo não conhece banco, comissões nem inventário."""
    assert not hasattr(menu, 'db') and not hasattr(menu, 'comissoes') and not hasattr(menu, 'inventario')
