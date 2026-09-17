"""Árvore de navegação: quem vê cada item e qual item fica marcado como o atual."""
import pytest

import db
import menu
import usuarios as u
from tests.conftest import SENHA_PADRAO, logar


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


# --- HTML renderizado (Tarefa 3 de 5C): base.html + menu-estado.js -------------------------------------

def _nav(resposta):
    """Recorta só a árvore de navegação da página (evita falsos positivos no resto do HTML)."""
    return resposta.data.split(b'id="main-navigation"')[1].split(b"menu-footer")[0]


def _grupo(nav, rotulo):
    """Isola o <div class="menu-folder">...</div> cujo título (primeiro <span class="content">) é `rotulo`.
    Sem divs aninhadas dentro de um grupo: o primeiro `</div>` depois do título fecha o próprio grupo."""
    for parte in nav.split(b'<div class="menu-folder')[1:]:
        titulo = parte.split(b'<span class="content">', 2)[1].split(b'</span>')[0]
        if titulo == rotulo.encode():
            fim = parte.index(b'</div>')
            return b'<div class="menu-folder' + parte[:fim + len(b'</div>')]
    raise AssertionError(f"grupo {rotulo!r} não encontrado no menu")


def test_nenhum_grupo_vazio_e_renderizado(cliente):
    """<ul role="group"></ul> nunca aparece: menu.montar já descarta grupos sem filhos permitidos."""
    nav = _nav(cliente.get("/"))
    assert b'<ul role="group"></ul>' not in nav


def test_exatamente_um_item_marcado_como_atual(cliente):
    for rota in ("/", "/cadastros/responsaveis", "/textos", "/ajuda"):
        nav = _nav(cliente.get(rota))
        assert nav.count(b'aria-current="page"') == 1, rota


def test_grupo_da_tela_atual_fica_expandido_e_ativo(cliente):
    nav = _nav(cliente.get("/cadastros/responsaveis"))
    grupo = _grupo(nav, "Cadastros")
    assert b'<div class="menu-folder active">' in grupo
    assert b'aria-expanded="true"' in grupo
    assert grupo.count(b'aria-current="page"') == 1


def test_inventariante_sem_vinculo_com_o_evento_nao_ve_link_do_evento(cliente):
    """Evento aberto com Fulano e Beltrana; uma terceira inventariante, fora da comissão, só vê Eventos."""
    conn = db.conectar()
    u.criar(conn, "carla", "Carla", SENHA_PADRAO, ["inventariante"], trocar_senha=False)
    fulano, beltrana = (u.por_login(conn, login)["id"] for login in ("admin", "beltrana"))
    cliente.post("/inventario/abrir", data={"nome": "Inv", "usuarios": [fulano, beltrana], "escopo": "todas"})
    cliente.post("/sair")
    logar(cliente, "carla", SENHA_PADRAO)
    nav = _nav(cliente.get("/inventario"))
    grupo = _grupo(nav, "Inventário")
    assert b">Eventos<" in grupo and b">Inv<" not in grupo and b"/painel" not in grupo


def test_consulta_inventarios_ve_painel_e_relatorio_do_evento_aberto(cliente):
    """`consulta_inventarios` enxerga qualquer evento e ganha Painel/Relatório (RELATORIOS)."""
    conn = db.conectar()
    u.criar(conn, "chefe", "Chefe", SENHA_PADRAO, ["consulta_inventarios"], trocar_senha=False)
    fulano, beltrana = (u.por_login(conn, login)["id"] for login in ("admin", "beltrana"))
    cliente.post("/inventario/abrir", data={"nome": "Inv", "usuarios": [fulano, beltrana], "escopo": "todas"})
    cliente.post("/sair")
    logar(cliente, "chefe", SENHA_PADRAO)
    nav = _nav(cliente.get("/inventario"))
    grupo = _grupo(nav, "Inventário")
    assert b">Painel<" in grupo and b">Relat" in grupo


def test_script_de_estado_do_menu_esta_presente(cliente):
    assert b'src="/static/js/menu-estado.js"' in cliente.get("/").data
