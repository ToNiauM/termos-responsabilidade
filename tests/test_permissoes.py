"""Matriz de permissões: unidade (permitido) e cobertura de todas as rotas do app."""
import pytest

import permissoes
import usuarios
from tests.conftest import logar

NEGADO = "Seu usuário não tem permissão para esta ação."


def test_permitido_por_funcao():
    assert usuarios.permitido(["consulta"], "home") and usuarios.permitido(["inventariante"], "home")
    assert usuarios.permitido(["consulta"], "termo") and usuarios.permitido(["consulta"], "termo_documento")
    assert not usuarios.permitido(["inventariante"], "termo") and not usuarios.permitido(["inventariante"], "centro_custos")
    assert not usuarios.permitido(["consulta"], "termo_docx") and not usuarios.permitido(["consulta"], "termo_registrar", "POST")
    assert usuarios.permitido(["consulta"], "termo_devolucao") and not usuarios.permitido(["consulta"], "termo_devolucao", "POST")
    assert usuarios.permitido(["operador"], "termo_devolucao", "POST")
    assert usuarios.permitido(["operador"], "responsaveis_editar", "POST") and not usuarios.permitido(["operador"], "responsaveis_excluir", "POST")
    assert usuarios.permitido(["operador"], "upload", "POST") and not usuarios.permitido(["operador"], "importar_cadastros", "POST")
    assert not usuarios.permitido(["inventariante"], "cadastros") and not usuarios.permitido(["inventariante"], "textos_tela")
    assert usuarios.permitido(["inventariante"], "inventario.ler", "POST") and not usuarios.permitido(["consulta"], "inventario.ler", "POST")
    assert not usuarios.permitido(["operador"], "inventario.abrir", "POST") and usuarios.permitido(["admin"], "inventario.excluir", "POST")
    assert not usuarios.permitido(["operador"], "usuarios.lista") and usuarios.permitido(["admin"], "usuarios.lista")
    for funcao in usuarios.FUNCOES:
        assert usuarios.permitido([funcao], "usuarios.senha", "POST") and usuarios.permitido([funcao], "usuarios.sair", "POST")
    assert usuarios.permitido(["admin"], "termo_docx", "HEAD")                 # HEAD conta como GET
    assert not usuarios.permitido(["admin"], "rota_que_nao_existe")            # fora da matriz = negado a todos
    assert not usuarios.permitido(["chefe"], "home") and not usuarios.permitido([], "home")
    assert not usuarios.permitido("consulta", "home")                          # texto não é coleção de funções
    somadas = ["inventariante", "consulta"]                                    # as funções se somam
    assert usuarios.permitido(somadas, "bem") and usuarios.permitido(somadas, "inventario.ler", "POST")
    assert not usuarios.permitido(somadas, "termo_docx")


def test_toda_rota_do_app_esta_na_matriz(dados):
    """Cada (endpoint, método) do app tem regra própria; não há regra geral por endpoint.
    A recíproca não vale: a matriz já registra `ajuda`, `analise` e `analise_xlsx`, rotas dos planos 5B/5C."""
    from app import app
    faltam = []
    for regra in app.url_map.iter_rules():
        if regra.endpoint == "static":
            continue
        for metodo in regra.methods - {"HEAD", "OPTIONS"}:
            if (regra.endpoint, metodo) not in permissoes.PERMISSOES:
                faltam.append(f"{regra.endpoint} [{metodo}]")
    assert not faltam, "Rotas sem regra em permissoes.PERMISSOES: " + ", ".join(sorted(faltam))


def test_rotas_de_escrita_do_inventario_saem_da_matriz(dados):
    """app.ESCRITA_INVENTARIO é derivada de CONFERENCIA: uma lista só, sem cópia para desencontrar."""
    import app as web
    assert web.ESCRITA_INVENTARIO == {
        "inventario.ler", "inventario.atualizar_leitura", "inventario.lote", "inventario.foto_leitura",
        "inventario.foto_excluir", "inventario.sobra", "inventario.sobra_excluir"}
    assert all(permissoes.PERMISSOES[(ep, "POST")] is permissoes.CONFERENCIA for ep in web.ESCRITA_INVENTARIO)


# (endpoint GET exemplar por área, e o que cada função isolada deve receber; 302 = redirecionado ao destino inicial)
_ROTAS_GET = {
    "/": {"admin": 200, "operador": 200, "inventariante": 302, "consulta": 200},
    "/usuarios": {"admin": 200, "operador": 403, "inventariante": 403, "consulta": 403},
    "/bem?numero=1001": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 200},
    "/centro-custos": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 200},
    "/termo/ccusto/CCI": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 200},
    "/termo/ccusto/CCI/documento": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 200},
    "/termo/ccusto/CCI/planilha": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 403},
    "/cadastros/responsaveis": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 403},
    "/textos": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 403},
    "/upload": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 403},
    "/inventario": {"admin": 200, "operador": 403, "inventariante": 200, "consulta": 403},
}
_ROTAS_POST = {
    "/gerar": {"admin": 302, "operador": 302, "inventariante": 403, "consulta": 403},
    "/termo_devolucao": {"admin": 302, "operador": 302, "inventariante": 403, "consulta": 403},
    # exclusão exige o token de revisão assinado (fluxo de 2 passos, ver confirmar_revisao em conftest.py); com
    # um único POST simples quem tem permissão cai na tela de confirmação (200) em vez de executar (302)
    "/cadastros/responsaveis/excluir": {"admin": 200, "operador": 403, "inventariante": 403, "consulta": 403},
    "/importar-cadastros": {"admin": 302, "operador": 403, "inventariante": 403, "consulta": 403},
    "/inventario/abrir": {"admin": 302, "operador": 403, "inventariante": 403, "consulta": 403},
}


@pytest.mark.parametrize("funcao", ["admin", "operador", "inventariante", "consulta"])
def test_rotas_por_funcao(cliente, usuarios_exemplo, funcao):
    login, senha = usuarios_exemplo[funcao]
    cliente.post("/sair")
    assert logar(cliente, login, senha).status_code == 302
    for rota, esperado in _ROTAS_GET.items():
        r = cliente.get(rota)
        assert r.status_code == esperado[funcao], f"{funcao} GET {rota}: {r.status_code} != {esperado[funcao]}"
        if esperado[funcao] == 403:
            assert NEGADO.encode() in r.data and b"main-navigation" in r.data
    for rota, esperado in _ROTAS_POST.items():
        r = cliente.post(rota, data={"ccusto": "CCI", "nome": "ANA SILVA", "limpar": "1"})
        assert r.status_code == esperado[funcao], f"{funcao} POST {rota}: {r.status_code} != {esperado[funcao]}"


def test_403_em_json_para_rotas_do_leitor(cliente, usuarios_exemplo):
    cliente.post("/sair")
    logar(cliente, *usuarios_exemplo["consulta"])
    r = cliente.post("/inventario/1/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    assert r.status_code == 403 and r.get_json()["erro"] == NEGADO
    r = cliente.post("/termo/ccusto/CCI/registrar")
    assert r.status_code == 403 and r.get_json()["erro"] == NEGADO


def test_options_segue_a_matriz(cliente, usuarios_exemplo):
    """A resposta automática de OPTIONS do Flask só sai se algum método real da rota for permitido."""
    r = cliente.open("/upload", method="OPTIONS")
    assert r.status_code == 200 and "GET" in r.headers["Allow"] and "POST" in r.headers["Allow"]
    cliente.post("/sair"); logar(cliente, *usuarios_exemplo["consulta"])
    assert cliente.open("/upload", method="OPTIONS").status_code == 403      # não pode nem GET nem POST
    assert cliente.open("/bem", method="OPTIONS").status_code == 200         # pode o GET da rota


def test_menu_por_funcao(cliente, usuarios_exemplo):
    def menu(rota="/"):
        return cliente.get(rota).data.split(b'id="main-navigation"')[1].split(b"menu-footer")[0]
    m = menu()
    assert b">Usu\xc3\xa1rios<" in m and b">Textos<" in m and b">Cadastros<" in m and b">In\xc3\xadcio<" in m
    cliente.post("/sair"); logar(cliente, *usuarios_exemplo["operador"])
    m = menu()
    assert b">Usu\xc3\xa1rios<" not in m and b">Textos<" in m and b">Cadastros<" in m and b">Atualizar base<" in m
    assert b">Invent\xc3\xa1rio<" not in m                                       # operador não tem função de inventário
    cliente.post("/sair"); logar(cliente, *usuarios_exemplo["inventariante"])
    m = menu("/inventario")
    assert b">Cadastros<" not in m and b">Termo por centro de custo<" not in m and b">Invent\xc3\xa1rio<" in m
    assert b">Recorte<" not in m and b">In\xc3\xadcio<" not in m                  # sem acesso ao acervo, sem Início
    cliente.post("/sair"); logar(cliente, *usuarios_exemplo["consulta"])
    m = menu()
    assert b">Termo por centro de custo<" in m and b">Textos<" not in m and b">Atualizar base<" not in m
    assert b">In\xc3\xadcio<" in m and b">Invent\xc3\xa1rio<" not in m


def test_menu_do_inventario_separa_conferencia_de_relatorios(cliente, dados):
    """Inventariante vê Eventos e o evento aberto de que participa; Painel/Relatório são de quem tem RELATORIOS."""
    import usuarios as u
    from tests.conftest import SENHA_PADRAO
    beltrana = u.por_login(dados, "beltrana")["id"]
    cliente.post("/inventario/abrir", data={"nome": "Inv", "usuarios": [beltrana], "escopo": "todas"})
    u.criar(dados, "chefe", "Chefe", SENHA_PADRAO, ["consulta_inventarios"], trocar_senha=False)
    cliente.post("/sair"); logar(cliente, "beltrana", SENHA_PADRAO)
    m = cliente.get("/inventario").data.split(b'id="main-navigation"')[1].split(b"menu-footer")[0]
    assert b">Eventos<" in m and b">Inv<" in m and b">Painel<" not in m and b">Relat\xc3\xb3rio<" not in m
    cliente.post("/sair"); logar(cliente, "chefe", SENHA_PADRAO)
    m = cliente.get("/inventario").data.split(b'id="main-navigation"')[1].split(b"menu-footer")[0]
    assert b">Eventos<" in m and b">Inv<" in m and b">Painel<" in m and b">Relat\xc3\xb3rio<" in m


def test_consulta_nao_ve_botoes_de_emissao(cliente, usuarios_exemplo):
    cliente.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "T", "numero_sei": "1111", "vigente": "1"})
    assert b"Copiar para o SEI" in cliente.get("/termo/ccusto/CCI").data
    assert b"Gerar termo" in cliente.get("/centro-custos").data
    assert b"Gerar termo" in cliente.get("/termos-individuais").data
    assert cliente.get("/termo/ccusto/CCI/docx").status_code == 200          # gera e registra a emissão (id 1)
    r = cliente.get("/termos-emitidos/1")
    assert r.status_code == 200 and b"Salvar" in r.data
    cliente.post("/sair"); logar(cliente, *usuarios_exemplo["consulta"])
    r = cliente.get("/termo/ccusto/CCI")
    assert r.status_code == 200 and b"Copiar para o SEI" not in r.data and b"Baixar .docx" not in r.data and b"Baixar planilha" not in r.data
    assert b"Processo SEI 1111" in r.data
    r = cliente.get("/centro-custos")
    assert r.status_code == 200 and b"Gerar termo" not in r.data
    r = cliente.get("/termos-individuais")
    assert r.status_code == 200 and b"Gerar termo" not in r.data
    r = cliente.get("/termos-emitidos/1")
    assert r.status_code == 200 and b"Salvar" not in r.data


def test_modo_desktop_sem_tela_de_usuarios(cliente_local):
    assert cliente_local.get("/usuarios").status_code == 404
    m = cliente_local.get("/").data.split(b'id="main-navigation"')[1].split(b"menu-footer")[0]
    assert b">Usu\xc3\xa1rios<" not in m and b">Textos<" in m
