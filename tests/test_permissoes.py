"""Matriz de permissões: unidade (permitido) e cobertura de todas as rotas do app."""
import pytest

import usuarios
from tests.conftest import logar


def test_permitido_por_perfil():
    assert usuarios.permitido("consulta", "home") and usuarios.permitido("inventariante", "bem")
    assert usuarios.permitido("consulta", "termo") and usuarios.permitido("consulta", "termo_documento")
    assert not usuarios.permitido("inventariante", "termo") and not usuarios.permitido("inventariante", "centro_custos")
    assert not usuarios.permitido("consulta", "termo_docx") and not usuarios.permitido("consulta", "termo_registrar", "POST")
    assert usuarios.permitido("consulta", "termo_devolucao") and not usuarios.permitido("consulta", "termo_devolucao", "POST")
    assert usuarios.permitido("operador", "termo_devolucao", "POST")
    assert usuarios.permitido("operador", "responsaveis_editar", "POST") and not usuarios.permitido("operador", "responsaveis_excluir", "POST")
    assert usuarios.permitido("operador", "upload", "POST") and not usuarios.permitido("operador", "importar_cadastros", "POST")
    assert not usuarios.permitido("inventariante", "cadastros") and not usuarios.permitido("inventariante", "textos_tela")
    assert usuarios.permitido("inventariante", "inventario.ler", "POST") and not usuarios.permitido("consulta", "inventario.ler", "POST")
    assert not usuarios.permitido("operador", "inventario.abrir", "POST") and usuarios.permitido("admin", "inventario.excluir", "POST")
    assert not usuarios.permitido("operador", "usuarios.lista") and usuarios.permitido("admin", "usuarios.lista")
    for perfil in usuarios.PERFIS:
        assert usuarios.permitido(perfil, "usuarios.senha", "POST") and usuarios.permitido(perfil, "usuarios.sair", "POST")
    assert usuarios.permitido("admin", "termo_docx", "HEAD")                 # HEAD conta como GET
    assert not usuarios.permitido("admin", "rota_que_nao_existe")            # fora da matriz = negado a todos
    assert not usuarios.permitido("chefe", "home")


def test_toda_rota_do_app_esta_na_matriz():
    from app import app
    faltam = []
    for regra in app.url_map.iter_rules():
        if regra.endpoint == "static":
            continue
        for metodo in regra.methods - {"HEAD", "OPTIONS"}:
            chave = regra.endpoint if metodo == "GET" else f"{regra.endpoint}:{metodo}"
            if chave not in usuarios.PERMISSOES and regra.endpoint not in usuarios.PERMISSOES:
                faltam.append(f"{regra.endpoint} [{metodo}]")
    assert not faltam, "Rotas sem regra em usuarios.PERMISSOES: " + ", ".join(sorted(faltam))


# (endpoint GET exemplar por área, e o que cada perfil deve receber)
# Nota: "/usuarios" só existe na Task 6 (rota ainda não criada); fica de fora até lá.
_ROTAS_GET = {
    "/": {"admin": 200, "operador": 200, "inventariante": 200, "consulta": 200},
    "/bem?numero=1001": {"admin": 200, "operador": 200, "inventariante": 200, "consulta": 200},
    "/centro-custos": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 200},
    "/termo/ccusto/CCI": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 200},
    "/termo/ccusto/CCI/documento": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 200},
    "/termo/ccusto/CCI/planilha": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 403},
    "/cadastros/responsaveis": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 403},
    "/textos": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 403},
    "/upload": {"admin": 200, "operador": 200, "inventariante": 403, "consulta": 403},
    "/inventario": {"admin": 200, "operador": 200, "inventariante": 200, "consulta": 200},
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


@pytest.mark.parametrize("perfil", ["admin", "operador", "inventariante", "consulta"])
def test_rotas_por_perfil(cliente, usuarios_exemplo, perfil):
    login, senha = usuarios_exemplo[perfil]
    cliente.post("/sair")
    assert logar(cliente, login, senha).status_code == 302
    for rota, esperado in _ROTAS_GET.items():
        r = cliente.get(rota)
        assert r.status_code == esperado[perfil], f"{perfil} GET {rota}: {r.status_code} != {esperado[perfil]}"
        if esperado[perfil] == 403:
            assert "Seu perfil não tem acesso a isso".encode() in r.data and b"main-navigation" in r.data
    for rota, esperado in _ROTAS_POST.items():
        r = cliente.post(rota, data={"ccusto": "CCI", "nome": "ANA SILVA", "limpar": "1"})
        assert r.status_code == esperado[perfil], f"{perfil} POST {rota}: {r.status_code} != {esperado[perfil]}"


def test_403_em_json_para_rotas_do_leitor(cliente, usuarios_exemplo):
    cliente.post("/sair")
    logar(cliente, *usuarios_exemplo["consulta"])
    r = cliente.post("/inventario/1/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    assert r.status_code == 403 and r.get_json()["erro"].startswith("Seu perfil")
    r = cliente.post("/termo/ccusto/CCI/registrar")
    assert r.status_code == 403 and r.get_json()["erro"].startswith("Seu perfil")


def test_menu_por_perfil(cliente, usuarios_exemplo):
    def menu():
        return cliente.get("/").data.split(b'id="main-navigation"')[1].split(b"menu-footer")[0]
    m = menu()
    # "Usuários" ainda não entra: a rota usuarios.lista só existe na Task 6 (item comentado em app.py até lá).
    assert b">Usu\xc3\xa1rios<" not in m and b">Textos<" in m and b">Cadastros<" in m
    cliente.post("/sair"); logar(cliente, *usuarios_exemplo["operador"])
    m = menu()
    assert b">Usu\xc3\xa1rios<" not in m and b">Textos<" in m and b">Cadastros<" in m and b">Atualizar base<" in m
    cliente.post("/sair"); logar(cliente, *usuarios_exemplo["inventariante"])
    m = menu()
    assert b">Cadastros<" not in m and b">Termo por centro de custo<" not in m and b">Invent\xc3\xa1rio<" in m and b">Recorte<" in m
    cliente.post("/sair"); logar(cliente, *usuarios_exemplo["consulta"])
    m = menu()
    assert b">Termo por centro de custo<" in m and b">Textos<" not in m and b">Atualizar base<" not in m


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
