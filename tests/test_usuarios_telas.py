"""Telas de usuários (só admin): lista, novo, editar, nova senha temporária."""
import re

import usuarios
from tests.conftest import ADMIN_LOGIN, SENHA_PADRAO, logar


def test_lista_busca_filtro_e_inativos(cliente, dados, usuarios_exemplo):
    r = cliente.get("/usuarios")
    assert r.status_code == 200 and b">admin<" in r.data and b">op<" in r.data and b"Operador" in r.data
    usuarios.editar(dados, usuarios.por_login(dados, "leitor")["id"], "Consulta Teste", ["consulta"], ativo=False)
    assert b">leitor<" not in cliente.get("/usuarios").data
    assert b">leitor<" in cliente.get("/usuarios?inativos=1").data
    r = cliente.get("/usuarios?q=oper")
    assert b">op<" in r.data and b">admin<" not in r.data
    r = cliente.get("/usuarios?perfil=inventariante")
    assert b">beltrana<" in r.data and b">op<" not in r.data


def test_novo_usuario_valida_e_cria(cliente, dados):
    r = cliente.get("/usuarios/novo")
    assert r.status_code == 200 and b'name="login"' in r.data and b'name="trocar_senha"' in r.data
    r = cliente.post("/usuarios/incluir", data={"login": "Novo Login", "nome": "Novo", "perfil": "operador", "senha": "Senha!234", "confirmacao": "Senha!234"})
    assert r.status_code == 200 and "Login inválido".encode() in r.data and b'value="Novo"' in r.data      # dados mantidos
    r = cliente.post("/usuarios/incluir", data={"login": "novo", "nome": "Novo", "perfil": "operador", "senha": "Senha!234", "confirmacao": "Outra!234"})
    assert r.status_code == 200 and "confirmação".encode() in r.data
    r = cliente.post("/usuarios/incluir", data={"login": "novo", "nome": "Novo", "perfil": "operador", "senha": "Senha!234", "confirmacao": "Senha!234", "trocar_senha": "1"}, follow_redirects=True)
    assert "Usuário novo criado".encode() in r.data and b">novo<" in r.data
    u = usuarios.por_login(dados, "novo")
    assert u["perfil"] == "operador" and u["trocar_senha"] == 1
    r = cliente.post("/usuarios/incluir", data={"login": "semtroca", "nome": "S", "perfil": "consulta", "senha": "Senha!234", "confirmacao": "Senha!234"}, follow_redirects=True)
    assert usuarios.por_login(dados, "semtroca")["trocar_senha"] == 0


def test_editar_perfil_ativo_e_travas(cliente, dados, usuarios_exemplo):
    op = usuarios.por_login(dados, "op")["id"]
    r = cliente.get(f"/usuarios/{op}/editar")
    assert r.status_code == 200 and b">op<" in r.data and b'name="login"' not in r.data                       # login não editável
    r = cliente.post(f"/usuarios/{op}/editar", data={"nome": "Operador Editado", "perfil": "inventariante", "ativo": "1"}, follow_redirects=True)
    assert b"Operador Editado" in r.data and usuarios.por_id(dados, op)["perfil"] == "inventariante"
    r = cliente.post(f"/usuarios/{op}/editar", data={"nome": "Operador Editado", "perfil": "inventariante"}, follow_redirects=True)
    assert usuarios.por_id(dados, op)["ativo"] == 0 and b"inativo" in r.data.lower()
    eu = usuarios.por_login(dados, ADMIN_LOGIN)["id"]
    r = cliente.post(f"/usuarios/{eu}/editar", data={"nome": "Fulano", "perfil": "consulta", "ativo": "1"})
    assert r.status_code == 200 and "própria conta".encode() in r.data
    assert usuarios.por_id(dados, eu)["perfil"] == "admin"
    assert cliente.get("/usuarios/999/editar").status_code == 404


def test_nova_senha_temporaria_aparece_uma_vez(cliente, dados, usuarios_exemplo):
    op = usuarios.por_login(dados, "op")["id"]
    r = cliente.post(f"/usuarios/{op}/nova-senha")
    assert r.status_code == 200 and b"Anote agora" in r.data
    assert r.headers["Cache-Control"] == "no-store"
    temp = re.search(rb'id="senha-temporaria">([A-Za-z0-9]{10})<', r.data).group(1).decode()
    assert temp not in cliente.get(f"/usuarios/{op}/editar").data.decode()
    assert usuarios.por_id(dados, op)["trocar_senha"] == 1
    cliente.post("/sair")
    assert logar(cliente, "op", temp).status_code == 302
    assert cliente.get("/").headers["Location"] == "/senha"


def test_lista_preserva_busca_ao_voltar(cliente, dados, usuarios_exemplo):
    op = usuarios.por_login(dados, "op")["id"]
    r = cliente.get(f"/usuarios/{op}/editar?retorno=%2Fusuarios%3Fq%3Doper")
    assert b'href="/usuarios?q=oper"' in r.data
    r = cliente.post(f"/usuarios/{op}/editar?retorno=%2Fusuarios%3Fq%3Doper", data={"nome": "Operador Teste", "perfil": "operador", "ativo": "1"})
    assert r.headers["Location"] == "/usuarios?q=oper"


def test_email_no_cadastro_e_na_lista(cliente, dados):
    r = cliente.get("/usuarios/novo")
    assert b'name="email"' in r.data
    r = cliente.post("/usuarios/incluir", data={"login": "novo", "nome": "Novo", "perfil": "operador", "senha": "Senha!234", "confirmacao": "Senha!234", "email": "Novo@cfc.org.br"}, follow_redirects=True)
    assert b"novo@cfc.org.br" in r.data                                             # coluna E-mail na lista
    r = cliente.post("/usuarios/incluir", data={"login": "outro", "nome": "Outro", "perfil": "operador", "senha": "Senha!234", "confirmacao": "Senha!234", "email": "novo@cfc.org.br"})
    assert r.status_code == 200 and "e-mail já".encode() in r.data and b'value="Outro"' in r.data and b'value="novo@cfc.org.br"' in r.data
    uid = usuarios.por_login(dados, "novo")["id"]
    r = cliente.get(f"/usuarios/{uid}/editar")
    assert b'name="email"' in r.data and b'value="novo@cfc.org.br"' in r.data
    cliente.post(f"/usuarios/{uid}/editar", data={"nome": "Novo", "perfil": "operador", "ativo": "1", "email": ""})
    assert usuarios.por_id(dados, uid)["email"] is None
