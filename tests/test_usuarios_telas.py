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
    r = cliente.get("/usuarios?funcao=inventariante")
    assert b">beltrana<" in r.data and b">op<" not in r.data
    assert b'name="funcao" type="radio" value="consulta_inventarios"' in r.data          # filtro oferece todas
    assert b'value="inventariante" checked="checked"' in r.data


def test_novo_usuario_valida_e_cria(cliente, dados):
    r = cliente.get("/usuarios/novo")
    assert r.status_code == 200 and b'name="login"' in r.data and b'name="trocar_senha"' in r.data
    assert b'name="funcoes" type="checkbox"' in r.data and b"checked" not in r.data.split(b"</fieldset>")[0]   # nada pré-marcado
    assert b'name="funcoes" type="checkbox" value="admin" required' not in r.data                             # sem required por caixa
    r = cliente.post("/usuarios/incluir", data={"login": "Novo Login", "nome": "Novo", "funcoes": ["operador"], "senha": "Senha!234", "confirmacao": "Senha!234"})
    assert r.status_code == 200 and "Login inválido".encode() in r.data and b'value="Novo"' in r.data      # dados mantidos
    assert b'value="operador" checked' in r.data                                                          # e a seleção também
    r = cliente.post("/usuarios/incluir", data={"login": "novo", "nome": "Novo", "funcoes": ["operador"], "senha": "Senha!234", "confirmacao": "Outra!234"})
    assert r.status_code == 200 and "confirmação".encode() in r.data
    r = cliente.post("/usuarios/incluir", data={"login": "novo", "nome": "Novo", "funcoes": ["operador"], "senha": "Senha!234", "confirmacao": "Senha!234", "trocar_senha": "1"}, follow_redirects=True)
    assert "Usuário novo criado".encode() in r.data and b">novo<" in r.data
    u = usuarios.por_login(dados, "novo")
    assert u["funcoes"] == ("operador",) and u["trocar_senha"] == 1
    r = cliente.post("/usuarios/incluir", data={"login": "semtroca", "nome": "S", "funcoes": ["consulta"], "senha": "Senha!234", "confirmacao": "Senha!234"}, follow_redirects=True)
    assert usuarios.por_login(dados, "semtroca")["trocar_senha"] == 0


def test_criar_com_varias_funcoes_e_filtrar_sem_duplicar(cliente, dados):
    r = cliente.post("/usuarios/incluir", data={
        "login": "multiplas", "nome": "Múltiplas", "senha": "Senha!234", "confirmacao": "Senha!234",
        "funcoes": ["inventariante", "consulta"],
    })
    assert r.status_code == 302
    assert set(usuarios.por_login(dados, "multiplas")["funcoes"]) == {"inventariante", "consulta"}
    html = cliente.get("/usuarios?funcao=consulta").get_data(as_text=True)
    assert html.count("<code>multiplas</code>") == 1
    r = cliente.post("/usuarios/incluir", data={
        "login": "semfuncao", "nome": "Sem função", "senha": "Senha!234", "confirmacao": "Senha!234",
    })
    assert r.status_code == 200
    assert usuarios.por_login(dados, "semfuncao") is None


def test_editar_funcoes_ativo_e_travas(cliente, dados, usuarios_exemplo):
    op = usuarios.por_login(dados, "op")["id"]
    r = cliente.get(f"/usuarios/{op}/editar")
    assert r.status_code == 200 and b">op<" in r.data and b'name="login"' not in r.data                       # login não editável
    assert b'value="operador" checked' in r.data
    r = cliente.post(f"/usuarios/{op}/editar", data={"nome": "Operador Editado", "funcoes": ["inventariante", "consulta"], "ativo": "1"}, follow_redirects=True)
    assert b"Operador Editado" in r.data and usuarios.por_id(dados, op)["funcoes"] == ("consulta", "inventariante")
    r = cliente.post(f"/usuarios/{op}/editar", data={"nome": "Operador Editado", "funcoes": ["inventariante"]}, follow_redirects=True)
    assert usuarios.por_id(dados, op)["ativo"] == 0 and b"inativo" in r.data.lower()
    eu = usuarios.por_login(dados, ADMIN_LOGIN)["id"]
    r = cliente.post(f"/usuarios/{eu}/editar", data={"nome": "Fulano", "funcoes": ["consulta"], "ativo": "1"})
    assert r.status_code == 200 and "própria conta".encode() in r.data
    assert usuarios.por_id(dados, eu)["funcoes"] == ("admin",)
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
    r = cliente.post(f"/usuarios/{op}/editar?retorno=%2Fusuarios%3Fq%3Doper", data={"nome": "Operador Teste", "funcoes": ["operador"], "ativo": "1"})
    assert r.headers["Location"] == "/usuarios?q=oper"


def test_email_no_cadastro_e_na_lista(cliente, dados):
    r = cliente.get("/usuarios/novo")
    assert b'name="email"' in r.data
    r = cliente.post("/usuarios/incluir", data={"login": "novo", "nome": "Novo", "funcoes": ["operador"], "senha": "Senha!234", "confirmacao": "Senha!234", "email": "Novo@cfc.org.br"}, follow_redirects=True)
    assert b"novo@cfc.org.br" in r.data                                             # coluna E-mail na lista
    r = cliente.post("/usuarios/incluir", data={"login": "outro", "nome": "Outro", "funcoes": ["operador"], "senha": "Senha!234", "confirmacao": "Senha!234", "email": "novo@cfc.org.br"})
    assert r.status_code == 200 and "e-mail já".encode() in r.data and b'value="Outro"' in r.data and b'value="novo@cfc.org.br"' in r.data
    uid = usuarios.por_login(dados, "novo")["id"]
    r = cliente.get(f"/usuarios/{uid}/editar")
    assert b'name="email"' in r.data and b'value="novo@cfc.org.br"' in r.data
    cliente.post(f"/usuarios/{uid}/editar", data={"nome": "Novo", "funcoes": ["operador"], "ativo": "1", "email": ""})
    assert usuarios.por_id(dados, uid)["email"] is None
