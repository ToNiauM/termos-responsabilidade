"""Usuários, senhas e permissões. Só dados e regras: toda função recebe `conn` primeiro e não importa Flask
(padrão de db.py). Hash de senha com werkzeug.security (scrypt), que já vem com o Flask."""
import re
import secrets
from datetime import datetime, timedelta

from werkzeug.security import check_password_hash, generate_password_hash

from db import ErroDeNegocio, _agora, _obrigatorio, _todos, _um

PERFIS = ("admin", "operador", "inventariante", "consulta")
ROTULO_PERFIL = {"admin": "Administrador", "operador": "Operador", "inventariante": "Inventariante", "consulta": "Consulta"}
PERFIS_COMISSAO = ("admin", "operador", "inventariante")
SENHA_MINIMA = 8
MAX_FALHAS = 5
BLOQUEIO_MINUTOS = 15
ALFABETO_TEMP = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789"
_LOGIN = re.compile(r"[a-z0-9._-]{2,30}")
_FORMATO = "%Y-%m-%d %H:%M:%S"

# Modo desktop (sem TERMOS_LOGIN): quem usa o programa Windows é o administrador da instalação.
USUARIO_LOCAL = {"id": None, "login": "local", "nome": "Administrador local", "perfil": "admin", "ativo": 1, "trocar_senha": 0}

_COLUNAS_LISTA = "id, login, nome, perfil, ativo, trocar_senha, falhas, bloqueado_ate, criado_em, ultimo_acesso"


def _agora_dt() -> datetime:
    """Separado para os testes simularem o relógio (bloqueio de 15 min)."""
    return datetime.now()


def _validar_login(login) -> str:
    v = str(login or "").strip().lower()
    if not _LOGIN.fullmatch(v):
        raise ErroDeNegocio("Login inválido: use de 2 a 30 caracteres entre letras minúsculas, números, ponto, hífen e sublinhado.")
    return v


def _validar_senha(senha) -> str:
    s = str(senha or "")
    if len(s) < SENHA_MINIMA:
        raise ErroDeNegocio(f"A senha precisa ter ao menos {SENHA_MINIMA} caracteres.")
    return s


def _validar_perfil(perfil) -> str:
    if perfil not in PERFIS:
        raise ErroDeNegocio("Perfil inválido.")
    return perfil


def criar(conn, login, nome, senha, perfil, trocar_senha=True) -> int:
    login = _validar_login(login)
    nome = _obrigatorio(nome, "Nome")
    senha = _validar_senha(senha)
    perfil = _validar_perfil(perfil)
    if por_login(conn, login):
        raise ErroDeNegocio(f"O login {login} já existe.")
    cur = conn.execute("INSERT INTO usuarios (login, nome, senha_hash, perfil, trocar_senha, criado_em) VALUES (?,?,?,?,?,?)",
                       (login, nome, generate_password_hash(senha), perfil, 1 if trocar_senha else 0, _agora()))
    conn.commit()
    return cur.lastrowid


def por_id(conn, id) -> dict | None:
    return _um(conn, "SELECT * FROM usuarios WHERE id = ?", id) if id is not None else None


def por_login(conn, login) -> dict | None:
    return _um(conn, "SELECT * FROM usuarios WHERE login = ?", str(login or "").strip().lower())


def listar(conn, busca="", perfil=None, inativos=False) -> list[dict]:
    """Sem senha_hash. busca casa em login e nome (sem caixa); inativos=False esconde os inativos."""
    sql, p = f"SELECT {_COLUNAS_LISTA} FROM usuarios WHERE 1=1", []
    if not inativos:
        sql += " AND ativo = 1"
    if perfil:
        sql += " AND perfil = ?"
        p.append(perfil)
    if busca and busca.strip():
        sql += " AND (lower(login) LIKE ? OR lower(nome) LIKE ?)"
        termo = f"%{busca.strip().lower()}%"
        p += [termo, termo]
    return _todos(conn, sql + " ORDER BY login", *p)


def _admins_ativos(conn) -> int:
    return conn.execute("SELECT count(*) FROM usuarios WHERE perfil = 'admin' AND ativo = 1").fetchone()[0]


def editar(conn, id, nome, perfil, ativo, logado_id=None) -> None:
    """Nome, perfil e ativo. Login não muda. Travas: o próprio logado não se inativa nem se rebaixa; o último
    administrador ativo não é inativado nem rebaixado."""
    u = por_id(conn, id)
    if not u:
        raise ErroDeNegocio("Usuário não encontrado.")
    nome = _obrigatorio(nome, "Nome")
    perfil = _validar_perfil(perfil)
    ativo = 1 if ativo else 0
    perde_admin = u["perfil"] == "admin" and u["ativo"] and (perfil != "admin" or not ativo)
    if perde_admin and logado_id is not None and int(logado_id) == int(id):
        raise ErroDeNegocio("Você não pode rebaixar nem inativar a própria conta.")
    if perde_admin and _admins_ativos(conn) <= 1:
        raise ErroDeNegocio("Este é o último administrador ativo: não pode ser rebaixado nem inativado.")
    conn.execute("UPDATE usuarios SET nome = ?, perfil = ?, ativo = ? WHERE id = ?", (nome, perfil, ativo, id))
    conn.commit()


def elegiveis_comissao(conn) -> list[dict]:
    """Usuários ativos que podem compor a comissão de um inventário (admin, operador, inventariante), por nome."""
    marcas = ",".join("?" * len(PERFIS_COMISSAO))
    return _todos(conn, f"SELECT {_COLUNAS_LISTA} FROM usuarios WHERE ativo = 1 AND perfil IN ({marcas}) ORDER BY nome, login",
                  *PERFIS_COMISSAO)


# ---------------------------------------------------------------- permissões (nega por padrão)
TODOS = frozenset(PERFIS)
GESTAO = frozenset({"admin", "operador"})
ADMIN = frozenset({"admin"})
LEITURA = frozenset({"admin", "operador", "inventariante"})     # ler no inventário: exige ainda estar na comissão
TERMOS_VER = frozenset({"admin", "operador", "consulta"})       # telas de termos: inventariante não vê

# Chave = endpoint Flask; "<endpoint>:POST" quando o POST tem regra diferente do GET. Rota ausente = 403 para todos
# (tests/test_permissoes.py garante que toda rota do app está aqui).
PERMISSOES = {
    # consulta geral
    "home": TODOS, "bem": TODOS, "pesquisa": TODOS, "recorte": TODOS, "recorte_xlsx": TODOS,
    # termos: ver para admin, operador e consulta; emitir/registrar só gestão
    "centro_custos": TERMOS_VER, "termos_individuais": TERMOS_VER, "termo": TERMOS_VER, "termo_documento": TERMOS_VER,
    "termo_devolucao": TERMOS_VER, "termo_devolucao:POST": GESTAO,
    "termos_emitidos_tela": TERMOS_VER, "termo_emitido_tela": TERMOS_VER,
    "gerar": GESTAO, "gerar_individual": GESTAO, "termo_docx": GESTAO, "termo_planilha": GESTAO, "termo_registrar": GESTAO,
    "termo_emitido_documento": GESTAO, "termo_emitido_email": GESTAO,
    # cadastros: gestão inclui/edita; só admin exclui
    "cadastros": GESTAO, "cadastro_novo": GESTAO,
    "responsaveis_incluir": GESTAO, "responsaveis_editar": GESTAO, "responsaveis_excluir": ADMIN,
    "pessoas_incluir": GESTAO, "pessoas_editar": GESTAO, "pessoas_excluir": ADMIN,
    "pessoas_atribuir": GESTAO, "pessoas_desatribuir": GESTAO,
    "localizacoes_incluir": GESTAO, "localizacoes_alterar": GESTAO, "localizacoes_mover": GESTAO, "localizacoes_excluir": ADMIN,
    "processos_incluir": GESTAO, "processos_vigente": GESTAO, "processos_encerrar": GESTAO, "processos_excluir": ADMIN,
    "cadastros_exportar": GESTAO, "importar_cadastros": ADMIN,
    # textos e base
    "textos_tela": GESTAO, "textos_salvar": GESTAO,
    "upload": GESTAO, "bens_exportar": GESTAO, "importacao_tela": GESTAO,
    # inventário
    "inventario.eventos_tela": TODOS, "inventario.evento_tela": TODOS, "inventario.sala_tela": TODOS,
    "inventario.relatorio_tela": TODOS, "inventario.xlsx": TODOS, "inventario.painel_tela": TODOS,
    "inventario.abrir": ADMIN, "inventario.encerrar": ADMIN, "inventario.comissao": ADMIN, "inventario.excluir": ADMIN,
    "inventario.ler": LEITURA, "inventario.atualizar_leitura": LEITURA, "inventario.lote": LEITURA,
    "inventario.foto_leitura": LEITURA, "inventario.foto_excluir": LEITURA, "inventario.sobra": LEITURA, "inventario.sobra_excluir": LEITURA,
    "inventario.integrante": LEITURA,  # temporário: rota sai na Task 7
    # conta e usuários
    "usuarios.login": TODOS, "usuarios.sair": TODOS, "usuarios.senha": TODOS,
    "usuarios.lista": ADMIN, "usuarios.novo": ADMIN, "usuarios.incluir": ADMIN, "usuarios.editar": ADMIN, "usuarios.nova_senha": ADMIN,
}


def permitido(perfil, endpoint, metodo="GET") -> bool:
    metodo = "GET" if metodo in ("GET", "HEAD") else metodo
    regra = PERMISSOES.get(f"{endpoint}:{metodo}") if metodo != "GET" else None
    if regra is None:
        regra = PERMISSOES.get(endpoint)
    return regra is not None and perfil in regra


# ---------------------------------------------------------------- autenticação e senhas
_INVALIDO = "Usuário ou senha inválidos."


def autenticar(conn, login, senha) -> dict:
    """Devolve o usuário. Mensagem única para inexistente, inativo e senha errada. 5 falhas seguidas bloqueiam
    por 15 minutos (a senha certa também falha nesse período e não conta como falha)."""
    u = por_login(conn, login)
    if not u or not u["ativo"]:
        raise ErroDeNegocio(_INVALIDO)
    agora = _agora_dt()
    if u["bloqueado_ate"] and agora.strftime(_FORMATO) < u["bloqueado_ate"]:
        raise ErroDeNegocio(f"Muitas tentativas. Aguarde {BLOQUEIO_MINUTOS} minutos.")
    if not check_password_hash(u["senha_hash"], str(senha or "")):
        falhas = u["falhas"] + 1
        bloqueado = (agora + timedelta(minutes=BLOQUEIO_MINUTOS)).strftime(_FORMATO) if falhas >= MAX_FALHAS else None
        conn.execute("UPDATE usuarios SET falhas = ?, bloqueado_ate = ? WHERE id = ?", (falhas, bloqueado, u["id"]))
        conn.commit()
        raise ErroDeNegocio(_INVALIDO)
    conn.execute("UPDATE usuarios SET falhas = 0, bloqueado_ate = NULL, ultimo_acesso = ? WHERE id = ?",
                 (agora.strftime(_FORMATO), u["id"]))
    conn.commit()
    return por_id(conn, u["id"])


def nova_senha_temporaria(conn, id) -> str:
    """Gera, grava o hash e devolve a senha em claro UMA vez; obriga a troca e limpa bloqueio."""
    if not por_id(conn, id):
        raise ErroDeNegocio("Usuário não encontrado.")
    senha = "".join(secrets.choice(ALFABETO_TEMP) for _ in range(10))
    conn.execute("UPDATE usuarios SET senha_hash = ?, trocar_senha = 1, falhas = 0, bloqueado_ate = NULL WHERE id = ?",
                 (generate_password_hash(senha), id))
    conn.commit()
    return senha


def trocar_senha(conn, id, atual, nova, confirmacao) -> None:
    u = por_id(conn, id)
    if not u:
        raise ErroDeNegocio("Usuário não encontrado.")
    if not check_password_hash(u["senha_hash"], str(atual or "")):
        raise ErroDeNegocio("A senha atual não confere.")
    nova = _validar_senha(nova)
    if nova == str(atual):
        raise ErroDeNegocio("A nova senha precisa ser diferente da atual.")
    if nova != str(confirmacao or ""):
        raise ErroDeNegocio("A confirmação não confere com a nova senha.")
    conn.execute("UPDATE usuarios SET senha_hash = ?, trocar_senha = 0 WHERE id = ?", (generate_password_hash(nova), id))
    conn.commit()


def criar_admin(conn, login, nome, senha) -> int:
    """Primeiro administrador e socorro: se o login já existe, redefine a senha, volta a admin, reativa e
    desbloqueia (o nome não muda)."""
    u = por_login(conn, login)
    if not u:
        return criar(conn, login, nome, senha, "admin", trocar_senha=False)
    senha = _validar_senha(senha)
    conn.execute("""UPDATE usuarios SET senha_hash = ?, perfil = 'admin', ativo = 1, trocar_senha = 0, falhas = 0,
                    bloqueado_ate = NULL WHERE id = ?""", (generate_password_hash(senha), u["id"]))
    conn.commit()
    return u["id"]


def main(argv, ler_senha=None) -> int:
    """`python usuarios.py criar-admin <login> "<Nome>"`: pede a senha duas vezes no terminal."""
    import getpass
    import sys

    import db
    ler_senha = ler_senha or getpass.getpass
    if len(argv) != 3 or argv[0] != "criar-admin":
        print('Uso: python usuarios.py criar-admin <login> "<Nome completo>"', file=sys.stderr)
        return 2
    senha, repetida = ler_senha("Senha: "), ler_senha("Repita a senha: ")
    if senha != repetida:
        print("As senhas não conferem.", file=sys.stderr)
        return 1
    db.inicializar()
    conn = db.conectar()
    try:
        uid = criar_admin(conn, argv[1], argv[2], senha)
    except ErroDeNegocio as e:
        print(str(e), file=sys.stderr)
        return 1
    finally:
        conn.close()
    print(f"Administrador '{argv[1].lower()}' pronto (id {uid}).")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main(sys.argv[1:]))
