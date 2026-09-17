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
