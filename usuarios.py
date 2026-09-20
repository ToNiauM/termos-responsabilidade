"""Usuários, senhas e permissões. Só dados e regras: toda função recebe `conn` primeiro e não importa Flask
(padrão de db.py). Hash de senha com werkzeug.security (scrypt), que já vem com o Flask."""
import re
import secrets
from datetime import datetime, timedelta

from werkzeug.security import check_password_hash, generate_password_hash

from db import ErroDeNegocio, _agora, _obrigatorio, _todos, _um
from permissoes import FUNCOES, ROTULOS, permitido

SENHA_MINIMA = 8
MAX_FALHAS = 5
BLOQUEIO_MINUTOS = 15
ALFABETO_TEMP = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789"
_LOGIN = re.compile(r"[a-z0-9._-]{2,30}")
_FORMATO = "%Y-%m-%d %H:%M:%S"

# Modo desktop (sem TERMOS_LOGIN): quem usa o programa Windows é o administrador da instalação.
USUARIO_LOCAL = {"id": None, "login": "local", "nome": "Administrador local", "funcoes": ("admin",), "ativo": 1, "trocar_senha": 0}

# Hash fictício calculado uma vez na importação: usado para gastar o mesmo tempo de um check_password_hash
# real quando o login não existe (ou está inativo), para não dar pista por tempo de resposta.
_HASH_FALSO = generate_password_hash("senha-falsa-para-tempo-constante")

_COLUNAS_LISTA = ("id, login, email, nome, ativo, trocar_senha, falhas, bloqueado_ate, criado_em, ultimo_acesso, "
                   "sei_login, sei_atualizado_em, spw_login, spw_atualizado_em")
_COLUNAS_CONTA = _COLUNAS_LISTA + ", senha_hash"
_MANTER = object()     # editar(): "não mexer no e-mail"


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


def _validar_email(email) -> str | None:
    """Vazio → None. Regra mínima: algo@algo, sem espaços, minúsculo."""
    v = str(email or "").strip().lower()
    if not v:
        return None
    if " " in v or v.count("@") != 1 or not all(v.split("@")):
        raise ErroDeNegocio("E-mail inválido.")
    return v


def _email_livre(conn, email, excluir_id=None) -> None:
    if email is None:
        return
    outro = por_email(conn, email)
    if outro and outro["id"] != excluir_id:
        raise ErroDeNegocio(f"Este e-mail já está em uso por {outro['login']}.")


def _validar_funcoes(funcoes):
    if isinstance(funcoes, str):
        raise ErroDeNegocio("Selecione as funções do usuário.")
    valores = set(funcoes or ())
    if not valores or not valores <= set(FUNCOES):
        raise ErroDeNegocio("Selecione ao menos uma função válida.")
    return tuple(f for f in FUNCOES if f in valores)


def _com_funcoes(conn, u):
    if u is None:
        return None
    atribuida = {r[0] for r in conn.execute(
        "SELECT funcao FROM usuarios_funcoes WHERE usuario_id=?", (u["id"],))}
    return dict(u, funcoes=tuple(f for f in FUNCOES if f in atribuida))


def _gravar_funcoes(conn, uid, funcoes):
    conn.execute("DELETE FROM usuarios_funcoes WHERE usuario_id=?", (uid,))
    conn.executemany("INSERT INTO usuarios_funcoes VALUES (?,?)", [(uid, f) for f in funcoes])


def por_id(conn, id) -> dict | None:
    return _com_funcoes(conn, _um(conn, f"SELECT {_COLUNAS_CONTA} FROM usuarios WHERE id=?", id)) if id is not None else None


def por_login(conn, login) -> dict | None:
    return _com_funcoes(conn, _um(conn, f"SELECT {_COLUNAS_CONTA} FROM usuarios WHERE login=?", str(login or "").strip().lower()))


def por_email(conn, email) -> dict | None:
    v = str(email or "").strip().lower()
    return _com_funcoes(conn, _um(conn, f"SELECT {_COLUNAS_CONTA} FROM usuarios WHERE email=?", v)) if v else None


def criar(conn, login, nome, senha, funcoes, trocar_senha=True, email=None) -> int:
    login, nome = _validar_login(login), _obrigatorio(nome, "Nome")
    senha, funcoes = _validar_senha(senha), _validar_funcoes(funcoes)
    email = _validar_email(email)
    _email_livre(conn, email)
    if por_login(conn, login):
        raise ErroDeNegocio(f"O login {login} já existe.")
    with conn:
        cur = conn.execute("""INSERT INTO usuarios
          (login,email,nome,senha_hash,trocar_senha,criado_em) VALUES (?,?,?,?,?,?)""",
          (login, email, nome, generate_password_hash(senha), int(bool(trocar_senha)), _agora()))
        _gravar_funcoes(conn, cur.lastrowid, funcoes)
    return cur.lastrowid


def listar(conn, busca="", funcao=None, inativos=False) -> list[dict]:
    """Sem senha_hash. busca casa em login, e-mail e nome (sem caixa); inativos=False esconde os inativos;
    funcao filtra por presença dessa função (um usuário pode ter mais de uma)."""
    sql, p = f"SELECT {_COLUNAS_LISTA} FROM usuarios WHERE 1=1", []
    if not inativos:
        sql += " AND ativo=1"
    if funcao:
        sql += " AND EXISTS(SELECT 1 FROM usuarios_funcoes f WHERE f.usuario_id=usuarios.id AND f.funcao=?)"
        p.append(funcao)
    if busca and busca.strip():
        sql += " AND (lower(login) LIKE ? OR lower(nome) LIKE ? OR lower(coalesce(email,'')) LIKE ?)"
        termo = f"%{busca.strip().lower()}%"
        p += [termo, termo, termo]
    return [_com_funcoes(conn, u) for u in _todos(conn, sql + " ORDER BY login", *p)]


def _admins_ativos(conn) -> int:
    return conn.execute("""SELECT count(*) FROM usuarios u WHERE u.ativo=1
      AND EXISTS(SELECT 1 FROM usuarios_funcoes f WHERE f.usuario_id=u.id AND f.funcao='admin')""").fetchone()[0]


def editar(conn, id, nome, funcoes, ativo, logado_id=None, email=_MANTER) -> None:
    """Nome, funções, ativo e e-mail (e-mail omitido = mantém; vazio = apaga). Login não muda. Travas: o próprio
    logado não se inativa nem perde a função admin; o último administrador ativo não é inativado nem rebaixado.

    Transação: exige conexão SEM transação pendente. `editar` abre sua própria `BEGIN IMMEDIATE` (reserva a
    escrita antes de contar admins, para duas requisições concorrentes não conseguirem remover o último admin
    ao mesmo tempo) e finaliza com commit/rollback próprios. O sqlite3 do Python abre uma transação implícita
    a partir do primeiro DML de uma conexão (isolation_level padrão); quem grava algo nessa mesma conexão antes
    de chamar `editar` (inclusive fixtures/rotas de teste que fazem `conn.execute('UPDATE ...')` fora deste
    módulo) precisa dar `conn.commit()` antes, senão o BEGIN IMMEDIATE falha com "cannot start a transaction
    within a transaction"."""
    funcoes = _validar_funcoes(funcoes)
    nome = _obrigatorio(nome, "Nome")
    conn.execute("BEGIN IMMEDIATE")
    try:
        u = por_id(conn, id)
        if not u:
            raise ErroDeNegocio("Usuário não encontrado.")
        perde_admin = "admin" in u["funcoes"] and u["ativo"] and ("admin" not in funcoes or not ativo)
        if perde_admin and logado_id is not None and int(logado_id) == int(id):
            raise ErroDeNegocio("Você não pode rebaixar nem inativar a própria conta.")
        if perde_admin and _admins_ativos(conn) <= 1:
            raise ErroDeNegocio("Este é o último administrador ativo: não pode ser rebaixado nem inativado.")
        email = u["email"] if email is _MANTER else _validar_email(email)
        _email_livre(conn, email, excluir_id=id)
        conn.execute("UPDATE usuarios SET nome=?,ativo=?,email=? WHERE id=?", (nome, int(bool(ativo)), email, id))
        _gravar_funcoes(conn, id, funcoes)
        conn.commit()
    except Exception:
        conn.rollback()
        raise


def elegiveis_comissao(conn) -> list[dict]:
    """Usuários ativos que podem compor a comissão de um inventário (têm função admin ou inventariante), por nome."""
    return sorted((u for u in listar(conn) if set(u["funcoes"]) & {"admin", "inventariante"}),
                  key=lambda u: (u["nome"], u["login"]))


# ---------------------------------------------------------------- permissões (nega por padrão)
# FUNCOES, ROTULOS e permitido vêm de permissoes.py (matriz por endpoint Flask).


# ---------------------------------------------------------------- autenticação e senhas
_INVALIDO = "Usuário ou senha inválidos."


def autenticar(conn, login, senha) -> dict:
    """Devolve o usuário. `login` pode ser o apelido ou o e-mail. Mensagem única para inexistente, inativo e senha
    errada. 5 falhas seguidas bloqueiam por 15 minutos (a senha certa também falha nesse período e não conta como falha)."""
    u = por_login(conn, login)
    if not u and "@" in str(login or ""):
        u = por_email(conn, login)
    if not u or not u["ativo"]:
        check_password_hash(_HASH_FALSO, str(senha or ""))    # tempo constante: não denuncia login inexistente/inativo
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


# ---------------------------------------------------------------- acesso ao SEI
def acesso_sei(conn, id) -> dict | None:
    """Login, unidade e data do acesso ao SEI de um usuário — nunca a senha."""
    r = _um(conn, "SELECT sei_login, sei_unidade, sei_atualizado_em FROM usuarios WHERE id=?", id)
    if not r or not (r["sei_login"] and r["sei_unidade"]):
        return None
    return {"login": r["sei_login"], "unidade": r["sei_unidade"], "atualizado_em": r["sei_atualizado_em"]}


def salvar_acesso_sei(conn, id, login, senha, unidade) -> None:
    """Senha vazia mantém a atual (erro se não há atual). Cifra com cofre.py; grava a data."""
    import cofre
    login = " ".join(str(login or "").split())
    unidade = " ".join(str(unidade or "").split()).upper()
    senha = str(senha or "")
    if not login:
        raise ErroDeNegocio("Informe o usuário do SEI.")
    if not unidade:
        raise ErroDeNegocio("Informe a sigla da unidade no SEI.")
    atual = _um(conn, "SELECT sei_senha FROM usuarios WHERE id=?", id)
    if atual is None:
        raise ErroDeNegocio("Usuário não encontrado.")
    if not senha and not atual["sei_senha"]:
        raise ErroDeNegocio("Informe a senha do SEI.")
    cifrada = cofre.cifrar(senha) if senha else atual["sei_senha"]
    with conn:
        conn.execute("UPDATE usuarios SET sei_login=?, sei_senha=?, sei_unidade=?, sei_atualizado_em=? WHERE id=?",
                     (login, cifrada, unidade, _agora(), id))


def apagar_acesso_sei(conn, id) -> None:
    with conn:
        conn.execute("UPDATE usuarios SET sei_login=NULL, sei_senha=NULL, sei_unidade=NULL, sei_atualizado_em=NULL WHERE id=?", (id,))


def credencial_sei(conn, login_sistema) -> dict | None:
    """Só para o trabalhador: credencial decifrada de quem pediu a emissão (None se não há acesso)."""
    import cofre
    r = _um(conn, "SELECT sei_login, sei_senha, sei_unidade FROM usuarios WHERE login=?", str(login_sistema or "").strip().lower())
    if not r or not (r["sei_login"] and r["sei_senha"] and r["sei_unidade"]):
        return None
    return {"SEI_USUARIO": r["sei_login"], "SEI_SENHA": cofre.decifrar(r["sei_senha"]), "SEI_UNIDADE": r["sei_unidade"]}


# ---------------------------------------------------------------- acesso ao SPW
def acesso_spw(conn, id) -> dict | None:
    """Login e data do acesso ao SPW de um usuário — nunca a senha."""
    r = _um(conn, "SELECT spw_login, spw_atualizado_em FROM usuarios WHERE id=?", id)
    if not r or not r["spw_login"]:
        return None
    return {"login": r["spw_login"], "atualizado_em": r["spw_atualizado_em"]}


def salvar_acesso_spw(conn, id, login, senha) -> None:
    """Senha vazia mantém a atual (erro se não há atual). Cifra com cofre.py; grava a data."""
    import cofre
    login = " ".join(str(login or "").split())
    senha = str(senha or "")
    if not login:
        raise ErroDeNegocio("Informe o usuário do SPW.")
    atual = _um(conn, "SELECT spw_senha FROM usuarios WHERE id=?", id)
    if atual is None:
        raise ErroDeNegocio("Usuário não encontrado.")
    if not senha and not atual["spw_senha"]:
        raise ErroDeNegocio("Informe a senha do SPW.")
    cifrada = cofre.cifrar(senha) if senha else atual["spw_senha"]
    with conn:
        conn.execute("UPDATE usuarios SET spw_login=?, spw_senha=?, spw_atualizado_em=? WHERE id=?",
                     (login, cifrada, _agora(), id))


def apagar_acesso_spw(conn, id) -> None:
    with conn:
        conn.execute("UPDATE usuarios SET spw_login=NULL, spw_senha=NULL, spw_atualizado_em=NULL WHERE id=?", (id,))


def credencial_spw(conn, login_sistema) -> dict | None:
    """Só para o trabalhador: credencial decifrada de quem pediu a atualização (None se não há acesso)."""
    import cofre
    r = _um(conn, "SELECT spw_login, spw_senha FROM usuarios WHERE login=?", str(login_sistema or "").strip().lower())
    if not r or not (r["spw_login"] and r["spw_senha"]):
        return None
    return {"SPW_USUARIO": r["spw_login"], "SPW_SENHA": cofre.decifrar(r["spw_senha"])}


def apagar_acessos(conn, id) -> None:
    """Apaga o acesso ao SEI e ao SPW de uma vez (botão "Apagar acessos" do admin)."""
    with conn:
        conn.execute("""UPDATE usuarios SET sei_login=NULL, sei_senha=NULL, sei_unidade=NULL, sei_atualizado_em=NULL,
                        spw_login=NULL, spw_senha=NULL, spw_atualizado_em=NULL WHERE id=?""", (id,))


def criar_admin(conn, login, nome, senha, email=None) -> int:
    """Primeiro administrador e socorro: se o login já existe, redefine a senha, garante a função admin
    (somada às que a conta já tinha, ex.: inventariante — não perde nenhuma), reativa e desbloqueia
    (o nome não muda; o e-mail só muda se informado)."""
    u = por_login(conn, login)
    if not u:
        return criar(conn, login, nome, senha, ["admin"], trocar_senha=False, email=email)
    senha = _validar_senha(senha)
    email = _validar_email(email) or u["email"]
    _email_livre(conn, email, excluir_id=u["id"])
    with conn:
        conn.execute("""UPDATE usuarios SET senha_hash = ?, ativo = 1, trocar_senha = 0, falhas = 0,
                        bloqueado_ate = NULL, email = ? WHERE id = ?""", (generate_password_hash(senha), email, u["id"]))
        _gravar_funcoes(conn, u["id"], set(u["funcoes"]) | {"admin"})
    return u["id"]


def main(argv, ler_senha=None) -> int:
    """`python usuarios.py criar-admin <login> "<Nome>" [e-mail]`: pede a senha duas vezes no terminal."""
    import getpass
    import sys

    import db
    ler_senha = ler_senha or getpass.getpass
    if len(argv) not in (3, 4) or argv[0] != "criar-admin":
        print('Uso: python usuarios.py criar-admin <login> "<Nome completo>" [e-mail]', file=sys.stderr)
        return 2
    senha, repetida = ler_senha("Senha: "), ler_senha("Repita a senha: ")
    if senha != repetida:
        print("As senhas não conferem.", file=sys.stderr)
        return 1
    db.inicializar()
    conn = db.conectar()
    try:
        uid = criar_admin(conn, argv[1], argv[2], senha, email=argv[3] if len(argv) == 4 else None)
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
