"""Rotas de conta e de usuários (blueprint `usuarios`). Regras em usuarios.py; g.usuario é resolvido em app.py."""
import secrets
from urllib.parse import urlsplit

from flask import Blueprint, abort, flash, g, make_response, redirect, render_template, request, session, url_for

import config
import db
import inventario
import usuarios

usuarios_bp = Blueprint("usuarios", __name__)


def _conn():
    if "conn" not in g:
        g.conn = db.conectar()
    return g.conn


def _proximo_seguro(valor: str | None) -> str:
    """Só caminhos relativos do próprio site: '/x?y=1'. '//host', 'http://…' e '/\\host' caem em '/'."""
    v = (valor or "").strip()
    if not v.startswith("/") or v.startswith("//") or v.startswith("/\\"):
        return url_for("home")
    partes = urlsplit(v)
    if partes.scheme or partes.netloc:
        return url_for("home")
    if any(c < " " or c == "\x7f" for c in v):     # \r, \n etc: caem em '/' em vez de estourar o redirect
        return url_for("home")
    return v


@usuarios_bp.before_request
def _so_com_login_ligado():
    """No desktop (sem TERMOS_LOGIN) não há conta: estas telas não existem."""
    if not config.exigir_login():
        abort(404)


@usuarios_bp.route("/login", methods=["GET", "POST"])
def login():
    conn = _conn()
    if g.usuario and request.method == "GET":
        return redirect(url_for("home"))
    proximo = request.args.get("proximo")
    sem_usuarios = conn.execute("SELECT count(*) FROM usuarios").fetchone()[0] == 0
    if request.method == "POST" and not sem_usuarios:
        try:
            u = usuarios.autenticar(conn, request.form.get("login", ""), request.form.get("senha", ""))
        except db.ErroDeNegocio as e:
            return render_template("login.html", erro=str(e), login=request.form.get("login", ""), proximo=proximo, sem_usuarios=False)
        session.clear()
        session["usuario_id"] = u["id"]
        session["csrf"] = secrets.token_urlsafe(32)
        session.permanent = True
        return redirect(_proximo_seguro(proximo))
    return render_template("login.html", erro=None, login="", proximo=proximo, sem_usuarios=sem_usuarios)


@usuarios_bp.route("/sair", methods=["POST"])
def sair():
    session.clear()
    return redirect(url_for("usuarios.login"))


@usuarios_bp.route("/senha", methods=["GET", "POST"])
def senha():
    if request.method == "POST":
        try:
            usuarios.trocar_senha(_conn(), g.usuario["id"], request.form.get("atual"), request.form.get("nova"), request.form.get("confirmacao"))
        except db.ErroDeNegocio as e:
            return render_template("senha.html", erro=str(e), obrigatoria=bool(g.usuario["trocar_senha"]), trilha=[("Trocar senha", None)])
        flash("Senha alterada.", "success")
        return redirect(url_for("home"))
    return render_template("senha.html", erro=None, obrigatoria=bool(g.usuario["trocar_senha"]), trilha=[("Trocar senha", None)])


def _retorno() -> str:
    return _proximo_seguro(request.args.get("retorno")) if request.args.get("retorno") else url_for("usuarios.lista")


def _form_usuario(u, valores, erro, senha_temporaria=None):
    return render_template("usuarios/formulario.html", u=u, valores=valores, erro=erro, funcoes=usuarios.FUNCOES,
                           rotulos=usuarios.ROTULOS, retorno=_retorno(), senha_temporaria=senha_temporaria,
                           trilha=[("Usuários", url_for("usuarios.lista")), ((u["login"] if u else "Novo usuário"), None)])


@usuarios_bp.route("/usuarios")
def lista():
    q, funcao, inativos = request.args.get("q", ""), request.args.get("funcao") or None, request.args.get("inativos") == "1"
    return render_template("usuarios/lista.html", lista=usuarios.listar(_conn(), q, funcao, inativos), q=q, funcao=funcao,
                           inativos=inativos, funcoes=usuarios.FUNCOES, rotulos=usuarios.ROTULOS, trilha=[("Usuários", None)])


@usuarios_bp.route("/usuarios/novo")
def novo():
    return _form_usuario(None, {"funcoes": [], "trocar_senha": "1"}, None)


@usuarios_bp.route("/usuarios/incluir", methods=["POST"])
def incluir():
    f = request.form
    valores = {k: f.get(k, "") for k in ("login", "email", "nome", "trocar_senha")}
    valores["funcoes"] = f.getlist("funcoes")
    try:
        if f.get("senha", "") != f.get("confirmacao", ""):
            raise db.ErroDeNegocio("A confirmação não confere com a senha.")
        usuarios.criar(_conn(), f.get("login"), f.get("nome"), f.get("senha"), f.getlist("funcoes"),
                       trocar_senha=bool(f.get("trocar_senha")), email=f.get("email"))
    except db.ErroDeNegocio as e:
        return _form_usuario(None, valores, str(e))
    flash(f"Usuário {f.get('login', '').strip().lower()} criado.", "success")
    return redirect(_retorno())


@usuarios_bp.route("/usuarios/<int:id>/editar", methods=["GET", "POST"])
def editar(id):
    conn = _conn()
    u = usuarios.por_id(conn, id) or abort(404)
    if request.method == "POST":
        f = request.form
        valores = {"nome": f.get("nome", ""), "funcoes": f.getlist("funcoes"), "ativo": f.get("ativo"), "email": f.get("email", "")}
        try:
            usuarios.editar(conn, id, f.get("nome"), f.getlist("funcoes"), ativo=bool(f.get("ativo")), logado_id=g.usuario["id"],
                            email=f.get("email", ""))
        except db.ErroDeNegocio as e:
            return _form_usuario(u, valores, str(e))
        nome_novo = " ".join(str(f.get("nome") or "").split())
        if nome_novo != u["nome"]:
            # Nome mudou: mantém o usuário na comissão do evento aberto (leituras já feitas ficam com o nome antigo).
            inventario.renomear_integrante(conn, u["nome"], nome_novo)
        flash(f"Usuário {u['login']} salvo" + ("" if f.get("ativo") else " (inativo)") + ".", "success")
        return redirect(_retorno())
    return _form_usuario(u, {"nome": u["nome"], "funcoes": u["funcoes"], "ativo": "1" if u["ativo"] else None, "email": u["email"] or ""}, None)


@usuarios_bp.route("/usuarios/<int:id>/nova-senha", methods=["POST"])
def nova_senha(id):
    conn = _conn()
    u = usuarios.por_id(conn, id) or abort(404)
    senha = usuarios.nova_senha_temporaria(conn, id)
    resp = make_response(_form_usuario(u, {"nome": u["nome"], "funcoes": u["funcoes"], "ativo": "1" if u["ativo"] else None, "email": u["email"] or ""},
                                       None, senha_temporaria=senha))
    resp.headers["Cache-Control"] = "no-store"     # a senha em claro aparece uma vez; não pode ficar no cache/histórico
    return resp
