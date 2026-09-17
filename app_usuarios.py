"""Rotas de conta e de usuários (blueprint `usuarios`). Regras em usuarios.py; g.usuario é resolvido em app.py."""
from urllib.parse import urlsplit

from flask import Blueprint, abort, flash, g, redirect, render_template, request, session, url_for

import config
import db
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
