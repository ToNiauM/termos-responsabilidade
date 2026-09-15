"""Rotas do módulo de inventário (blueprint /inventario). Dados em inventario.py; fotos em fotos.py.
A conexão por request e o errorhandler de ErroDeNegocio são os de app.py (g.conn e handler global)."""
import io

from flask import Blueprint, abort, flash, g, jsonify, redirect, render_template, request, send_file, session, url_for

import db
import fotos
import inventario

inventario_bp = Blueprint("inventario", __name__, url_prefix="/inventario")


def _conn():
    if "conn" not in g:
        g.conn = db.conectar()
    return g.conn


def _evento_ou_404(conn, id):
    return inventario.evento(conn, id) or abort(404)


def _trilha(e=None, *resto):
    t = [("Inventário", url_for("inventario.eventos_tela"))]
    if e:
        t.append((e["nome"], url_for("inventario.evento_tela", id=e["id"])))
    t += list(resto)
    t[-1] = (t[-1][0], None)
    return t


@inventario_bp.route("")
def eventos_tela():
    conn = _conn()
    aberto = inventario.evento_aberto(conn)
    return render_template("inventario_eventos.html", aberto=inventario.evento(conn, aberto["id"]) if aberto else None,
                           eventos=[e for e in inventario.eventos(conn) if e["encerrado_em"]],
                           salas_ativas=db.localizacoes_ativas(conn), trilha=_trilha())


@inventario_bp.route("/abrir", methods=["POST"])
def abrir():
    f = request.form
    salas = None if f.get("escopo", "todas") == "todas" else f.getlist("salas")
    eid = inventario.abrir_evento(_conn(), f.get("nome", ""), f.get("descricao", ""), f.get("integrantes", "").splitlines(), salas)
    flash("Evento aberto. Escolha o integrante e comece pelas salas.", "success")
    return redirect(url_for("inventario.evento_tela", id=eid))


@inventario_bp.route("/<int:id>")
def evento_tela(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    return render_template("inventario_evento.html", e=e, salas=inventario.salas(conn, id),
                           integrante=session.get("integrante"), confirmar=request.args.get("confirmar"), trilha=_trilha(e))


@inventario_bp.route("/<int:id>/encerrar", methods=["POST"])
def encerrar(id):
    conn = _conn()
    _evento_ou_404(conn, id)
    if not request.form.get("confirmar"):
        return redirect(url_for("inventario.evento_tela", id=id, confirmar="encerrar"))
    inventario.encerrar_evento(conn, id)
    flash("Evento encerrado. As leituras ficam congeladas; relatório e planilha continuam disponíveis.", "success")
    return redirect(url_for("inventario.evento_tela", id=id))


@inventario_bp.route("/<int:id>/integrante", methods=["POST"])
def integrante(id):
    e = _evento_ou_404(_conn(), id)
    nome = request.form.get("integrante", "")
    if nome not in e["integrantes"]:
        raise db.ErroDeNegocio("Integrante não está na comissão deste evento.")
    session["integrante"] = nome
    volta = request.form.get("volta") or url_for("inventario.evento_tela", id=id)
    return redirect(volta if volta.startswith("/") else url_for("inventario.evento_tela", id=id))


@inventario_bp.route("/<int:id>/sala/<path:localizacao>")
def sala_tela(id, localizacao):
    return redirect(url_for("inventario.evento_tela", id=id))      # completada na Task 7


@inventario_bp.route("/<int:id>/relatorio")
def relatorio_tela(id):
    return redirect(url_for("inventario.evento_tela", id=id))      # completada na Task 8
