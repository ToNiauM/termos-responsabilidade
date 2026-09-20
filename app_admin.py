"""Tela Administração (só admin): aba Inventários aqui; a aba Usuários é a tela /usuarios de app_usuarios.py."""
from flask import Blueprint, g, render_template, request

import db
import inventario
import usuarios

admin_bp = Blueprint("admin", __name__)


def _conn():
    if "conn" not in g:
        g.conn = db.conectar()
    return g.conn


def _local() -> bool:
    return g.usuario["id"] is None


def _elegiveis(conn) -> list[dict]:
    """Quem pode compor a comissão: web = todos os ativos; desktop = administrador local + elegíveis por nome."""
    if _local():
        return [dict(g.usuario)] + usuarios.elegiveis_comissao(conn)
    return usuarios.ativos_para_comissao(conn)


@admin_bp.route("/administracao")
def tela():
    conn = _conn()
    eventos = [inventario.evento(conn, e["id"]) for e in inventario.eventos(conn)]
    return render_template("administracao.html", aba="inventarios", eventos=eventos,
                           ha_aberto=any(e["estado"] == "aberto" for e in eventos),
                           salas_ativas=db.localizacoes_ativas(conn), elegiveis=_elegiveis(conn), local=_local(),
                           confirmar=request.args.get("confirmar", type=int),
                           trilha=[("Administração", None)])
