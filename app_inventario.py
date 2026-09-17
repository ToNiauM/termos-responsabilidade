"""Rotas do módulo de inventário (blueprint /inventario). Dados em inventario.py; fotos em fotos.py.
A conexão por request e o errorhandler de ErroDeNegocio são os de app.py (g.conn e handler global)."""
import io

from flask import Blueprint, abort, flash, g, jsonify, redirect, render_template, request, send_file, session, url_for

import db
import fotos
import inventario
import painel_inventario

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
    return redirect(volta if (volta.startswith("/") and not volta.startswith("//") and not volta.startswith("/\\"))
                     else url_for("inventario.evento_tela", id=id))


def _json_erro(e, status=409):
    return jsonify({"erro": str(e)}), status


def _corpo_json() -> dict | None:
    """Corpo JSON das rotas de leitura: precisa ser um objeto; qualquer outra coisa é 400."""
    dados = request.get_json(silent=True)
    return dados if isinstance(dados, dict) else None


def _numero_lido(texto) -> int | None:
    t = str(texto or "").strip()
    return int(t) if t.isdigit() and t.strip("0") else None


def _bem_json(r):
    b = r["bem"]
    return {"situacao": r["situacao"], "numero": b["numero"], "descricao": b["descricao"], "complemento": b["complemento"],
            "situacao_bem": b["situacao"], "ativo": r["ativo"], "cadastrado_em": r["cadastrado_em"], "reler": r["reler"],
            "leitura_anterior": r["leitura_anterior"], "lido_em": r["lido_em"], "integrante": r["integrante"]}


_ORDEM_SITUACAO = {"pendente": 0, "localizado": 1, "divergente": 2}


@inventario_bp.route("/<int:id>/sala/<path:localizacao>")
def sala_tela(id, localizacao):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    salas = inventario.salas(conn, id)
    sala = next((s for s in salas if s["localizacao"] == localizacao), None)
    if not sala:
        abort(404)
    d = inventario.bens_da_sala(conn, id, localizacao)
    d["bens"].sort(key=lambda b: (_ORDEM_SITUACAO[b["situacao_inv"]], b["numero"]))
    return render_template("inventario_sala.html", e=e, sala=sala, localizacao=localizacao, integrante=session.get("integrante"),
                           conservacao=inventario.CONSERVACAO, fotos_ativas=fotos.configurado(), **d,
                           trilha=_trilha(e, (localizacao, None)))


@inventario_bp.route("/<int:id>/sala/<path:localizacao>/ler", methods=["POST"])
def ler(id, localizacao):
    conn = _conn()
    dados = _corpo_json()
    if dados is None:
        return jsonify({"erro": "Envie um objeto JSON."}), 400
    numero = _numero_lido(dados.get("numero"))
    if numero is None:
        return jsonify({"erro": "Número inválido.", "numero": None}), 404
    try:
        r = inventario.ler(conn, id, localizacao, numero, session.get("integrante") or "")
    except inventario.BemNaoEncontrado as e:
        return jsonify({"erro": str(e), "numero": e.numero}), 404
    except db.ErroDeNegocio as e:
        return _json_erro(e)
    return jsonify(_bem_json(r))


@inventario_bp.route("/<int:id>/leitura/<int:numero>", methods=["POST"])
def atualizar_leitura(id, numero):
    dados = _corpo_json()
    if dados is None:
        return jsonify({"erro": "Envie um objeto JSON."}), 400
    try:
        inventario.atualizar_leitura(_conn(), id, numero, **{k: v for k, v in dados.items() if k in ("conservacao", "quem_usa", "observacao")})
    except db.ErroDeNegocio as e:
        return _json_erro(e)
    return jsonify({"ok": True})


def _foto_processada():
    """Valida e comprime a foto enviada em request.files['foto']; ErroDeNegocio se faltar ou fotos desativadas."""
    if not fotos.configurado():
        raise db.ErroDeNegocio("Fotos desativadas: bucket não configurado.")
    arquivo = request.files.get("foto")
    if not arquivo or not arquivo.filename:
        raise db.ErroDeNegocio("Envie a foto.")
    return fotos.comprimir(fotos.validar(arquivo))


@inventario_bp.route("/<int:id>/leitura/<int:numero>/foto", methods=["POST"])
def foto_leitura(id, numero):
    conn = _conn()
    try:
        inventario._evento_aberto_ou_erro(conn, id)
        dados = _foto_processada()
        try:
            url = fotos.enviar(fotos.nome_bem(id, numero), dados)
        except Exception:
            return _json_erro("Falha ao enviar a foto.")
        inventario.atualizar_leitura(conn, id, numero, foto_url=url)
    except db.ErroDeNegocio as e:
        return _json_erro(e)
    return jsonify({"foto_url": url})


@inventario_bp.route("/<int:id>/leitura/<int:numero>/foto/excluir", methods=["POST"])
def foto_excluir(id, numero):
    conn = _conn()
    atual = conn.execute("SELECT foto_url, localizacao FROM inventario_leituras WHERE evento_id = ? AND numero = ?", (id, numero)).fetchone()
    if not atual:
        abort(404)
    inventario.atualizar_leitura(conn, id, numero, foto_url="")
    fotos.apagar(atual["foto_url"])
    flash("Foto removida.", "success")
    return redirect(url_for("inventario.sala_tela", id=id, localizacao=request.form.get("volta") or atual["localizacao"]))


@inventario_bp.route("/<int:id>/sala/<path:localizacao>/sobra", methods=["POST"])
def sobra(id, localizacao):
    conn = _conn()
    f = request.form
    integrante = session.get("integrante") or ""
    exigir = fotos.configurado()
    dados = None
    if exigir:
        if not (request.files.get("foto") and request.files["foto"].filename):
            raise db.ErroDeNegocio("A sobra precisa de foto.")
        dados = fotos.comprimir(fotos.validar(request.files["foto"]))
    sid = inventario.registrar_sobra(conn, id, localizacao, f.get("descricao", ""), f.get("complemento", ""),
                                     f.get("observacao", ""), "", integrante, exigir_foto=False)
    if exigir:
        try:
            url = fotos.enviar(fotos.nome_sobra(id, sid), dados)
        except Exception:
            inventario.excluir_sobra(conn, id, sid)
            raise db.ErroDeNegocio("Falha ao enviar a foto; sobra não registrada. Tente de novo.")
        inventario.definir_foto_sobra(conn, sid, url)
    flash("Sobra registrada.", "success")
    return redirect(url_for("inventario.sala_tela", id=id, localizacao=localizacao))


@inventario_bp.route("/<int:id>/sobra/<int:sobra_id>/excluir", methods=["POST"])
def sobra_excluir(id, sobra_id):
    s = inventario.excluir_sobra(_conn(), id, sobra_id)
    fotos.apagar(s["foto_url"])
    flash("Sobra excluída.", "success")
    return redirect(url_for("inventario.sala_tela", id=id, localizacao=s["localizacao"]))


@inventario_bp.route("/<int:id>/relatorio")
def relatorio_tela(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    loc, sit = request.args.get("localizacao") or None, request.args.get("situacao") or None
    return render_template("inventario_relatorio.html", e=e, linhas=inventario.relatorio(conn, id, loc, sit), localizacao=loc, situacao=sit,
                           salas=[s["localizacao"] for s in inventario.salas(conn, id)], rotulos=inventario.ROTULO_SITUACAO,
                           trilha=_trilha(e, ("Relatório", None)))


@inventario_bp.route("/<int:id>/xlsx")
def xlsx(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    loc = request.args.get("localizacao") or None
    arquivo = inventario.exportar_xlsx(conn, id, io.BytesIO(), loc)
    arquivo.seek(0)
    nome = f"inventario_{id}_{''.join(c if c.isalnum() else '_' for c in (loc or 'todas'))}.xlsx"
    return send_file(arquivo, as_attachment=True, download_name=nome)


@inventario_bp.route("/<int:id>/painel")
def painel_tela(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    andar = request.args.get("andar") or None
    dados = inventario.painel(conn, id, andar)
    return render_template("inventario_painel.html", e=e, r=dados["resumo"], andar=andar,
                           cards=painel_inventario.cards(dados, id, andar), trilha=_trilha(e, ("Painel", None)))
