"""Rotas do módulo de inventário (blueprint /inventario). Dados em inventario.py; fotos em fotos.py.
A conexão por request e o errorhandler de ErroDeNegocio são os de app.py (g.conn e handler global)."""
import io

from flask import Blueprint, abort, flash, g, jsonify, redirect, render_template, request, send_file, url_for

import comissoes
import db
import fotos
import inventario
import painel_inventario
import usuarios

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


def _local() -> bool:
    """Modo desktop: o administrador local não tem linha em `usuarios`, então a comissão vai por nome."""
    return g.usuario["id"] is None


def _elegiveis(conn) -> list[dict]:
    """Usuários que podem compor a comissão. No desktop o administrador local entra sempre."""
    lista = usuarios.elegiveis_comissao(conn)
    if _local():
        lista = [dict(g.usuario)] + lista
    return lista


def _nomes_para_comissao(marcados: list) -> list[str]:
    """Só no modo desktop: nomes marcados + o administrador local, que sempre compõe a comissão."""
    nomes = list(marcados)
    if g.usuario["nome"] not in nomes:
        nomes.append(g.usuario["nome"])
    return nomes


def _na_comissao(conn, e) -> bool:
    return comissoes.pode_conferir(conn, g.usuario, e["id"])


def _exigir_comissao(conn, id):
    """Quem não está na comissão do evento não altera leituras, fotos nem sobras (mesma regra de ler).
    O vínculo é por identidade (comissoes), não pelo nome digitado na comissão."""
    e = _evento_ou_404(conn, id)
    if not comissoes.pode_conferir(conn, g.usuario, id):
        raise db.ErroDeNegocio(inventario.FORA_DA_COMISSAO)
    return e


def _pode(endpoint, metodo="GET") -> bool:
    return usuarios.permitido(g.usuario["funcoes"], endpoint, metodo)


@inventario_bp.route("")
def eventos_tela():
    """Só os eventos visíveis a quem pediu: o inventariante vê apenas aqueles de que participa."""
    conn = _conn()
    visiveis = comissoes.eventos_visiveis(conn, g.usuario)
    aberto = next((e for e in visiveis if not e["encerrado_em"]), None)
    pode_abrir = _pode("inventario.abrir", "POST")
    return render_template("inventario_eventos.html", aberto=inventario.evento(conn, aberto["id"]) if aberto else None,
                           eventos=[e for e in visiveis if e["encerrado_em"]],
                           salas_ativas=db.localizacoes_ativas(conn) if pode_abrir else [],
                           elegiveis=_elegiveis(conn) if pode_abrir else [], local=_local(),
                           pode_abrir=pode_abrir, pode_relatorios=_pode("inventario.relatorio_tela"),
                           trilha=_trilha())


@inventario_bp.route("/abrir", methods=["POST"])
def abrir():
    conn = _conn()
    f = request.form
    salas = None if f.get("escopo", "todas") == "todas" else f.getlist("salas")
    if _local():
        eid = inventario.abrir_evento(conn, f.get("nome", ""), f.get("descricao", ""), _nomes_para_comissao(f.getlist("integrantes")),
                                      salas, elegiveis=[u["nome"] for u in _elegiveis(conn)])
    else:
        eid = comissoes.abrir(conn, f.get("nome", ""), f.get("descricao", ""), f.getlist("usuarios"), salas)
    flash("Evento aberto. Comece pelas salas.", "success")
    return redirect(url_for("inventario.evento_tela", id=eid))


@inventario_bp.route("/<int:id>")
def evento_tela(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    return render_template("inventario_evento.html", e=e, salas=inventario.salas(conn, id),
                           na_comissao=_na_comissao(conn, e), confirmar=request.args.get("confirmar"), trilha=_trilha(e))


@inventario_bp.route("/<int:id>/encerrar", methods=["POST"])
def encerrar(id):
    conn = _conn()
    _evento_ou_404(conn, id)
    if not request.form.get("confirmar"):
        return redirect(url_for("inventario.evento_tela", id=id, confirmar="encerrar"))
    inventario.encerrar_evento(conn, id)
    flash("Evento encerrado. As leituras ficam congeladas; relatório e planilha continuam disponíveis.", "success")
    return redirect(url_for("inventario.evento_tela", id=id))


@inventario_bp.route("/<int:id>/comissao", methods=["GET", "POST"])
def comissao(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    inventario._evento_aberto_ou_erro(conn, id)
    if request.method == "POST":
        if _local():
            # Desktop: sem contas para vincular; nomes já na comissão continuam aceitos, nome novo só se for elegível.
            inventario.editar_comissao(conn, id, _nomes_para_comissao(request.form.getlist("integrantes")),
                                       elegiveis=[u["nome"] for u in _elegiveis(conn)] + e["integrantes"])
        else:
            comissoes.definir(conn, id, request.form.getlist("usuarios"))
        flash("Comissão atualizada.", "success")
        return redirect(url_for("inventario.evento_tela", id=id))
    com_leituras = {r[0] for r in conn.execute("SELECT DISTINCT integrante FROM inventario_leituras WHERE evento_id = ?", (id,))}
    ligados = comissoes.vinculos(conn, id)
    return render_template("inventario_comissao.html", e=e, elegiveis=_elegiveis(conn), com_leituras=com_leituras,
                           local=_local(), selecionados={v["usuario_id"] for v in ligados},
                           sem_vinculo=[n for n in e["integrantes"] if n not in {v["nome_na_comissao"] for v in ligados}],
                           trilha=_trilha(e, ("Comissão", None)))


def _apagar_fotos_do_evento(url):
    try:
        fotos.apagar(url)
    except Exception:
        raise db.ErroDeNegocio("Não foi possível apagar as fotos no bucket; o evento foi mantido. Tente de novo.")


@inventario_bp.route("/<int:id>/excluir", methods=["GET", "POST"])
def excluir(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    if request.method == "POST":
        try:
            inventario.excluir_evento(conn, id, request.form.get("nome", ""), apagar=_apagar_fotos_do_evento)
        except db.ErroDeNegocio as erro:
            flash(str(erro), "error")
            return redirect(url_for("inventario.excluir", id=id))
        flash(f"Evento {e['nome']} excluído.", "success")
        return redirect(url_for("inventario.eventos_tela"))
    return render_template("inventario_excluir.html", e=e, c=inventario.contagem_para_exclusao(conn, id), trilha=_trilha(e, ("Excluir", None)))


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
    return render_template("inventario_sala.html", e=e, sala=sala, localizacao=localizacao,
                           na_comissao=_na_comissao(conn, e), integrante=g.usuario["nome"],
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
        _exigir_comissao(conn, id)      # vínculo por ID: o homônimo de um integrante não lê pelo nome
        r = inventario.ler(conn, id, localizacao, numero, g.usuario["nome"])
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
    conn = _conn()
    try:
        _exigir_comissao(conn, id)
        inventario.atualizar_leitura(conn, id, numero, **{k: v for k, v in dados.items() if k in ("conservacao", "quem_usa", "observacao")})
    except db.ErroDeNegocio as e:
        return _json_erro(e)
    return jsonify({"ok": True})


@inventario_bp.route("/<int:id>/sala/<path:localizacao>/lote", methods=["POST"])
def lote(id, localizacao):
    """Marcar como localizados (leitura sem plaqueta) ou desmarcar (apaga a leitura) os bens selecionados."""
    conn = _conn()
    _exigir_comissao(conn, id)
    volta = redirect(url_for("inventario.sala_tela", id=id, localizacao=localizacao))
    numeros = [int(n) for n in request.form.getlist("numeros") if n.strip().isdecimal()]
    if not numeros:
        flash("Selecione ao menos um bem.", "warning")
        return volta
    if request.form.get("acao") == "desmarcar":
        urls, apagadas = inventario.desfazer_leituras(conn, id, numeros, apagar=_apagar_no_bucket)
        if apagadas:
            flash("Leitura(s) desfeita(s): os bens voltaram a pendentes.", "success")
        else:
            flash("Nenhuma leitura para desfazer.", "warning")
        return volta
    r = inventario.ler_lote(conn, id, localizacao, numeros, g.usuario["nome"])
    msg = f"{r['lidos']} bem(ns) marcado(s) como localizado(s)."
    if r["nao_encontrados"]:
        msg += " Não encontrado(s): " + ", ".join(str(n) for n in r["nao_encontrados"]) + "."
    flash(msg, "success" if r["lidos"] else "warning")
    return volta


def _apagar_no_bucket(url):
    """Apaga a foto no bucket ANTES de o registro sair do banco; falha vira ErroDeNegocio e nada muda."""
    try:
        fotos.apagar(url)
    except Exception:
        raise db.ErroDeNegocio("Não foi possível apagar a foto no bucket; nada foi alterado. Tente de novo.")


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
    """Mais uma foto do bem neste evento. Devolve a lista completa (nfoto, url) para a tela redesenhar a célula."""
    conn = _conn()
    try:
        inventario._evento_aberto_ou_erro(conn, id)
        _exigir_comissao(conn, id)
        dados = _foto_processada()

        def enviar(chave):
            try:
                return fotos.enviar(chave, dados)
            except Exception:
                raise db.ErroDeNegocio("Falha ao enviar a foto.")

        lista = inventario.adicionar_foto(conn, id, numero, enviar)
    except db.ErroDeNegocio as e:
        return _json_erro(e)
    return jsonify({"fotos": lista})


@inventario_bp.route("/<int:id>/leitura/<int:numero>/foto/<int:nfoto>/excluir", methods=["POST"])
def foto_excluir(id, numero, nfoto):
    conn = _conn()
    _exigir_comissao(conn, id)
    leitura = conn.execute("SELECT localizacao FROM inventario_leituras WHERE evento_id = ? AND numero = ?", (id, numero)).fetchone()
    if not leitura:
        abort(404)
    url = inventario.apagar_foto(conn, id, numero, nfoto, apagar=_apagar_no_bucket)
    if url is None:
        abort(404)
    flash("Foto removida.", "success")
    return redirect(url_for("inventario.sala_tela", id=id, localizacao=request.form.get("volta") or leitura["localizacao"]))


@inventario_bp.route("/<int:id>/sala/<path:localizacao>/sobra", methods=["POST"])
def sobra(id, localizacao):
    conn = _conn()
    _exigir_comissao(conn, id)          # vínculo por ID: o homônimo de um integrante não registra sobra pelo nome
    f = request.form
    integrante = g.usuario["nome"]
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
            url = fotos.enviar(fotos.chave_sobra(inventario.pasta_do_evento(conn, id), sid), dados)
        except Exception:
            inventario.excluir_sobra(conn, id, sid)
            raise db.ErroDeNegocio("Falha ao enviar a foto; sobra não registrada. Tente de novo.")
        inventario.definir_foto_sobra(conn, sid, url)
    flash("Sobra registrada.", "success")
    return redirect(url_for("inventario.sala_tela", id=id, localizacao=localizacao))


@inventario_bp.route("/<int:id>/sobra/<int:sobra_id>/excluir", methods=["POST"])
def sobra_excluir(id, sobra_id):
    conn = _conn()
    _exigir_comissao(conn, id)
    s = inventario.excluir_sobra(conn, id, sobra_id, apagar=_apagar_no_bucket)
    flash("Sobra excluída.", "success")
    return redirect(url_for("inventario.sala_tela", id=id, localizacao=s["localizacao"]))


def _filtros_relatorio() -> dict:
    return {k: (request.args.get(k) or None) for k in inventario.FILTROS_RELATORIO}


@inventario_bp.route("/<int:id>/relatorio")
def relatorio_tela(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    f = _filtros_relatorio()
    linhas = inventario.relatorio(conn, id, **f)
    return render_template("inventario_relatorio.html", e=e, linhas=linhas, f=f, ativos={k: v for k, v in f.items() if v},
                           descricao=inventario.descrever_filtros(f), n_fotos=inventario.contar_fotos(linhas),
                           salas=[s["localizacao"] for s in inventario.salas(conn, id)], rotulos=inventario.ROTULO_SITUACAO,
                           conservacao=inventario.CONSERVACAO, colunas_ordem=inventario.COLUNAS_ORDEM, trilha=_trilha(e, ("Relatório", None)))


@inventario_bp.route("/<int:id>/xlsx")
def xlsx(id):
    conn = _conn()
    _evento_ou_404(conn, id)
    f = _filtros_relatorio()
    arquivo = inventario.exportar_xlsx(conn, id, io.BytesIO(), fotos=request.args.get("fotos") == "1", **f)
    arquivo.seek(0)
    nome = f"inventario_{id}_{''.join(c if c.isalnum() else '_' for c in (f['localizacao'] or 'todas'))}.xlsx"
    return send_file(arquivo, as_attachment=True, download_name=nome)


@inventario_bp.route("/<int:id>/painel")
def painel_tela(id):
    conn = _conn()
    e = _evento_ou_404(conn, id)
    andar = request.args.get("andar") or None
    dados = inventario.painel(conn, id, andar)
    return render_template("inventario_painel.html", e=e, r=dados["resumo"], andar=andar,
                           cards=painel_inventario.cards(dados, id, andar), trilha=_trilha(e, ("Painel", None)))
