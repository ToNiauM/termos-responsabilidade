"""Termos de Responsabilidade — CFC. Rotas Flask; dados em db.py; documentos em termos_html.py e nos geradores."""
from pathlib import Path

from flask import Flask, abort, flash, g, redirect, render_template, request, send_file, session, url_for

import config
import db
import termos_html
from Script_Termo_Individual import criar_termo_responsabilidade
from Termo_de_Responsabilidade import gerar_planilha_centro, gerar_termo_centro
from termo_devolucao import gerar_termo_devolucao

app = Flask(__name__, template_folder=str(config.pasta_recursos() / "templates"),
            static_folder=str(config.pasta_recursos() / "static"))
app.secret_key = "termos-cfc-local"  # sessão só guarda seleção de bens; programa roda em 127.0.0.1

DSGOV = {"ORGAO": "Conselho Federal de Contabilidade", "SISTEMA": "Termos de Responsabilidade",
         "SUBTITULO": "Setor de Patrimônio"}


@app.context_processor
def contexto_dsgov():
    return {"DSGOV": DSGOV, "MENU": [
        ("Início", "fa-home", url_for("home")),
        ("Termo por centro de custo", "fa-building", url_for("centro_custos")),
        ("Termo individual", "fa-user-check", url_for("termos_individuais")),
        ("Termo de devolução", "fa-box-open", url_for("termo_devolucao")),
        ("Cadastros", "fa-address-book", url_for("cadastros", aba="responsaveis")),
        ("Atualizar base", "fa-upload", url_for("upload")),
    ]}


def obter_conn():
    if "conn" not in g:
        g.conn = db.conectar()
    return g.conn


@app.teardown_appcontext
def fechar_conn(_exc):
    conn = g.pop("conn", None)
    if conn is not None:
        conn.close()


@app.errorhandler(db.ErroDeNegocio)
def erro_de_negocio(e):
    flash(str(e), "error")
    return redirect(request.referrer or url_for("home"))


def _nome_arquivo(s: str) -> str:
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in s)


# ---------------------------------------------------------------- início e ficha do bem
@app.route("/")
def home():
    return render_template("index.html", trilha=[])


@app.route("/bem")
def bem():
    numero = request.args.get("numero", "").strip()
    ficha = db.ficha_do_bem(obter_conn(), int(numero)) if numero.isdigit() else None
    if not ficha:
        flash(f"Bem {numero or '(vazio)'} não encontrado.", "error")
        return redirect(url_for("home"))
    return render_template("bem.html", bem=ficha, trilha=[(f"Bem {numero}", None)])


# ---------------------------------------------------------------- termos
def _bens_do_termo(conn, tipo, chave):
    """Devolve (titulo, corpo_html, bens, extra) do termo pedido; 404 se não existir."""
    if tipo == "ccusto":
        resp = db.responsavel(conn, chave) or abort(404)
        bens = db.bens_do_centro(conn, chave)
        return f"Termo de Responsabilidade - {chave}", termos_html.corpo_ccusto(chave, resp, bens), bens, resp
    if tipo == "individual":
        if chave not in db.pessoas(conn):
            abort(404)
        bens = db.bens_da_pessoa(conn, chave)
        return f"Termo de Responsabilidade - {chave}", termos_html.corpo_individual(chave, bens), bens, None
    if tipo == "devolucao":
        numeros = session.get("bens_selecionados", [])
        bens = [b for b in (db.buscar_bem(conn, int(n)) for n in numeros) if b]
        return f"Termo de Devolução - {chave}", termos_html.corpo_devolucao(chave, bens), bens, None
    abort(404)


@app.route("/centro-custos")
def centro_custos():
    return render_template("centro_custos.html", centros=db.centros(obter_conn()),
                           trilha=[("Termo por centro de custo", None)])


@app.route("/gerar", methods=["POST"])
def gerar():
    return redirect(url_for("termo", tipo="ccusto", chave=request.form["ccusto"]))


@app.route("/termos-individuais")
def termos_individuais():
    return render_template("termos_individuais.html", nomes=db.pessoas(obter_conn()),
                           trilha=[("Termo individual", None)])


@app.route("/gerar-individual", methods=["POST"])
def gerar_individual():
    return redirect(url_for("termo", tipo="individual", chave=request.form["nome"]))


@app.route("/termo/<tipo>/<chave>")
def termo(tipo, chave):
    titulo, _, bens, _ = _bens_do_termo(obter_conn(), tipo, chave)
    return render_template("termo.html", tipo=tipo, chave=chave, titulo=titulo, quantidade=len(bens),
                           trilha=[(titulo, None)])


@app.route("/termo/<tipo>/<chave>/documento")
def termo_documento(tipo, chave):
    titulo, corpo, _, _ = _bens_do_termo(obter_conn(), tipo, chave)
    html = termos_html.documento(titulo, corpo)
    (config.pasta_saida() / f"{_nome_arquivo(titulo)}.html").write_text(html, encoding="utf8")
    return html


@app.route("/termo/<tipo>/<chave>/docx")
def termo_docx(tipo, chave):
    conn = obter_conn()
    _, _, bens, extra = _bens_do_termo(conn, tipo, chave)
    saida = config.pasta_saida()
    if tipo == "ccusto":
        destino = gerar_termo_centro(chave, extra, bens, saida / f"Termo_de_Responsabilidade_{_nome_arquivo(chave)}.docx")
    elif tipo == "individual":
        destino = criar_termo_responsabilidade(chave, bens, saida / f"Termo_{_nome_arquivo(chave)}.docx")
    else:
        destino = gerar_termo_devolucao(chave, bens, saida / f"Termo_Devolucao_{_nome_arquivo(chave)}.docx")
        if destino is None:
            flash("Nenhum bem selecionado.", "error")
            return redirect(url_for("termo_devolucao"))
    return send_file(destino, as_attachment=True, download_name=destino.name)


@app.route("/termo/ccusto/<chave>/planilha")
def termo_planilha(chave):
    _, _, bens, _ = _bens_do_termo(obter_conn(), "ccusto", chave)
    destino = gerar_planilha_centro(bens, config.pasta_saida() / f"planilha_{_nome_arquivo(chave)}.xlsx")
    return send_file(destino, as_attachment=True, download_name=destino.name)


# ---------------------------------------------------------------- atualizar base
@app.route("/upload", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        arquivo = request.files.get("arquivo")
        if not arquivo or not arquivo.filename.lower().endswith(".xlsx"):
            flash("Envie o export do sistema em .xlsx.", "error")
            return redirect(url_for("upload"))
        resumo = db.importar_bens(obter_conn(), arquivo.stream)
        flash(f"{resumo['total']} bens importados ({resumo['ativos']} ativos).", "success")
        return redirect(url_for("upload"))
    return render_template("upload.html", sem_centro=db.localizacoes_sem_centro(obter_conn()),
                           trilha=[("Atualizar base", None)])


# ---------------------------------------------------------------- rotas provisórias (Tasks 11 e 12)
@app.route("/termo_devolucao")
def termo_devolucao():
    return redirect(url_for("home"))


@app.route("/cadastros/<aba>")
def cadastros(aba):
    return redirect(url_for("home"))


if __name__ == "__main__":
    db.inicializar()
    app.run(host="127.0.0.1", port=5000, debug=True)
