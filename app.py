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


# ---------------------------------------------------------------- termo de devolução
@app.route("/termo_devolucao", methods=["GET", "POST"])
def termo_devolucao():
    conn = obter_conn()
    nome = request.form.get("nome") or session.get("nome_devolucao")
    selecionados = session.setdefault("bens_selecionados", [])
    if request.method == "POST":
        if nome:
            session["nome_devolucao"] = nome
        if request.form.get("limpar"):
            session["bens_selecionados"] = []
        elif request.form.get("gerar"):
            if not nome or not selecionados:
                flash("Escolha a pessoa e adicione ao menos um bem.", "error")
            else:
                return redirect(url_for("termo", tipo="devolucao", chave=nome))
        elif request.form.get("remover"):
            session["bens_selecionados"] = [n for n in selecionados if n != request.form["remover"]]
        else:
            numero = request.form.get("numero_bem", "").strip()
            if not numero.isdigit() or not db.buscar_bem(conn, int(numero)):
                flash(f"Bem {numero or '(vazio)'} não encontrado. Verifique o número digitado.", "error")
            elif numero not in selecionados:
                session["bens_selecionados"] = selecionados + [numero]
        session.modified = True
        return redirect(url_for("termo_devolucao"))
    bens = [b for b in (db.buscar_bem(conn, int(n)) for n in selecionados) if b]
    return render_template("termo_devolucao.html", nomes=db.pessoas(conn), nome=nome, bens=bens,
                           total=sum(b["valor_atual"] or 0 for b in bens), trilha=[("Termo de devolução", None)])


# ---------------------------------------------------------------- cadastros
ABAS = ("responsaveis", "localizacoes", "pessoas")


@app.route("/cadastros/<aba>")
def cadastros(aba):
    if aba not in ABAS:
        abort(404)
    conn = obter_conn()
    nome = request.args.get("nome") or None
    return render_template(
        "cadastros.html", aba=aba, trilha=[("Cadastros", None)],
        centros=db.centros(conn), mapeadas=db.localizacoes_mapeadas(conn), pendentes=db.localizacoes_sem_centro(conn),
        pessoas=db.pessoas(conn), nome=nome, bens_pessoa=db.bens_da_pessoa(conn, nome) if nome else [],
        confirmar=request.args.get("confirmar"))


def _volta(aba, **args):
    return redirect(url_for("cadastros", aba=aba, **args))


@app.route("/cadastros/responsaveis/incluir", methods=["POST"])
def responsaveis_incluir():
    db.incluir_responsavel(obter_conn(), request.form)
    flash("Responsável incluído.", "success")
    return _volta("responsaveis")


@app.route("/cadastros/responsaveis/excluir", methods=["POST"])
def responsaveis_excluir():
    db.excluir_responsavel(obter_conn(), request.form["ccustos"])
    flash("Centro de custo excluído.", "success")
    return _volta("responsaveis")


@app.route("/cadastros/responsaveis/renomear", methods=["POST"])
def responsaveis_renomear():
    db.renomear_centro(obter_conn(), request.form["antigo"], request.form["novo"])
    flash(f"{request.form['antigo']} renomeado para {request.form['novo'].upper()}; localizações atualizadas.", "success")
    return _volta("responsaveis")


@app.route("/cadastros/localizacoes/incluir", methods=["POST"])
def localizacoes_incluir():
    db.incluir_localizacao(obter_conn(), request.form["localizacao"], request.form["ccustos"])
    flash("Localização mapeada.", "success")
    return _volta("localizacoes")


@app.route("/cadastros/localizacoes/excluir", methods=["POST"])
def localizacoes_excluir():
    db.excluir_localizacao(obter_conn(), request.form["localizacao"])
    flash("Mapeamento removido.", "success")
    return _volta("localizacoes")


@app.route("/cadastros/pessoas/incluir", methods=["POST"])
def pessoas_incluir():
    nome = db.incluir_pessoa(obter_conn(), request.form["nome"])
    return _volta("pessoas", nome=nome)


@app.route("/cadastros/pessoas/excluir", methods=["POST"])
def pessoas_excluir():
    nome = request.form["nome"]
    if not request.form.get("confirmar"):
        flash(f"Excluir {nome} remove também os bens atribuídos a ela. Clique em confirmar para prosseguir.", "warning")
        return _volta("pessoas", nome=nome, confirmar="excluir")
    db.excluir_pessoa(obter_conn(), nome)
    flash("Pessoa excluída.", "success")
    return _volta("pessoas")


@app.route("/cadastros/pessoas/atribuir", methods=["POST"])
def pessoas_atribuir():
    nome, numero = request.form["nome"], request.form.get("numero", "").strip()
    if not numero.isdigit():
        raise db.ErroDeNegocio("Digite o número do bem.")
    try:
        db.atribuir(obter_conn(), nome, int(numero), confirmar=bool(request.form.get("confirmar")))
    except db.JaAtribuido as e:
        flash(f"O bem {numero} está com {e.pessoa}. Clique em confirmar para transferir a {nome}.", "warning")
        return _volta("pessoas", nome=nome, confirmar=numero)
    flash(f"Bem {numero} atribuído a {nome}.", "success")
    return _volta("pessoas", nome=nome)


@app.route("/cadastros/pessoas/desatribuir", methods=["POST"])
def pessoas_desatribuir():
    db.desatribuir(obter_conn(), request.form["nome"], int(request.form["numero"]))
    flash("Atribuição removida; o bem volta a responder pelo setor.", "success")
    return _volta("pessoas", nome=request.form["nome"])


if __name__ == "__main__":
    db.inicializar()
    app.run(host="127.0.0.1", port=5000, debug=True)
