"""Termos de Responsabilidade — CFC. Rotas Flask; dados em db.py; documentos em termos_html.py e nos geradores."""
import io
import re

from flask import Flask, abort, flash, g, redirect, render_template, request, send_file, session, url_for

import config
import db
import painel
import termos_html
import textos
from Script_Termo_Individual import criar_termo_responsabilidade
from Termo_de_Responsabilidade import gerar_planilha_centro, gerar_termo_centro
from termo_devolucao import gerar_termo_devolucao

app = Flask(__name__, template_folder=str(config.pasta_recursos() / "templates"),
            static_folder=str(config.pasta_recursos() / "static"))
app.secret_key = "termos-cfc-local"  # sessão só guarda seleção de bens; programa roda em 127.0.0.1

DSGOV_FIXO = {"SISTEMA": "Termos de Responsabilidade", "SUBTITULO": "Setor de Patrimônio"}


@app.context_processor
def contexto_dsgov():
    dsgov = dict(DSGOV_FIXO, ORGAO=textos.obter(obter_conn())["orgao_nome"])
    return {"DSGOV": dsgov, "MENU": [
        ("Início", "fa-home", url_for("home")),
        ("Termo por centro de custo", "fa-building", url_for("centro_custos")),
        ("Termo individual", "fa-user-check", url_for("termos_individuais")),
        ("Termo de devolução", "fa-box-open", url_for("termo_devolucao")),
        ("Termos emitidos", "fa-history", url_for("termos_emitidos_tela")),
        ("Recorte", "fa-filter", url_for("recorte")),
        ("Cadastros", "fa-address-book", url_for("cadastros", aba="responsaveis")),
        ("Textos", "fa-file-signature", url_for("textos_tela")),
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


def _baixar(arquivo: io.BytesIO, nome: str):
    """Envia um arquivo gerado em memória; nada é gravado em disco."""
    arquivo.seek(0)
    return send_file(arquivo, as_attachment=True, download_name=nome)


# ---------------------------------------------------------------- início e ficha do bem
@app.route("/")
def home():
    p = db.painel(obter_conn())
    f = {"situacao": "ATIVO"}
    return render_template("index.html", p=p, cards=painel.cards_graficos(p["dimensoes"], f), f=f,
                           moeda=painel.moeda, url_recorte=painel.url_recorte, trilha=[])


def _decimal(v: str) -> str:
    """'1.000,50' → '1000.50'; '1.000' → '1000' (ponto de milhar); '1000.5' → '1000.5' (ponto decimal)."""
    v = v.replace("R$", "").replace(" ", "")
    if "," in v:
        v = v.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+", v):
        v = v.replace(".", "")
    try:
        float(v)
    except ValueError:
        raise db.ErroDeNegocio(f"Valor inválido: {v}")
    return v


def _filtros_recorte() -> dict:
    """Filtros da query string. situacao ausente = ATIVO; situacao vazia (campo enviado em branco) = todas.
    Valores em R$ aceitam vírgula decimal e ponto de milhar."""
    f = {k: request.args.get(k, "").strip() for k in db.FILTROS}
    if "situacao" not in request.args:
        f["situacao"] = "ATIVO"
    for k in ("valor_de", "valor_ate"):
        if f[k]:
            f[k] = _decimal(f[k])
    return {k: v for k, v in f.items() if v}


@app.route("/recorte")
def recorte():
    conn = obter_conn()
    f = _filtros_recorte()
    r = db.recorte(conn, f)
    omitir = tuple(k for k in ("situacao", "ccusto", "classificacao", "localizacao", "idade", "ano", "faixa", "pessoa")
                   if f.get(k) and f.get(k) not in ("imoveis", "sem-imoveis"))
    nomes = {"idade": dict(db.FAIXAS_IDADE).get(f.get("idade")), "faixa": dict(db.FAIXAS_VALOR).get(f.get("faixa"))}
    termo_de = None
    if f.get("pessoa"):
        termo_de = ("individual", f["pessoa"], db.situacao_termo(conn, "individual", f["pessoa"], db.bens_da_pessoa(conn, f["pessoa"])))
    elif f.get("ccusto") and f["ccusto"] != "-" and db.responsavel(conn, f["ccusto"]):
        termo_de = ("ccusto", f["ccusto"], db.situacao_termo(conn, "ccusto", f["ccusto"], db.bens_do_centro(conn, f["ccusto"])))
    opcoes = {
        "situacoes": [r[0] for r in conn.execute("SELECT DISTINCT situacao FROM bens ORDER BY 1")],
        "classificacoes": [r[0] for r in conn.execute("SELECT DISTINCT classificacao FROM bens WHERE classificacao <> '' ORDER BY 1")],
        "localizacoes": [r[0] for r in conn.execute("SELECT DISTINCT localizacao FROM bens WHERE localizacao <> '' ORDER BY 1")],
        "centros": [c["ccustos"] for c in db.centros(conn)], "pessoas": db.pessoas(conn), "idades": db.FAIXAS_IDADE,
    }
    return render_template("recorte.html", f=f, r=r, cards=painel.cards_graficos(r["dimensoes"], f, omitir), termo_de=termo_de,
                           descricao=painel.descrever(f, nomes), opcoes=opcoes, moeda=painel.moeda,
                           trilha=[("Recorte", None)])


@app.route("/recorte/xlsx")
def recorte_xlsx():
    return _baixar(db.exportar_recorte(obter_conn(), _filtros_recorte(), io.BytesIO()), "recorte.xlsx")


@app.route("/bem")
def bem():
    numero = request.args.get("numero", "").strip()
    ficha = db.ficha_do_bem(obter_conn(), int(numero)) if numero.isdigit() else None
    if not ficha:
        flash(f"Bem {numero or '(vazio)'} não encontrado.", "error")
        return redirect(url_for("home"))
    return render_template("bem.html", bem=ficha, historico=db.historico_do_bem(obter_conn(), int(numero)), rotulos=db.ROTULO_TIPO,
                           trilha=[(f"Bem {numero}", None)])


@app.route("/pesquisa")
def pesquisa():
    """Busca rápida (cabeçalho). ?q= procura em centros, pessoas e bens; número existente abre a ficha.
    ?ccusto= ou ?pessoa= listam só os bens daquele centro/pessoa (mesma regra dos termos)."""
    conn, q = obter_conn(), request.args.get("q", "").strip()
    if request.args.get("ccusto"):
        chave = request.args["ccusto"]
        return render_template("pesquisa.html", q=chave, filtro="ccusto", chave=chave, bens=db.bens_do_centro(conn, chave),
                               trilha=[("Pesquisa", url_for("pesquisa", q=chave)), (f"Bens de {chave}", None)])
    if request.args.get("pessoa"):
        chave = request.args["pessoa"]
        return render_template("pesquisa.html", q=chave, filtro="individual", chave=chave, bens=db.bens_da_pessoa(conn, chave),
                               trilha=[("Pesquisa", url_for("pesquisa", q=chave)), (f"Bens de {chave}", None)])
    if q.isdigit() and db.buscar_bem(conn, int(q)):
        return redirect(url_for("bem", numero=q))
    r = db.pesquisar(conn, q) if q else {"centros": [], "pessoas": [], "bens": [], "truncado": False}
    return render_template("pesquisa.html", q=q, filtro=None, trilha=[("Pesquisa", None)], **r)


# ---------------------------------------------------------------- termos
def _bens_do_termo(conn, tipo, chave):
    """Devolve (titulo, corpo_html, bens, extra) do termo pedido; 404 se não existir."""
    t = textos.obter(conn)
    if tipo == "ccusto":
        resp = db.responsavel(conn, chave) or abort(404)
        bens = db.bens_do_centro(conn, chave)
        return f"Termo de Responsabilidade - {chave}", termos_html.corpo_ccusto(chave, resp, bens, textos=t), bens, resp
    if tipo == "individual":
        if chave not in db.pessoas(conn):
            abort(404)
        bens = db.bens_da_pessoa(conn, chave)
        return f"Termo de Responsabilidade - {chave}", termos_html.corpo_individual(chave, bens, textos=t), bens, None
    if tipo == "devolucao":
        numeros = session.get("bens_selecionados", [])
        bens = [b for b in (db.buscar_bem(conn, int(n)) for n in numeros) if b]
        return f"Termo de Devolução - {chave}", termos_html.corpo_devolucao(chave, bens, textos=t), bens, None
    abort(404)


@app.route("/centro-custos")
def centro_custos():
    conn = obter_conn()
    return render_template("centro_custos.html", centros=db.situacoes_centros(conn), trilha=[("Termo por centro de custo", None)])


@app.route("/gerar", methods=["POST"])
def gerar():
    return redirect(url_for("termo", tipo="ccusto", chave=request.form["ccusto"]))


@app.route("/termos-individuais")
def termos_individuais():
    conn = obter_conn()
    return render_template("termos_individuais.html", nomes=db.pessoas(conn), pessoas=db.situacoes_pessoas(conn),
                           trilha=[("Termo individual", None)])


@app.route("/gerar-individual", methods=["POST"])
def gerar_individual():
    return redirect(url_for("termo", tipo="individual", chave=request.form["nome"]))


def _exigir_processo(conn, tipo, chave):
    """Sem processo SEI vigente do tipo não há emissão: flash + volta à tela do termo."""
    if db.processo_vigente(conn, tipo):
        return None
    flash(f"Cadastre um processo SEI vigente para {db.ROTULO_TIPO[tipo]} em Cadastros → Processos SEI.", "error")
    return redirect(url_for("termo", tipo=tipo, chave=chave))


@app.route("/termo/<tipo>/<chave>")
def termo(tipo, chave):
    conn = obter_conn()
    titulo, _, bens, _ = _bens_do_termo(conn, tipo, chave)
    if tipo == "devolucao":
        situacao = {"estado": None, "ultimo": db.ultimo_termo(conn, tipo, chave), "entraram": 0, "sairam": 0}
    else:
        situacao = db.situacao_termo(conn, tipo, chave, bens)
    return render_template("termo.html", tipo=tipo, chave=chave, titulo=titulo, quantidade=len(bens),
                           processo=db.processo_vigente(conn, tipo), situacao=situacao,
                           rotulo_tipo=db.ROTULO_TIPO[tipo], trilha=[(titulo, None)])


@app.route("/termo/<tipo>/<chave>/documento")
def termo_documento(tipo, chave):
    titulo, corpo, _, _ = _bens_do_termo(obter_conn(), tipo, chave)
    return termos_html.documento(titulo, corpo)


@app.route("/termo/<tipo>/<chave>/docx")
def termo_docx(tipo, chave):
    conn = obter_conn()
    if (volta := _exigir_processo(conn, tipo, chave)):
        return volta
    _, _, bens, extra = _bens_do_termo(conn, tipo, chave)
    t = textos.obter(conn)
    arquivo = io.BytesIO()   # gerado em memória: nada fica gravado no servidor
    if tipo == "ccusto":
        gerar_termo_centro(chave, extra, bens, arquivo, textos=t)
        nome = f"Termo_de_Responsabilidade_{_nome_arquivo(chave)}.docx"
    elif tipo == "individual":
        criar_termo_responsabilidade(chave, bens, arquivo, textos=t)
        nome = f"Termo_{_nome_arquivo(chave)}.docx"
    else:
        if gerar_termo_devolucao(chave, bens, arquivo, textos=t) is None:
            flash("Nenhum bem selecionado.", "error")
            return redirect(url_for("termo_devolucao"))
        nome = f"Termo_Devolucao_{_nome_arquivo(chave)}.docx"
    db.registrar_emissao(conn, tipo, chave, bens)
    return _baixar(arquivo, nome)


@app.route("/termo/<tipo>/<chave>/registrar", methods=["POST"])
def termo_registrar(tipo, chave):
    """Chamado pelo botão Copiar depois da cópia dar certo. Responde JSON."""
    conn = obter_conn()
    _, _, bens, _ = _bens_do_termo(conn, tipo, chave)
    if not db.processo_vigente(conn, tipo):
        return {"erro": f"Cadastre um processo SEI vigente para {db.ROTULO_TIPO[tipo]}."}, 409
    t = db.registrar_emissao(conn, tipo, chave, bens)
    return {"id": t["id"], "emitido_em": t["emitido_em"]}


@app.route("/termo/ccusto/<chave>/planilha")
def termo_planilha(chave):
    _, _, bens, _ = _bens_do_termo(obter_conn(), "ccusto", chave)
    return _baixar(gerar_planilha_centro(bens, io.BytesIO()), f"planilha_{_nome_arquivo(chave)}.xlsx")


@app.route("/termos-emitidos/<int:id>")
def termo_emitido_tela(id):
    t = db.termo_emitido(obter_conn(), id) or abort(404)
    return render_template("termo_emitido.html", t=t, rotulos=db.ROTULO_TIPO, trilha=[("Termos emitidos", url_for("termos_emitidos_tela")), (f"Registro {id}", None)])


@app.route("/termos-emitidos")
def termos_emitidos_tela():
    tipo, chave = request.args.get("tipo") or None, request.args.get("chave", "").strip() or None
    return render_template("termos_emitidos.html", termos=db.termos_emitidos(obter_conn(), tipo, chave),
                           tipo=tipo, chave=chave, rotulos=db.ROTULO_TIPO, trilha=[("Termos emitidos", None)])


@app.route("/termos-emitidos/<int:id>/documento", methods=["POST"])
def termo_emitido_documento(id):
    conn = obter_conn()
    db.termo_emitido(conn, id) or abort(404)
    db.salvar_documento_sei(conn, id, request.form.get("documento_sei", ""))
    flash("Documento SEI salvo.", "success")
    return redirect(url_for("termo_emitido_tela", id=id))


# ---------------------------------------------------------------- atualizar base
@app.route("/upload", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        arquivo = request.files.get("arquivo")
        if not arquivo or not arquivo.filename.lower().endswith(".xlsx"):
            flash("Envie o export do sistema em .xlsx.", "error")
            return redirect(url_for("upload"))
        resumo = db.importar_bens(obter_conn(), arquivo.stream, nome_arquivo=arquivo.filename)
        flash(f"{resumo['total']} bens importados ({resumo['ativos']} ativos): {resumo['novos']} novo(s), "
              f"{resumo['removidos']} removido(s), {resumo['movidos']} movido(s), {resumo['situacao']} com situação alterada.", "success")
        return redirect(url_for("upload"))
    return render_template("upload.html", sem_centro=db.localizacoes_sem_centro(obter_conn()),
                           importacoes=db.importacoes(obter_conn()), trilha=[("Atualizar base", None)])


@app.route("/bens/exportar")
def bens_exportar():
    return _baixar(db.exportar_bens(obter_conn(), io.BytesIO()), "bens.xlsx")


@app.route("/importacoes/<int:id>")
def importacao_tela(id):
    i = db.importacao(obter_conn(), id) or abort(404)
    return render_template("importacao.html", i=i, trilha=[("Atualizar base", url_for("upload")), (f"Importação de {i['importado_em'][:10]}", None)])


# ---------------------------------------------------------------- termo de devolução
@app.route("/termo_devolucao", methods=["GET", "POST"])
def termo_devolucao():
    conn = obter_conn()
    nome = request.form.get("nome") or session.get("nome_devolucao")
    selecionados = session.setdefault("bens_selecionados", [])
    if request.method == "POST":
        if nome and nome != session.get("nome_devolucao"):
            session["bens_selecionados"] = selecionados = []
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
            else:
                numero = str(int(numero))
                if numero not in selecionados:
                    session["bens_selecionados"] = selecionados + [numero]
        session.modified = True
        return redirect(url_for("termo_devolucao"))
    bens = [b for b in (db.buscar_bem(conn, int(n)) for n in selecionados) if b]
    return render_template("termo_devolucao.html", nomes=db.pessoas(conn), nome=nome, bens=bens,
                           total=sum(b["valor_atual"] or 0 for b in bens), trilha=[("Termo de devolução", None)])


# ---------------------------------------------------------------- textos do termo
@app.route("/textos")
def textos_tela():
    return render_template("textos.html", valores=textos.obter(obter_conn()), grupos=textos.GRUPOS,
                           textarea=textos.TEXTAREA, rotulos=textos.ROTULOS, marcadores=textos.MARCADORES,
                           padrao=textos.PADRAO, trilha=[("Textos", None)])


@app.route("/textos", methods=["POST"])
def textos_salvar():
    conn = obter_conn()
    chave = request.form.get("restaurar")
    if chave:
        textos.restaurar(conn, chave)
        flash(f"Padrão restaurado: {textos.ROTULOS.get(chave, chave)}.", "success")
        return redirect(url_for("textos_tela"))
    novos = {c: request.form.get(c, "").replace("\r\n", "\n") for c in textos.PADRAO}
    for c, v in novos.items():
        textos.validar(c, v)          # tudo validado antes de gravar qualquer coisa
    for c, v in novos.items():
        textos.salvar(conn, c, v)
    flash("Textos salvos.", "success")
    return redirect(url_for("textos_tela"))


# ---------------------------------------------------------------- cadastros
ABAS = ("responsaveis", "localizacoes", "pessoas", "processos")


@app.route("/cadastros/<aba>")
def cadastros(aba):
    if aba not in ABAS:
        abort(404)
    conn = obter_conn()
    nome = request.args.get("nome") or None
    if nome and nome not in db.pessoas(conn):
        flash("Pessoa não encontrada.", "error")
        nome = None
    return render_template(
        "cadastros.html", aba=aba, trilha=[("Cadastros", None)],
        centros=db.centros(conn), mapeadas=db.localizacoes_mapeadas(conn), pendentes=db.localizacoes_sem_centro(conn),
        pessoas=db.pessoas(conn), nome=nome, bens_pessoa=db.bens_da_pessoa(conn, nome) if nome else [],
        confirmar=request.args.get("confirmar"), excluir=request.args.get("excluir"),
        processos=db.processos(conn), tipos=list(db.ROTULO_TIPO.items()))


def _volta(aba, **args):
    return redirect(url_for("cadastros", aba=aba, **args))


@app.route("/cadastros/responsaveis/incluir", methods=["POST"])
def responsaveis_incluir():
    db.incluir_responsavel(obter_conn(), request.form)
    flash("Responsável incluído.", "success")
    return _volta("responsaveis")


@app.route("/cadastros/responsaveis/<ccustos>/editar", methods=["GET", "POST"])
def responsaveis_editar(ccustos):
    conn = obter_conn()
    c = db.responsavel(conn, ccustos) or abort(404)
    if request.method == "POST":
        db._obrigatorio(request.form.get("responsavel"), "Responsável")  # valida antes de renomear: nada fica pela metade
        nova = " ".join(request.form.get("ccustos", "").split()).upper()
        if nova and nova != ccustos:
            db.renomear_centro(conn, ccustos, nova)
            ccustos = nova
        db.atualizar_responsavel(conn, ccustos, request.form)
        flash(f"Centro de custo {ccustos} atualizado.", "success")
        return _volta("responsaveis")
    locais = [l["localizacao"] for l in db.localizacoes_mapeadas(conn) if l["ccustos"] == ccustos]
    return render_template("editar_responsavel.html", c=c, locais=locais,
                           trilha=[("Cadastros", url_for("cadastros", aba="responsaveis")), (f"Editar {ccustos}", None)])


@app.route("/cadastros/responsaveis/excluir", methods=["POST"])
def responsaveis_excluir():
    conn, sigla = obter_conn(), request.form["ccustos"]
    db.checar_exclusao_centro(conn, sigla)   # bloqueia já aqui se houver bens sob guarda
    if not request.form.get("confirmar"):
        n = sum(1 for l in db.localizacoes_mapeadas(conn) if l["ccustos"] == sigla)
        flash(f"Excluir {sigla} remove também o mapeamento de {n} localização(ões), que voltam a pendentes. "
              "Clique em confirmar para prosseguir.", "warning")
        return _volta("responsaveis", excluir=sigla)
    db.excluir_responsavel(conn, sigla)
    flash(f"Centro de custo {sigla} excluído.", "success")
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


@app.route("/cadastros/localizacoes/mover", methods=["POST"])
def localizacoes_mover():
    destino = request.form.get("ccustos_destino", "")
    n = db.mover_localizacoes(obter_conn(), request.form.getlist("localizacoes"), destino)
    flash(f"{n} localização(ões) movida(s) para {destino}.", "success")
    return _volta("localizacoes")


@app.route("/cadastros/pessoas/<nome>/editar", methods=["GET", "POST"])
def pessoas_editar(nome):
    conn = obter_conn()
    if nome not in db.pessoas(conn):
        abort(404)
    if request.method == "POST":
        novo = db.renomear_pessoa(conn, nome, request.form.get("nome", ""))
        flash(f"Pessoa renomeada para {novo}.", "success")
        return _volta("pessoas", nome=novo)
    return render_template("editar_pessoa.html", nome=nome,
                           trilha=[("Cadastros", url_for("cadastros", aba="pessoas")), (f"Editar {nome}", None)])


@app.route("/cadastros/pessoas/incluir", methods=["POST"])
def pessoas_incluir():
    nome = db.incluir_pessoa(obter_conn(), request.form["nome"])
    flash(f"{nome} incluída.", "success")
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
        db.atribuir(obter_conn(), nome, int(numero), confirmar=(request.form.get("confirmar") == numero))
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


@app.route("/cadastros/processos/incluir", methods=["POST"])
def processos_incluir():
    db.incluir_processo(obter_conn(), request.form.get("tipo", ""), request.form.get("descricao", ""),
                        request.form.get("numero_sei", ""), vigente=bool(request.form.get("vigente")))
    flash("Processo incluído.", "success")
    return _volta("processos")


@app.route("/cadastros/processos/vigente", methods=["POST"])
def processos_vigente():
    db.marcar_vigente(obter_conn(), int(request.form["id"]))
    flash("Processo marcado como vigente.", "success")
    return _volta("processos")


@app.route("/cadastros/processos/encerrar", methods=["POST"])
def processos_encerrar():
    db.encerrar_processo(obter_conn(), int(request.form["id"]))
    flash("Processo encerrado; nenhum termo desse tipo será emitido até marcar outro como vigente.", "success")
    return _volta("processos")


@app.route("/cadastros/processos/excluir", methods=["POST"])
def processos_excluir():
    db.excluir_processo(obter_conn(), int(request.form["id"]))
    flash("Processo excluído.", "success")
    return _volta("processos")


@app.route("/cadastros/exportar")
def cadastros_exportar():
    return _baixar(db.exportar_cadastros(obter_conn(), io.BytesIO()), "cadastros.xlsx")


@app.route("/importar-cadastros", methods=["POST"])
def importar_cadastros():
    arquivo = request.files.get("arquivo")
    if not arquivo or not arquivo.filename.lower().endswith(".xlsx"):
        flash("Envie a planilha de cadastros em .xlsx.", "error")
        return redirect(url_for("upload"))
    r = db.importar_cadastros(obter_conn(), arquivo.stream)
    flash(f"Cadastros importados: {r['responsaveis']} centro(s) de custo, {r['localizacoes']} localização(ões), "
          f"{r['pessoas']} pessoa(s), {r['atribuicoes']} atribuição(ões).", "success")
    return redirect(url_for("upload"))


if __name__ == "__main__":
    db.inicializar()
    app.run(host="127.0.0.1", port=12345, debug=True)
