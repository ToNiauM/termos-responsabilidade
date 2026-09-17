"""Fluxos de manutenção de cadastros, mantendo os endpoints públicos do aplicativo."""
import hashlib
import json
import re
from urllib.parse import parse_qs, urlsplit

from flask import abort, flash, redirect, render_template, request, url_for
from itsdangerous import BadSignature, URLSafeTimedSerializer

import db

AREAS = {
    "responsaveis": ("Centros de custo", "Centros, responsáveis e localizações vinculadas", "Novo centro de custo"),
    "localizacoes": ("Localizações", "Vincule as localizações da base aos centros de custo", "Vincular localização"),
    "pessoas": ("Pessoas", "Pessoas cadastradas e bens sob sua responsabilidade", "Nova pessoa"),
    "processos": ("Processos SEI", "Processos usados na emissão dos termos", "Novo processo"),
}
PARAMETROS = ("q", "centro", "situacao", "tipo", "ordem", "direcao", "pagina", "por_pagina", "nome")


def registrar_cadastros(app, obter_conn):
    def ancora(chave):
        return "registro-" + hashlib.sha256(str(chave).encode()).hexdigest()[:16]

    def retorno(aba):
        """Reconstrói somente destinos conhecidos; nunca redireciona para URL recebida."""
        bruto = request.values.get("retorno", "")
        u = urlsplit(bruto)
        if not u.scheme and not u.netloc and u.path == url_for("cadastros", aba=aba):
            valores = parse_qs(u.query)
            fragmento = u.fragment if re.fullmatch(r"registro-[a-f0-9]{16}", u.fragment) else None
            return url_for("cadastros", aba=aba, _anchor=fragmento,
                           **{k: valores[k][0] for k in PARAMETROS if k in valores})
        return url_for("cadastros", aba=aba)

    def contexto(aba, titulo=None, **kwargs):
        return dict(aba=aba, areas=AREAS, titulo=titulo or AREAS[aba][0], retorno=retorno(aba),
                    trilha=[("Cadastros", url_for("cadastros", aba=aba)), (titulo or AREAS[aba][0], None)],
                    **kwargs)

    def voltar(aba, chave=None, pessoa=False):
        u = urlsplit(retorno(aba))
        args = {k: v[0] for k, v in parse_qs(u.query).items() if k in PARAMETROS}
        if pessoa or (aba == "pessoas" and args.get("nome")):
            args["nome"] = chave
        if chave is not None:
            args["registro"] = chave
        return redirect(url_for("cadastros", aba=aba, **args, _anchor=ancora(chave) if chave else None))

    def formulario(aba, valores, erros=None, chave=None):
        conn = obter_conn()
        titulo = ("Editar centro de custo" if aba == "responsaveis" else "Editar pessoa") if chave else AREAS[aba][2]
        locais = [l["localizacao"] for l in db.localizacoes_mapeadas(conn) if l["ccustos"] == chave] if aba == "responsaveis" else []
        return render_template("cadastros/formulario.html", **contexto(aba, titulo), valores=valores,
                               erros=erros or {}, chave=chave, locais=locais,
                               centros=db.centros(conn), pendentes=db.localizacoes_sem_centro(conn),
                               tipos=list(db.ROTULO_TIPO.items()))

    def validar(aba, valores, chave=None):
        obrigatorios = {"responsaveis": [("ccustos", "Sigla"), ("responsavel", "Responsável")],
                       "pessoas": [("nome", "Nome")], "localizacoes": [("localizacao", "Localização"), ("ccustos", "Centro de custo")],
                       "processos": [("tipo", "Tipo de termo"), ("descricao", "Descrição"), ("numero_sei", "Número do processo SEI")]}
        erros = {k: f"{rotulo} é obrigatório." for k, rotulo in obrigatorios[aba] if not valores.get(k, "").strip()}
        conn = obter_conn()
        if aba == "responsaveis":
            sigla = " ".join(valores.get("ccustos", "").split()).upper()
            if sigla != chave and db.responsavel(conn, sigla):
                erros["ccustos"] = "Já existe um centro de custo com esta sigla."
        if aba == "pessoas":
            nome = " ".join(valores.get("nome", "").split()).upper()
            if nome != chave and nome in db.pessoas(conn):
                erros["nome"] = "Esta pessoa já está cadastrada. Busque o nome na lista para editá-la."
        if aba in ("responsaveis", "pessoas"):
            email = valores.get("email", "").strip()
            if email and not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email):
                erros["email"] = "Informe um e-mail válido."
        if aba == "localizacoes" and not db.responsavel(conn, valores.get("ccustos", "")):
            erros["ccustos"] = "Escolha um centro de custo cadastrado."
        if aba == "processos" and valores.get("tipo") not in db.TIPOS_TERMO:
            erros["tipo"] = "Escolha o tipo de termo."
        return erros

    def confirmar(aba, titulo, preparar, executar, mensagem, chave=None, pessoa=False):
        """Revisão assinada, com revalidação do estado sob trava de escrita antes de gravar."""
        conn = obter_conn()
        token = request.form.get("revisao")
        if token:
            conn.execute("BEGIN IMMEDIATE")
        try:
            descricao, linhas, estado = preparar(conn)
            campos = [(k, v) for k, v in request.form.items(multi=True) if k not in ("revisao", "confirmar", "csrf")]
            # O retorno também é assinado, mas reconstruído antes de qualquer redirecionamento.
            assinatura = hashlib.sha256(json.dumps([request.endpoint, campos, estado], sort_keys=True,
                                                     ensure_ascii=False).encode()).hexdigest()
            serializador = URLSafeTimedSerializer(app.secret_key, salt="cadastros-revisao")
            revisado = False
            if token:
                try:
                    revisado = serializador.loads(token, max_age=1800) == assinatura
                except BadSignature:
                    pass
            if revisado:
                executar(conn)
                conn.commit()
                flash(mensagem, "success")
                return voltar(aba, chave, pessoa)
            conn.rollback()
            erro = "Os dados mudaram ou a revisão expirou. Confira o resumo e confirme novamente." if token else None
            return render_template("cadastros/confirmar.html", **contexto(aba, titulo),
                                   descricao=descricao, linhas=linhas, campos=campos,
                                   revisao=serializador.dumps(assinatura), erro=erro,
                                   acao=request.path, bloqueado=False,
                                   botao=("Confirmar exclusão" if "exclusão" in titulo else
                                          "Confirmar remoção" if "remoção" in titulo else
                                          "Confirmar transferência" if "transferência" in titulo else
                                          "Confirmar atribuição" if "atribuição" in titulo else "Confirmar alteração"))
        except db.ErroDeNegocio as e:
            conn.rollback()
            return render_template("cadastros/confirmar.html", **contexto(aba, titulo),
                                   descricao=str(e), linhas=[], campos=[], bloqueado=True, erro=None)
        except Exception:
            conn.rollback()
            raise

    @app.route("/cadastros/<aba>")
    def cadastros(aba):
        if aba not in AREAS:
            abort(404)
        conn = obter_conn()
        filtros = {k: request.args.get(k, "") for k in PARAMETROS}
        lista = db.listar_cadastros(conn, aba, filtros)
        filtros.update({k: str(lista[k]) for k in ("ordem", "direcao", "pagina", "por_pagina")})
        origem = url_for("cadastros", aba=aba, **{k: v for k, v in filtros.items() if v})

        def link(endpoint, **kwargs):
            chave = kwargs.get("ccustos") or kwargs.get("nome") or kwargs.get("localizacao")
            return url_for(endpoint, retorno=origem + ("#" + ancora(chave) if chave else ""), **kwargs)

        def pagina_url(**kwargs):
            params = {k: v for k, v in filtros.items() if v}
            params.update(kwargs)
            return url_for("cadastros", aba=aba, **params)

        nome = filtros.get("nome") if aba == "pessoas" else None
        if nome and nome not in db.pessoas(conn):
            flash("Pessoa não encontrada.", "error")
            return redirect(url_for("cadastros", aba=aba))
        registro = request.args.get("registro")
        salvo = None
        if registro and registro not in [str(i["chave"]) for i in lista["itens"]]:
            if aba == "responsaveis":
                c = db.responsavel(conn, registro)
                if c:
                    salvo = (registro, link("responsaveis_editar", ccustos=registro))
            elif aba == "pessoas" and registro in db.pessoas(conn):
                salvo = (registro, pagina_url(nome=registro))
            elif aba == "localizacoes" and registro in {l["localizacao"] for l in db.localizacoes_mapeadas(conn)}:
                salvo = (registro, url_for("cadastros", aba=aba, q=registro))
        return render_template("cadastros.html", **contexto(aba), lista=lista, filtros=filtros, origem=origem,
                               link=link, pagina_url=pagina_url, ancora=ancora, salvo=salvo, registro=registro,
                               centros=db.centros(conn), pendentes=db.localizacoes_sem_centro(conn),
                               tipos=list(db.ROTULO_TIPO.items()), nome=nome, pessoa=db.pessoa(conn, nome) if nome else None,
                               bens_pessoa=db.bens_da_pessoa(conn, nome) if nome else [],
                               vigentes={tipo: db.processo_vigente(conn, tipo) for tipo in db.TIPOS_TERMO})

    @app.route("/cadastros/<aba>/novo")
    def cadastro_novo(aba):
        if aba not in AREAS:
            abort(404)
        valores = {"vigente": "1"} if aba == "processos" else {}
        if aba == "localizacoes":
            valores["localizacao"] = request.args.get("localizacao", "")
        return formulario(aba, valores)

    @app.route("/cadastros/responsaveis/incluir", methods=["POST"])
    def responsaveis_incluir():
        valores = request.form.to_dict()
        erros = validar("responsaveis", valores)
        if erros:
            return formulario("responsaveis", valores, erros)
        db.incluir_responsavel(obter_conn(), valores)
        sigla = " ".join(valores["ccustos"].split()).upper()
        flash(f"Centro de custo {sigla} cadastrado. Você já pode vincular suas localizações.", "success")
        return voltar("responsaveis", sigla)

    @app.route("/cadastros/responsaveis/<ccustos>/editar", methods=["GET", "POST"])
    def responsaveis_editar(ccustos):
        c = db.responsavel(obter_conn(), ccustos) or abort(404)
        if request.method == "GET":
            return formulario("responsaveis", c, chave=ccustos)
        valores = request.form.to_dict()
        erros = validar("responsaveis", valores, ccustos)
        if erros:
            return formulario("responsaveis", valores, erros, ccustos)
        sigla = db.salvar_centro(obter_conn(), ccustos, valores)
        flash(f"Centro de custo {sigla} atualizado. As localizações vinculadas foram mantidas.", "success")
        return voltar("responsaveis", sigla)

    @app.route("/cadastros/responsaveis/excluir", methods=["POST"])
    def responsaveis_excluir():
        sigla = request.form.get("ccustos", "")
        def preparar(conn):
            db.checar_exclusao_centro(conn, sigla)
            locais = [l for l in db.localizacoes_mapeadas(conn) if l["ccustos"] == sigla]
            return (f"Excluir {sigla} remove o vínculo de {len(locais)} localização(ões). As que tiverem bens ativos ficarão sem centro de custo. Os bens permanecem na base.",
                    [(l["localizacao"], sigla, "Sem centro de custo") for l in locais],
                    [db.responsavel(conn, sigla), locais])
        return confirmar("responsaveis", "Confirmar exclusão do centro", preparar,
                         lambda c: db.excluir_responsavel(c, sigla), f"Centro de custo {sigla} excluído.")

    @app.route("/cadastros/pessoas/incluir", methods=["POST"])
    def pessoas_incluir():
        valores = request.form.to_dict()
        erros = validar("pessoas", valores)
        if erros:
            return formulario("pessoas", valores, erros)
        nome = db.incluir_pessoa(obter_conn(), valores["nome"], valores.get("email"), valores.get("matricula"))
        flash(f"{nome} cadastrada. Consulte um patrimônio para atribuir bens a esta pessoa.", "success")
        return voltar("pessoas", nome, pessoa=True)

    @app.route("/cadastros/pessoas/<nome>/editar", methods=["GET", "POST"])
    def pessoas_editar(nome):
        if nome not in db.pessoas(obter_conn()):
            abort(404)
        if request.method == "GET":
            return formulario("pessoas", dict(db.pessoa(obter_conn(), nome)), chave=nome)
        valores = request.form.to_dict()
        erros = validar("pessoas", valores, nome)
        if erros:
            return formulario("pessoas", valores, erros, nome)
        novo = db.salvar_pessoa(obter_conn(), nome, valores)
        flash(f"Pessoa {novo} atualizada. Os bens atribuídos foram mantidos.", "success")
        return voltar("pessoas", novo, pessoa=not bool(request.form.get("retorno")))

    @app.route("/cadastros/pessoas/excluir", methods=["POST"])
    def pessoas_excluir():
        nome = request.form.get("nome", "")
        def preparar(conn):
            if nome not in db.pessoas(conn):
                raise db.ErroDeNegocio("Pessoa não encontrada.")
            bens = db.bens_da_pessoa(conn, nome)
            return (f"Excluir {nome} remove {len(bens)} atribuição(ões). Os bens permanecem na base e passam a seguir o centro definido pela localização, quando houver.",
                    [(f"{b['numero']} — {b['descricao']}", nome, "Responsabilidade pela localização") for b in bens], bens)
        return confirmar("pessoas", "Confirmar exclusão da pessoa", preparar,
                         lambda c: db.excluir_pessoa(c, nome), "Pessoa excluída.")

    @app.route("/cadastros/pessoas/atribuir", methods=["POST"])
    def pessoas_atribuir():
        nome, numero = request.form.get("nome", ""), request.form.get("numero", "").strip()
        if nome not in db.pessoas(obter_conn()):
            abort(404)
        erro = None
        if not numero.isdigit():
            erro = "Digite o número do patrimônio."
        elif not db.buscar_bem(obter_conn(), int(numero)):
            erro = f"Bem {numero} não encontrado na base."
        if erro:
            return render_template("cadastros/atribuir.html", **contexto("pessoas", "Consultar patrimônio"),
                                   nome=nome, valores={"numero": numero}, erros={"numero": erro})
        def preparar(conn):
            if nome not in db.pessoas(conn):
                raise db.ErroDeNegocio("Pessoa não encontrada.")
            if not numero.isdigit():
                raise db.ErroDeNegocio("Digite o número do patrimônio.")
            bem = db.ficha_do_bem(conn, int(numero))
            if not bem:
                raise db.ErroDeNegocio(f"Bem {numero} não encontrado na base.")
            atual = bem["pessoa"] or bem["ccustos"] or "Sem responsável definido"
            return (f"Confira o patrimônio antes de confirmar a atribuição a {nome}. Localização: {bem['localizacao']}. Situação: {bem['situacao']}.",
                    [(f"{bem['numero']} — {bem['descricao']} — {bem['complemento'] or ''}", atual, nome)], bem)
        return confirmar("pessoas", "Confirmar atribuição do patrimônio", preparar,
                         lambda c: db.atribuir(c, nome, int(numero), confirmar=True),
                         f"Bem {numero} atribuído a {nome}.", nome, pessoa=True)

    @app.route("/cadastros/pessoas/desatribuir", methods=["POST"])
    def pessoas_desatribuir():
        nome, numero = request.form.get("nome", ""), request.form.get("numero", "")
        def preparar(conn):
            bem = db.ficha_do_bem(conn, int(numero)) if numero.isdigit() else None
            if not bem or bem["pessoa"] != nome:
                raise db.ErroDeNegocio("Este patrimônio não está mais atribuído à pessoa. Volte à lista e confira os vínculos.")
            return ("O bem permanece na base. Ao remover o vínculo, a responsabilidade passa a seguir sua localização, quando houver centro vinculado.",
                    [(f"{numero} — {bem['descricao']}", nome, bem["ccustos"] or "Sem centro de custo")], bem)
        return confirmar("pessoas", "Confirmar remoção do vínculo", preparar,
                         lambda c: db.desatribuir(c, nome, int(numero)), "Atribuição removida.", nome, pessoa=True)

    @app.route("/cadastros/localizacoes/incluir", methods=["POST"])
    def localizacoes_incluir():
        valores = request.form.to_dict()
        erros = validar("localizacoes", valores)
        if not erros and valores["localizacao"] not in db.localizacoes_sem_centro(obter_conn()):
            erros["localizacao"] = "Escolha uma localização com bens ativos sem centro de custo."
        if erros:
            return formulario("localizacoes", valores, erros)
        db.incluir_localizacao(obter_conn(), valores["localizacao"], valores["ccustos"])
        flash("Localização vinculada ao centro de custo.", "success")
        return voltar("localizacoes", valores["localizacao"])

    @app.route("/cadastros/localizacoes/alterar")
    def localizacoes_alterar():
        localizacao = request.args.get("localizacao", "")
        atual = next((l for l in db.localizacoes_mapeadas(obter_conn()) if l["localizacao"] == localizacao), None)
        if not atual:
            abort(404)
        return render_template("cadastros/mover.html", **contexto("localizacoes", "Alterar centro da localização"),
                               locais=[atual], valores={}, erros={}, centros=db.centros(obter_conn()))

    @app.route("/cadastros/localizacoes/mover", methods=["POST"])
    def localizacoes_mover():
        locais = sorted(set(request.form.getlist("localizacoes")))
        destino = request.form.get("ccustos_destino", "")
        if locais and not db.responsavel(obter_conn(), destino):
            mapeadas = {l["localizacao"]: l for l in db.localizacoes_mapeadas(obter_conn())}
            return render_template("cadastros/mover.html", **contexto("localizacoes", "Alterar centro das localizações"),
                                   locais=[mapeadas.get(l, {"localizacao": l, "ccustos": "Sem vínculo"}) for l in locais],
                                   valores=request.form.to_dict(), centros=db.centros(obter_conn()),
                                   erros={"ccustos_destino": "Escolha um centro de custo de destino cadastrado."})
        def preparar(conn):
            if not locais:
                raise db.ErroDeNegocio("Selecione ao menos uma localização.")
            centro = db.responsavel(conn, destino)
            if not centro:
                raise db.ErroDeNegocio("Escolha um centro de custo de destino cadastrado.")
            mapeadas = {l["localizacao"]: l["ccustos"] for l in db.localizacoes_mapeadas(conn)}
            if any(l not in mapeadas for l in locais):
                raise db.ErroDeNegocio("Uma localização não está mais vinculada. Volte à lista e refaça a seleção.")
            linhas = [(l, mapeadas[l], destino) for l in locais]
            # Inclui os bens sob guarda: se a base mudar, exige uma nova revisão.
            bens = [dict(b) for b in conn.execute(
                "SELECT numero, situacao, localizacao FROM bens WHERE localizacao IN (" + ",".join("?" for _ in locais) + ") ORDER BY numero", locais)]
            return (f"Confira {len(locais)} localização(ões) antes de confirmar. Os bens continuam nas mesmas localizações; a responsabilidade do setor passa a seguir o centro de destino. Atribuições individuais são mantidas.",
                    linhas, [linhas, centro, bens])
        return confirmar("localizacoes", "Confirmar transferência de localizações", preparar,
                         lambda c: db.mover_localizacoes(c, locais, destino),
                         f"{len(locais)} localização(ões) movida(s) para {destino}.")

    @app.route("/cadastros/localizacoes/excluir", methods=["POST"])
    def localizacoes_excluir():
        loc = request.form.get("localizacao", "")
        def preparar(conn):
            atual = next((l for l in db.localizacoes_mapeadas(conn) if l["localizacao"] == loc), None)
            if not atual:
                raise db.ErroDeNegocio("O vínculo desta localização já foi removido.")
            return ("Remover este vínculo não exclui os bens nem altera sua localização na base importada. Se houver bens ativos, a localização ficará sem centro de custo.",
                    [(loc, atual["ccustos"], "Sem centro de custo")], atual)
        return confirmar("localizacoes", "Confirmar remoção do vínculo", preparar,
                         lambda c: db.excluir_localizacao(c, loc), "Vínculo da localização removido.")

    @app.route("/cadastros/processos/incluir", methods=["POST"])
    def processos_incluir():
        valores = request.form.to_dict()
        erros = validar("processos", valores)
        if erros:
            return formulario("processos", valores, erros)
        tipo = valores["tipo"]
        vigente = bool(valores.get("vigente"))
        def executar(conn):
            db.incluir_processo(conn, tipo, valores["descricao"], valores["numero_sei"], vigente=vigente)
        # A checagem e a inclusão compartilham a trava para não substituir um
        # processo criado por outra requisição sem a revisão correspondente.
        conn = obter_conn()
        conn.execute("BEGIN IMMEDIATE")
        atual = db.processo_vigente(conn, tipo)
        if vigente and (atual or request.form.get("revisao")):
            conn.rollback()
            def preparar(conn):
                p = db.processo_vigente(conn, tipo)
                return ("O novo processo será usado nas próximas emissões deste tipo. Os termos já registrados preservam seu processo original.",
                        [(db.ROTULO_TIPO[tipo], p["numero_sei"] if p else "Sem processo vigente", valores["numero_sei"])], p)
            return confirmar("processos", "Confirmar substituição do processo vigente", preparar, executar, "Processo incluído como vigente.")
        executar(conn)
        flash("Processo incluído.", "success")
        return voltar("processos")

    def processo_acao(acao):
        valor = request.form.get("id", "")
        if not valor.isdigit():
            abort(404)
        ident = int(valor)
        def preparar(conn):
            p = db._processo(conn, ident)
            termos = conn.execute("SELECT COUNT(*) FROM termos_emitidos WHERE processo_id=?", (ident,)).fetchone()[0]
            atual = db.processo_vigente(conn, p["tipo"])
            if acao == "excluir" and termos:
                raise db.ErroDeNegocio("Este processo tem termos registrados; encerre-o em vez de excluir.")
            if acao == "vigente":
                descricao = "Este processo será usado nas próximas emissões. O processo vigente anterior será encerrado; os termos já emitidos são preservados."
                linhas = [(db.ROTULO_TIPO[p["tipo"]], atual["numero_sei"] if atual else "Sem processo vigente", p["numero_sei"])]
            else:
                descricao = f"Processo {p['numero_sei']}: {p['descricao']}. "
                descricao += ("Este tipo de termo ficará sem processo vigente e sua emissão ficará indisponível até definir outro." if p["vigente"] else "O processo vigente deste tipo não será alterado.")
                linhas = [(db.ROTULO_TIPO[p["tipo"]], p["numero_sei"], "Excluído" if acao == "excluir" else "Encerrado")]
            return descricao, linhas, [p, atual, termos]
        titulos = {"vigente": "Confirmar processo vigente", "encerrar": "Confirmar encerramento do processo", "excluir": "Confirmar exclusão do processo"}
        funcoes = {"vigente": db.marcar_vigente, "encerrar": db.encerrar_processo, "excluir": db.excluir_processo}
        mensagens = {"vigente": "Processo marcado como vigente.", "encerrar": "Processo encerrado.", "excluir": "Processo excluído."}
        return confirmar("processos", titulos[acao], preparar, lambda c: funcoes[acao](c, ident), mensagens[acao])

    @app.route("/cadastros/processos/vigente", methods=["POST"])
    def processos_vigente():
        return processo_acao("vigente")

    @app.route("/cadastros/processos/encerrar", methods=["POST"])
    def processos_encerrar():
        return processo_acao("encerrar")

    @app.route("/cadastros/processos/excluir", methods=["POST"])
    def processos_excluir():
        return processo_acao("excluir")
