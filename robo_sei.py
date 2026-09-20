"""Emissão de termos no SEI: cria o documento no processo vigente, cola o termo e inclui no bloco
"Termos {UNIDADE}". Roda dentro do container `robo` do compose (alvo `robo` do Dockerfile, com
Playwright), chamado por atender_pedidos.py; a suíte usa um SEI falso.

A credencial (usuário, senha e unidade) vem do pedido — de quem clicou em Emitir, montada por
atender_pedidos.executar_sei com usuarios.credencial_sei — e não deste módulo; sei.env só guarda a
URL de login e o órgão (CHAVES_ENV), comuns a todo mundo.

Seletores provados em 2026-09-20 (docs/superpowers/notes/2026-09-20-spike-sei-escrita/README.md).
`playwright` é importado dentro de `abrir_sei`: o módulo importa sem ele.
"""
import re
import time
import urllib.parse
from contextlib import contextmanager

import config
import db
import segredos
import textos

BASE = "https://sei.cfc.org.br/sei/"
HOSTS = {"sei.cfc.org.br", "sip.cfc.org.br"}
ARQUIVO_ENV = segredos.PASTA / "sei.env"
CHAVES_ENV = ("SEI_LOGIN_URL", "SEI_ORGAO")
PASTA_SEI = "sei"                     # dentro de config.pasta_dados(): erro.png
TIMEOUT_S = 30
RE_ROTULO = re.compile(r"^(?P<rotulo>.*?)\s*\((?P<numero>\d{6,8})\)\s*$")
JS_ARVORE = """
() => {
  const out = [];
  for (const a of document.querySelectorAll("a[id^='anchor']")) {
    const m = /^anchor(\\d+)$/.exec(a.id);
    if (!m) continue;
    out.push({id: m[1], texto: (a.innerText || '').replace(/\\s+/g, ' ').trim()});
  }
  return {anchors: out, aguarde: document.querySelectorAll("a[id^='anchorAGUARDE']").length};
}
"""
JS_CORPO = """() => { const f = document.querySelector('iframe[title="Corpo do Texto"]');
  const box = f && f.closest('[id^="cke_txaEditor_"]'); return box ? box.id.replace('cke_','') : null; }"""


class RoboErro(Exception):
    """Erro previsto; a mensagem vai para robo_pedidos.mensagem e para a tela do termo."""


class TempoEsgotado(RoboErro):
    """O SEI não respondeu a tempo (seletor nunca apareceu). A mensagem interna cita seletores/frames
    só para quem depura o robô; o usuário só vê a versão genérica montada por `_mensagem_de`."""


def esperar(condicao, timeout_s=TIMEOUT_S, intervalo=0.3, erro="tempo esgotado"):
    fim = time.time() + timeout_s
    while time.time() < fim:
        r = condicao()
        if r:
            return r
        time.sleep(intervalo)
    raise TempoEsgotado(erro)


class SEI:
    """Uma sessão do SEI num contexto do Chromium. Os iframes do SEI têm URL pouco confiável: os frames
    são localizados pelo conteúdo (`_frame_com`) ou pelo id do elemento (`_frame`)."""

    def __init__(self, context, timeout_s: int = TIMEOUT_S):
        self.ctx, self.t = context, timeout_s
        self.avisos: list[str] = []                  # alert/confirm que o SEI mostrou (descartados, mas guardados p/ diagnóstico)
        context.on("page", self._vigiar_dialogos)
        self.p = context.new_page()
        self.p.set_default_timeout(timeout_s * 1000)
        self._saiu = False

    def _vigiar_dialogos(self, page) -> None:
        def dialogo(d):
            self.avisos.append(d.message.strip())
            d.dismiss()
        page.on("dialog", dialogo)

    # --- infraestrutura ---
    def anotar(self, texto: str, nome: str = "erro.txt") -> None:
        """Erro bruto (exceção, passo, avisos do SEI) ao lado de erro.png — só para quem depura o robô."""
        try:
            pasta = config.pasta_dados() / PASTA_SEI
            pasta.mkdir(parents=True, exist_ok=True)
            (pasta / nome).write_text(texto, encoding="utf-8")
        except OSError:
            pass

    def foto(self, nome: str = "erro.png", page=None) -> None:
        try:
            pasta = config.pasta_dados() / PASTA_SEI
            pasta.mkdir(parents=True, exist_ok=True)
            (page or self.p).screenshot(path=str(pasta / nome), full_page=True)
        except Exception:
            pass

    def _frame(self, nome):
        el = self.p.locator(f"#{nome}").first.element_handle(timeout=self.t * 1000)
        f = el.content_frame() if el else None
        if f is None:
            raise TempoEsgotado(f"frame {nome} ausente")
        return f

    def _frame_com(self, seletor, timeout_s=None):
        def achar():
            for f in self.p.frames:
                try:
                    if f.locator(seletor).count():
                        return f
                except Exception:
                    pass
        return esperar(achar, timeout_s or self.t, erro=f"tela com {seletor} não apareceu")

    def _arvore_bruta(self):
        try:
            r = self._frame("ifrArvore").evaluate(JS_ARVORE)
        except Exception:
            return None
        return r if r["anchors"] else None

    def _arvore(self):
        r = self._arvore_bruta()
        return r if r and not r["aguarde"] else None

    def _abrir_todas_pastas(self):
        """Processo com muitos documentos: o SEI agrupa em pastas fechadas (nó AGUARDE)."""
        f = self._frame("ifrArvore")
        botao = f.locator("img[title='Abrir todas as Pastas']")
        if not botao.count():
            return self._arvore()
        botao.first.click()
        contagens = []

        def estavel():
            r = self._arvore()
            if not r:
                return None
            contagens.append(len(r["anchors"]))
            return r if len(contagens) >= 2 and contagens[-1] == contagens[-2] else None
        return esperar(estavel, self.t, intervalo=0.7, erro="pastas da árvore não abriram")

    def arvore(self):
        return self._arvore() or self._abrir_todas_pastas()

    def _anchors(self):
        """Nós da árvore; lista vazia se a árvore ainda não carregou (ex.: logo após o SEI recarregar
        a ifrArvore, quando `arvore()` pode devolver None por uma fração de segundo)."""
        return (self.arvore() or {"anchors": []})["anchors"]

    # --- sessão ---
    def login(self, env: dict) -> dict:
        url = env.get("SEI_LOGIN_URL", BASE)
        partes = urllib.parse.urlsplit(url)
        if partes.scheme != "https" or partes.hostname not in HOSTS:
            raise RoboErro("SEI_LOGIN_URL inesperada em secrets/sei.env")
        self.p.goto(url, wait_until="domcontentloaded")
        vistos = len(self.avisos)
        self.p.select_option("#selOrgao", label=env.get("SEI_ORGAO", "CFC"))
        self.p.fill("#txtUsuario", env["SEI_USUARIO"])
        self.p.fill("#pwdSenha", env["SEI_SENHA"])
        self.p.click("#sbmAcessar")
        logado = False
        fim = time.time() + self.t
        while time.time() < fim:
            if any("inválid" in a.casefold() for a in self.avisos[vistos:]):
                break                                # usuário/senha recusados: o diálogo já respondeu, não vale esperar o timeout inteiro
            try:
                if self.p.locator("#txtPesquisaRapida").count():
                    logado = True
                    break
            except Exception:
                pass
            time.sleep(0.2)
        if not logado:
            return {"autenticado": False, "unidade": ""}
        unidade = ""
        try:
            unidade = self.p.locator("#lnkInfraUnidade").first.inner_text(timeout=3000).strip()
        except Exception:
            pass
        return {"autenticado": urllib.parse.urlsplit(self.p.url).hostname in HOSTS, "unidade": unidade}

    def trocar_unidade(self, sigla: str) -> str:
        """Troca a unidade corrente (mostrada em #lnkInfraUnidade) para `sigla`, ex.: GESERV."""
        oc = self.p.locator("#lnkInfraUnidade").first.get_attribute("onclick") or ""
        m = re.search(r"href='([^']+)'", oc)
        if m:
            self.p.goto(urllib.parse.urljoin(BASE, m.group(1)), wait_until="domcontentloaded")
        linha = self.p.locator("tr").filter(has=self.p.locator("td", has_text=re.compile(rf"^{re.escape(sigla)}$")))
        if not linha.count():
            raise RoboErro(f"Unidade {sigla} não está disponível para este usuário no SEI.")
        with self.p.expect_navigation(wait_until="domcontentloaded"):
            linha.first.locator("label.infraRadioLabel").first.click()
        atual = self.p.locator("#lnkInfraUnidade").first.inner_text(timeout=3000).strip()
        if atual != sigla:
            raise RoboErro(f"Unidade {sigla} não está disponível para este usuário no SEI.")
        return atual

    def logout(self) -> None:
        if self._saiu:
            return                                   # enviar_termo já saiu; abrir_sei chamaria de novo (5s de clique perdidos)
        self._saiu = True
        for pg in self.ctx.pages[1:]:
            try:
                pg.close()
            except Exception:
                pass
        try:
            self.p.click("#lnkInfraSairSistema", timeout=5000)
        except Exception:
            pass

    # --- processo e árvore ---
    def abrir_processo(self, numero: str) -> dict:
        campo = self.p.locator("#txtPesquisaRapida")
        campo.fill(numero)
        with self.p.expect_navigation(wait_until="domcontentloaded"):
            campo.press("Enter")
        if self.p.title().strip() != f"SEI - {numero}":
            raise RoboErro(f"Processo {numero} não abriu no SEI; nada foi criado.")
        esperar(self._arvore_bruta, self.t, erro="árvore do processo não carregou")
        return {"titulo_confere": True, "nos": len(self._anchors())}

    def documento_na_arvore(self, rotulo: str) -> str | None:
        for a in self._anchors():
            m = RE_ROTULO.match(a["texto"])
            if m and m.group("rotulo") == rotulo:
                return m.group("numero")
        return None

    def _selecionar_raiz(self):
        r = esperar(self.arvore, self.t, erro="árvore do processo não carregou")
        self._frame("ifrArvore").click(f"#anchor{r['anchors'][0]['id']}")
        return self._frame_com("img[title='Incluir Documento']")

    # --- escrita ---
    def incluir_documento(self, tipo_nome: str, nome_arvore: str, html: str, rotulo: str) -> str:
        fr = self._selecionar_raiz()
        time.sleep(0.8)
        for tentativa in (1, 2):                            # o clique se perde se o painel ainda estava recarregando (como no bloco)
            fr.locator("a:has(img[title='Incluir Documento'])").first.click()
            try:
                fr = self._frame_com("#ancExibirSeries", 10)
                break
            except RoboErro:
                if tentativa == 2:
                    raise
                fr = self._frame_com("img[title='Incluir Documento']")
        time.sleep(0.8)
        fr.click("#ancExibirSeries")                       # "Exibir todos os tipos" (a lista inicial é parcial)
        time.sleep(1.0)
        tipos = fr.locator("a[onclick^='escolher']").filter(has_text=re.compile(f"^{re.escape(tipo_nome)}$"))
        if not tipos.count():
            raise RoboErro(f"Tipo de documento '{tipo_nome}' não existe no SEI; corrija em Textos.")
        tipos.first.click()
        fr = self._frame_com("#txtNomeArvore")
        time.sleep(0.8)
        fr.fill("#txtNomeArvore", nome_arvore)
        gravado = fr.input_value("#txtNomeArvore")
        if gravado != nome_arvore:                          # o campo tem maxlength: nome de pessoa longo é cortado pelo SEI
            rotulo = f"{tipo_nome} {gravado}"
        fr.click("label[for=optPublico]")
        fr.wait_for_function("() => document.getElementById('optPublico').checked")
        with self.ctx.expect_page(timeout=self.t * 1000) as nova:
            fr.click("#btnSalvar")
        ed = nova.value
        ed.wait_for_load_state("domcontentloaded")
        ed.wait_for_function("() => typeof CKEDITOR !== 'undefined' && document.querySelector('iframe[title=\"Corpo do Texto\"]')")
        time.sleep(1.5)
        inst = ed.evaluate(JS_CORPO)
        if not inst:
            raise RoboErro("editor do SEI sem o Corpo do Texto")
        ed.evaluate(f"h => {{ const e = CKEDITOR.instances['{inst}']; e.setData(h); e.fire('change'); }}", html)
        time.sleep(0.5)
        ed.locator("a[title^='Salvar']:visible").first.click()
        time.sleep(2.5)
        ed.close()
        return esperar(lambda: self.documento_na_arvore(rotulo), self.t, intervalo=1.0,
                       erro="documento salvo não apareceu na árvore")

    def _linha_do_documento(self, fr, numero: str) -> str | None:
        return fr.evaluate("n => { const tr = [...document.querySelectorAll('tr')].find(tr => tr.innerText.includes(n));"
                           " return tr ? tr.innerText.replace(/\\s+/g,' ').trim() : null; }", numero)

    def incluir_em_bloco(self, numero: str, nome_bloco: str) -> str:
        no = None
        for a in self._anchors():
            m = RE_ROTULO.match(a["texto"])
            if m and m.group("numero") == numero:
                no = a["id"]
        if no is None:
            raise RoboErro(f"documento {numero} não está na árvore")
        self._frame("ifrArvore").click(f"#anchor{no}")
        self._frame_com("img[title='Incluir em Bloco de Assinatura']")
        for f in self.p.frames:                             # documento grande: esperar renderizar antes do clique
            if f.name == "ifrVisualizacao":
                try:
                    f.wait_for_load_state("load", timeout=60000)
                except Exception:
                    pass
        fr = None
        for tentativa in (1, 2):
            self._frame_com("img[title='Incluir em Bloco de Assinatura']").locator(
                "a:has(img[title='Incluir em Bloco de Assinatura'])").first.click()
            try:
                fr = self._frame_com("#selBloco", 20)
                break
            except RoboErro:
                if tentativa == 2:
                    raise
        time.sleep(0.8)
        opcoes = fr.evaluate("() => [...document.getElementById('selBloco').options].map(o => o.text)")
        alvo = [o for o in opcoes if o.endswith(" - " + nome_bloco)]
        if not alvo:
            raise RoboErro(f"Bloco '{nome_bloco}' não existe no SEI; crie o bloco e clique em Incluir no bloco.")
        numero_bloco = alvo[0].split(" - ")[0]
        linha = self._linha_do_documento(fr, numero)
        if linha and linha.endswith(" " + numero_bloco):
            return numero_bloco                             # já estava: não clica de novo
        fr.select_option("#selBloco", label=alvo[0])
        time.sleep(0.5)
        fr.click("#sbmIncluir")
        time.sleep(2)
        fr = self._frame_com("#selBloco")
        linha = self._linha_do_documento(fr, numero)
        if not (linha and linha.endswith(" " + numero_bloco)):
            raise RoboErro(f"o SEI não confirmou o documento {numero} no bloco {numero_bloco}")
        return numero_bloco


@contextmanager
def abrir_sei(env: dict):
    """Chromium headless com um SEI; fecha tudo ao sair, mesmo em erro."""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context(viewport={"width": 1400, "height": 900})
        sei = SEI(ctx)
        try:
            yield sei
        finally:
            try:
                sei.logout()
            finally:
                browser.close()


AO_PASSO = {"login": "entrar no SEI", "documento": "criar o documento", "bloco": "incluir no bloco de assinatura"}


def _mensagem_de(exc: Exception, passo: str, avisos=()) -> str:
    if isinstance(exc, TempoEsgotado) or type(exc).__name__ == "TimeoutError":
        m = f"O SEI não respondeu a tempo ao {AO_PASSO[passo]}."
        if avisos:                                      # o SEI mostrou um alert/confirm que o robô descartou: é a causa provável
            m += f" O SEI avisou: «{avisos[-1][:200]}»"
        return m
    return (str(exc) or type(exc).__name__)[:500]


def enviar_termo(conn, pedido: dict, abrir=None, env: dict | None = None) -> dict:
    """Passos login → documento → bloco, gravando cada um em robo_pedidos antes de começar. Nunca levanta.
    `abrir(env)` é um context manager que devolve um SEI (os testes injetam um falso)."""
    abrir = abrir or abrir_sei
    pid = pedido["id"]
    termo = db.termo_emitido(conn, pedido["termo_id"])
    passo, avisos = "login", []
    try:
        if termo is None:
            raise RoboErro("Termo emitido não existe mais.")
        if not env:
            raise RoboErro("Credencial do SEI não informada ao robô.")
        tipo_nome = textos.obter(conn)[f"sei_tipo_{termo['tipo']}"]
        nome_arvore = f"{termo['numero_termo']} - {termo['chave']}"      # "01/2026 - GELAI" (centro de custo) ou "01/2026 - NOME DA PESSOA"
        rotulo = f"{tipo_nome} {nome_arvore}"
        nome_bloco = f"Termos {termo['unidade_sei']}"
        db.marcar_passo(conn, pid, "login")
        with abrir(env) as sei:
            try:
                resultado_login = sei.login(env)
                if not resultado_login.get("autenticado"):
                    raise RoboErro("O SEI recusou seu usuário ou senha; atualize em Meus acessos.")
                unidade = env.get("SEI_UNIDADE", "").strip()
                if unidade and unidade != resultado_login.get("unidade"):
                    sei.trocar_unidade(unidade)
                passo = "documento"
                db.marcar_passo(conn, pid, "documento")
                sei.abrir_processo(termo["numero_sei"])
                numero = termo["documento_sei"]
                if not numero:
                    numero = sei.documento_na_arvore(rotulo) or sei.incluir_documento(tipo_nome, nome_arvore, pedido["html"] or "", rotulo)
                    db.salvar_documento_sei(conn, termo["id"], numero, termo["bloco_sei"] or "")
                passo = "bloco"
                db.marcar_passo(conn, pid, "bloco")
                bloco = sei.incluir_em_bloco(numero, nome_bloco)
                db.salvar_documento_sei(conn, termo["id"], numero, bloco)
            except Exception as exc:
                getattr(sei, "foto", lambda *_: None)("erro.png")
                avisos = getattr(sei, "avisos", [])
                getattr(sei, "anotar", lambda *_: None)(
                    f"{db._agora()} pedido {pid} termo {termo['id']} passo: {passo}\n{exc!r}\n"
                    + "".join(f"aviso do SEI: {a}\n" for a in avisos))
                raise
            finally:
                sei.logout()
        mensagem = f"documento {numero} no bloco {bloco}"
        db.marcar_passo(conn, pid, "concluido", mensagem)
        return {"passo": "concluido", "mensagem": mensagem, "documento_sei": numero, "bloco_sei": bloco}
    except Exception as exc:
        mensagem = _mensagem_de(exc, passo, avisos)
        db.marcar_passo(conn, pid, "erro", mensagem)
        t = db.termo_emitido(conn, pedido["termo_id"]) or {}
        return {"passo": "erro", "mensagem": mensagem, "documento_sei": t.get("documento_sei"), "bloco_sei": t.get("bloco_sei")}
