"""Base do spike: login e abertura de processo (copiados do coletor do PCA). Código descartável."""
import json, os, re, sys, time, urllib.parse
from pathlib import Path

BASE = "https://sei.cfc.org.br/sei/"
HOSTS = {"sei.cfc.org.br", "sip.cfc.org.br"}
SAIDA = Path(os.environ.get("SPIKE_SAIDA", "/tmp/claude-1003/-opt-web-termos-responsabilidade/233f7156-8544-4b7a-a36c-a5d39d8168f9/scratchpad/sei"))
PROCESSO = os.environ.get("SPIKE_PROCESSO", "90796110000022.000059/2026-88")

JS_ARVORE = """
() => {
  const out = [];
  for (const a of document.querySelectorAll("a[id^='anchor']")) {
    const m = /^anchor(\\d+)$/.exec(a.id);
    if (!m) continue;
    out.push({id: m[1], texto: (a.innerText || '').replace(/\\s+/g, ' ').trim(), href: a.getAttribute('href') || ''});
  }
  return {anchors: out, aguarde: document.querySelectorAll("a[id^='anchorAGUARDE']").length};
}
"""

def env_pca():
    env = {}
    for linha in Path("/opt/web/pca-cfc/.env").read_text().splitlines():
        if linha.startswith("SEI_"):
            k, v = linha.split("=", 1); env[k] = v.strip().strip("'\"")
    return env

def log(*a):
    print(time.strftime("%H:%M:%S"), *a, flush=True)

def esperar(cond, timeout_s=30, intervalo=0.3, erro="tempo esgotado"):
    fim = time.time() + timeout_s
    while time.time() < fim:
        r = cond()
        if r: return r
        time.sleep(intervalo)
    raise RuntimeError(erro)

class SEI:
    def __init__(self, ctx, t=30):
        self.ctx, self.t = ctx, t
        self.p = ctx.new_page(); self.p.set_default_timeout(t * 1000)
        self.n = 0

    def foto(self, nome, page=None):
        self.n += 1
        (page or self.p).screenshot(path=str(SAIDA / f"{self.n:02d}-{nome}.png"), full_page=True)
        log("foto", f"{self.n:02d}-{nome}.png")

    def frame(self, nome):
        el = self.p.locator(f"#{nome}").first.element_handle(timeout=self.t * 1000)
        f = el.content_frame() if el else None
        if f is None: raise RuntimeError(f"frame {nome} ausente")
        return f

    def login(self, cfg):
        self.p.goto(cfg["SEI_LOGIN_URL"], wait_until="domcontentloaded")
        self.p.select_option("#selOrgao", label=cfg.get("SEI_ORGAO", "CFC"))
        self.p.fill("#txtUsuario", cfg["SEI_USUARIO"]); self.p.fill("#pwdSenha", cfg["SEI_SENHA"])
        self.p.click("#sbmAcessar"); self.p.wait_for_selector("#txtPesquisaRapida")
        unidade = self.p.locator("#lnkInfraUnidade").first.inner_text(timeout=3000).strip()
        log("login ok, unidade:", unidade); return unidade

    def logout(self):
        try: self.p.click("#lnkInfraSairSistema", timeout=5000)
        except Exception: pass

    def arvore(self):
        try:
            r = self.frame("ifrArvore").evaluate(JS_ARVORE)
        except Exception: return None
        return r if r["anchors"] and not r["aguarde"] else None

    def arvore_parcial(self):
        try: r = self.frame("ifrArvore").evaluate(JS_ARVORE)
        except Exception: return None
        return r if r["anchors"] else None

    def abrir_todas_pastas(self):
        """Processo com muitos documentos: SEI agrupa em pastas fechadas (nó AGUARDE). Abre todas e espera estabilizar."""
        f = self.frame("ifrArvore")
        botao = f.locator("img[title='Abrir todas as Pastas']")
        if not botao.count(): return self.arvore()
        botao.first.click()
        contagens = []
        def estavel():
            r = self.arvore()
            if not r: return None
            contagens.append(len(r["anchors"]))
            return r if len(contagens) >= 2 and contagens[-1] == contagens[-2] else None
        return esperar(estavel, self.t, intervalo=0.7, erro="pastas não abriram")

    def abrir_processo(self, numero):
        campo = self.p.locator("#txtPesquisaRapida"); campo.fill(numero)
        with self.p.expect_navigation(wait_until="domcontentloaded"): campo.press("Enter")
        log("título:", self.p.title(), "| url:", self.p.url.split("?")[0])
        esperar(self.arvore_parcial, self.t, erro="árvore não carregou")
        arv = self.arvore() or self.abrir_todas_pastas()
        log("árvore:", len(arv["anchors"]), "nós")
        return arv

def frame_por_url(s, trecho, timeout=30):
    def achar():
        for f in s.p.frames:
            if trecho in f.url:
                try: f.evaluate("1")
                except Exception: return None
                return f
    return esperar(achar, timeout, erro=f"frame com '{trecho}' não apareceu")

def selecionar_raiz(s, arv):
    s.frame("ifrArvore").click(f"#anchor{arv['anchors'][0]['id']}")
    return frame_por_url(s, "acao=arvore_visualizar")

def campos(fr):
    return fr.evaluate("""() => [...document.querySelectorAll('input,select,textarea,button')].map(e => ({tag: e.tagName, type: e.type, id: e.id, name: e.name, value: (e.value||'').slice(0,60), checked: e.checked, label: (document.querySelector('label[for="'+e.id+'"]')||{}).innerText||'', opts: e.tagName=='SELECT' ? [...e.options].slice(0,30).map(o=>o.text) : undefined}))""")

def frame_com(s, seletor, timeout=30):
    """Frame que contém `seletor` (as URLs dos iframes do SEI não são confiáveis)."""
    def achar():
        for f in s.p.frames:
            try:
                if f.locator(seletor).count(): return f
            except Exception: pass
    return esperar(achar, timeout, erro=f"nenhum frame com {seletor}")
