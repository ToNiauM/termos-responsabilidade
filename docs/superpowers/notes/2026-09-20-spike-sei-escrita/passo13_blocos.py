"""Troca para GESERV, lista blocos existentes e mapeia a tela Novo Bloco (sem criar)."""
from sei_base import *
from playwright.sync_api import sync_playwright
import re as _re
UNIDADE = os.environ.get("UNIDADE", "GESERV")
def trocar_unidade(s, sigla):
    oc = s.p.locator("#lnkInfraUnidade").first.get_attribute("onclick") or ""
    url = urllib.parse.urljoin(BASE, _re.search(r"href='([^']+)'", oc).group(1))
    s.p.goto(url, wait_until="domcontentloaded"); time.sleep(1)
    linha = s.p.locator("tr", has=s.p.locator(f"td:text-is('{sigla}')")).first
    print("linha:", linha.evaluate("tr => ({oc: (tr.getAttribute('onclick')||'').slice(0,160), html: tr.innerHTML.slice(0,400)})"))
    with s.p.expect_navigation(wait_until="domcontentloaded"): linha.locator("label.infraRadioLabel").first.click()
    time.sleep(1.5)
    atual = s.p.locator("#lnkInfraUnidade").first.inner_text().strip()
    print("unidade atual:", atual); return atual
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    try:
        s.login(env_pca()); trocar_unidade(s, UNIDADE)
        s.p.locator("a:has-text('Blocos')").first.click(); time.sleep(0.8)
        with s.p.expect_navigation(wait_until="domcontentloaded"): s.p.locator("a[href*='acao=bloco_assinatura_listar']").first.click()
        time.sleep(1.5)
        print("URL blocos:", s.p.url[:90]); s.foto("blocos-geserv")
        print("blocos existentes:", s.p.evaluate("() => [...document.querySelectorAll('tr')].map(tr => tr.innerText.replace(/\\s+/g,' ').trim()).filter(t => /Termos|Número|Estado/.test(t)).slice(0,30)"))
        print("botões:", s.p.evaluate("() => [...document.querySelectorAll('button, input[type=button], input[type=submit]')].map(e => ({id: e.id, v: e.value || e.innerText, oc: (e.getAttribute('onclick')||'').slice(0,120)}))"))
        print("links acao:", s.p.evaluate("() => [...document.querySelectorAll('a[href*=bloco_assinatura]')].map(a => ({t: a.innerText.trim().slice(0,30), href: (a.getAttribute('href')||'').slice(0,90)}))"))
        s.p.locator("#btnNovo, button:has-text('Novo'), input[value='Novo'], a:has-text('Novo Bloco'), a:has-text('Novo')").first.click(); time.sleep(1.5); s.foto("novo-bloco")
        print("URL novo:", s.p.url[:100])
        print("CAMPOS:"); [print("  ", json.dumps(c, ensure_ascii=False)[:300]) for c in campos(s.p) if c["type"] != "hidden"]
        print("LABELS:", s.p.evaluate("() => [...document.querySelectorAll('label')].map(l => l.innerText.trim()).filter(Boolean)"))
    finally: s.logout(); b.close()
