from sei_base import *
from playwright.sync_api import sync_playwright
exec(open("passo13_blocos.py").read().split("with sync_playwright()")[0].split("UNIDADE = ")[0])
import re as _re
def trocar_unidade(s, sigla):
    oc = s.p.locator("#lnkInfraUnidade").first.get_attribute("onclick") or ""
    s.p.goto(urllib.parse.urljoin(BASE, _re.search(r"href='([^']+)'", oc).group(1)), wait_until="domcontentloaded"); time.sleep(1)
    linha = s.p.locator("tr", has=s.p.locator(f"td:text-is('{sigla}')")).first
    with s.p.expect_navigation(wait_until="domcontentloaded"): linha.locator("label.infraRadioLabel").first.click()
    time.sleep(1.5)
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    try:
        s.login(env_pca())
        for unidade in ("GELIC", "GESERV"):
            if unidade != "GELIC": trocar_unidade(s, unidade)
            print(unidade, "submenu Blocos:", s.p.evaluate("() => [...document.querySelectorAll('a[href*=bloco]')].map(a => a.innerText.trim() + ' -> ' + (a.getAttribute('href')||'').split('&')[0])"))
            href = s.p.locator("a[href*='acao=bloco_assinatura_listar']").first.get_attribute("href")
            r = s.p.goto(urllib.parse.urljoin(BASE, href), wait_until="domcontentloaded"); time.sleep(1.5)
            print(unidade, "status:", r.status, "| url:", s.p.url.split("?")[1][:40], "| título h1:", s.p.evaluate("() => (document.querySelector('h1')||{}).innerText"))
            print(unidade, "aviso:", s.p.evaluate("() => (document.body.innerText.match(/(n[aã]o (tem|possui)[^.]*\\.)/i)||[''])[0]"))
    finally: s.logout(); b.close()
