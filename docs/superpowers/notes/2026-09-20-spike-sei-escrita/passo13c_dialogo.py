from sei_base import *
from playwright.sync_api import sync_playwright
import re as _re
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    msgs = []
    s.p.on("dialog", lambda d: (msgs.append(d.message), d.dismiss()))
    try:
        s.login(env_pca())
        oc = s.p.locator("#lnkInfraUnidade").first.get_attribute("onclick") or ""
        s.p.goto(urllib.parse.urljoin(BASE, _re.search(r"href='([^']+)'", oc).group(1)), wait_until="domcontentloaded"); time.sleep(1)
        for sigla in ("GESERV", "GECER", "SEGED", "CAE-A"):
            linha = s.p.locator("tr", has=s.p.locator(f"td:text-is('{sigla}')")).first
            with s.p.expect_navigation(wait_until="domcontentloaded"): linha.locator("label.infraRadioLabel").first.click()
            time.sleep(1)
            href = s.p.locator("a[href*='acao=bloco_assinatura_listar']").first.get_attribute("href")
            resp = s.p.goto(urllib.parse.urljoin(BASE, href), wait_until="domcontentloaded"); time.sleep(1)
            print(sigla, "→", s.p.evaluate("() => (document.querySelector('h1')||{}).innerText"), "| dialogs:", msgs[-1:] , "| redirects:", [r.url.split('acao=')[1][:30] for r in ([resp.request.redirected_from] if resp and resp.request.redirected_from else [])])
            oc = s.p.locator("#lnkInfraUnidade").first.get_attribute("onclick") or ""
            s.p.goto(urllib.parse.urljoin(BASE, _re.search(r"href='([^']+)'", oc).group(1)), wait_until="domcontentloaded"); time.sleep(1)
    finally: s.logout(); b.close()
