from sei_base import *
from playwright.sync_api import sync_playwright
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    try:
        print("login em:", s.login(env_pca()))
        href = s.p.locator("a[href*='acao=bloco_assinatura_listar']").first.get_attribute("href")
        s.p.goto(urllib.parse.urljoin(BASE, href), wait_until="domcontentloaded"); time.sleep(1.2)
        for l in s.p.evaluate("() => [...document.querySelectorAll('tr')].map(tr => [...tr.querySelectorAll('td')].map(td => td.innerText.trim()).filter(Boolean).join(' | ')).filter(t => /Termos /.test(t))"): print("  ", l)
        s.foto("blocos-gelic")
    finally: s.logout(); b.close()
