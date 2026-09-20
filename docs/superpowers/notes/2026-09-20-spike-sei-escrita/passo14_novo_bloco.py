"""GELIC: lista os blocos e mapeia a tela Novo Bloco de Assinatura (sem criar)."""
from sei_base import *
from playwright.sync_api import sync_playwright
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    try:
        s.login(env_pca())
        href = s.p.locator("a[href*='acao=bloco_assinatura_listar']").first.get_attribute("href")
        s.p.goto(urllib.parse.urljoin(BASE, href), wait_until="domcontentloaded"); time.sleep(1.5)
        print("h1:", s.p.evaluate("() => (document.querySelector('h1')||{}).innerText"))
        print("blocos:", s.p.evaluate("() => [...document.querySelectorAll('tr')].map(tr => tr.innerText.replace(/\\s+/g,' ').trim()).filter(t => /Termos/.test(t)).slice(0,40)"))
        print("botões:", s.p.evaluate("() => [...document.querySelectorAll('button, input[type=button], input[type=submit]')].map(e => ({id: e.id, v: e.value || e.innerText, oc: (e.getAttribute('onclick')||'').slice(0,120)}))"))
        s.p.locator("#btnNovo, button:has-text('Novo')").first.click(); time.sleep(1.5); s.foto("novo-bloco")
        print("URL novo:", s.p.url.split("?")[1][:60])
        print("CAMPOS:"); [print("  ", json.dumps(c, ensure_ascii=False)[:300]) for c in campos(s.p) if c["type"] != "hidden"]
        print("LABELS:", s.p.evaluate("() => [...document.querySelectorAll('label')].map(l => l.innerText.trim()).filter(Boolean)"))
    finally: s.logout(); b.close()
