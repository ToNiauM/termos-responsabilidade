from sei_base import *
from playwright.sync_api import sync_playwright
import re as _re
def ir_trocar_unidade(s):
    oc = s.p.locator("#lnkInfraUnidade").first.get_attribute("onclick") or ""
    m = _re.search(r"href='([^']+)'", oc); url = urllib.parse.urljoin(BASE, m.group(1))
    s.p.goto(url, wait_until="domcontentloaded"); time.sleep(1.5)
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    try:
        s.login(env_pca()); ir_trocar_unidade(s); s.foto("trocar-unidade")
        print("URL:", s.p.url[:100])
        print("links unidades:", s.p.evaluate("() => [...document.querySelectorAll('a')].filter(a => /^[A-Z][A-Z-]{2,12}$/.test(a.innerText.trim())).map(a => ({t: a.innerText.trim(), href: (a.getAttribute('href')||'').slice(0,120), oc: (a.getAttribute('onclick')||'').slice(0,100)}))"))
        print("selects:", s.p.evaluate("() => [...document.querySelectorAll('select')].map(s => ({id: s.id, opts: [...s.options].map(o => o.text).slice(0,40)}))"))
        print("tabela:", s.p.evaluate("() => [...document.querySelectorAll('tr')].map(tr => tr.innerText.replace(/\\s+/g,' ').trim()).slice(0,20)"))
    finally: s.logout(); b.close()
