from sei_acoes import *
from playwright.sync_api import sync_playwright
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    try:
        s.login(env_pca()); arv = s.abrir_processo(PROCESSO)
        fr = selecionar_raiz(s, arv); time.sleep(0.8)
        fr.locator("a:has(img[title='Incluir Documento'])").first.click()
        fr = frame_com(s, "text=Escolha o Tipo do Documento"); time.sleep(1)
        print("links perto do título:", fr.evaluate("() => [...document.querySelectorAll('a, img')].filter(e => /todos|exibir|mais|\\+/i.test((e.title||'')+(e.alt||'')+(e.id||''))).map(e => ({tag: e.tagName, id: e.id, title: e.title, alt: e.alt, oc: (e.getAttribute('onclick')||'').slice(0,80), href: (e.getAttribute('href')||'').slice(0,60)}))"))
        antes = len(fr.evaluate("() => [...document.querySelectorAll('a[onclick^=\"escolher\"]')]"))
        fr.locator("#ancTodos, img[title*='Todos'], a[title*='Todos'], img[title*='todos']").first.click(); time.sleep(1.5)
        todos = fr.evaluate("() => [...document.querySelectorAll('a[onclick^=\"escolher\"]')].map(a => a.innerText.trim())")
        print("antes:", antes, "depois:", len(todos)); print([t for t in todos if "evolu" in t.lower()])
        (SAIDA / "tipos-todos.json").write_text(json.dumps(todos, ensure_ascii=False, indent=1))
        fr.fill("#txtFiltro", "Devolu") if fr.locator("#txtFiltro").count() else None; time.sleep(1); s.foto("tipos-todos-devolucao")
    finally: s.logout(); b.close()
