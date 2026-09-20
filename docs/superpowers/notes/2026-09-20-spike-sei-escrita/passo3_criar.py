from sei_base import *
from playwright.sync_api import sync_playwright
NUMERO, NOME = os.environ.get("NUM", "01/2026"), os.environ.get("NOME", "TESTE")
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    try:
        s.login(env_pca()); arv = s.abrir_processo(PROCESSO)
        fr = selecionar_raiz(s, arv); time.sleep(1)
        fr.locator("a:has(img[title='Incluir Documento'])").first.click()
        fr = frame_com(s, "text=Escolha o Tipo do Documento"); time.sleep(1)
        todos = fr.evaluate("() => [...document.querySelectorAll('a[onclick^=\"escolher\"]')].map(a => a.innerText.trim())")
        (SAIDA / "tipos.json").write_text(json.dumps(todos, ensure_ascii=False, indent=1))
        print("tipos:", len(todos), "| com Devolu:", [t for t in todos if "evolu" in t])
        fr.locator("a").filter(has_text="Termo de Responsabilidade").first.click()
        fr = frame_com(s, "#txtNomeArvore"); time.sleep(1)
        fr.fill("#txtNomeArvore", NOME); fr.click("label[for=optPublico]"); fr.wait_for_function("() => document.getElementById('optPublico').checked")
        s.foto("form-preenchido")
        with ctx.expect_page(timeout=30000) as nova:
            fr.click("#btnSalvar")
        ed = nova.value; ed.wait_for_load_state("domcontentloaded"); time.sleep(3)
        print("EDITOR url:", ed.url[:160]); print("EDITOR title:", ed.title())
        s.foto("editor", ed)
        info = ed.evaluate("""() => ({ck: typeof CKEDITOR !== 'undefined', inst: typeof CKEDITOR !== 'undefined' ? Object.keys(CKEDITOR.instances) : [],
            iframes: [...document.querySelectorAll('iframe')].map(f => ({id: f.id, cls: f.className, title: f.title})),
            botoes: [...document.querySelectorAll('a,button,img')].map(e => (e.title||e.innerText||e.alt||'').trim()).filter(t => /salv|assin|fechar/i.test(t)),
            ids: [...document.querySelectorAll('[id]')].map(e => e.id).filter(i => /salv|btn|editor|txa/i.test(i)).slice(0,40)})""")
        print(json.dumps(info, ensure_ascii=False, indent=1))
        (SAIDA / "editor.html").write_text(ed.content())
        # árvore depois de criar
        time.sleep(2); arv2 = s.arvore() or s.abrir_processo(PROCESSO)
        print("ÁRVORE:", [a["texto"] for a in arv2["anchors"]])
        s.foto("arvore-depois")
        ed.close()
    finally: s.logout(); b.close()
