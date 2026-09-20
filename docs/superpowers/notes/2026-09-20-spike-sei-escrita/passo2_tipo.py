from sei_base import *
from playwright.sync_api import sync_playwright
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    try:
        s.login(env_pca()); arv = s.abrir_processo(PROCESSO)
        fr = selecionar_raiz(s, arv); time.sleep(1)
        fr.locator("a:has(img[title='Incluir Documento'])").first.click()
        fr = frame_com(s, "text=Escolha o Tipo do Documento"); time.sleep(1)
        print("CAMPOS:"); [print("  ", json.dumps(c, ensure_ascii=False)) for c in campos(fr)]
        tipos = fr.evaluate("""() => [...document.querySelectorAll('a')].map(a => ({t: a.innerText.trim(), href: (a.getAttribute('href')||'').slice(0,100), oc: (a.getAttribute('onclick')||'').slice(0,80)})).filter(x => x.t)""")
        print("TIPOS:", len(tipos)); [print("  ", json.dumps(t, ensure_ascii=False)) for t in tipos if "Termo" in t["t"]]
        fr.locator("a").filter(has_text="Termo de Responsabilidade").first.click()
        fr = frame_com(s, "#txtDescricao, #btnSalvar, input[value='Salvar']"); time.sleep(1.5)
        s.foto("form-gerar")
        print("FORM:"); [print("  ", json.dumps(c, ensure_ascii=False)) for c in campos(fr)]
        print("LABELS:", fr.evaluate("() => [...document.querySelectorAll('label')].map(l => l.innerText.trim()).filter(Boolean)"))
    finally: s.logout(); b.close()
