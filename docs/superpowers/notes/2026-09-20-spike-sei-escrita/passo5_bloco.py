from sei_base import *
from playwright.sync_api import sync_playwright
NOME = os.environ.get("NOME", "02/2026 - TESTE"); BLOCO = os.environ.get("BLOCO", "Termos TESTE"); FAZER = os.environ.get("FAZER") == "1"
def frame_doc(s):
    def achar():
        for f in s.p.frames:
            try:
                if f.locator("img[title='Incluir em Bloco de Assinatura']").count(): return f
            except Exception: pass
    return esperar(achar, 30, erro="barra do documento não apareceu")
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    try:
        s.login(env_pca()); arv = s.abrir_processo(PROCESSO)
        no = [a for a in arv["anchors"] if NOME in a["texto"]][0]; print("nó:", no["texto"])
        s.frame("ifrArvore").click(f"#anchor{no['id']}")
        fr = frame_doc(s); time.sleep(1)
        bot = fr.evaluate("""() => [...document.querySelectorAll('a')].filter(a => a.querySelector('img') && a.querySelector('img').title).map(a => ({img: a.querySelector('img').title, href: (a.getAttribute('href')||'').slice(0,80), oc: (a.getAttribute('onclick')||'').slice(0,80), target: a.target}))""")
        print("BOTÕES DO DOCUMENTO:"); [print("  ", json.dumps(x, ensure_ascii=False)) for x in bot]
        fr.locator("a:has(img[title='Incluir em Bloco de Assinatura'])").first.click()
        fr = frame_com(s, "#selBloco", 30); time.sleep(1)
        s.foto("incluir-bloco")
        print("CAMPOS:"); [print("  ", json.dumps(c, ensure_ascii=False)) for c in campos(fr) if c["type"] not in ("hidden",)]
        print("TEXTO:", fr.evaluate("() => document.body.innerText.slice(0,700)"))
        if FAZER:
            fr.select_option("#selBloco", label=[o for o in fr.evaluate("() => [...document.getElementById('selBloco').options].map(o => o.text)") if o.endswith(" - " + BLOCO)][0]); time.sleep(1)
            fr.click("#sbmIncluir"); time.sleep(3)
            s.foto("bloco-incluido")
            print("DEPOIS:", frame_com(s, "body").evaluate("() => document.body.innerText.slice(0,800)"))
    finally: s.logout(); b.close()
