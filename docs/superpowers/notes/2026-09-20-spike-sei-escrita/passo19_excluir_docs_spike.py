"""Exclui do processo de rascunho os documentos 'Termo de Responsabilidade 01/2026 - <CC>' criados pelo spike (21)."""
import sys; sys.path.insert(0, "/opt/web/termos-responsabilidade")
from sei_acoes import *
from playwright.sync_api import sync_playwright
import db
conn = db.conectar(Path("/opt/web/termos-responsabilidade/dados/termos.db"))
ALVOS = {f"Termo de Responsabilidade 01/2026 - {r[0]}" for r in conn.execute("select ccustos from responsaveis")}
def frame_doc(s):
    def achar():
        for f in s.p.frames:
            try:
                if f.locator("img[title='Excluir']").count(): return f
            except Exception: pass
    return esperar(achar, 30, erro="barra do documento não apareceu")
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    dialogos = []
    s.p.on("dialog", lambda d: (dialogos.append(d.message), d.accept()))
    try:
        u = s.login(env_pca())
        if u != "GELIC":
            import re as _re
            oc = s.p.locator("#lnkInfraUnidade").first.get_attribute("onclick") or ""
            s.p.goto(urllib.parse.urljoin(BASE, _re.search(r"href='([^']+)'", oc).group(1)), wait_until="domcontentloaded"); time.sleep(1)
            with s.p.expect_navigation(wait_until="domcontentloaded"): s.p.locator("tr", has=s.p.locator("td:text-is('GELIC')")).first.locator("label.infraRadioLabel").first.click()
        s.abrir_processo(PROCESSO)
        excluidos = []
        for _ in range(30):
            arv = s.arvore() or s.abrir_todas_pastas()
            alvo = next((a for a in arv["anchors"] if RE_ROTULO.match(a["texto"]) and RE_ROTULO.match(a["texto"]).group("rotulo") in ALVOS), None)
            if not alvo: break
            s.frame("ifrArvore").click(f"#anchor{alvo['id']}")
            fr = frame_doc(s); time.sleep(1)
            fr.locator("a:has(img[title='Excluir'])").first.click(); time.sleep(2.5)
            excluidos.append(alvo["texto"]); log("excluído:", alvo["texto"], "| diálogo:", dialogos[-1:])
            s.abrir_processo(PROCESSO)
        arv = s.arvore() or s.abrir_todas_pastas()
        print("excluídos:", len(excluidos)); print("árvore agora:", [a["texto"] for a in arv["anchors"]])
    finally: s.logout(); b.close()
