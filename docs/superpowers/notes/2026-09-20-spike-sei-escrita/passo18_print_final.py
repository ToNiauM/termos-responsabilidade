from sei_acoes import *
from playwright.sync_api import sync_playwright
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    try:
        s.login(env_pca()); s.abrir_processo(PROCESSO)
        no = no_da_arvore(s, "Termo de Responsabilidade 04/2026 - TESTE"); print("nó:", no)
        s.frame("ifrArvore").click(f"#anchor{no['id']}"); time.sleep(4); s.foto("prova-final")
    finally: s.logout(); b.close()
