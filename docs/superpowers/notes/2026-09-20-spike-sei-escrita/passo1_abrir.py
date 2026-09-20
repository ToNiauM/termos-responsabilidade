"""Passo 1: login, abrir o processo de rascunho, listar árvore e a barra de botões do processo."""
from sei_base import *
from playwright.sync_api import sync_playwright

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900})
    s = SEI(ctx)
    try:
        s.login(env_pca())
        arv = s.abrir_processo(PROCESSO)
        print(json.dumps(arv, ensure_ascii=False, indent=1)[:3000])
        s.foto("processo")
        v = s.frame("ifrVisualizacao")
        botoes = v.evaluate("""() => [...document.querySelectorAll('#divArvoreAcoes a, .divArvoreAcoes a, a[href*="acao="]')].map(a => ({texto: (a.innerText||'').trim(), title: a.title, img: (a.querySelector('img')||{}).title||'', href: (a.getAttribute('href')||'').slice(0,160), onclick: (a.getAttribute('onclick')||'').slice(0,160)}))""")
        print("BOTÕES:"); [print(" ", json.dumps(x, ensure_ascii=False)) for x in botoes[:40]]
        (SAIDA / "visualizacao.html").write_text(v.content())
    finally:
        s.logout(); b.close()
