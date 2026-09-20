"""Testa se o editor do SEI preserva estilo inline: justificado + recuo nos parágrafos, tabela 90%."""
import sys; sys.path.insert(0, "/opt/web/termos-responsabilidade")
from sei_base import *
from playwright.sync_api import sync_playwright
import db, termos_html, textos
from sei_acoes import *
conn = db.conectar(Path("/opt/web/termos-responsabilidade/dados/termos.db")); tx = textos.obter(conn)
CC = os.environ.get("CC", "GECONT"); NOME = os.environ.get("NOME", "03/2026 - TESTE")
html = termos_html.corpo_ccusto(CC, db.responsavel(conn, CC), db.bens_do_centro(conn, CC), textos=tx)
ESTILO = {"semrecuo": "text-align:justify;text-indent:0;margin:0 0 7pt", "centro": "text-align:center;text-indent:0;margin:0 0 7pt",
          "direita": "text-align:right;text-indent:0;margin:0 0 7pt", "assinatura": "text-align:center;text-indent:0;margin-top:20pt"}
for cls, st in ESTILO.items(): html = html.replace(f'class="{cls}"', f'style="{st}"')
html = re.sub(r"<p>", '<p style="text-align:justify;text-indent:1.25cm;margin:0 0 7pt">', html)
html = html.replace("width:100%", "width:90%").replace("width:80%", "width:90%")
print("classes restantes:", len(re.findall(r'class="', html)), "| p com style:", len(re.findall(r'<p style', html)))
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    try:
        s.login(env_pca()); s.abrir_processo(PROCESSO)
        rotulo = f"Termo de Responsabilidade {NOME}"
        if not no_da_arvore(s, rotulo):
            criar_documento(s, ctx, NOME, html)
        no = esperar(lambda: no_da_arvore(s, rotulo), 30, erro="não apareceu"); print("nó:", no)
        s.frame("ifrArvore").click(f"#anchor{no['id']}"); time.sleep(4)
        # foto do documento como o SEI mostra
        for f in s.p.frames:
            if f.name == "ifrVisualizacao":
                try: f.wait_for_load_state("load", timeout=30000)
                except Exception: pass
        s.foto("formatacao-sei")
        for f in s.p.frames:
            try:
                if f.locator("table").count() and "Termo de Responsabilidade - " in f.evaluate("() => document.body.innerText"):
                    print("estilos preservados:", f.evaluate("() => ({p: [...document.querySelectorAll('p')].slice(0,4).map(p => p.getAttribute('style')), table: [...document.querySelectorAll('table')].map(t => t.getAttribute('style'))})"))
            except Exception: pass
    finally: s.logout(); b.close()
