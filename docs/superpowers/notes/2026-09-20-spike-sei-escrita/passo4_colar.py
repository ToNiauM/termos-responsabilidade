"""Cria doc NN/AAAA - TESTE, injeta o HTML do nosso termo no Corpo do Texto, salva, reabre e confere."""
import sys; sys.path.insert(0, "/opt/web/termos-responsabilidade")
from sei_base import *
from playwright.sync_api import sync_playwright
import db, termos_html, textos
NOME = os.environ.get("NOME", "02/2026 - TESTE"); CC = os.environ.get("CC", "GECER")
conn = db.conectar(Path("/opt/web/termos-responsabilidade/dados/termos.db"))
html = termos_html.corpo_ccusto(CC, db.responsavel(conn, CC), db.bens_do_centro(conn, CC), textos=textos.obter(conn))
print("HTML do termo:", len(html), "chars; bens:", len(db.bens_do_centro(conn, CC)))

JS_CORPO = """() => { const f = document.querySelector('iframe[title="Corpo do Texto"]'); const box = f && f.closest('[id^="cke_txaEditor_"]'); return box ? box.id.replace('cke_','') : null; }"""

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    try:
        s.login(env_pca()); arv = s.abrir_processo(PROCESSO)
        fr = selecionar_raiz(s, arv); time.sleep(1)
        fr.locator("a:has(img[title='Incluir Documento'])").first.click()
        fr = frame_com(s, "text=Escolha o Tipo do Documento"); time.sleep(1)
        fr.locator("a").filter(has_text="Termo de Responsabilidade").first.click()
        fr = frame_com(s, "#txtNomeArvore"); time.sleep(1)
        fr.fill("#txtNomeArvore", NOME); fr.click("label[for=optPublico]"); fr.wait_for_function("() => document.getElementById('optPublico').checked")
        with ctx.expect_page(timeout=30000) as nova: fr.click("#btnSalvar")
        ed = nova.value; ed.wait_for_load_state("domcontentloaded")
        ed.wait_for_function("() => typeof CKEDITOR !== 'undefined' && document.querySelector('iframe[title=\"Corpo do Texto\"]')"); time.sleep(2)
        inst = ed.evaluate(JS_CORPO); print("instância do corpo:", inst)
        antes = ed.evaluate(f"() => CKEDITOR.instances['{inst}'].getData().length"); print("modelo tinha", antes, "chars")
        ed.evaluate(f"h => {{ const e = CKEDITOR.instances['{inst}']; e.setData(h); e.fire('change'); }}", html)
        time.sleep(1); s.foto("editor-colado", ed)
        # salvar: o botão visível da barra compartilhada
        ed.locator("a[title^='Salvar']:visible").first.click(); time.sleep(3)
        print("após salvar, url:", ed.url[:120]); s.foto("editor-salvo", ed)
        msg = ed.evaluate("() => (document.body.innerText||'').slice(0,300)")
        ed.close()
        # reabrir o documento na árvore e ler o conteúdo
        arv2 = esperar(lambda: (lambda a: a if a and any(NOME in x["texto"] for x in a["anchors"]) else None)(s.arvore()), 30, erro="doc não apareceu na árvore")
        no = [a for a in arv2["anchors"] if NOME in a["texto"]][0]; print("nó:", no["texto"])
        s.frame("ifrArvore").click(f"#anchor{no['id']}"); time.sleep(3)
        vis = frame_com(s, "text=TERMO DE RESPONSABILIDADE", 30)
        texto = vis.evaluate("() => document.body.innerText")
        print("conteúdo reaberto tem", len(texto), "chars; tem nome do responsável:", db.responsavel(conn, CC)["responsavel"] in texto, "| tem 'OBSERVAÇÕES' do modelo:", "OBSERVAÇÕES" in texto)
        s.foto("doc-reaberto")
        (SAIDA / "doc-reaberto.html").write_text(vis.content())
    finally: s.logout(); b.close()
