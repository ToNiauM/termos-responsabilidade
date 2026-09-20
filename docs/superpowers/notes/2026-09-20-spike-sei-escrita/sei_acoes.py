"""Ações de escrita no SEI provadas pelo spike: criar documento, incluir em bloco, achar nó pelo rótulo."""
from sei_base import *

JS_CORPO = """() => { const f = document.querySelector('iframe[title="Corpo do Texto"]'); const box = f && f.closest('[id^="cke_txaEditor_"]'); return box ? box.id.replace('cke_','') : null; }"""
RE_ROTULO = re.compile(r"^(?P<rotulo>.*?)\s*\((?P<numero>\d{6,8})\)\s*$")

def frame_doc(s):
    def achar():
        for f in s.p.frames:
            try:
                if f.locator("img[title='Incluir em Bloco de Assinatura']").count(): return f
            except Exception: pass
    return esperar(achar, 30, erro="barra do documento não apareceu")

def no_da_arvore(s, rotulo):
    a = s.arvore() or s.abrir_todas_pastas()
    if not a: return None
    for x in a["anchors"]:
        m = RE_ROTULO.match(x["texto"])
        if m and m.group("rotulo") == rotulo: return {"id": x["id"], "numero": m.group("numero"), "texto": x["texto"]}

def criar_documento(s, ctx, nome, html):
    fr = selecionar_raiz(s, esperar(s.arvore, 30, erro='árvore não pronta')); time.sleep(0.8)
    fr.locator("a:has(img[title='Incluir Documento'])").first.click()
    fr = frame_com(s, "text=Escolha o Tipo do Documento"); time.sleep(0.8)
    fr.locator("a[onclick^='escolher']").filter(has_text="Termo de Responsabilidade").first.click()
    fr = frame_com(s, "#txtNomeArvore"); time.sleep(0.8)
    fr.fill("#txtNomeArvore", nome); fr.click("label[for=optPublico]"); fr.wait_for_function("() => document.getElementById('optPublico').checked")
    with ctx.expect_page(timeout=30000) as nova: fr.click("#btnSalvar")
    ed = nova.value; ed.wait_for_load_state("domcontentloaded")
    ed.wait_for_function("() => typeof CKEDITOR !== 'undefined' && document.querySelector('iframe[title=\"Corpo do Texto\"]')"); time.sleep(1.5)
    inst = ed.evaluate(JS_CORPO)
    ed.evaluate(f"h => {{ const e = CKEDITOR.instances['{inst}']; e.setData(h); e.fire('change'); }}", html); time.sleep(0.5)
    ed.locator("a[title^='Salvar']:visible").first.click(); time.sleep(2.5)
    ed.close()

def incluir_em_bloco(s, no, bloco):
    s.frame("ifrArvore").click(f"#anchor{no['id']}")
    fr = frame_doc(s); time.sleep(0.8)
    for f in s.p.frames:
        if f.name == "ifrVisualizacao":
            try: f.wait_for_load_state("load", timeout=60000)
            except Exception: pass
    for tentativa in (1, 2):
        frame_doc(s).locator("a:has(img[title='Incluir em Bloco de Assinatura'])").first.click()
        try:
            fr = frame_com(s, "#selBloco", 20); break
        except RuntimeError:
            if tentativa == 2: raise
    time.sleep(0.8)
    opcoes = fr.evaluate("() => [...document.getElementById('selBloco').options].map(o => o.text)")
    alvo = [o for o in opcoes if o.endswith(" - " + bloco)]
    if not alvo: raise RuntimeError(f"Bloco '{bloco}' não existe (opções: {opcoes[:5]}...)")
    numero_bloco = alvo[0].split(" - ")[0]
    linha_antes = fr.evaluate(f"() => {{ const tr = [...document.querySelectorAll('tr')].find(tr => tr.innerText.includes('{no['numero']}')); return tr ? tr.innerText.replace(/\\s+/g,' ').trim() : null; }}")
    if linha_antes and linha_antes.endswith(" " + numero_bloco): return linha_antes + " (já estava)"
    fr.select_option("#selBloco", label=alvo[0]); time.sleep(0.5)
    fr.click("#sbmIncluir"); time.sleep(2)
    fr = frame_com(s, "#selBloco")
    linha = fr.evaluate(f"() => {{ const tr = [...document.querySelectorAll('tr')].find(tr => tr.innerText.includes('{no['numero']}')); return tr ? tr.innerText.replace(/\\s+/g,' ').trim() : null; }}")
    return linha

