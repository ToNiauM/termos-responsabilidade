"""Exclui os blocos 'Termos {CC}' criados por engano em CAE-A (69767–69787), só se vazios e com esse nome."""
import sys; sys.path.insert(0, "/opt/web/termos-responsabilidade")
from sei_base import *
from playwright.sync_api import sync_playwright
import re as _re
def trocar_unidade(s, sigla):
    oc = s.p.locator("#lnkInfraUnidade").first.get_attribute("onclick") or ""
    s.p.goto(urllib.parse.urljoin(BASE, _re.search(r"href='([^']+)'", oc).group(1)), wait_until="domcontentloaded"); time.sleep(1)
    linha = s.p.locator("tr", has=s.p.locator(f"td:text-is('{sigla}')")).first
    with s.p.expect_navigation(wait_until="domcontentloaded"): linha.locator("label.infraRadioLabel").first.click()
    time.sleep(1); return s.p.locator("#lnkInfraUnidade").first.inner_text().strip()
def ir_blocos(s):
    href = s.p.locator("a[href*='acao=bloco_assinatura_listar']").first.get_attribute("href")
    s.p.goto(urllib.parse.urljoin(BASE, href), wait_until="domcontentloaded"); time.sleep(1.2)
ALVO = set(range(69767, 69788))
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    dialogos = []
    s.p.on("dialog", lambda d: (dialogos.append(d.message), d.accept()))
    try:
        atual = s.login(env_pca())
        if atual != "CAE-A": print("unidade agora:", trocar_unidade(s, "CAE-A"))
        ir_blocos(s); print("h1:", s.p.evaluate("() => (document.querySelector('h1')||{}).innerText"))
        excluidos = []
        for _ in range(30):
            linhas = s.p.evaluate("() => [...document.querySelectorAll('tr')].map(tr => { const tds = [...tr.querySelectorAll('td')].map(td => td.innerText.trim()); return {tds, icones: [...tr.querySelectorAll('img[title]')].map(i => i.title)}; }).filter(r => r.tds.length > 5)")
            if _ == 0: print("exemplo de linha:", linhas[0] if linhas else None)
            def numero(r): return next((int(t) for t in r["tds"] if t.isdigit()), None)
            alvo = [l for l in linhas if numero(l) in ALVO and any(t.startswith("Termos ") for t in l["tds"]) and "CAE-A" in l["tds"] and any("xcluir" in i for i in l["icones"])]
            for l in alvo: l["n"] = str(numero(l)); l["d"] = next(t for t in l["tds"] if t.startswith("Termos "))
            if not alvo: break
            l = alvo[0]
            ok = s.p.evaluate("n => { const tr = [...document.querySelectorAll('tr')].find(tr => [...tr.querySelectorAll('td')].some(td => td.innerText.trim() === n)); const img = tr && tr.querySelector('img[title=\"Excluir Bloco\"]'); if (!img) return false; (img.closest('a') || img).click(); return true; }", l["n"])
            if not ok: print("não achei o ícone de", l["n"]); break
            time.sleep(2)
            excluidos.append((l["n"].strip(), l["d"].strip())); log("excluído:", excluidos[-1], "| diálogo:", dialogos[-1:] )
            ir_blocos(s)
        restantes = s.p.evaluate("() => [...document.querySelectorAll('tr')].map(tr => tr.innerText.replace(/\\s+/g,' ').trim()).filter(t => /Termos /.test(t))")
        print("excluídos:", len(excluidos)); print("restantes com 'Termos' em CAE-A:", restantes); s.foto("caea-depois")
    finally: s.logout(); b.close()
