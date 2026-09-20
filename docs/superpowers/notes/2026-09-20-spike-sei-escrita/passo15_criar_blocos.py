"""Cria os blocos 'Termos {CC}' na unidade atual do login (GELIC), um por centro de custo. Idempotente."""
import sys; sys.path.insert(0, "/opt/web/termos-responsabilidade")
from sei_base import *
from playwright.sync_api import sync_playwright
import db, re as _re
UNIDADE = os.environ.get("UNIDADE", "GELIC")
def trocar_unidade(s, sigla):
    oc = s.p.locator("#lnkInfraUnidade").first.get_attribute("onclick") or ""
    s.p.goto(urllib.parse.urljoin(BASE, _re.search(r"href='([^']+)'", oc).group(1)), wait_until="domcontentloaded"); time.sleep(1)
    linha = s.p.locator("tr", has=s.p.locator(f"td:text-is('{sigla}')")).first
    with s.p.expect_navigation(wait_until="domcontentloaded"): linha.locator("label.infraRadioLabel").first.click()
    time.sleep(1); return s.p.locator("#lnkInfraUnidade").first.inner_text().strip()
conn = db.conectar(Path("/opt/web/termos-responsabilidade/dados/termos.db"))
centros = [r[0] for r in conn.execute("select ccustos from responsaveis order by ccustos")]
def lista_blocos(s):
    href = s.p.locator("a[href*='acao=bloco_assinatura_listar']").first.get_attribute("href")
    s.p.goto(urllib.parse.urljoin(BASE, href), wait_until="domcontentloaded"); time.sleep(1.2)
    return s.p.evaluate("() => [...document.querySelectorAll('tr')].map(tr => [...tr.querySelectorAll('td')].map(td => td.innerText.trim())).filter(c => c.length > 2)")
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    try:
        atual = s.login(env_pca())
        if atual != UNIDADE: atual = trocar_unidade(s, UNIDADE)
        assert atual == UNIDADE, f"unidade errada: {atual}"; log("unidade:", atual)
        existentes = {" ".join(c) for c in lista_blocos(s)}
        criados, pulados = [], []
        for cc in centros:
            nome = f"Termos {cc}"
            if any(nome in e for e in existentes): pulados.append(nome); continue
            s.p.click("#btnNovo"); s.p.wait_for_selector("#txaDescricao"); time.sleep(0.5)
            s.p.fill("#txaDescricao", nome)
            with s.p.expect_navigation(wait_until="domcontentloaded"): s.p.click("button[name=sbmCadastrarBloco]")
            time.sleep(1); criados.append(nome)
            log("criado:", nome, "| url:", s.p.url.split("?")[1][:40])
            lista_blocos(s)
        final = lista_blocos(s); s.foto("blocos-gelic-final")
        termos = [c for c in final if any(x.startswith("Termos ") for x in c)]
        print("criados:", len(criados), "pulados:", pulados)
        print("blocos 'Termos' na lista:", len(termos)); [print("  ", c[:4]) for c in termos]
    finally: s.logout(); b.close()
