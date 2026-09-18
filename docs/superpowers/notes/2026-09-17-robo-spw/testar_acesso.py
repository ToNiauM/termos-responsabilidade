"""Teste de acesso ao SPW, só leitura: login DevExpress, abre a consulta, captura telas e lista controles."""
import os, json
from playwright.sync_api import sync_playwright
env = dict(l.rstrip("\n").split("=", 1) for l in open("/opt/web/termos-responsabilidade/secrets/spw.env") if "=" in l)
S = os.path.dirname(os.path.abspath(__file__))
P = "#ContentPlaceHolder1_ASPxRoundPanel1_"
def controles(pg):
    return pg.evaluate("""() => Array.from(document.querySelectorAll('input,select,button,a,div[id]')).map(e => ({
        tag:e.tagName, type:e.type||'', id:e.id, vis:!!(e.offsetWidth||e.offsetHeight), text:(e.innerText||e.value||'').trim().slice(0,40), href:(e.href||'').slice(0,80)}))
        .filter(e => e.vis && e.type!=='hidden' && (e.text || e.type) && !/^dx|_DDD|_PW|_CD$|_I$/.test(e.id||'x'))""")
with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page(viewport={"width": 1280, "height": 900})
    erros = []; pg.on("pageerror", lambda e: erros.append(str(e)))
    pg.goto(env["SPW_LOGIN_URL"], wait_until="networkidle", timeout=60000)
    pg.fill(P + "txtUsuario_I", env["SPW_USUARIO"])
    pg.click(P + "txtSenha_I_CLND"); pg.wait_for_timeout(300)
    pg.fill(P + "txtSenha_I", env["SPW_SENHA"], force=True)
    with pg.expect_navigation(wait_until="networkidle", timeout=60000):
        pg.click(P + "btnEntrar")
    print("APOS LOGIN title:", pg.title(), "| url:", pg.url)
    pg.screenshot(path=f"{S}/02_apos_login.png", full_page=True)
    print("APOS LOGIN controles:", json.dumps(controles(pg), ensure_ascii=False)[:2500])
    pg.goto(env["SPW_CONSULTA_URL"], wait_until="networkidle", timeout=60000)
    print("CONSULTA title:", pg.title(), "| url:", pg.url)
    pg.screenshot(path=f"{S}/03_consulta.png", full_page=True)
    print("CONSULTA controles:", json.dumps(controles(pg), ensure_ascii=False)[:5000])
    print("frames:", [f.url for f in pg.frames][:6])
    print("pageerrors:", erros[:3])
    b.close()
