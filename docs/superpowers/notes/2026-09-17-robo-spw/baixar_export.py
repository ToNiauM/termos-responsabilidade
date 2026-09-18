import os, json
from playwright.sync_api import sync_playwright
env = dict(l.rstrip("\n").split("=", 1) for l in open("/opt/web/termos-responsabilidade/secrets/spw.env") if "=" in l)
S = os.path.dirname(os.path.abspath(__file__)); P = "#ContentPlaceHolder1_ASPxRoundPanel1_"
with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page(viewport={"width": 1280, "height": 900}, accept_downloads=True)
    pg.goto(env["SPW_LOGIN_URL"], wait_until="networkidle", timeout=60000)
    pg.fill(P + "txtUsuario_I", env["SPW_USUARIO"]); pg.click(P + "txtSenha_I_CLND"); pg.wait_for_timeout(300)
    pg.fill(P + "txtSenha_I", env["SPW_SENHA"], force=True)
    with pg.expect_navigation(wait_until="networkidle", timeout=60000): pg.click(P + "btnEntrar")
    pg.goto(env["SPW_CONSULTA_URL"], wait_until="networkidle", timeout=60000)
    pg.click("#ContentPlaceHolder1_ASPxButton1"); pg.wait_for_timeout(1500)
    popup = pg.evaluate("""() => { const d=document.querySelector('[id*=PCExportacao]'); return d? d.innerText.trim().slice(0,400) : null }""")
    print("popup texto:", json.dumps(popup, ensure_ascii=False))
    print("popup inputs:", json.dumps(pg.evaluate("""() => Array.from(document.querySelectorAll('[id*=PCExportacao] input, [id*=PCExportacao] select, [id*=PCExportacao] label, [id*=PCExportacao] td.dxeEditArea')).filter(e=>e.offsetWidth).map(e=>({tag:e.tagName,type:e.type||'',id:e.id,value:e.value||'',text:(e.innerText||'').trim().slice(0,30),checked:e.checked}))"""), ensure_ascii=False)[:2000])
    pg.select_option("#ContentPlaceHolder1_PCExportacao_cboArquivo", label="Excel"); pg.select_option("#ContentPlaceHolder1_PCExportacao_cboModeloExportacao", label="Detalhado"); pg.wait_for_timeout(500)
    with pg.expect_download(timeout=300000) as dl:
        pg.click("#ContentPlaceHolder1_PCExportacao_imgExportar")
    d = dl.value; destino = f"{S}/export_{d.suggested_filename}"; d.save_as(destino)
    print("download:", d.suggested_filename, os.path.getsize(destino), "bytes ->", destino)
    b.close()
