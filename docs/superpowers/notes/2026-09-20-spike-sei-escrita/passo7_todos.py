"""Missão: um Termo de Responsabilidade por centro de custo no processo de rascunho, rótulo 01/2026 - CC,
incluído no bloco Termos TESTE. Idempotente: pula documento que já está na árvore com o mesmo rótulo."""
import sys; sys.path.insert(0, "/opt/web/termos-responsabilidade")
from sei_base import *
from playwright.sync_api import sync_playwright
import db, termos_html, textos
BLOCO = os.environ.get("BLOCO", "Termos TESTE"); ANO = "2026"; SO = os.environ.get("SO")
conn = db.conectar(Path("/opt/web/termos-responsabilidade/dados/termos.db")); tx = textos.obter(conn)
centros = [r[0] for r in conn.execute("select ccustos from responsaveis order by ccustos")]
if SO: centros = SO.split(",")
from sei_acoes import *

resultados = []
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=True); ctx = b.new_context(viewport={"width": 1400, "height": 900}); s = SEI(ctx)
    try:
        s.login(env_pca()); s.abrir_processo(PROCESSO)
        for cc in centros:
            t0 = time.time(); rotulo = f"Termo de Responsabilidade 01/{ANO} - {cc}"; r = {"cc": cc}
            try:
                bens = db.bens_do_centro(conn, cc); r["bens"] = len(bens)
                no = no_da_arvore(s, rotulo)
                if no: r["ja_existia"] = True
                else:
                    html = termos_html.corpo_ccusto(cc, db.responsavel(conn, cc), bens, textos=tx); r["html"] = len(html)
                    criar_documento(s, ctx, f"01/{ANO} - {cc}", html)
                    no = esperar(lambda: no_da_arvore(s, rotulo), 30, erro="documento não apareceu na árvore")
                r["numero_sei"] = no["numero"]; r["t_doc"] = round(time.time() - t0, 1)
                linha = incluir_em_bloco(s, no, BLOCO); r["bloco_linha"] = linha
                r["t_total"] = round(time.time() - t0, 1); r["ok"] = True
            except Exception as e:
                r["erro"] = f"{type(e).__name__}: {str(e)[:200]}"; s.foto(f"erro-{cc}")
                for pg in ctx.pages[1:]:
                    try: pg.close()
                    except Exception: pass
                try: s.abrir_processo(PROCESSO); esperar(s.arvore, 30)
                except Exception as e2: log("recuperação falhou:", e2)
            log(json.dumps(r, ensure_ascii=False)); resultados.append(r)
            (SAIDA / "todos.json").write_text(json.dumps(resultados, ensure_ascii=False, indent=1))
        s.foto("arvore-final")
    finally: s.logout(); b.close()
