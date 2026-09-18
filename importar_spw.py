"""Robô do SPW: entra no sistema de patrimônio, exporta a relação de bens e importa em termos.db.

Roda no host, fora do container, com o venv que tem Playwright + xlrd:

    .venv-robo/bin/python importar_spw.py

Cron (dias úteis, 4h): ver README "Robô do SPW". Segredos em secrets/spw.env (SPW_USUARIO, SPW_SENHA,
SPW_LOGIN_URL, SPW_CONSULTA_URL). Resultado de cada execução fica em robo_execucoes e aparece no Início.
playwright e xlrd são importados dentro das funções: os testes rodam sem eles.
"""
import hashlib
import io
import sqlite3
import sys
import time
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook

import config
import db

RAIZ = Path(__file__).resolve().parent
ARQUIVO_ENV = RAIZ / "secrets" / "spw.env"
CHAVES_ENV = ("SPW_USUARIO", "SPW_SENHA", "SPW_LOGIN_URL", "SPW_CONSULTA_URL")
NOME_ARQUIVO = "SPW automático"
PASTA_SPW = "spw"            # dentro de config.pasta_dados(): ultimo.xls e erro.png


class RoboErro(Exception):
    """Erro previsto do robô; a mensagem vai para robo_execucoes e para o Início."""


def ler_env(caminho: Path = ARQUIVO_ENV) -> dict:
    if not Path(caminho).exists():
        raise RoboErro(f"secrets/spw.env não encontrado ou incompleto ({caminho})")
    env = {}
    for linha in Path(caminho).read_text(encoding="utf-8").splitlines():
        if "=" in linha and not linha.lstrip().startswith("#"):
            chave, valor = linha.split("=", 1)
            env[chave.strip()] = valor.strip()
    faltando = [c for c in CHAVES_ENV if not env.get(c)]
    if faltando:
        raise RoboErro(f"secrets/spw.env não encontrado ou incompleto ({caminho}): falta " + ", ".join(faltando))
    return env


# o ramo numérico (repr(float(v))) é proposital: faz o float do xlrd e o int do openpyxl darem o
# mesmo hash; não trocar por db._texto (que formataria os dois como texto de forma diferente).
def _normalizar(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, (int, float)):
        return repr(float(v))
    return " ".join(str(v).split())


def hash_linhas(linhas: list[list]) -> str:
    """SHA-256 das linhas (cabeçalho incluído) com os valores normalizados como db._texto faria."""
    h = hashlib.sha256()
    for linha in linhas:
        h.update("\t".join(_normalizar(c) for c in linha).encode("utf-8"))
        h.update(b"\n")
    return h.hexdigest()


def linhas_para_xlsx(linhas: list[list]) -> io.BytesIO:
    """Planilha no formato que db.importar_bens aceita: cabeçalho na linha 1, uma aba."""
    wb = Workbook()
    ws = wb.active
    for linha in linhas:
        ws.append(list(linha))
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def executar(conn, baixar=None, agora=None) -> dict:
    """Fluxo completo de uma execução; nunca levanta: devolve {"resultado", "mensagem", "importacao_id"}.
    baixar() devolve as linhas do export (cabeçalho primeiro); os testes injetam listas prontas."""
    baixar = baixar or baixar_e_ler
    agora = agora or db._agora
    iniciado = agora()
    try:
        linhas = baixar()
        cabecalho = [db._texto(c) for c in (linhas[0] if linhas else [])]
        faltando = [c for c in db.COLUNAS_EXPORT if c not in cabecalho]
        if faltando:
            raise RoboErro("cabeçalho do SPW mudou: faltam " + ", ".join(faltando))
        atuais = conn.execute("SELECT count(*) FROM bens").fetchone()[0]
        if atuais and len(linhas) - 1 < atuais * 0.9:
            raise RoboErro(f"export do SPW veio curto: {len(linhas) - 1} bens contra {atuais} na base; "
                           "confira no SPW e use Atualizar base se a queda for real")
        h = hash_linhas(linhas)
        if h == db.ultimo_hash_robo(conn):
            mensagem = f"{len(linhas) - 1} bens no export, nada mudou"
            db.registrar_execucao_robo(conn, iniciado, "sem_mudanca", hash=h, mensagem=mensagem)
            return {"resultado": "sem_mudanca", "mensagem": mensagem, "importacao_id": None}
        r = db.importar_bens(conn, linhas_para_xlsx(linhas), nome_arquivo=NOME_ARQUIVO)
        mensagem = (f"{r['total']} bens ({r['ativos']} ativos): {r['novos']} novo(s), {r['removidos']} removido(s), "
                    f"{r['movidos']} movido(s), {r['situacao']} com situação alterada")
        db.registrar_execucao_robo(conn, iniciado, "importado", hash=h, importacao_id=r["importacao_id"], mensagem=mensagem)
        return {"resultado": "importado", "mensagem": mensagem, "importacao_id": r["importacao_id"]}
    except Exception as exc:
        mensagem = (str(exc) or type(exc).__name__)[:500]
        try:
            db.registrar_execucao_robo(conn, iniciado, "erro", mensagem=mensagem)
        except sqlite3.Error:
            pass                                    # banco travado: fica só no log do cron
        return {"resultado": "erro", "mensagem": mensagem, "importacao_id": None}


def baixar_export(env: dict, destino: Path) -> Path:
    """Login no SPW, abre a consulta de bens e exporta Excel/Detalhado em `destino`.
    Seletores provados em 2026-09-17 (docs/superpowers/notes/2026-09-17-robo-spw). Em erro salva erro.png ao lado."""
    from playwright.sync_api import sync_playwright
    P = "#ContentPlaceHolder1_ASPxRoundPanel1_"
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        navegador = p.chromium.launch()
        pagina = navegador.new_page(viewport={"width": 1280, "height": 900}, accept_downloads=True)
        try:
            pagina.goto(env["SPW_LOGIN_URL"], wait_until="networkidle", timeout=60000)
            pagina.fill(P + "txtUsuario_I", env["SPW_USUARIO"])
            pagina.click(P + "txtSenha_I_CLND")
            pagina.wait_for_timeout(300)
            pagina.fill(P + "txtSenha_I", env["SPW_SENHA"], force=True)
            with pagina.expect_navigation(wait_until="networkidle", timeout=60000):
                pagina.click(P + "btnEntrar")
            if "MenuChamador" not in pagina.url:
                raise RoboErro("login no SPW não chegou ao menu (usuário/senha?): " + pagina.url)
            pagina.goto(env["SPW_CONSULTA_URL"], wait_until="networkidle", timeout=60000)
            pagina.click("#ContentPlaceHolder1_ASPxButton1")
            pagina.wait_for_selector("#ContentPlaceHolder1_PCExportacao_cboArquivo", state="visible", timeout=30000)
            # o painel reinicializa os combos ~1,5s depois de ficar visível, sobrescrevendo qualquer seleção
            # feita antes disso (volta para "PDF"); esperar aqui e depois conferir que "Excel" pegou.
            pagina.wait_for_timeout(2500)
            pagina.select_option("#ContentPlaceHolder1_PCExportacao_cboArquivo", label="Excel")
            pagina.select_option("#ContentPlaceHolder1_PCExportacao_cboModeloExportacao", label="Detalhado")
            pagina.wait_for_timeout(500)
            if pagina.eval_on_selector("#ContentPlaceHolder1_PCExportacao_cboArquivo", "e => e.value") != "0":
                raise RoboErro("painel de exportação do SPW não aceitou 'Excel' (voltou para PDF)")
            with pagina.expect_download(timeout=300000) as dl:
                pagina.click("#ContentPlaceHolder1_PCExportacao_imgExportar")
            dl.value.save_as(destino)
        except Exception:
            try:
                pagina.screenshot(path=str(destino.parent / "erro.png"), full_page=True)
            except Exception:
                pass
            raise
        finally:
            navegador.close()
    if destino.stat().st_size < 1000:
        raise RoboErro(f"export do SPW veio vazio ({destino.stat().st_size} bytes)")
    return destino


def ler_xls(caminho: Path) -> list[list]:
    """Lê o .xls (BIFF) do SPW e devolve as linhas a partir do cabeçalho 'Número Bem'."""
    import xlrd
    wb = xlrd.open_workbook(str(caminho))
    ws = wb.sheet_by_index(0)
    linhas = []
    for i in range(ws.nrows):
        linha = []
        for c in ws.row(i):
            if c.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                linha.append(None)
            elif c.ctype == xlrd.XL_CELL_DATE:
                linha.append(xlrd.xldate_as_datetime(c.value, wb.datemode))
            elif c.ctype == xlrd.XL_CELL_NUMBER:
                linha.append(c.value)
            else:
                linha.append(str(c.value))
        linhas.append(linha)
    for i, linha in enumerate(linhas):
        if linha and db._texto(linha[0]) == "Número Bem":
            return linhas[i:]
    raise RoboErro("export do SPW sem a linha de cabeçalho 'Número Bem'")


def baixar_e_ler() -> list[list]:
    pasta = config.pasta_dados() / PASTA_SPW
    return ler_xls(baixar_export(ler_env(), pasta / "ultimo.xls"))


def main() -> int:
    inicio = time.monotonic()
    conn = db.conectar()
    try:
        iniciado = db._agora()
        try:
            db.criar_esquema(conn)           # idempotente; garante robo_execucoes mesmo antes do rebuild
            r = executar(conn)
        except Exception as exc:
            mensagem = (str(exc) or type(exc).__name__)[:500]
            try:
                db.registrar_execucao_robo(conn, iniciado, "erro", mensagem=mensagem)
            except sqlite3.Error:
                pass                          # robo_execucoes pode nem existir ainda: fica só no log do cron
            r = {"resultado": "erro", "mensagem": mensagem, "importacao_id": None}
    finally:
        conn.close()
    print(f"{db._agora()} {r['resultado']} {r['mensagem']} ({time.monotonic() - inicio:.0f}s)", flush=True)
    return 0 if r["resultado"] != "erro" else 1


if __name__ == "__main__":
    sys.exit(main())
