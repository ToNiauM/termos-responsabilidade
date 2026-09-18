"""Robô do SPW: entra no sistema de patrimônio, exporta a relação de bens e importa em termos.db.

Roda no host, fora do container, com o venv que tem Playwright + xlrd:

    .venv-robo/bin/python importar_spw.py

Cron (dias úteis, 4h): ver README "Robô do SPW". Segredos em secrets/spw.env (SPW_USUARIO, SPW_SENHA,
SPW_LOGIN_URL, SPW_CONSULTA_URL). Resultado de cada execução fica em robo_execucoes e aparece no Início.
playwright e xlrd são importados dentro das funções: os testes rodam sem eles.
"""
import hashlib
import io
import sys
import time
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook

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
        raise RoboErro("spw.env não encontrado ou incompleto")
    env = {}
    for linha in Path(caminho).read_text().splitlines():
        if "=" in linha and not linha.lstrip().startswith("#"):
            chave, valor = linha.split("=", 1)
            env[chave.strip()] = valor.strip()
    faltando = [c for c in CHAVES_ENV if not env.get(c)]
    if faltando:
        raise RoboErro("spw.env não encontrado ou incompleto: falta " + ", ".join(faltando))
    return env


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
        except Exception:
            pass                                    # banco travado: fica só no log do cron
        return {"resultado": "erro", "mensagem": mensagem, "importacao_id": None}


def baixar_export(env: dict, destino: Path) -> Path:
    raise NotImplementedError        # Tarefa 4


def ler_xls(caminho: Path) -> list[list]:
    raise NotImplementedError        # Tarefa 4


def baixar_e_ler() -> list[list]:
    raise NotImplementedError        # Tarefa 4
