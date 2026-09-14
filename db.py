"""Banco SQLite: esquema, importação de bens, consultas e cadastros.

Todas as funções recebem a conexão como primeiro argumento; quem abre e fecha é o chamador
(o Flask, por request; os testes, por fixture). Nenhuma função aqui usa Flask.
"""
import sqlite3
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

import config

ESQUEMA = """
CREATE TABLE IF NOT EXISTS bens (
  numero        INTEGER PRIMARY KEY,
  situacao      TEXT NOT NULL,
  descricao     TEXT NOT NULL,
  complemento   TEXT,
  classificacao TEXT,
  localizacao   TEXT,
  data_entrada  TEXT,
  valor_compra  REAL,
  valor_atual   REAL
);
CREATE TABLE IF NOT EXISTS responsaveis (
  ccustos     TEXT PRIMARY KEY,
  tratamento  TEXT,
  responsavel TEXT NOT NULL,
  email       TEXT,
  matricula   TEXT,
  funcao      TEXT
);
CREATE TABLE IF NOT EXISTS localizacoes (
  localizacao TEXT PRIMARY KEY,
  ccustos     TEXT NOT NULL REFERENCES responsaveis(ccustos) ON UPDATE CASCADE
);
CREATE TABLE IF NOT EXISTS pessoas (
  nome TEXT PRIMARY KEY
);
CREATE TABLE IF NOT EXISTS atribuicoes (
  nome   TEXT    NOT NULL REFERENCES pessoas(nome) ON UPDATE CASCADE ON DELETE CASCADE,
  numero INTEGER NOT NULL REFERENCES bens(numero) DEFERRABLE INITIALLY DEFERRED,
  PRIMARY KEY (nome, numero)
);
"""


class ErroDeNegocio(Exception):
    """Erro que vira mensagem para o usuário (flash), não traceback."""


def conectar(caminho: Path | None = None) -> sqlite3.Connection:
    conn = sqlite3.connect(str(caminho or config.caminho_db()))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def criar_esquema(conn: sqlite3.Connection) -> None:
    conn.executescript(ESQUEMA)


def inicializar() -> None:
    """Primeira execução do programa: pastas + esquema vazio."""
    config.preparar_pastas()
    conn = conectar()
    try:
        criar_esquema(conn)
    finally:
        conn.close()


COLUNAS_EXPORT = {
    "Número Bem": "numero",
    "Situação": "situacao",
    "Descrição": "descricao",
    "Complemento": "complemento",
    "Classificação Contábil": "classificacao",
    "Localização": "localizacao",
    "Data Entrada": "data_entrada",
    "Valor Compra": "valor_compra",
    "Valor Atual": "valor_atual",
}


class ImportacaoInvalida(ErroDeNegocio):
    pass


def _texto(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    return " ".join(str(v).split())


def _numero(v):
    try:
        return float(v) if v not in (None, "") else None
    except (TypeError, ValueError):
        return None


def _aba_do_export(wb):
    for ws in wb.worksheets:
        cabecalho = [_texto(c) for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())]
        if "Número Bem" in cabecalho:
            return ws, cabecalho
    raise ImportacaoInvalida("Nenhuma aba com a coluna 'Número Bem'. Envie o export do sistema de patrimônio.")


def importar_bens(conn: sqlite3.Connection, arquivo) -> dict:
    """Substitui a tabela `bens` pelo conteúdo do export. Tudo ou nada.

    Devolve {"total", "ativos", "sem_centro"}. Levanta ImportacaoInvalida (e não altera nada)
    se faltar coluna ou se algum bem atribuído a pessoa deixar de existir.
    """
    try:
        wb = load_workbook(arquivo, read_only=True, data_only=True)
    except Exception:
        raise ImportacaoInvalida("Arquivo inválido: envie o export do sistema em .xlsx.")

    ws, cabecalho = _aba_do_export(wb)
    faltando = [c for c in COLUNAS_EXPORT if c not in cabecalho]
    if faltando:
        raise ImportacaoInvalida("Colunas ausentes no export: " + ", ".join(faltando))
    indice = {campo: cabecalho.index(col) for col, campo in COLUNAS_EXPORT.items()}

    linhas = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        num = _numero(r[indice["numero"]])
        if num is None:
            continue
        linhas.append((
            int(num), _texto(r[indice["situacao"]]), _texto(r[indice["descricao"]]),
            _texto(r[indice["complemento"]]), _texto(r[indice["classificacao"]]),
            _texto(r[indice["localizacao"]]), _texto(r[indice["data_entrada"]]),
            _numero(r[indice["valor_compra"]]), _numero(r[indice["valor_atual"]]),
        ))
    wb.close()

    try:
        conn.execute("DELETE FROM bens")
        conn.executemany("INSERT INTO bens VALUES (?,?,?,?,?,?,?,?,?)", linhas)
        orfaos = [str(r[0]) for r in conn.execute(
            "SELECT numero FROM atribuicoes WHERE numero NOT IN (SELECT numero FROM bens) ORDER BY numero")]
        if orfaos:
            raise ImportacaoInvalida(
                "O export não traz bens que estão atribuídos a pessoas: " + ", ".join(orfaos)
                + ". Remova a atribuição na aba Pessoas ou use um export completo.")
        conn.commit()
    except sqlite3.IntegrityError:
        conn.rollback()
        raise ImportacaoInvalida("O export tem número de bem repetido; corrija a planilha e envie de novo.")
    except Exception:
        conn.rollback()
        raise

    total = conn.execute("SELECT count(*) FROM bens").fetchone()[0]
    ativos = conn.execute("SELECT count(*) FROM bens WHERE situacao='ATIVO'").fetchone()[0]
    return {"total": total, "ativos": ativos, "sem_centro": localizacoes_sem_centro(conn)}


def localizacoes_sem_centro(conn: sqlite3.Connection) -> list[str]:
    """Localizações de bens ATIVOS que não têm centro de custo mapeado."""
    return [r[0] for r in conn.execute(
        "SELECT DISTINCT localizacao FROM bens WHERE situacao='ATIVO' AND localizacao <> '' "
        "AND localizacao NOT IN (SELECT localizacao FROM localizacoes) ORDER BY localizacao")]
