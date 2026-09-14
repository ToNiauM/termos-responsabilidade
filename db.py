"""Banco SQLite: esquema, importação de bens, consultas e cadastros.

Todas as funções recebem a conexão como primeiro argumento; quem abre e fecha é o chamador
(o Flask, por request; os testes, por fixture). Nenhuma função aqui usa Flask.
"""
import sqlite3
from pathlib import Path

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
