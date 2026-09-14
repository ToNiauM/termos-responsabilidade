import sqlite3

import pytest

import db
from tests.conftest import semear


def test_esquema_cria_cinco_tabelas(dados):
    nomes = {r["name"] for r in dados.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    assert {"bens", "responsaveis", "localizacoes", "pessoas", "atribuicoes"} <= nomes


def test_esquema_e_idempotente(dados):
    db.criar_esquema(dados)  # segunda vez não pode falhar


def test_foreign_keys_ligadas(dados):
    semear(dados)
    with pytest.raises(sqlite3.IntegrityError):
        dados.execute("INSERT INTO localizacoes VALUES ('02 - X', 'NAO_EXISTE')")
