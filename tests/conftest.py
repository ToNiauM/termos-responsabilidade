"""Fixtures compartilhadas: banco SQLite temporário com dados de exemplo."""
import pytest


@pytest.fixture
def dados(tmp_path, monkeypatch):
    """Pasta de dados isolada + conexão com esquema criado. Nada de dados ainda."""
    monkeypatch.setenv("TERMOS_DADOS", str(tmp_path))
    import config
    import db
    config.preparar_pastas()
    conn = db.conectar()
    db.criar_esquema(conn)
    yield conn
    conn.close()


def semear(conn):
    """Cenário mínimo: 1 centro (CCI), 1 sala mapeada, 4 bens, 1 pessoa com 1 bem atribuído."""
    conn.execute("INSERT INTO responsaveis VALUES ('CCI','Prezada','JAQUELINE PORTELA','j@cfc.org.br','46','coordenadora')")
    conn.execute("INSERT INTO localizacoes VALUES ('01 - SALA CCI','CCI')")
    conn.executemany(
        "INSERT INTO bens VALUES (?,?,?,?,?,?,?,?,?)",
        [
            (1001, "ATIVO", "CADEIRA", "GIRATÓRIA", "MÓVEIS", "01 - SALA CCI", "31/12/1996", 75.94, 64.54),
            (1002, "ATIVO", "NOTEBOOK", "DELL", "EQUIPAMENTOS", "01 - SALA CCI", "06/12/2012", 3000.0, 1500.0),
            (1003, "BAIXADO", "MESA", "ANTIGA", "MÓVEIS", "01 - SALA CCI", "06/12/2012", 100.0, 10.0),
            (1004, "ATIVO", "ARMÁRIO", "AÇO", "MÓVEIS", "99 - SEM MAPA", "06/12/2012", 500.0, 250.5),
        ],
    )
    conn.execute("INSERT INTO pessoas VALUES ('ANA SILVA')")
    conn.execute("INSERT INTO atribuicoes VALUES ('ANA SILVA', 1002)")
    conn.commit()
