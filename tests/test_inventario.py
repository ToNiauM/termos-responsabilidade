"""Módulo de inventário: eventos, salas, leituras, sobras, relatório, xlsx."""
import pytest

import db
import inventario
from tests.conftest import semear


def semear_inventario(conn):
    """semear() + 2ª sala com 2 bens e 1 sala vazia de escopo; devolve o id do evento aberto."""
    semear(conn)
    conn.execute("INSERT INTO bens VALUES (2001,'ATIVO','MONITOR','LG','EQUIPAMENTOS','02 - SALA B','01/01/2020',900,800)")
    conn.execute("INSERT INTO bens VALUES (2002,'ATIVO','IMPRESSORA','HP','EQUIPAMENTOS','02 - SALA B','01/01/2020',1200,1000)")
    conn.commit()
    return inventario.abrir_evento(conn, "Inventário 2026", "Portaria 1/2026", ["Fulano", "Beltrana"])


def test_localizacoes_ativas(dados):
    semear(dados)
    assert db.localizacoes_ativas(dados) == ["01 - SALA CCI", "99 - SEM MAPA"]   # 1003 é BAIXADO, não muda nada


def test_abrir_evento_todas_as_salas_e_integrantes(dados):
    eid = semear_inventario(dados)
    e = inventario.evento(dados, eid)
    assert e["nome"] == "Inventário 2026" and e["encerrado_em"] is None and e["integrantes"] == ["Beltrana", "Fulano"]
    assert [s["localizacao"] for s in inventario.salas(dados, eid)] == ["01 - SALA CCI", "02 - SALA B", "99 - SEM MAPA"]
    assert inventario.evento_aberto(dados)["id"] == eid
    assert e["resumo"] == {"salas": 3, "salas_iniciadas": 0, "bens": 5, "lidos": 0, "divergentes": 0,
                           "pendentes": 5, "sobras": 0, "pct_bens": 0.0}


def test_abrir_evento_amostragem_e_validacoes(dados):
    semear(dados)
    with pytest.raises(db.ErroDeNegocio):
        inventario.abrir_evento(dados, "", "", ["A"])
    with pytest.raises(db.ErroDeNegocio):
        inventario.abrir_evento(dados, "X", "", [" ", ""])
    with pytest.raises(db.ErroDeNegocio):
        inventario.abrir_evento(dados, "X", "", ["A"], salas=["NÃO EXISTE"])
    eid = inventario.abrir_evento(dados, "Amostra", None, ["A", "A ", "b"], salas=["99 - SEM MAPA"])
    assert [s["localizacao"] for s in inventario.salas(dados, eid)] == ["99 - SEM MAPA"]
    assert inventario.evento(dados, eid)["integrantes"] == ["A", "b"]
    with pytest.raises(db.ErroDeNegocio):
        inventario.abrir_evento(dados, "Outro", "", ["A"])          # já há aberto
    inventario.encerrar_evento(dados, eid)
    assert inventario.evento_aberto(dados) is None
    e2 = inventario.abrir_evento(dados, "Outro", "", ["A"])
    assert [x["id"] for x in inventario.eventos(dados)] == [e2, eid]   # aberto primeiro


def test_salas_contadores_e_resumo(dados):
    eid = semear_inventario(dados)
    s = {x["localizacao"]: x for x in inventario.salas(dados, eid)}
    assert s["01 - SALA CCI"]["total"] == 2 and s["01 - SALA CCI"]["ccustos"] == "CCI" and s["02 - SALA B"]["ccustos"] is None
    assert (s["01 - SALA CCI"]["localizados"], s["01 - SALA CCI"]["pendentes"], s["01 - SALA CCI"]["divergentes"]) == (0, 2, 0)
    assert "concluida_em" not in s["01 - SALA CCI"]
    dados.execute("INSERT INTO inventario_leituras (evento_id, numero, localizacao, lido_em, integrante) VALUES (?,?,?,?,?)",
                  (eid, 2001, "02 - SALA B", "2026-09-15 10:00:00", "Fulano"))
    dados.commit()
    r = inventario.resumo(dados, eid)
    assert (r["salas_iniciadas"], r["lidos"], r["pendentes"], r["pct_bens"]) == (1, 1, 4, 20.0)
    inventario.encerrar_evento(dados, eid)
    assert inventario.evento(dados, eid)["encerrado_em"] is not None
    inventario.encerrar_evento(dados, eid)                                        # idempotente
    with pytest.raises(db.ErroDeNegocio):
        inventario.encerrar_evento(dados, 999)
