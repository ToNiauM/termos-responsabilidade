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


def test_ler_localizado_divergente_reler_e_erros(dados):
    eid = semear_inventario(dados)
    r = inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    assert r["situacao"] == "localizado" and r["reler"] is False and r["ativo"] and r["bem"]["descricao"] == "CADEIRA"
    r = inventario.ler(dados, eid, "01 - SALA CCI", 2001, "Fulano")             # MONITOR é da SALA B
    assert r["situacao"] == "divergente" and r["cadastrado_em"] == "02 - SALA B"
    r = inventario.ler(dados, eid, "02 - SALA B", 2001, "Beltrana")             # reler: atualiza a mesma linha
    assert r["situacao"] == "localizado" and r["reler"] and r["leitura_anterior"]["localizacao"] == "01 - SALA CCI"
    assert dados.execute("SELECT count(*) FROM inventario_leituras WHERE evento_id = ?", (eid,)).fetchone()[0] == 2
    r = inventario.ler(dados, eid, "01 - SALA CCI", 1003, "Fulano")             # BAIXADO: registra, avisa
    assert r["ativo"] is False and r["situacao"] == "localizado"
    with pytest.raises(inventario.BemNaoEncontrado) as e:
        inventario.ler(dados, eid, "01 - SALA CCI", 99999, "Fulano")
    assert e.value.numero == 99999
    with pytest.raises(db.ErroDeNegocio):
        inventario.ler(dados, eid, "SALA QUE NÃO EXISTE", 1001, "Fulano")
    with pytest.raises(db.ErroDeNegocio):
        inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Ninguém")
    s = {x["localizacao"]: x for x in inventario.salas(dados, eid)}
    assert (s["01 - SALA CCI"]["localizados"], s["01 - SALA CCI"]["pendentes"]) == (1, 1)
    assert (s["02 - SALA B"]["localizados"], s["02 - SALA B"]["divergentes"]) == (1, 0)
    assert s["01 - SALA CCI"]["divergentes"] == 0

    inventario.ler(dados, eid, "02 - SALA B", 1003, "Fulano")                   # BAIXADO relido em outra sala
    s = {x["localizacao"]: x for x in inventario.salas(dados, eid)}
    assert s["02 - SALA B"]["divergentes"] == 0                                # BAIXADO não conta nem aqui

    inventario.ler(dados, eid, "02 - SALA B", 1001, "Fulano")                  # controle: ATIVO de outra sala conta
    s = {x["localizacao"]: x for x in inventario.salas(dados, eid)}
    assert s["02 - SALA B"]["divergentes"] == 1
    assert (s["01 - SALA CCI"]["localizados"], s["01 - SALA CCI"]["pendentes"]) == (0, 2)   # 1001 saiu da CCI
    assert s["02 - SALA B"]["localizados"] == 1                                             # 2001 continua localizado aqui

    inventario.encerrar_evento(dados, eid)
    with pytest.raises(db.ErroDeNegocio):
        inventario.ler(dados, eid, "01 - SALA CCI", 1002, "Fulano")


def test_bens_da_sala_e_atualizar_leitura(dados):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.ler(dados, eid, "01 - SALA CCI", 2002, "Fulano")                 # trazido da SALA B
    inventario.ler(dados, eid, "02 - SALA B", 1002, "Fulano")                   # bem da CCI lido na SALA B
    inventario.ler(dados, eid, "01 - SALA CCI", 1003, "Fulano")                 # BAIXADO da própria sala
    d = inventario.bens_da_sala(dados, eid, "01 - SALA CCI")
    por = {b["numero"]: b for b in d["bens"]}
    assert set(por) == {1001, 1002}                                              # só ativos da sala
    assert por[1001]["situacao_inv"] == "localizado" and por[1002]["situacao_inv"] == "divergente" and por[1002]["lido_em_sala"] == "02 - SALA B"
    assert [t["numero"] for t in d["trazidos"]] == [1003, 2002] and d["sobras"] == []
    inventario.atualizar_leitura(dados, eid, 1001, conservacao="Ruim", quem_usa="Ciclana", observacao="pé quebrado")
    b = {x["numero"]: x for x in inventario.bens_da_sala(dados, eid, "01 - SALA CCI")["bens"]}[1001]
    assert (b["conservacao"], b["quem_usa"], b["observacao"]) == ("Ruim", "Ciclana", "pé quebrado")
    inventario.atualizar_leitura(dados, eid, 1001, conservacao="")                # limpa
    assert {x["numero"]: x for x in inventario.bens_da_sala(dados, eid, "01 - SALA CCI")["bens"]}[1001]["conservacao"] is None
    with pytest.raises(db.ErroDeNegocio):
        inventario.atualizar_leitura(dados, eid, 1001, conservacao="Ótimo")
    with pytest.raises(db.ErroDeNegocio):
        inventario.atualizar_leitura(dados, eid, 2001, quem_usa="x")             # sem leitura


def test_sobras(dados):
    eid = semear_inventario(dados)
    with pytest.raises(db.ErroDeNegocio):
        inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "", "", "achado", "http://x/1.webp", "Fulano")
    with pytest.raises(db.ErroDeNegocio):
        inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "VENTILADOR", "", "", "http://x/1.webp", "Fulano")
    with pytest.raises(db.ErroDeNegocio):
        inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "VENTILADOR", "", "achado", "", "Fulano")
    with pytest.raises(db.ErroDeNegocio):
        inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "VENTILADOR", "", "achado", "", "Ninguém", exigir_foto=False)
    sid = inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "VENTILADOR", "", "achado", "", "Fulano", exigir_foto=False)
    inventario.definir_foto_sobra(dados, sid, "http://x/1.webp")
    s = inventario.bens_da_sala(dados, eid, "01 - SALA CCI")["sobras"]
    assert len(s) == 1 and s[0]["foto_url"] == "http://x/1.webp" and inventario.resumo(dados, eid)["sobras"] == 1
    with pytest.raises(db.ErroDeNegocio):
        inventario.excluir_sobra(dados, eid, 999)
    assert inventario.excluir_sobra(dados, eid, sid)["descricao"] == "VENTILADOR"
    assert inventario.resumo(dados, eid)["sobras"] == 0
    sid2 = inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "TABLET", "", "achado", "", "Fulano", exigir_foto=False)
    inventario.encerrar_evento(dados, eid)
    with pytest.raises(db.ErroDeNegocio):
        inventario.definir_foto_sobra(dados, sid2, "http://x/2.webp")
