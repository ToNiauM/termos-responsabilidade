"""Laço de atendimento com executores falsos: FIFO, um por vez, erro não derruba, órfão volta à fila."""
import pytest

import atender_pedidos as ap
import db
from tests.conftest import semear


def _sei_pedido(conn, chave="CCI"):
    db.incluir_processo(conn, "ccusto", "T", "1111") if not db.processo_vigente(conn, "ccusto") else None
    t = db.registrar_emissao(conn, "ccusto", chave, db.bens_do_centro(conn, chave))
    t = db.preparar_envio_sei(conn, t["id"])
    return db.enfileirar_pedido(conn, "sei", termo_id=t["id"], html="<p>x</p>")


def test_atende_em_ordem_um_por_vez(dados):
    semear(dados)
    ordem = []
    a = _sei_pedido(dados)
    b = db.enfileirar_pedido(dados, "spw")

    def sei(conn, pedido):
        ordem.append(("sei", pedido["id"])); db.marcar_passo(conn, pedido["id"], "concluido", "ok")
        return {"passo": "concluido"}

    def spw(conn, pedido):
        ordem.append(("spw", pedido["id"])); db.marcar_passo(conn, pedido["id"], "concluido", "ok")
        return {"resultado": "sem_mudanca"}
    ex = {"sei": sei, "spw": spw}
    assert ap.atender_um(dados, ex)["id"] == a
    assert ap.atender_um(dados, ex)["id"] == b
    assert ap.atender_um(dados, ex) is None
    assert ordem == [("sei", a), ("spw", b)]


def test_excecao_do_executor_vira_erro_e_o_laco_segue(dados, tmp_path, monkeypatch):
    semear(dados)
    monkeypatch.setattr(ap, "ARQUIVO_LOG", tmp_path / "robo_pedidos.log")
    a = _sei_pedido(dados)
    b = db.enfileirar_pedido(dados, "spw")

    def explode(conn, pedido):
        raise RuntimeError("frame ausente")
    ex = {"sei": explode, "spw": lambda c, p: db.marcar_passo(c, p["id"], "concluido", "ok")}
    r = ap.atender_um(dados, ex)
    assert r["id"] == a and r["passo"] == "erro" and r["mensagem"] == "frame ausente"
    assert "RuntimeError" in (tmp_path / "robo_pedidos.log").read_text()
    assert ap.atender_um(dados, ex)["id"] == b


def test_laco_recoloca_orfaos_e_para_quando_mandado(dados):
    semear(dados)
    a = db.enfileirar_pedido(dados, "spw")
    db.marcar_passo(dados, a, "rodando")
    dados.execute("UPDATE robo_pedidos SET iniciado_em = '2020-01-01 00:00:00' WHERE id = ?", (a,)); dados.commit()
    voltas = []
    atendidos = []

    def spw(conn, pedido):
        atendidos.append(pedido["id"]); db.marcar_passo(conn, pedido["id"], "concluido", "ok")
    ap.laco(dados, intervalo=0, executores={"spw": spw, "sei": None}, continuar=lambda: len(voltas) < 3,
            dormir=lambda s: voltas.append(s))
    assert atendidos == [a] and voltas == [0, 0, 0]


def test_executar_spw_usa_importar_spw_e_grava_passos(dados, monkeypatch):
    semear(dados)
    import importar_spw
    monkeypatch.setattr(importar_spw, "executar", lambda conn, **k: {"resultado": "importado", "mensagem": "7 bens", "importacao_id": 1})
    pid = db.enfileirar_pedido(dados, "spw")
    r = ap.executar_spw(dados, db.pedido(dados, pid))
    p = db.pedido(dados, pid)
    assert r["resultado"] == "importado" and p["passo"] == "concluido" and p["mensagem"] == "7 bens"
    monkeypatch.setattr(importar_spw, "executar", lambda conn, **k: {"resultado": "erro", "mensagem": "SPW fora do ar", "importacao_id": None})
    pid = db.enfileirar_pedido(dados, "spw")
    ap.executar_spw(dados, db.pedido(dados, pid))
    assert db.pedido(dados, pid)["passo"] == "erro" and db.pedido(dados, pid)["mensagem"] == "SPW fora do ar"


def test_travar_e_exclusivo(tmp_path):
    caminho = tmp_path / "robo.lock"
    with ap.travar(caminho):
        import fcntl
        with open(caminho, "w") as f:
            with pytest.raises(BlockingIOError):
                fcntl.flock(f, fcntl.LOCK_EX | fcntl.LOCK_NB)
    with open(caminho, "w") as f:
        fcntl.flock(f, fcntl.LOCK_EX | fcntl.LOCK_NB)        # liberado depois do with
