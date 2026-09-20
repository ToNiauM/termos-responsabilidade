"""Trabalhador do host: atende a fila robo_pedidos (emissão no SEI e atualização com o SPW), um pedido por vez.

    .venv-robo/bin/python atender_pedidos.py

Serviço systemd em ops/termos-robo.service (ver README). Lock dados/robo.lock compartilhado com
atualizar_base.sh (cron do SPW). Log de exceções em dados/robo_pedidos.log. Sem este processo o site
funciona: os pedidos ficam "aguardando a vez" e a tela avisa depois de 2 minutos.
"""
import fcntl
import sys
import time
import traceback
from contextlib import contextmanager
from pathlib import Path

import config
import db
import importar_spw
import robo_sei

INTERVALO_S = 3.0
ARQUIVO_LOG = None  # None = usa config.pasta_dados() no momento do registro (respeita TERMOS_DADOS nos testes);
                    # os testes podem sobrescrever com monkeypatch.setattr(ap, "ARQUIVO_LOG", caminho)
ARQUIVO_LOCK = config.pasta_dados() / "robo.lock"


@contextmanager
def travar(caminho: Path = ARQUIVO_LOCK):
    """flock exclusivo (bloqueante) enquanto um pedido é atendido; o cron do SPW espera por ele."""
    Path(caminho).parent.mkdir(parents=True, exist_ok=True)
    with open(caminho, "w") as f:
        fcntl.flock(f, fcntl.LOCK_EX)
        try:
            yield
        finally:
            fcntl.flock(f, fcntl.LOCK_UN)


def _registrar(texto: str) -> None:
    caminho = ARQUIVO_LOG or (config.pasta_dados() / "robo_pedidos.log")
    try:
        Path(caminho).parent.mkdir(parents=True, exist_ok=True)
        with open(caminho, "a", encoding="utf-8") as f:
            f.write(f"{db._agora()} {texto}\n")
    except OSError:
        pass


def executar_spw(conn, pedido: dict) -> dict:
    db.marcar_passo(conn, pedido["id"], "rodando")
    r = importar_spw.executar(conn)
    db.marcar_passo(conn, pedido["id"], "erro" if r["resultado"] == "erro" else "concluido", r["mensagem"])
    return r


EXECUTORES = {"sei": robo_sei.enviar_termo, "spw": executar_spw}


def atender_um(conn, executores: dict | None = None, lock: Path | None = None) -> dict | None:
    """Pega o pedido mais antigo em aguardando e executa; devolve o pedido como ficou (ou None se a fila está vazia)."""
    executores = executores or EXECUTORES
    p = db.proximo_pedido(conn)
    if not p:
        return None
    cm = travar(lock) if lock else _sem_lock()
    with cm:
        try:
            executores[p["tipo"]](conn, p)
        except Exception as exc:
            _registrar(f"pedido {p['id']} ({p['tipo']}): {traceback.format_exc()}")
            db.marcar_passo(conn, p["id"], "erro", (str(exc) or type(exc).__name__)[:500])
    return db.pedido(conn, p["id"])


@contextmanager
def _sem_lock():
    yield


def laco(conn, intervalo: float = INTERVALO_S, executores: dict | None = None, continuar=None,
         dormir=time.sleep, lock: Path | None = None) -> None:
    continuar = continuar or (lambda: True)
    reenfileirados = db.pedidos_orfaos_para_aguardando(conn)
    if reenfileirados:
        _registrar(f"{reenfileirados} pedido(s) órfão(s) voltaram à fila")
    while continuar():
        if atender_um(conn, executores, lock) is None:
            dormir(intervalo)


def main() -> int:
    conn = db.conectar()
    db.criar_esquema(conn)
    _registrar("trabalhador iniciado")
    try:
        laco(conn, lock=ARQUIVO_LOCK)
    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
