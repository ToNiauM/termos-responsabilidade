"""Trabalhador da fila robo_pedidos (emissão no SEI e atualização com o SPW), um pedido por vez.

Roda dentro do container `robo` do compose (alvo `robo` do Dockerfile, com Playwright):

    docker compose up -d robo

Lock dados/robo.lock compartilhado com atualizar_base.sh (mesmo arquivo em host e container, pelo volume
./dados:/app/dados). Log de exceções em dados/robo_pedidos.log. Sem este processo o site funciona: os
pedidos ficam "aguardando a vez" e a tela avisa depois de 2 minutos.
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
import segredos
import usuarios

INTERVALO_S = 3.0
# None = usa config.pasta_dados() no momento do uso (respeita TERMOS_DADOS nos testes); os testes
# podem sobrescrever com monkeypatch.setattr(ap, "ARQUIVO_LOG"/"ARQUIVO_LOCK", caminho)
ARQUIVO_LOG = None
ARQUIVO_LOCK = None


def _caminho_lock() -> Path:
    return ARQUIVO_LOCK or (config.pasta_dados() / "robo.lock")


def _caminho_log() -> Path:
    return ARQUIVO_LOG or (config.pasta_dados() / "robo_pedidos.log")


@contextmanager
def travar(caminho: Path):
    """flock exclusivo (bloqueante) enquanto um pedido é atendido; o cron do SPW espera por ele."""
    Path(caminho).parent.mkdir(parents=True, exist_ok=True)
    with open(caminho, "w") as f:
        fcntl.flock(f, fcntl.LOCK_EX)
        try:
            yield
        finally:
            fcntl.flock(f, fcntl.LOCK_UN)


def _registrar(texto: str) -> None:
    caminho = _caminho_log()
    try:
        Path(caminho).parent.mkdir(parents=True, exist_ok=True)
        with open(caminho, "a", encoding="utf-8") as f:
            f.write(f"{db._agora()} {texto}\n")
    except OSError:
        pass


def executar_sei(conn, pedido: dict, enviar=None) -> dict:
    """Monta a credencial de quem pediu (usuarios.credencial_sei) + URL/órgão do sei.env e chama o robô."""
    enviar = enviar or robo_sei.enviar_termo
    try:
        env = segredos.ler_env(robo_sei.ARQUIVO_ENV, robo_sei.CHAVES_ENV)
        credencial = usuarios.credencial_sei(conn, pedido.get("criado_por"))
    except (segredos.SegredoAusente, db.ErroDeNegocio) as e:
        db.marcar_passo(conn, pedido["id"], "erro", str(e)[:500])
        return {"passo": "erro", "mensagem": str(e)}
    if not credencial:
        mensagem = "Quem pediu a emissão não tem acesso ao SEI cadastrado; cadastre em Meus acessos e emita de novo."
        db.marcar_passo(conn, pedido["id"], "erro", mensagem)
        return {"passo": "erro", "mensagem": mensagem}
    return enviar(conn, pedido, env={**env, **credencial})


def executar_spw(conn, pedido: dict, executar=None) -> dict:
    """Sem `criado_por` (cron / ./atualizar_base.sh) usa o spw.env inteiro, como hoje. Pela fila do site,
    troca SPW_USUARIO/SPW_SENHA pela credencial de quem pediu (usuarios.credencial_spw), mantendo as
    URLs de spw.env."""
    executar = executar or importar_spw.executar
    db.marcar_passo(conn, pedido["id"], "rodando")
    criado_por = pedido.get("criado_por")
    if not criado_por:
        r = executar(conn, env=None)
    else:
        try:
            env_arquivo = importar_spw.ler_env()
            credencial = usuarios.credencial_spw(conn, criado_por)
        except (segredos.SegredoAusente, db.ErroDeNegocio, importar_spw.RoboErro) as e:
            mensagem = str(e)
            db.marcar_passo(conn, pedido["id"], "erro", mensagem[:500])
            return {"resultado": "erro", "mensagem": mensagem}
        if not credencial:
            mensagem = "Quem pediu a atualização não tem acesso ao SPW cadastrado; cadastre em Meus acessos e peça de novo."
            db.marcar_passo(conn, pedido["id"], "erro", mensagem)
            return {"resultado": "erro", "mensagem": mensagem}
        r = executar(conn, env={**env_arquivo, **credencial})
    if criado_por and r["resultado"] == "erro" and importar_spw.MSG_LOGIN_RECUSADO in r["mensagem"]:
        r["mensagem"] = r["mensagem"] + "; atualize em Meus acessos"
    db.marcar_passo(conn, pedido["id"], "erro" if r["resultado"] == "erro" else "concluido", r["mensagem"])
    return r


EXECUTORES = {"sei": executar_sei, "spw": executar_spw}


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
            # Fila ociosa: este trabalhador não tem nada em andamento, então qualquer pedido preso em
            # passo intermediário há mais de 10 min é órfão de verdade (reinício rápido demais para o
            # requeue do início pegar). O prazo ainda protege uma segunda instância rodando à mão.
            reenfileirados = db.pedidos_orfaos_para_aguardando(conn)
            if reenfileirados:
                _registrar(f"{reenfileirados} pedido(s) órfão(s) voltaram à fila")
            dormir(intervalo)


def main() -> int:
    conn = db.conectar()
    db.criar_esquema(conn)
    _registrar("trabalhador iniciado")
    try:
        laco(conn, lock=_caminho_lock())
    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
