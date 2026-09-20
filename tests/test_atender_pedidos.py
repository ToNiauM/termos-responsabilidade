"""Laço de atendimento com executores falsos: FIFO, um por vez, erro não derruba, órfão volta à fila."""
from datetime import datetime, timedelta

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


def test_laco_recoloca_orfao_surgido_durante_a_espera_ociosa(dados, monkeypatch):
    """O requeue do início do laço só roda uma vez: um pedido que só completa os 10 min depois (reinício
    rápido demais) tem que ser pego numa volta ociosa seguinte, não só no arranque."""
    semear(dados)
    pid = _sei_pedido(dados)
    db.marcar_passo(dados, pid, "rodando")
    limite = (datetime.now() - timedelta(minutes=11)).strftime("%Y-%m-%d %H:%M:%S")
    dados.execute("UPDATE robo_pedidos SET iniciado_em = ? WHERE id = ?", (limite, pid)); dados.commit()

    chamadas = []
    original = db.pedidos_orfaos_para_aguardando

    def requeue_so_na_segunda_chamada(conn, *a, **k):
        # simula o requeue do início do laço não pegando o órfão ainda (a folga de 10 min só se
        # completa depois); a chamada de verdade só acontece daqui pra frente.
        chamadas.append(1)
        return 0 if len(chamadas) == 1 else original(conn, *a, **k)
    monkeypatch.setattr(db, "pedidos_orfaos_para_aguardando", requeue_so_na_segunda_chamada)

    atendidos = []

    def sei(conn, pedido):
        atendidos.append(pedido["id"]); db.marcar_passo(conn, pedido["id"], "concluido", "ok")
    voltas = []
    ap.laco(dados, intervalo=0, executores={"sei": sei, "spw": None}, continuar=lambda: len(voltas) < 2,
            dormir=lambda s: voltas.append(s))
    assert atendidos == [pid]
    assert db.pedido(dados, pid)["passo"] == "concluido"


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


def test_executar_sei_monta_env_de_quem_pediu(dados, chave, tmp_path, monkeypatch):
    import robo_sei, segredos, usuarios
    semear(dados)
    (tmp_path / "sei.env").write_text("SEI_LOGIN_URL=https://sei.cfc.org.br/sei/\nSEI_ORGAO=CFC\n")
    monkeypatch.setattr(robo_sei, "ARQUIVO_ENV", tmp_path / "sei.env")
    uid = usuarios.criar(dados, "maria", "Maria", "Senha!234", ["operador"])
    usuarios.salvar_acesso_sei(dados, uid, "maria.silva", "S3nha", "GECONT")
    db.incluir_processo(dados, "ccusto", "T", "1111")
    t = db.preparar_envio_sei(dados, db.registrar_emissao(dados, "ccusto", "CCI", db.bens_do_centro(dados, "CCI"))["id"])
    pid = db.enfileirar_pedido(dados, "sei", termo_id=t["id"], html="<p>x</p>", criado_por="maria")
    recebido = {}
    def enviar(conn, pedido, env=None):
        recebido.update(env); db.marcar_passo(conn, pedido["id"], "concluido", "ok"); return {"passo": "concluido"}
    ap.executar_sei(dados, db.pedido(dados, pid), enviar=enviar)
    assert recebido == {"SEI_LOGIN_URL": "https://sei.cfc.org.br/sei/", "SEI_ORGAO": "CFC", "SEI_USUARIO": "maria.silva", "SEI_SENHA": "S3nha", "SEI_UNIDADE": "GECONT"}
    # sem credencial (acesso apagado depois de clicar)
    usuarios.apagar_acesso_sei(dados, uid)
    pid2 = db.enfileirar_pedido(dados, "sei", termo_id=t["id"], html="<p>x</p>", criado_por="maria")
    chamado = []
    ap.executar_sei(dados, db.pedido(dados, pid2), enviar=lambda *a, **k: chamado.append(1))
    p = db.pedido(dados, pid2)
    assert p["passo"] == "erro" and p["mensagem"] == "Quem pediu a emissão não tem acesso ao SEI cadastrado; cadastre em Meus acessos e emita de novo." and not chamado
    # sem chave do cofre; caminho com o mesmo nome "chaves.env" (mensagem cita o nome real, como test_cofre.py),
    # mas numa pasta que não existe: a fixture `chave` já criou tmp_path/chaves.env de verdade
    usuarios.salvar_acesso_sei(dados, uid, "maria.silva", "S3nha", "GECONT")
    import cofre
    monkeypatch.setattr(cofre, "ARQUIVO_CHAVE", tmp_path / "sem-chave" / "chaves.env")
    pid3 = db.enfileirar_pedido(dados, "sei", termo_id=t["id"], html="<p>x</p>", criado_por="maria")
    ap.executar_sei(dados, db.pedido(dados, pid3), enviar=lambda *a, **k: chamado.append(1))
    assert db.pedido(dados, pid3)["mensagem"].startswith("secrets/chaves.env não encontrado") and not chamado


def test_executar_spw_usa_credencial_de_quem_pediu(dados, chave, tmp_path, monkeypatch):
    import importar_spw, usuarios
    semear(dados)
    (tmp_path / "spw.env").write_text(
        "SPW_USUARIO=robo\nSPW_SENHA=robo123\nSPW_LOGIN_URL=https://spw.cfc.org.br/login\nSPW_CONSULTA_URL=https://spw.cfc.org.br/consulta\n")
    monkeypatch.setattr(importar_spw, "ARQUIVO_ENV", tmp_path / "spw.env")
    uid = usuarios.criar(dados, "maria", "Maria", "Senha!234", ["operador"])
    usuarios.salvar_acesso_spw(dados, uid, "maria.spw", "S3nha")
    pid = db.enfileirar_pedido(dados, "spw", criado_por="maria")
    recebido = {}
    def executar(conn, env=None):
        recebido.update(env or {})
        return {"resultado": "importado", "mensagem": "7 bens", "importacao_id": 1}
    ap.executar_spw(dados, db.pedido(dados, pid), executar=executar)
    assert recebido == {"SPW_USUARIO": "maria.spw", "SPW_SENHA": "S3nha",
                        "SPW_LOGIN_URL": "https://spw.cfc.org.br/login", "SPW_CONSULTA_URL": "https://spw.cfc.org.br/consulta"}
    # sem credencial (acesso apagado depois de clicar)
    usuarios.apagar_acesso_spw(dados, uid)
    pid2 = db.enfileirar_pedido(dados, "spw", criado_por="maria")
    chamado = []
    ap.executar_spw(dados, db.pedido(dados, pid2), executar=lambda *a, **k: chamado.append(1))
    p = db.pedido(dados, pid2)
    assert p["passo"] == "erro" and p["mensagem"] == "Quem pediu a atualização não tem acesso ao SPW cadastrado; cadastre em Meus acessos e peça de novo." and not chamado
    # sem criado_por (cron / scripts/atualizar_base.sh): repassa env=None; quem lê o spw.env inteiro é o
    # baixar_e_ler() de dentro de importar_spw.executar de verdade, não este módulo
    pid3 = db.enfileirar_pedido(dados, "spw")
    recebido2 = []
    def executar2(conn, env=None):
        recebido2.append(env)
        return {"resultado": "sem_mudanca", "mensagem": "nada mudou", "importacao_id": None}
    ap.executar_spw(dados, db.pedido(dados, pid3), executar=executar2)
    assert recebido2 == [None]
    assert db.pedido(dados, pid3)["passo"] == "concluido"


def test_executar_spw_erro_de_login_pelo_site_aponta_para_meus_acessos(dados, chave, tmp_path, monkeypatch):
    import importar_spw, usuarios
    semear(dados)
    (tmp_path / "spw.env").write_text(
        "SPW_USUARIO=robo\nSPW_SENHA=robo123\nSPW_LOGIN_URL=https://spw.cfc.org.br/login\nSPW_CONSULTA_URL=https://spw.cfc.org.br/consulta\n")
    monkeypatch.setattr(importar_spw, "ARQUIVO_ENV", tmp_path / "spw.env")
    uid = usuarios.criar(dados, "maria", "Maria", "Senha!234", ["operador"])
    usuarios.salvar_acesso_spw(dados, uid, "maria.spw", "senha-errada")

    def executar(conn, env=None):
        return {"resultado": "erro", "mensagem": "login no SPW não chegou ao menu (usuário/senha?): https://x", "importacao_id": None}

    # pela fila do site (criado_por): a mensagem ganha a dica de onde corrigir
    pid = db.enfileirar_pedido(dados, "spw", criado_por="maria")
    ap.executar_spw(dados, db.pedido(dados, pid), executar=executar)
    assert db.pedido(dados, pid)["mensagem"] == "login no SPW não chegou ao menu (usuário/senha?): https://x; atualize em Meus acessos"
    # sem criado_por (cron): mensagem sem alteração
    pid2 = db.enfileirar_pedido(dados, "spw")
    ap.executar_spw(dados, db.pedido(dados, pid2), executar=executar)
    assert db.pedido(dados, pid2)["mensagem"] == "login no SPW não chegou ao menu (usuário/senha?): https://x"


def test_travar_e_exclusivo(tmp_path):
    caminho = tmp_path / "robo.lock"
    with ap.travar(caminho):
        import fcntl
        with open(caminho, "w") as f:
            with pytest.raises(BlockingIOError):
                fcntl.flock(f, fcntl.LOCK_EX | fcntl.LOCK_NB)
    with open(caminho, "w") as f:
        fcntl.flock(f, fcntl.LOCK_EX | fcntl.LOCK_NB)        # liberado depois do with
