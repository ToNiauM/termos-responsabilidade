"""scripts/zerar_inventario.py (a parte que roda no container): copia as fotos, apaga no bucket e depois no banco.
Bucket falso: nenhum teste fala com o R2."""
import importlib.util
from pathlib import Path

import pytest

import db
import inventario
from tests.test_inventario import semear_inventario

RAIZ = Path(__file__).resolve().parents[1]
_spec = importlib.util.spec_from_file_location("zerar_inventario", RAIZ / "scripts" / "zerar_inventario.py")
zi = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(zi)

BASE = "https://fotos.teste"


class ErroCliente(Exception):
    """Imita botocore ClientError só no que o script lê (response['Error']['Code'])."""
    def __init__(self, codigo):
        super().__init__(codigo)
        self.response = {"Error": {"Code": codigo}}


class BucketFalso:
    def __init__(self, objetos, falha_get=None, falha_delete=False):
        self.objetos, self.falha_get, self.falha_delete, self.apagados = dict(objetos), falha_get, falha_delete, []

    def get_object(self, Bucket, Key):
        if self.falha_get:
            raise ErroCliente(self.falha_get)
        if Key not in self.objetos:
            raise ErroCliente("NoSuchKey")
        corpo = self.objetos[Key]
        return {"Body": type("B", (), {"read": lambda self_: corpo})()}

    def delete_object(self, Bucket, Key):
        if self.falha_delete:
            raise RuntimeError("R2 fora do ar")
        self.apagados.append(Key)
        self.objetos.pop(Key, None)


class ConexaoSemFechar:
    """O script fecha a conexão no fim; a do teste precisa continuar aberta."""
    def __init__(self, conn):
        self._conn = conn

    def __getattr__(self, nome):
        return getattr(self._conn, nome)

    def close(self):
        pass


@pytest.fixture
def cenario(dados, monkeypatch):
    for v, valor in {"R2_ACCESS_KEY_ID": "x", "R2_SECRET_ACCESS_KEY": "x", "R2_ENDPOINT_URL": "https://r2.teste",
                     "R2_BUCKET_NAME": "fotos", "R2_PUBLIC_URL": BASE}.items():
        monkeypatch.setenv(v, valor)
    monkeypatch.setattr(zi, "ClientError", ErroCliente, raising=False)
    import botocore.exceptions
    monkeypatch.setattr(botocore.exceptions, "ClientError", ErroCliente)
    monkeypatch.setattr(db, "conectar", lambda *a, **k: ConexaoSemFechar(dados))
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    for n in (1, 2):
        inventario.adicionar_foto(dados, eid, 1001, lambda chave: f"{BASE}/{chave}")
    chaves = [u[len(BASE) + 1:] for u in inventario.urls_das_fotos(dados, eid)]
    assert len(chaves) == 2

    def bucket(**kw):
        b = BucketFalso({c: f"foto {c}".encode() for c in chaves}, **kw)
        monkeypatch.setattr(zi.fotos, "_cliente", lambda: b)
        return b
    return dados, eid, chaves, bucket


def _existe(conn, eid):
    return conn.execute("SELECT count(*) FROM inventario_eventos WHERE id = ?", (eid,)).fetchone()[0] == 1


def test_copia_apaga_no_bucket_e_depois_no_banco(cenario, tmp_path):
    conn, eid, chaves, bucket = cenario
    b = bucket()
    assert zi.main(["--copia", str(tmp_path), str(eid)]) == 0
    copiados = sorted(p.name for p in tmp_path.rglob("*.webp"))
    assert copiados == sorted(c.replace("/", "__") for c in chaves)
    assert sorted(b.apagados) == sorted(chaves) and not b.objetos
    assert not _existe(conn, eid)
    assert conn.execute("SELECT count(*) FROM inventario_leituras WHERE evento_id = ?", (eid,)).fetchone()[0] == 0


def test_falha_ao_copiar_nao_apaga_nada(cenario, tmp_path):
    conn, eid, chaves, bucket = cenario
    b = bucket(falha_get="AccessDenied")
    with pytest.raises(ErroCliente):
        zi.main(["--copia", str(tmp_path), str(eid)])
    assert b.apagados == [] and _existe(conn, eid)


def test_foto_que_ja_sumiu_do_bucket_nao_trava(cenario, tmp_path):
    conn, eid, chaves, bucket = cenario
    b = bucket()
    del b.objetos[chaves[0]]
    assert zi.main(["--copia", str(tmp_path), str(eid)]) == 0
    assert len(list(tmp_path.rglob("*.webp"))) == 1 and not _existe(conn, eid)


def test_bucket_falha_ao_apagar_banco_fica_intacto(cenario, tmp_path):
    conn, eid, chaves, bucket = cenario
    bucket(falha_delete=True)
    assert zi.main([str(eid)]) == 1
    assert _existe(conn, eid)
    assert len(inventario.urls_das_fotos(conn, eid)) == 2


def test_manter_fotos_apaga_so_o_banco(cenario):
    conn, eid, chaves, bucket = cenario
    b = bucket()
    assert zi.main(["--manter-fotos", str(eid)]) == 0
    assert b.apagados == [] and len(b.objetos) == 2 and not _existe(conn, eid)


def test_sem_bucket_configurado_recusa_apagar_fotos(cenario, monkeypatch):
    conn, eid, chaves, bucket = cenario
    monkeypatch.delenv("R2_BUCKET_NAME")
    assert zi.main([str(eid)]) == 1 and _existe(conn, eid)
