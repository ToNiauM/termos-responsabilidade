"""fotos.py: validação, compressão WebP e envio ao R2 (cliente falso; nada de rede)."""
import io

import pytest
from PIL import Image

import db
import fotos


class Arquivo:
    def __init__(self, nome, dados):
        self.filename, self._dados = nome, dados

    def read(self):
        return self._dados


def png(largura=100, altura=50):
    buf = io.BytesIO()
    Image.new("RGB", (largura, altura), (200, 10, 10)).save(buf, "PNG")
    return buf.getvalue()


def test_validar(monkeypatch):
    assert fotos.validar(Arquivo("a.PNG", png())) == png()
    with pytest.raises(db.ErroDeNegocio):
        fotos.validar(Arquivo("a.gif", png()))
    with pytest.raises(db.ErroDeNegocio):
        fotos.validar(Arquivo("a.jpg", b"isto nao e imagem"))
    monkeypatch.setattr(fotos, "TAMANHO_MAX", 10)
    with pytest.raises(db.ErroDeNegocio):
        fotos.validar(Arquivo("a.png", png()))


def test_comprimir_webp_e_redimensiona():
    saida = fotos.comprimir(png(4000, 1000))
    img = Image.open(io.BytesIO(saida))
    assert img.format == "WEBP" and img.size == (1920, 480)
    assert Image.open(io.BytesIO(fotos.comprimir(png()))).size == (100, 50)


class ClienteFalso:
    def __init__(self):
        self.enviados, self.apagados = [], []

    def put_object(self, Bucket, Key, Body, ContentType):
        self.enviados.append((Bucket, Key, len(Body), ContentType))

    def delete_object(self, Bucket, Key):
        self.apagados.append((Bucket, Key))


def test_pasta_e_chaves():
    assert fotos.pasta("Inventário 2026", 1) == "inventario2026"
    assert fotos.pasta("  INVENTÁRIO – Sede / Anexo 2 ", 1) == "inventariosedeanexo2"
    assert fotos.pasta("Ação Çedilha", 1) == "acaocedilha"
    assert fotos.pasta("???", 7) == "evento7"
    assert fotos.pasta("", 7) == "evento7"
    assert fotos.chave_bem("inventario2026", 1, 12334) == "inventario2026/1-12334.webp"
    assert fotos.chave_sobra("inventario2026", 7) == "inventario2026/sobra-7.webp"


def test_enviar_apagar_e_configuracao(monkeypatch):
    for v in fotos.VARIAVEIS:
        monkeypatch.delenv(v, raising=False)
    assert fotos.configurado() is False
    with pytest.raises(db.ErroDeNegocio):
        fotos.enviar("x.webp", b"...")
    monkeypatch.setenv("R2_ACCESS_KEY_ID", "k"); monkeypatch.setenv("R2_SECRET_ACCESS_KEY", "s")
    monkeypatch.setenv("R2_ENDPOINT_URL", "https://acc.r2.cloudflarestorage.com"); monkeypatch.setenv("R2_BUCKET_NAME", "fotos")
    monkeypatch.setenv("R2_PUBLIC_URL", "https://fotos.exemplo.org/")
    falso = ClienteFalso()
    monkeypatch.setattr(fotos, "_cliente", lambda: falso)
    assert fotos.configurado()
    url = fotos.enviar("inventario2026/1-1001.webp", b"webp")
    assert url == "https://fotos.exemplo.org/inventario2026/1-1001.webp"
    assert falso.enviados == [("fotos", "inventario2026/1-1001.webp", 4, "image/webp")]
    fotos.apagar(url)
    assert falso.apagados == [("fotos", "inventario2026/1-1001.webp")]
    # formato antigo (fotos gravadas antes da Fase 3) continua reconhecido
    fotos.apagar("https://fotos.exemplo.org/inventario/INV1_BEM_1001_20260915120000.webp")
    assert falso.apagados[-1] == ("fotos", "inventario/INV1_BEM_1001_20260915120000.webp")
    fotos.apagar("https://outro/sem-prefixo.webp")                       # ignora
    fotos.apagar(None)
    assert len(falso.apagados) == 2
    monkeypatch.delenv("R2_PUBLIC_URL")
    assert fotos.enviar("a/1-2.webp", b"1") == "https://acc.r2.cloudflarestorage.com/fotos/a/1-2.webp"
    fotos.apagar("https://acc.r2.cloudflarestorage.com/fotos/a/1-2.webp")
    assert falso.apagados[-1] == ("fotos", "a/1-2.webp")
    assert not hasattr(fotos, "nome_bem") and not hasattr(fotos, "PREFIXO")
