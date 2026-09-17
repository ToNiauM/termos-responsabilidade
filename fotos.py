"""Fotos do inventário no Cloudflare R2 (API S3): validação, compressão para WebP e envio.
Transcrito do sistema antigo (sga/cfc: app/services/storage.py e images.py), sem Flask.

Credenciais nas variáveis de ambiente R2_* (compose.yml lê secrets/.env). Sem elas, `configurado()` é
falso e as telas desativam os botões de foto; o resto do inventário funciona (programa Windows offline).

Chave `{pasta}/{nfoto}-{numero}.webp` (pasta = nome do evento normalizado); fotos anteriores à Fase 3
ficam em `inventario/INV...`."""
import io
import os
import re
import unicodedata

from db import ErroDeNegocio

VARIAVEIS = ("R2_ACCESS_KEY_ID", "R2_SECRET_ACCESS_KEY", "R2_ENDPOINT_URL", "R2_BUCKET_NAME")
EXTENSOES = {".jpg", ".jpeg", ".png", ".webp"}
TAMANHO_MAX = 5 * 1024 * 1024
LARGURA_MAX, ALTURA_MAX, QUALIDADE = 1920, 1080, 85


def configurado() -> bool:
    return all(os.environ.get(v) for v in VARIAVEIS)


def validar(arquivo) -> bytes:
    """arquivo: objeto com .filename e .read() (FileStorage do Flask). Devolve os bytes se for imagem válida."""
    nome = (getattr(arquivo, "filename", "") or "").lower()
    if os.path.splitext(nome)[1] not in EXTENSOES:
        raise ErroDeNegocio("Foto: envie um arquivo .jpg, .png ou .webp.")
    dados = arquivo.read()
    if len(dados) > TAMANHO_MAX:
        raise ErroDeNegocio("Foto maior que 5 MB.")
    from PIL import Image
    try:
        Image.open(io.BytesIO(dados)).verify()
    except Exception:
        raise ErroDeNegocio("O arquivo não é uma imagem válida.")
    return dados


def comprimir(dados: bytes) -> bytes:
    """WebP qualidade 85, no máximo 1920×1080 mantendo a proporção, orientação EXIF aplicada."""
    from PIL import Image, ImageOps
    img = ImageOps.exif_transpose(Image.open(io.BytesIO(dados)))
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGB")
    largura, altura = img.size
    if largura > LARGURA_MAX or altura > ALTURA_MAX:
        razao = min(LARGURA_MAX / largura, ALTURA_MAX / altura)
        img = img.resize((int(largura * razao), int(altura * razao)), Image.Resampling.LANCZOS)
    saida = io.BytesIO()
    img.save(saida, "WEBP", quality=QUALIDADE, optimize=True)
    return saida.getvalue()


def _cliente():
    import boto3
    from botocore.config import Config
    return boto3.client("s3", aws_access_key_id=os.environ["R2_ACCESS_KEY_ID"],
                        aws_secret_access_key=os.environ["R2_SECRET_ACCESS_KEY"], region_name="auto",
                        endpoint_url=os.environ["R2_ENDPOINT_URL"],
                        config=Config(signature_version="s3v4", s3={"addressing_style": "path"}))


def _url(chave: str) -> str:
    base = os.environ.get("R2_PUBLIC_URL")
    if base:
        return f"{base.rstrip('/')}/{chave}"
    return f"{os.environ['R2_ENDPOINT_URL'].rstrip('/')}/{os.environ['R2_BUCKET_NAME']}/{chave}"


def pasta(nome: str, evento_id: int) -> str:
    """Pasta das fotos do evento no bucket: nome em minúsculas, sem acento, só [a-z0-9]. Vazio → evento<id>."""
    base = unicodedata.normalize("NFD", nome or "")
    limpo = re.sub(r"[^a-z0-9]", "", "".join(c for c in base if not unicodedata.combining(c)).casefold())
    return limpo or f"evento{evento_id}"


def chave_bem(pasta: str, nfoto: int, numero: int) -> str:
    return f"{pasta}/{nfoto}-{numero}.webp"


def chave_sobra(pasta: str, sobra_id: int) -> str:
    return f"{pasta}/sobra-{sobra_id}.webp"


def enviar(chave: str, dados: bytes) -> str:
    """Grava `chave` no bucket exatamente como recebida e devolve a URL pública."""
    if not configurado():
        raise ErroDeNegocio("Fotos desativadas: bucket não configurado.")
    _cliente().put_object(Bucket=os.environ["R2_BUCKET_NAME"], Key=chave, Body=dados, ContentType="image/webp")
    return _url(chave)


_PREFIXO_ANTIGO = "inventario/"


def apagar(url: str | None) -> None:
    """Apaga o objeto pela chave contida na URL: URL nova = base pública + chave; URL antiga (antes da
    Fase 3) tem "inventario/" no meio. Erro do bucket é ignorado (a URL some do banco de qualquer jeito)."""
    if not url or not configurado():
        return
    base = _url("")
    if url.startswith(base):
        chave = url[len(base):]
    else:
        pos = url.find(_PREFIXO_ANTIGO)
        if pos < 0:
            return
        chave = url[pos:]
    if not chave:
        return
    try:
        _cliente().delete_object(Bucket=os.environ["R2_BUCKET_NAME"], Key=chave)
    except Exception:
        pass
