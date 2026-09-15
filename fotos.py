"""Fotos do inventário no Cloudflare R2 (API S3): validação, compressão para WebP e envio.
Transcrito do sistema antigo (sga/cfc: app/services/storage.py e images.py), sem Flask.

Credenciais nas variáveis de ambiente R2_* (compose.yml lê secrets/.env). Sem elas, `configurado()` é
falso e as telas desativam os botões de foto; o resto do inventário funciona (programa Windows offline)."""
import io
import os
from datetime import datetime

from db import ErroDeNegocio

VARIAVEIS = ("R2_ACCESS_KEY_ID", "R2_SECRET_ACCESS_KEY", "R2_ENDPOINT_URL", "R2_BUCKET_NAME")
EXTENSOES = {".jpg", ".jpeg", ".png", ".webp"}
TAMANHO_MAX = 5 * 1024 * 1024
LARGURA_MAX, ALTURA_MAX, QUALIDADE = 1920, 1080, 85
PREFIXO = "inventario/"


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


def enviar(nome: str, dados: bytes) -> str:
    """Grava `inventario/<nome>` no bucket e devolve a URL pública."""
    if not configurado():
        raise ErroDeNegocio("Fotos desativadas: bucket não configurado.")
    chave = PREFIXO + nome
    _cliente().put_object(Bucket=os.environ["R2_BUCKET_NAME"], Key=chave, Body=dados, ContentType="image/webp")
    return _url(chave)


def apagar(url: str | None) -> None:
    """Apaga o objeto pela chave contida na URL; erro só é ignorado (a URL some do banco de qualquer jeito)."""
    if not url or not configurado():
        return
    pos = url.find(PREFIXO)
    if pos < 0:
        return
    try:
        _cliente().delete_object(Bucket=os.environ["R2_BUCKET_NAME"], Key=url[pos:])
    except Exception:
        pass


def _carimbo() -> str:
    return datetime.now().strftime("%Y%m%d%H%M%S")


def nome_bem(evento_id: int, numero: int) -> str:
    return f"INV{evento_id}_BEM_{numero}_{_carimbo()}.webp"


def nome_sobra(evento_id: int, sobra_id: int) -> str:
    return f"INV{evento_id}_SOBRA_{sobra_id}_{_carimbo()}.webp"
