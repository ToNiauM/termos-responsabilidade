"""Onde os dados vivem. Tudo que lê ou grava arquivo passa por aqui.

Windows empacotado (PyInstaller): pasta `dados/` ao lado do .exe.
Desenvolvimento: `./dados/`. A variável TERMOS_DADOS sobrepõe (usada nos testes).
"""
import os
import secrets
import shutil
import sys
import time
from pathlib import Path


def pasta_recursos() -> Path:
    """Arquivos embutidos no programa (templates, static, timbrado.docx original)."""
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


def pasta_dados() -> Path:
    if os.environ.get("TERMOS_DADOS"):
        return Path(os.environ["TERMOS_DADOS"])
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / "dados"
    return Path(__file__).resolve().parent / "dados"


def caminho_db() -> Path:
    return pasta_dados() / "termos.db"


def caminho_timbrado() -> Path:
    return pasta_dados() / "timbrado.docx"


def preparar_pastas() -> None:
    """Cria a pasta de dados e copia o timbrado na primeira execução."""
    pasta_dados().mkdir(parents=True, exist_ok=True)
    if not caminho_timbrado().exists():
        shutil.copy(pasta_recursos() / "timbrado.docx", caminho_timbrado())


def chave_secreta() -> str:
    """Assina cookies e tokens de revisão. TERMOS_SEGREDO (web) ou dados/segredo.txt, criado na primeira
    execução. Dois processos ao mesmo tempo não brigam: O_EXCL garante um único criador; o outro lê."""
    if os.environ.get("TERMOS_SEGREDO"):
        return os.environ["TERMOS_SEGREDO"]
    caminho = pasta_dados() / "segredo.txt"
    if not caminho.exists():
        pasta_dados().mkdir(parents=True, exist_ok=True)
        try:
            fd = os.open(caminho, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        except FileExistsError:
            pass
        else:
            with os.fdopen(fd, "w") as f:
                f.write(secrets.token_hex(32))
    for _ in range(50):                     # o outro processo pode ainda estar escrevendo
        chave = caminho.read_text().strip()
        if chave:
            return chave
        time.sleep(0.02)
    raise RuntimeError(f"{caminho} continua vazio: apague o arquivo e abra o programa de novo.")
