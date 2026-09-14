"""Onde os dados vivem. Tudo que lê ou grava arquivo passa por aqui.

Windows empacotado (PyInstaller): pasta `dados/` ao lado do .exe.
Desenvolvimento: `./dados/`. A variável TERMOS_DADOS sobrepõe (usada nos testes).
"""
import os
import shutil
import sys
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
