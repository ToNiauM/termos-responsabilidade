"""Leitura dos arquivos secrets/*.env (CHAVE=valor, uma por linha; # comenta; aspas simples/duplas caem).
Usado pelos processos do host (importar_spw.py, robo_sei.py). Nunca registrar valores em log."""
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
PASTA = RAIZ / "secrets"


class SegredoAusente(Exception):
    """Arquivo ausente ou chave faltando; a mensagem cita secrets/<nome> e as chaves que faltam."""


def ler_env(caminho: Path, chaves: tuple[str, ...]) -> dict:
    caminho = Path(caminho)
    rotulo = f"secrets/{caminho.name} não encontrado ou incompleto ({caminho})"
    if not caminho.exists():
        raise SegredoAusente(rotulo)
    env = {}
    for linha in caminho.read_text(encoding="utf-8").splitlines():
        if "=" in linha and not linha.lstrip().startswith("#"):
            chave, valor = linha.split("=", 1)
            env[chave.strip()] = valor.strip().strip("'\"")
    faltando = [c for c in chaves if not env.get(c)]
    if faltando:
        raise SegredoAusente(f"{rotulo}: falta " + ", ".join(faltando))
    return env
