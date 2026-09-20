"""Cifra das senhas do SEI guardadas em usuarios.sei_senha. Fernet (cryptography) com a chave em
secrets/chaves.env (CHAVE_SENHAS=...). A mesma chave serve ao site (cifra ao salvar) e ao serviço robo (decifra
ao atender o pedido); os dois montam secrets/. Sem a chave, nada é salvo nem lido — e um backup do banco sozinho
não revela senha nenhuma. Gerar uma vez: python -c "import cofre; print(cofre.gerar_chave())"."""
from cryptography.fernet import Fernet, InvalidToken

import db
import segredos

ARQUIVO_CHAVE = segredos.PASTA / "chaves.env"


def gerar_chave() -> str:
    return Fernet.generate_key().decode()


def _fernet() -> Fernet:
    try:
        env = segredos.ler_env(ARQUIVO_CHAVE, ("CHAVE_SENHAS",))
    except segredos.SegredoAusente as e:
        raise db.ErroDeNegocio(str(e))
    try:
        return Fernet(env["CHAVE_SENHAS"].encode())
    except (ValueError, TypeError):
        raise db.ErroDeNegocio("CHAVE_SENHAS inválida em secrets/chaves.env.")


def cifrar(texto: str) -> str:
    return _fernet().encrypt(texto.encode("utf-8")).decode()


def decifrar(token: str) -> str:
    try:
        return _fernet().decrypt(token.encode()).decode("utf-8")
    except InvalidToken:
        raise db.ErroDeNegocio("Senha cifrada com outra chave; cadastre a senha do SEI de novo.")
