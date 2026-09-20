import pytest

import cofre
import db


@pytest.fixture
def chave(tmp_path, monkeypatch):
    arq = tmp_path / "chaves.env"
    arq.write_text(f"CHAVE_SENHAS={cofre.gerar_chave()}\n")
    monkeypatch.setattr(cofre, "ARQUIVO_CHAVE", arq)
    return arq


def test_ida_e_volta_e_tokens_diferentes(chave):
    a, b = cofre.cifrar("Segredo!1"), cofre.cifrar("Segredo!1")
    assert a != b and "Segredo" not in a and cofre.decifrar(a) == cofre.decifrar(b) == "Segredo!1"


def test_sem_chave_e_chave_invalida(tmp_path, monkeypatch):
    monkeypatch.setattr(cofre, "ARQUIVO_CHAVE", tmp_path / "chaves.env")
    with pytest.raises(db.ErroDeNegocio, match="secrets/chaves.env não encontrado ou incompleto"):
        cofre.cifrar("x")
    (tmp_path / "chaves.env").write_text("CHAVE_SENHAS=nao-e-uma-chave\n")
    with pytest.raises(db.ErroDeNegocio, match="CHAVE_SENHAS inválida"):
        cofre.cifrar("x")


def test_token_de_outra_chave(chave, tmp_path, monkeypatch):
    token = cofre.cifrar("x")
    outra = tmp_path / "outra.env"
    outra.write_text(f"CHAVE_SENHAS={cofre.gerar_chave()}\n")
    monkeypatch.setattr(cofre, "ARQUIVO_CHAVE", outra)
    with pytest.raises(db.ErroDeNegocio, match="cifrada com outra chave"):
        cofre.decifrar(token)
