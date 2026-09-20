from pathlib import Path

import pytest

import segredos


def test_ler_env_le_chaves_e_ignora_comentarios(tmp_path):
    arq = tmp_path / "sei.env"
    arq.write_text("# c\nSEI_USUARIO=u\nSEI_SENHA=s=com=igual\n\nSEI_LOGIN_URL='https://x/'\nSEI_ORGAO=CFC\n")
    env = segredos.ler_env(arq, ("SEI_USUARIO", "SEI_SENHA", "SEI_LOGIN_URL", "SEI_ORGAO"))
    assert env["SEI_SENHA"] == "s=com=igual" and env["SEI_LOGIN_URL"] == "https://x/"      # aspas simples caem


def test_ler_env_erros_com_nome_do_arquivo(tmp_path):
    with pytest.raises(segredos.SegredoAusente, match="secrets/sei.env não encontrado ou incompleto"):
        segredos.ler_env(tmp_path / "sei.env", ("SEI_USUARIO",))
    (tmp_path / "sei.env").write_text("SEI_USUARIO=u\n")
    with pytest.raises(segredos.SegredoAusente, match="falta SEI_SENHA"):
        segredos.ler_env(tmp_path / "sei.env", ("SEI_USUARIO", "SEI_SENHA"))


def test_pasta_padrao_e_secrets_do_projeto():
    assert segredos.PASTA == Path(segredos.__file__).resolve().parent / "secrets"
