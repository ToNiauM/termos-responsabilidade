from pathlib import Path

import pytest


def test_env_sobrepoe_pasta_dados(tmp_path, monkeypatch):
    monkeypatch.setenv("TERMOS_DADOS", str(tmp_path))
    import config
    assert config.pasta_dados() == tmp_path
    assert config.caminho_db() == tmp_path / "termos.db"
    assert config.caminho_timbrado() == tmp_path / "timbrado.docx"


def test_preparar_pastas_copia_timbrado(tmp_path, monkeypatch):
    monkeypatch.setenv("TERMOS_DADOS", str(tmp_path / "dados"))
    import config
    config.preparar_pastas()
    tmp_path = tmp_path / "dados"
    assert (tmp_path / "timbrado.docx").stat().st_size > 1000


def test_sem_env_usa_pasta_dados_do_projeto(monkeypatch):
    monkeypatch.delenv("TERMOS_DADOS", raising=False)
    import config
    assert config.pasta_dados() == Path(config.__file__).parent / "dados"


def test_chave_secreta_vem_da_variavel_de_ambiente(tmp_path, monkeypatch):
    monkeypatch.setenv("TERMOS_DADOS", str(tmp_path))
    monkeypatch.setenv("TERMOS_SEGREDO", "segredo-da-web")
    import config
    assert config.chave_secreta() == "segredo-da-web"
    assert not (tmp_path / "segredo.txt").exists()


def test_chave_secreta_e_gerada_uma_vez_por_instalacao(tmp_path, monkeypatch):
    monkeypatch.setenv("TERMOS_DADOS", str(tmp_path / "dados"))   # pasta ainda não existe
    monkeypatch.delenv("TERMOS_SEGREDO", raising=False)
    import config
    chave = config.chave_secreta()
    assert len(chave) == 64 and int(chave, 16) >= 0
    assert (tmp_path / "dados" / "segredo.txt").read_text() == chave
    assert config.chave_secreta() == chave                          # segunda chamada lê o arquivo


def test_chave_secreta_respeita_arquivo_ja_criado_por_outro_processo(tmp_path, monkeypatch):
    monkeypatch.setenv("TERMOS_DADOS", str(tmp_path))
    monkeypatch.delenv("TERMOS_SEGREDO", raising=False)
    (tmp_path / "segredo.txt").write_text("abc")
    import config
    assert config.chave_secreta() == "abc"


def test_chave_secreta_espera_o_outro_processo_terminar_de_escrever(tmp_path, monkeypatch):
    import threading
    monkeypatch.setenv("TERMOS_DADOS", str(tmp_path))
    monkeypatch.delenv("TERMOS_SEGREDO", raising=False)
    arquivo = tmp_path / "segredo.txt"
    arquivo.write_text("")                                   # criado, ainda não escrito
    threading.Timer(0.1, lambda: arquivo.write_text("abc")).start()
    import config
    assert config.chave_secreta() == "abc"


def test_chave_secreta_sem_permissao_explica_o_motivo(tmp_path, monkeypatch):
    import os
    if os.name != "posix" or os.geteuid() == 0:
        pytest.skip("chmod 0 não impede leitura aqui")
    monkeypatch.setenv("TERMOS_DADOS", str(tmp_path))
    monkeypatch.delenv("TERMOS_SEGREDO", raising=False)
    arquivo = tmp_path / "segredo.txt"
    arquivo.write_text("abc")
    arquivo.chmod(0)
    import config
    with pytest.raises(RuntimeError, match="TERMOS_SEGREDO"):
        config.chave_secreta()
    arquivo.chmod(0o600)
