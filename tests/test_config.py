from pathlib import Path


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
