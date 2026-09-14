import sqlite3

import pytest
from openpyxl import Workbook

import db
from tests.conftest import semear


def test_esquema_cria_cinco_tabelas(dados):
    nomes = {r["name"] for r in dados.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    assert {"bens", "responsaveis", "localizacoes", "pessoas", "atribuicoes"} <= nomes


def test_esquema_e_idempotente(dados):
    db.criar_esquema(dados)  # segunda vez não pode falhar


def test_foreign_keys_ligadas(dados):
    semear(dados)
    with pytest.raises(sqlite3.IntegrityError):
        dados.execute("INSERT INTO localizacoes VALUES ('02 - X', 'NAO_EXISTE')")


CABECALHO = ["Número Bem", "Situação", "Descrição", "Complemento", "Classificação Contábil",
             "Localização", "Data Entrada", "Valor Compra", "Valor Atual"]


def xlsx(tmp_path, linhas, cabecalho=CABECALHO, aba="base"):
    wb = Workbook()
    ws = wb.active
    ws.title = aba
    ws.append(cabecalho)
    for l in linhas:
        ws.append(l)
    caminho = tmp_path / "export.xlsx"
    wb.save(caminho)
    return caminho


def test_importar_substitui_bens_e_conta(dados, tmp_path):
    semear(dados)
    arq = xlsx(tmp_path, [
        [1002, "ATIVO", "NOTEBOOK", "DELL NOVO", "EQUIP", "01 - SALA CCI", "06/12/2012", 3000, 1400],
        [2001, "ATIVO", "MONITOR", "LG", "EQUIP", "01 - SALA CCI", "01/01/2020", 900, 800],
        [2002, "DOADO", "CADEIRA", "", "MÓVEIS", "CFC", "01/01/2000", 10, 1],
    ])
    resumo = db.importar_bens(dados, arq)
    assert resumo["total"] == 3 and resumo["ativos"] == 2
    assert [r["numero"] for r in dados.execute("SELECT numero FROM bens ORDER BY numero")] == [1002, 2001, 2002]
    assert dados.execute("SELECT complemento FROM bens WHERE numero=1002").fetchone()[0] == "DELL NOVO"
    # outras tabelas intactas
    assert dados.execute("SELECT count(*) FROM atribuicoes").fetchone()[0] == 1


def test_importar_cabecalho_errado_nao_altera_nada(dados, tmp_path):
    semear(dados)
    arq = xlsx(tmp_path, [[1, "ATIVO"]], cabecalho=["Patrimônio", "Situação"])
    with pytest.raises(db.ImportacaoInvalida):
        db.importar_bens(dados, arq)
    assert dados.execute("SELECT count(*) FROM bens").fetchone()[0] == 4


def test_importar_que_some_com_bem_atribuido_e_revertida(dados, tmp_path):
    semear(dados)  # ANA tem o 1002
    arq = xlsx(tmp_path, [[1001, "ATIVO", "CADEIRA", "", "MÓVEIS", "01 - SALA CCI", "x", 1, 1]])
    with pytest.raises(db.ImportacaoInvalida) as e:
        db.importar_bens(dados, arq)
    assert "1002" in str(e.value)
    assert dados.execute("SELECT count(*) FROM bens").fetchone()[0] == 4


def test_importar_usa_aba_pelo_cabecalho_e_converte_data(dados, tmp_path):
    from datetime import datetime
    wb = Workbook()
    wb.active.title = "outra"
    wb.active.append(["Nada", "aqui"])
    ws = wb.create_sheet("base")
    ws.append(CABECALHO)
    ws.append([3001, "ATIVO", "MESA", None, "MÓVEIS", "02 - X", datetime(2020, 3, 9), 100, 90])
    arq = tmp_path / "e.xlsx"
    wb.save(arq)
    db.importar_bens(dados, arq)
    r = dados.execute("SELECT data_entrada, complemento FROM bens WHERE numero=3001").fetchone()
    assert r["data_entrada"] == "09/03/2020" and r["complemento"] == ""


def test_localizacoes_sem_centro(dados):
    semear(dados)
    assert db.localizacoes_sem_centro(dados) == ["99 - SEM MAPA"]


def test_importar_arquivo_invalido_levanta_importacao_invalida(dados, tmp_path):
    semear(dados)
    arq = tmp_path / "x.xlsx"
    arq.write_bytes(b"nada")
    with pytest.raises(db.ImportacaoInvalida):
        db.importar_bens(dados, arq)
    assert dados.execute("SELECT count(*) FROM bens").fetchone()[0] == 4


def test_importar_numero_repetido_e_revertida(dados, tmp_path):
    semear(dados)
    arq = xlsx(tmp_path, [
        [1001, "ATIVO", "CADEIRA", "", "MÓVEIS", "01 - SALA CCI", "01/01/2000", 10, 5],
        [1001, "ATIVO", "CADEIRA CÓPIA", "", "MÓVEIS", "01 - SALA CCI", "01/01/2000", 10, 5],
    ])
    with pytest.raises(db.ImportacaoInvalida) as e:
        db.importar_bens(dados, arq)
    assert "repetido" in str(e.value)
    assert dados.execute("SELECT count(*) FROM bens").fetchone()[0] == 4


def test_bens_do_centro_exclui_baixados_atribuidos_e_sem_mapa(dados):
    semear(dados)
    assert [b["numero"] for b in db.bens_do_centro(dados, "CCI")] == [1001]


def test_bens_da_pessoa(dados):
    semear(dados)
    bens = db.bens_da_pessoa(dados, "ANA SILVA")
    assert [b["numero"] for b in bens] == [1002] and bens[0]["descricao"] == "NOTEBOOK"


def test_centros_responsavel_pessoas(dados):
    semear(dados)
    assert [c["ccustos"] for c in db.centros(dados)] == ["CCI"]
    assert db.responsavel(dados, "CCI")["responsavel"] == "JAQUELINE PORTELA"
    assert db.responsavel(dados, "XX") is None
    assert db.pessoas(dados) == ["ANA SILVA"]


def test_ficha_do_bem_com_setor_e_pessoa(dados):
    semear(dados)
    f = db.ficha_do_bem(dados, 1002)
    assert f["ccustos"] == "CCI" and f["responsavel"] == "JAQUELINE PORTELA" and f["pessoa"] == "ANA SILVA"
    f = db.ficha_do_bem(dados, 1004)
    assert f["ccustos"] is None and f["pessoa"] is None
    assert db.ficha_do_bem(dados, 9999) is None
    assert db.buscar_bem(dados, 1001)["descricao"] == "CADEIRA"
    assert db.localizacoes_mapeadas(dados) == [{"localizacao": "01 - SALA CCI", "ccustos": "CCI"}]
