import sqlite3

import pytest
from openpyxl import Workbook, load_workbook

import db
from tests.conftest import semear


def test_esquema_cria_seis_tabelas(dados):
    nomes = {r["name"] for r in dados.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    assert {"bens", "responsaveis", "localizacoes", "pessoas", "atribuicoes", "textos"} <= nomes


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


def test_renomear_centro_cascateia(dados):
    semear(dados)
    db.renomear_centro(dados, "CCI", "GESERV")
    assert db.localizacoes_mapeadas(dados)[0]["ccustos"] == "GESERV"
    assert db.responsavel(dados, "CCI") is None and db.responsavel(dados, "GESERV")


def test_excluir_centro_em_uso_falha(dados):
    semear(dados)  # CCI tem o bem 1001 ativo em "01 - SALA CCI"
    with pytest.raises(db.CentroEmUso) as e:
        db.excluir_responsavel(dados, "CCI")
    assert "1 bem" in str(e.value)
    db.desatribuir(dados, "ANA SILVA", 1002)  # agora 1001 e 1002 respondem pelo setor
    with pytest.raises(db.CentroEmUso) as e:
        db.excluir_responsavel(dados, "CCI")
    assert "2 bens" in str(e.value)


def test_incluir_responsavel_e_localizacao(dados):
    semear(dados)
    db.incluir_responsavel(dados, {"ccustos": " decom ", "tratamento": "Prezado", "responsavel": "THIAGO",
                                   "email": "t@cfc", "matricula": "481", "funcao": "gerente"})
    assert db.responsavel(dados, "DECOM")["responsavel"] == "THIAGO"
    db.incluir_localizacao(dados, "99 - SEM MAPA", "DECOM")
    assert db.localizacoes_sem_centro(dados) == []
    with pytest.raises(db.ErroDeNegocio):
        db.incluir_responsavel(dados, {"ccustos": "", "responsavel": "X"})


def test_atribuir_bem_livre_e_transferencia(dados):
    semear(dados)
    db.incluir_pessoa(dados, "  BRUNO LIMA ")
    db.atribuir(dados, "BRUNO LIMA", 1001)
    assert [b["numero"] for b in db.bens_da_pessoa(dados, "BRUNO LIMA")] == [1001]
    with pytest.raises(db.BemNaoEncontrado):
        db.atribuir(dados, "BRUNO LIMA", 9999)
    with pytest.raises(db.JaAtribuido) as e:
        db.atribuir(dados, "BRUNO LIMA", 1002)
    assert e.value.pessoa == "ANA SILVA"
    db.atribuir(dados, "BRUNO LIMA", 1002, confirmar=True)
    assert db.bens_da_pessoa(dados, "ANA SILVA") == []
    assert db.pessoa_do_bem(dados, 1002) == "BRUNO LIMA"
    db.atribuir(dados, "BRUNO LIMA", 1002)  # já é dele: não é erro


def test_desatribuir_e_excluir_pessoa(dados):
    semear(dados)
    db.desatribuir(dados, "ANA SILVA", 1002)
    assert [b["numero"] for b in db.bens_do_centro(dados, "CCI")] == [1001, 1002]
    db.atribuir(dados, "ANA SILVA", 1002)
    db.excluir_pessoa(dados, "ANA SILVA")
    assert db.pessoas(dados) == [] and db.pessoa_do_bem(dados, 1002) is None


def test_excluir_centro_sem_bens_apaga_mapeamentos(dados):
    semear(dados)
    db.incluir_responsavel(dados, {"ccustos": "VAZIO", "responsavel": "X"})
    db.incluir_localizacao(dados, "99 - SEM MAPA", "VAZIO")   # 1004 é ATIVO aqui...
    with pytest.raises(db.CentroEmUso):
        db.excluir_responsavel(dados, "VAZIO")
    dados.execute("UPDATE bens SET situacao='BAIXADO' WHERE numero=1004")
    dados.commit()
    db.excluir_responsavel(dados, "VAZIO")                       # ...sem bens ativos: some, e a sala volta a pendente
    assert db.responsavel(dados, "VAZIO") is None
    assert db.localizacoes_mapeadas(dados) == [{"localizacao": "01 - SALA CCI", "ccustos": "CCI"}]
    with pytest.raises(db.ErroDeNegocio):
        db.excluir_responsavel(dados, "NAO_EXISTE")


def test_atualizar_responsavel(dados):
    semear(dados)
    db.atualizar_responsavel(dados, "CCI", {"tratamento": "Prezado", "responsavel": " carlos ", "email": "c@cfc",
                                           "matricula": "99", "funcao": "gerente"})
    r = db.responsavel(dados, "CCI")
    assert (r["responsavel"], r["funcao"], r["matricula"]) == ("carlos", "gerente", "99")
    with pytest.raises(db.ErroDeNegocio):
        db.atualizar_responsavel(dados, "CCI", {"responsavel": ""})
    with pytest.raises(db.ErroDeNegocio):
        db.atualizar_responsavel(dados, "NAO_EXISTE", {"responsavel": "X"})


def test_mover_localizacoes(dados):
    semear(dados)
    db.incluir_responsavel(dados, {"ccustos": "PRES", "responsavel": "Y"})
    db.incluir_localizacao(dados, "99 - SEM MAPA", "CCI")
    n = db.mover_localizacoes(dados, ["01 - SALA CCI", "99 - SEM MAPA"], "PRES")
    assert n == 2 and {l["ccustos"] for l in db.localizacoes_mapeadas(dados)} == {"PRES"}
    assert [b["numero"] for b in db.bens_do_centro(dados, "PRES")] == [1001, 1004]
    with pytest.raises(db.ErroDeNegocio):
        db.mover_localizacoes(dados, [], "PRES")
    with pytest.raises(db.ErroDeNegocio):
        db.mover_localizacoes(dados, ["01 - SALA CCI"], "NAO_EXISTE")


def test_renomear_pessoa_mantem_atribuicoes(dados):
    semear(dados)
    assert db.renomear_pessoa(dados, "ANA SILVA", " ana  souza ") == "ANA SOUZA"
    assert db.pessoas(dados) == ["ANA SOUZA"] and db.pessoa_do_bem(dados, 1002) == "ANA SOUZA"
    db.incluir_pessoa(dados, "BRUNO")
    with pytest.raises(db.ErroDeNegocio):
        db.renomear_pessoa(dados, "ANA SOUZA", "bruno")
    with pytest.raises(db.ErroDeNegocio):
        db.renomear_pessoa(dados, "NINGUEM", "X")
    assert db.renomear_pessoa(dados, "ANA SOUZA", "ANA SOUZA") == "ANA SOUZA"


def test_exportar_cadastros_quatro_abas(dados, tmp_path):
    semear(dados)
    arq = db.exportar_cadastros(dados, tmp_path / "c.xlsx")
    wb = load_workbook(arq)
    assert wb.sheetnames == ["responsaveis", "localizacoes", "pessoas", "atribuicoes"]
    assert [c.value for c in wb["responsaveis"][1]] == ["ccustos", "tratamento", "responsavel", "email", "matricula", "funcao"]
    assert [c.value for c in wb["atribuicoes"][2]] == ["ANA SILVA", 1002]
    assert wb["localizacoes"].max_row == 2 and wb["pessoas"].max_row == 2


def test_importar_a_propria_exportacao_e_idempotente(dados, tmp_path):
    semear(dados)
    arq = db.exportar_cadastros(dados, tmp_path / "c.xlsx")
    resumo = db.importar_cadastros(dados, arq)
    assert resumo == {"responsaveis": 1, "localizacoes": 1, "pessoas": 1, "atribuicoes": 1, "sem_centro": ["99 - SEM MAPA"]}
    assert db.ficha_do_bem(dados, 1002)["pessoa"] == "ANA SILVA"


def cadastros_xlsx(tmp_path, **abas):
    wb = Workbook()
    wb.remove(wb.active)
    cabecalhos = {"responsaveis": ["ccustos", "tratamento", "responsavel", "email", "matricula", "funcao"],
                  "localizacoes": ["localizacao", "ccustos"], "pessoas": ["nome"], "atribuicoes": ["nome", "numero"]}
    for aba, cab in cabecalhos.items():
        ws = wb.create_sheet(aba)
        ws.append(cab)
        for linha in abas.get(aba, []):
            ws.append(linha)
    caminho = tmp_path / "cad.xlsx"
    wb.save(caminho)
    return caminho


def test_importar_cadastros_substitui_e_normaliza(dados, tmp_path):
    semear(dados)
    arq = cadastros_xlsx(tmp_path,
                         responsaveis=[["geserv ", "Prezado", " carlos", "c@cfc", 7, "gerente"], ["pres", "", "MARIA", "", "", ""]],
                         localizacoes=[["01 - SALA CCI", "GESERV"], ["99 - SEM MAPA", "pres"]],
                         pessoas=[[" bruno lima "]], atribuicoes=[["bruno lima", "1001"]])
    resumo = db.importar_cadastros(dados, arq)
    assert resumo["responsaveis"] == 2 and resumo["atribuicoes"] == 1 and resumo["sem_centro"] == []
    assert db.responsavel(dados, "CCI") is None and db.responsavel(dados, "GESERV")["matricula"] == "7"
    f = db.ficha_do_bem(dados, 1001)
    assert f["ccustos"] == "GESERV" and f["pessoa"] == "BRUNO LIMA"
    assert db.pessoas(dados) == ["BRUNO LIMA"]


@pytest.mark.parametrize("abas, trecho", [
    ({"localizacoes": [["01 - SALA CCI", "NAOEXISTE"]]}, "NAOEXISTE"),
    ({"atribuicoes": [["ANA SILVA", 1001]]}, "pessoas"),                     # pessoa fora da aba pessoas
    ({"pessoas": [["ANA"]], "atribuicoes": [["ANA", 9999]]}, "9999"),
    ({"pessoas": [["ANA"], ["BRUNO"]], "atribuicoes": [["ANA", 1001], ["BRUNO", 1001]]}, "repetido"),
    ({"responsaveis": [["", "", "X", "", "", ""]]}, "sigla"),
    ({"responsaveis": [["A", "", "", "", "", ""]]}, "responsável"),
    ({"responsaveis": [["A", "", "X", "", "", ""], ["a", "", "Y", "", "", ""]]}, "repetid"),
    ({"pessoas": [["ANA"]], "atribuicoes": [["ANA", "1001.5"]]}, "inválido"),
])
def test_importar_cadastros_invalidos_nao_alteram_nada(dados, tmp_path, abas, trecho):
    semear(dados)
    with pytest.raises(db.ImportacaoInvalida) as e:
        db.importar_cadastros(dados, cadastros_xlsx(tmp_path, **abas))
    assert trecho in str(e.value)
    assert db.centros(dados)[0]["ccustos"] == "CCI" and db.pessoas(dados) == ["ANA SILVA"]


def test_importar_cadastros_sem_aba_ou_coluna(dados, tmp_path):
    semear(dados)
    wb = Workbook()
    wb.active.title = "responsaveis"
    wb.active.append(["ccustos", "responsavel"])
    arq = tmp_path / "x.xlsx"
    wb.save(arq)
    with pytest.raises(db.ImportacaoInvalida) as e:
        db.importar_cadastros(dados, arq)
    assert "localizacoes" in str(e.value) or "coluna" in str(e.value)


def test_importar_planilha_truncada_levanta_importacao_invalida(dados, tmp_path):
    semear(dados)
    arq = xlsx(tmp_path, [[1002, "ATIVO", "NOTEBOOK", "DELL", "EQUIP", "01 - SALA CCI", "06/12/2012", 3000, 1400]])
    bruto = arq.read_bytes()
    truncado = tmp_path / "trunc.xlsx"
    truncado.write_bytes(bruto[: int(len(bruto) * 0.6)])
    with pytest.raises(db.ImportacaoInvalida):
        db.importar_bens(dados, truncado)
    with pytest.raises(db.ImportacaoInvalida):
        db.importar_cadastros(dados, truncado)
