from openpyxl import Workbook

import db
import migrar_inventario as mi
from tests.conftest import semear

COLUNAS = ["numero", "situacao", "descricao", "complemento", "classificacao_contabil", "localizacao_sistema", "data_entrada",
           "valor_compra", "valor_atual", "local_verificado", "estado_conservacao", "usuario", "data_hora", "usuario_bem",
           "imagem", "observacao", "contagem"]


def sepat(tmp_path, linhas):
    wb = Workbook()
    ws = wb.active
    ws.title = "BASE"
    ws.append(COLUNAS)
    for l in linhas:
        ws.append(l + [None] * (len(COLUNAS) - len(l)))
    wb.create_sheet("USUARIOS").append(["nome", "email"])
    p = tmp_path / "INVENTARIO_SEPAT .xlsx"
    wb.save(p)
    return p


def test_transformar_gera_evento_leituras_e_sobra(dados, tmp_path):
    semear(dados)
    linhas = [
        [1001, "ATIVO", "CADEIRA", "", "", "01 - SALA CCI", "", 0, 0, "01 - SALA CCI", "Bom", "Rafael Silvio", "25/08/2026 14:44:46"],
        [1002, "ATIVO", "NOTEBOOK", "", "", "01 - SALA CCI", "", 0, 0, "99 - SEM MAPA", "Regular", "Denise  Cristiane", "19/08/2026 10:53:02", "JOÃO", "http://f/1.webp", "tela riscada"],
        [1002, "ATIVO", "NOTEBOOK", "", "", "01 - SALA CCI", "", 0, 0, "01 - SALA CCI", "Bom", "Rafael Silvio", "28/08/2026 09:11:52"],   # releitura mais recente
        [1003, "BAIXADO", "MESA", "", "", "01 - SALA CCI", "", 0, 0, "01 - SALA CCI", "Ruim", "Rafael Silvio", "28/08/2026 09:12:00"],    # baixado lido: entra
        [1004, "ATIVO", "ARMÁRIO", "", "", "99 - SEM MAPA", "", 0, 0, None, None, None, None],                                        # não lido: não entra
        [None, None, "Cadeira Delic", "", "", "", "", 0, 0, "07 - DELIC", None, "Antônio Rodrigues", "18/08/2026 14:16:38", None, None, None],  # sobra sem foto/obs
    ]
    abas, r = mi.transformar(dados, mi.ler_base(sepat(tmp_path, linhas)))
    assert r["leituras"] == 3 and r["sobras"] == 1 and r["problemas"] == [] and r["inexistentes"] == []
    assert abas["inv_eventos"] == [[1, mi.EVENTO_NOME, mi.EVENTO_DESCRICAO, "2026-08-18 14:16:38", None]]
    assert abas["inv_integrantes"] == [[1, "Antônio Rodrigues"], [1, "Denise Cristiane"], [1, "Rafael Silvio"]]
    assert abas["inv_salas"] == [[1, "01 - SALA CCI"], [1, "07 - DELIC"], [1, "99 - SEM MAPA"]]
    l1002 = next(l for l in abas["inv_leituras"] if l[1] == 1002)
    assert l1002[2:6] == ["01 - SALA CCI", "2026-08-28 09:11:52", "Rafael Silvio", "Bom"]   # ficou a mais recente
    assert abas["inv_sobras"] == [[1, "07 - DELIC", "Cadeira Delic", None, mi.SEM_OBS, mi.SEM_FOTO, "Antônio Rodrigues", "2026-08-18 14:16:38"]]
    assert r["divergentes"] == 0 and r["por_integrante"] == {"Rafael Silvio": 3}


def test_bem_inexistente_interrompe_ou_e_pulado(dados, tmp_path):
    semear(dados)
    linhas = [[1001, "ATIVO", "CADEIRA", "", "", "", "", 0, 0, "01 - SALA CCI", "Bom", "R", "25/08/2026 14:44:46"],
              [14906, "ATIVO", "NOVO", "", "", "", "", 0, 0, "01 - SALA CCI", "Bom", "R", "25/08/2026 14:44:46"]]
    import pytest
    with pytest.raises(SystemExit) as e:
        mi.transformar(dados, mi.ler_base(sepat(tmp_path, linhas)))
    assert "14906" in str(e.value) and "SPW" in str(e.value)
    abas, r = mi.transformar(dados, mi.ler_base(sepat(tmp_path, linhas)), pular_inexistentes=True)
    assert r["leituras"] == 1 and r["inexistentes"] == [14906]


def test_arquivo_gerado_e_importado_pelo_sistema(dados, tmp_path):
    """Ponta a ponta: planilha antiga + cadastros exportados → arquivo → db.importar_cadastros."""
    semear(dados)
    cadastros = db.exportar_cadastros(dados, tmp_path / "cadastros.xlsx")
    linhas = [[1001, "ATIVO", "CADEIRA", "", "", "", "", 0, 0, "01 - SALA CCI", "Bom", "Rafael Silvio", "25/08/2026 14:44:46"],
              [1002, "ATIVO", "NOTEBOOK", "", "", "", "", 0, 0, "02 - OUTRA", "Bom", "Rafael Silvio", "26/08/2026 10:00:00"],
              [None, None, "Sobra X", "", "", "", "", 0, 0, "01 - SALA CCI", None, "Rafael Silvio", "26/08/2026 11:00:00", None, "http://f/s.webp", "sem plaqueta"]]
    abas, _ = mi.transformar(dados, mi.ler_base(sepat(tmp_path, linhas)))
    saida = mi.gravar(cadastros, tmp_path / "saida.xlsx", abas)
    resumo = db.importar_cadastros(dados, saida)
    assert resumo["inv_eventos"] == 1 and resumo["inv_leituras"] == 2 and resumo["inv_sobras"] == 1
    assert resumo["responsaveis"] == 1 and resumo["pessoas"] == 1 and resumo["atribuicoes"] == 1   # cadastros preservados
    import inventario
    e = inventario.evento_aberto(dados)
    assert e and e["nome"] == mi.EVENTO_NOME and e["aberto_em"] == "2026-08-25 14:44:46"
    assert dados.execute("SELECT COUNT(*) FROM inventario_salas").fetchone()[0] == 3    # 01, 99 (ativas) + 02 - OUTRA (lida)
    assert dados.execute("SELECT localizacao FROM inventario_leituras WHERE numero = 1002").fetchone()[0] == "02 - OUTRA"
    # rodar de novo substitui as abas inv_* sem duplicar
    saida2 = mi.gravar(saida, tmp_path / "saida2.xlsx", abas)
    assert db.importar_cadastros(dados, saida2)["inv_leituras"] == 2
