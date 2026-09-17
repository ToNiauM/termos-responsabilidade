"""Módulo de inventário: eventos, salas, leituras, sobras, relatório, xlsx."""
import pytest

import db
import inventario
from tests.conftest import semear


def semear_inventario(conn):
    """semear() + 2ª sala com 2 bens e 1 sala vazia de escopo; devolve o id do evento aberto."""
    semear(conn)
    conn.execute("INSERT INTO bens VALUES (2001,'ATIVO','MONITOR','LG','EQUIPAMENTOS','02 - SALA B','01/01/2020',900,800)")
    conn.execute("INSERT INTO bens VALUES (2002,'ATIVO','IMPRESSORA','HP','EQUIPAMENTOS','02 - SALA B','01/01/2020',1200,1000)")
    conn.commit()
    return inventario.abrir_evento(conn, "Inventário 2026", "Portaria 1/2026", ["Fulano", "Beltrana"])


def test_localizacoes_ativas(dados):
    semear(dados)
    assert db.localizacoes_ativas(dados) == ["01 - SALA CCI", "99 - SEM MAPA"]   # 1003 é BAIXADO, não muda nada


def test_abrir_evento_todas_as_salas_e_integrantes(dados):
    eid = semear_inventario(dados)
    e = inventario.evento(dados, eid)
    assert e["nome"] == "Inventário 2026" and e["encerrado_em"] is None and e["integrantes"] == ["Beltrana", "Fulano"]
    assert [s["localizacao"] for s in inventario.salas(dados, eid)] == ["01 - SALA CCI", "02 - SALA B", "99 - SEM MAPA"]
    assert inventario.evento_aberto(dados)["id"] == eid
    assert e["resumo"] == {"salas": 3, "salas_iniciadas": 0, "bens": 5, "lidos": 0, "divergentes": 0,
                           "pendentes": 5, "sobras": 0, "pct_bens": 0.0}


def test_abrir_evento_amostragem_e_validacoes(dados):
    semear(dados)
    with pytest.raises(db.ErroDeNegocio):
        inventario.abrir_evento(dados, "", "", ["A"])
    with pytest.raises(db.ErroDeNegocio):
        inventario.abrir_evento(dados, "X", "", [" ", ""])
    with pytest.raises(db.ErroDeNegocio):
        inventario.abrir_evento(dados, "X", "", ["A"], salas=["NÃO EXISTE"])
    eid = inventario.abrir_evento(dados, "Amostra", None, ["A", "A ", "b"], salas=["99 - SEM MAPA"])
    assert [s["localizacao"] for s in inventario.salas(dados, eid)] == ["99 - SEM MAPA"]
    assert inventario.evento(dados, eid)["integrantes"] == ["A", "b"]
    with pytest.raises(db.ErroDeNegocio):
        inventario.abrir_evento(dados, "Outro", "", ["A"])          # já há aberto
    inventario.encerrar_evento(dados, eid)
    assert inventario.evento_aberto(dados) is None
    e2 = inventario.abrir_evento(dados, "Outro", "", ["A"])
    assert [x["id"] for x in inventario.eventos(dados)] == [e2, eid]   # aberto primeiro


def test_salas_contadores_e_resumo(dados):
    eid = semear_inventario(dados)
    s = {x["localizacao"]: x for x in inventario.salas(dados, eid)}
    assert s["01 - SALA CCI"]["total"] == 2 and s["01 - SALA CCI"]["ccustos"] == "CCI" and s["02 - SALA B"]["ccustos"] is None
    assert (s["01 - SALA CCI"]["localizados"], s["01 - SALA CCI"]["pendentes"], s["01 - SALA CCI"]["divergentes"]) == (0, 2, 0)
    assert "concluida_em" not in s["01 - SALA CCI"]
    dados.execute("INSERT INTO inventario_leituras (evento_id, numero, localizacao, lido_em, integrante) VALUES (?,?,?,?,?)",
                  (eid, 2001, "02 - SALA B", "2026-09-15 10:00:00", "Fulano"))
    dados.commit()
    r = inventario.resumo(dados, eid)
    assert (r["salas_iniciadas"], r["lidos"], r["pendentes"], r["pct_bens"]) == (1, 1, 4, 20.0)
    inventario.encerrar_evento(dados, eid)
    assert inventario.evento(dados, eid)["encerrado_em"] is not None
    inventario.encerrar_evento(dados, eid)                                        # idempotente
    with pytest.raises(db.ErroDeNegocio):
        inventario.encerrar_evento(dados, 999)


def test_ler_localizado_divergente_reler_e_erros(dados):
    eid = semear_inventario(dados)
    r = inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    assert r["situacao"] == "localizado" and r["reler"] is False and r["ativo"] and r["bem"]["descricao"] == "CADEIRA"
    r = inventario.ler(dados, eid, "01 - SALA CCI", 2001, "Fulano")             # MONITOR é da SALA B
    assert r["situacao"] == "divergente" and r["cadastrado_em"] == "02 - SALA B"
    r = inventario.ler(dados, eid, "02 - SALA B", 2001, "Beltrana")             # reler: atualiza a mesma linha
    assert r["situacao"] == "localizado" and r["reler"] and r["leitura_anterior"]["localizacao"] == "01 - SALA CCI"
    assert dados.execute("SELECT count(*) FROM inventario_leituras WHERE evento_id = ?", (eid,)).fetchone()[0] == 2
    r = inventario.ler(dados, eid, "01 - SALA CCI", 1003, "Fulano")             # BAIXADO: registra, avisa
    assert r["ativo"] is False and r["situacao"] == "localizado"
    with pytest.raises(inventario.BemNaoEncontrado) as e:
        inventario.ler(dados, eid, "01 - SALA CCI", 99999, "Fulano")
    assert e.value.numero == 99999
    with pytest.raises(db.ErroDeNegocio):
        inventario.ler(dados, eid, "SALA QUE NÃO EXISTE", 1001, "Fulano")
    with pytest.raises(db.ErroDeNegocio):
        inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Ninguém")
    s = {x["localizacao"]: x for x in inventario.salas(dados, eid)}
    assert (s["01 - SALA CCI"]["localizados"], s["01 - SALA CCI"]["pendentes"]) == (1, 1)
    assert (s["02 - SALA B"]["localizados"], s["02 - SALA B"]["divergentes"]) == (1, 0)
    assert s["01 - SALA CCI"]["divergentes"] == 0

    inventario.ler(dados, eid, "02 - SALA B", 1003, "Fulano")                   # BAIXADO relido em outra sala
    s = {x["localizacao"]: x for x in inventario.salas(dados, eid)}
    assert s["02 - SALA B"]["divergentes"] == 0                                # BAIXADO não conta nem aqui

    inventario.ler(dados, eid, "02 - SALA B", 1001, "Fulano")                  # controle: ATIVO de outra sala conta
    s = {x["localizacao"]: x for x in inventario.salas(dados, eid)}
    assert s["02 - SALA B"]["divergentes"] == 1
    assert (s["01 - SALA CCI"]["localizados"], s["01 - SALA CCI"]["pendentes"]) == (0, 2)   # 1001 saiu da CCI
    assert s["02 - SALA B"]["localizados"] == 1                                             # 2001 continua localizado aqui

    inventario.encerrar_evento(dados, eid)
    with pytest.raises(db.ErroDeNegocio):
        inventario.ler(dados, eid, "01 - SALA CCI", 1002, "Fulano")


def test_bens_da_sala_e_atualizar_leitura(dados):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.ler(dados, eid, "01 - SALA CCI", 2002, "Fulano")                 # trazido da SALA B
    inventario.ler(dados, eid, "02 - SALA B", 1002, "Fulano")                   # bem da CCI lido na SALA B
    inventario.ler(dados, eid, "01 - SALA CCI", 1003, "Fulano")                 # BAIXADO da própria sala
    d = inventario.bens_da_sala(dados, eid, "01 - SALA CCI")
    por = {b["numero"]: b for b in d["bens"]}
    assert set(por) == {1001, 1002}                                              # só ativos da sala
    assert por[1001]["situacao_inv"] == "localizado" and por[1002]["situacao_inv"] == "divergente" and por[1002]["lido_em_sala"] == "02 - SALA B"
    assert [t["numero"] for t in d["trazidos"]] == [1003, 2002] and d["sobras"] == []
    inventario.atualizar_leitura(dados, eid, 1001, conservacao="Ruim", quem_usa="Ciclana", observacao="pé quebrado")
    b = {x["numero"]: x for x in inventario.bens_da_sala(dados, eid, "01 - SALA CCI")["bens"]}[1001]
    assert (b["conservacao"], b["quem_usa"], b["observacao"]) == ("Ruim", "Ciclana", "pé quebrado")
    inventario.atualizar_leitura(dados, eid, 1001, conservacao="")                # limpa
    assert {x["numero"]: x for x in inventario.bens_da_sala(dados, eid, "01 - SALA CCI")["bens"]}[1001]["conservacao"] is None
    with pytest.raises(db.ErroDeNegocio):
        inventario.atualizar_leitura(dados, eid, 1001, conservacao="Ótimo")
    with pytest.raises(db.ErroDeNegocio):
        inventario.atualizar_leitura(dados, eid, 2001, quem_usa="x")             # sem leitura


def test_sobras(dados):
    eid = semear_inventario(dados)
    with pytest.raises(db.ErroDeNegocio):
        inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "", "", "achado", "http://x/1.webp", "Fulano")
    with pytest.raises(db.ErroDeNegocio):
        inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "VENTILADOR", "", "", "http://x/1.webp", "Fulano")
    with pytest.raises(db.ErroDeNegocio):
        inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "VENTILADOR", "", "achado", "", "Fulano")
    with pytest.raises(db.ErroDeNegocio):
        inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "VENTILADOR", "", "achado", "", "Ninguém", exigir_foto=False)
    sid = inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "VENTILADOR", "", "achado", "", "Fulano", exigir_foto=False)
    inventario.definir_foto_sobra(dados, sid, "http://x/1.webp")
    s = inventario.bens_da_sala(dados, eid, "01 - SALA CCI")["sobras"]
    assert len(s) == 1 and s[0]["foto_url"] == "http://x/1.webp" and inventario.resumo(dados, eid)["sobras"] == 1
    with pytest.raises(db.ErroDeNegocio):
        inventario.excluir_sobra(dados, eid, 999)
    assert inventario.excluir_sobra(dados, eid, sid)["descricao"] == "VENTILADOR"
    assert inventario.resumo(dados, eid)["sobras"] == 0
    sid2 = inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "TABLET", "", "achado", "", "Fulano", exigir_foto=False)
    inventario.encerrar_evento(dados, eid)
    with pytest.raises(db.ErroDeNegocio):
        inventario.definir_foto_sobra(dados, sid2, "http://x/2.webp")


def test_relatorio_e_xlsx(dados, tmp_path):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.ler(dados, eid, "01 - SALA CCI", 2001, "Fulano")                     # divergente (é da SALA B)
    inventario.atualizar_leitura(dados, eid, 1001, conservacao="Bom", quem_usa="Ciclana")
    inventario.registrar_sobra(dados, eid, "02 - SALA B", "VENTILADOR", "ARNO", "sem plaqueta", "http://x/s.webp", "Beltrana")
    inventario.ler(dados, eid, "01 - SALA CCI", 1003, "Fulano")                     # BAIXADO lido na própria sala
    r = inventario.relatorio(dados, eid)
    b1003 = next(x for x in r if x["numero"] == 1003)
    assert b1003["situacao_bem"] == "BAIXADO" and b1003["situacao_inv"] == "localizado"
    assert [x["numero"] for x in inventario.relatorio(dados, eid, localizacao="01 - SALA CCI")] == [1001, 1002, 1003, 2001]
    with pytest.raises(db.ErroDeNegocio):
        inventario.relatorio(dados, 999)
    with pytest.raises(db.ErroDeNegocio):
        inventario.exportar_xlsx(dados, 999, tmp_path / "x.xlsx")
    assert [(x["numero"], x["situacao_inv"]) for x in r] == [(1001, "localizado"), (1002, "pendente"), (1003, "localizado"), (2001, "divergente"), (2002, "pendente"), (1004, "pendente")]
    assert r[3]["local_sistema"] == "02 - SALA B" and r[3]["local_inventario"] == "01 - SALA CCI"
    assert [x["numero"] for x in inventario.relatorio(dados, eid, situacao="pendente")] == [1002, 2002, 1004]
    from openpyxl import load_workbook
    wb = load_workbook(inventario.exportar_xlsx(dados, eid, tmp_path / "inv.xlsx"))
    assert wb.sheetnames == ["Bens", "Sobras"]
    linhas = list(wb["Bens"].iter_rows(values_only=True))
    assert linhas[0][0] == "Inventário 2026" and linhas[4] == tuple(inventario.COLUNAS_XLSX)
    assert linhas[5][:3] == (1001, "CADEIRA", "GIRATÓRIA") and linhas[5][6] == "Localizado" and linhas[5][7] == "Bom"
    sobras = list(wb["Sobras"].iter_rows(values_only=True))
    assert sobras[1][0] == "Sala" and sobras[2][:2] == ("02 - SALA B", "VENTILADOR") and sobras[2][6] == "http://x/s.webp"
    so_cci = load_workbook(inventario.exportar_xlsx(dados, eid, tmp_path / "cci.xlsx", localizacao="01 - SALA CCI"))["Bens"]
    assert so_cci.max_row == 5 + 4


def test_xlsx_cabecalho_filtros_e_fotos(dados, tmp_path):
    from openpyxl import load_workbook
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.atualizar_leitura(dados, eid, 1001, foto_url="https://x/1001.webp")
    inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "VENTILADOR", None, "sem plaqueta", "https://x/s.webp", "Fulano")
    wb = load_workbook(inventario.exportar_xlsx(dados, eid, tmp_path / "a.xlsx", situacao="localizado", fotos=True))
    ws = wb["Bens"]
    linhas = list(ws.iter_rows(values_only=True))
    assert linhas[0][0] == "Inventário 2026" and linhas[1][0].startswith("Gerado em ")
    assert linhas[2][0] == "Todas as salas · Situação Localizado" and linhas[3][0] == "Total de bens: 1"
    assert linhas[4] == tuple(inventario.COLUNAS_XLSX) and linhas[5][0] == 1001 and len(linhas) == 6
    assert linhas[5][12] == '=_xlfn.IMAGE("https://x/1001.webp")' and ws.row_dimensions[6].height == 60
    assert wb["Sobras"].cell(row=3, column=7).value == '=_xlfn.IMAGE("https://x/s.webp")'
    ws = load_workbook(inventario.exportar_xlsx(dados, eid, tmp_path / "b.xlsx", fotos=True))["Bens"]
    assert ws.cell(row=6, column=13).value.startswith("=_xlfn") and ws.cell(row=7, column=13).value == "-"   # 1002 sem foto
    ws = load_workbook(inventario.exportar_xlsx(dados, eid, tmp_path / "c.xlsx"))["Bens"]
    assert ws.cell(row=6, column=13).value == "https://x/1001.webp" and ws.row_dimensions[6].height is None


def test_planilha_de_cadastros_exporta_e_importa_abas_de_inventario(dados, tmp_path):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.registrar_sobra(dados, eid, "02 - SALA B", "VENTILADOR", None, "achado", "http://x/s.webp", "Fulano")
    from openpyxl import load_workbook
    caminho = db.exportar_cadastros(dados, tmp_path / "c.xlsx")
    wb = load_workbook(caminho)
    assert wb.sheetnames == ["responsaveis", "localizacoes", "pessoas", "atribuicoes", "inv_eventos", "inv_integrantes", "inv_salas", "inv_leituras", "inv_sobras", "inv_bens_encerrados"]
    assert list(wb["inv_leituras"].iter_rows(values_only=True))[1][:3] == (eid, 1001, "01 - SALA CCI")
    # editar: encerra o evento, acrescenta uma leitura migrada de outro sistema, e reimporta
    ws = wb["inv_eventos"]
    ws.cell(row=2, column=5, value="2026-01-31")                                       # encerrado_em só com data
    wb["inv_leituras"].append([eid, 2001, "02 - SALA B", "2026-01-20 10:00:00", "Antigo", "Regular", "", "migrado", ""])
    wb.save(tmp_path / "c2.xlsx")
    with open(tmp_path / "c2.xlsx", "rb") as f:
        r = db.importar_cadastros(dados, f)
    assert r["inv_leituras"] == 2 and r["inv_eventos"] == 1
    assert inventario.evento(dados, eid)["encerrado_em"] == "2026-01-31 00:00:00"
    assert {x["numero"]: x["situacao_inv"] for x in inventario.relatorio(dados, eid)}[2001] == "localizado"
    assert inventario.resumo(dados, eid)["sobras"] == 1


def test_planilha_sem_abas_de_inventario_nao_toca_nas_tabelas(dados, tmp_path):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    from openpyxl import load_workbook
    wb = load_workbook(db.exportar_cadastros(dados, tmp_path / "c.xlsx"))
    for aba in list(inventario.ABAS):
        wb.remove(wb[aba])
    wb.save(tmp_path / "so4.xlsx")
    with open(tmp_path / "so4.xlsx", "rb") as f:
        r = db.importar_cadastros(dados, f)
    assert "inv_leituras" not in r and inventario.resumo(dados, eid)["lidos"] == 1


def test_planilha_com_abas_de_inventario_incompletas_e_recusada(dados, tmp_path):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    from openpyxl import load_workbook
    wb = load_workbook(db.exportar_cadastros(dados, tmp_path / "c.xlsx"))
    for aba in ("inv_integrantes", "inv_salas", "inv_leituras", "inv_sobras"):   # sobra só inv_eventos
        wb.remove(wb[aba])
    wb.save(tmp_path / "parcial.xlsx")
    with open(tmp_path / "parcial.xlsx", "rb") as f, pytest.raises(db.ImportacaoInvalida) as e:
        db.importar_cadastros(dados, f)
    assert "incompletas" in str(e.value) and "inv_leituras" in str(e.value)
    assert inventario.resumo(dados, eid)["lidos"] == 1 and inventario.evento(dados, eid) is not None


def test_planilha_de_inventario_validacoes(dados, tmp_path):
    eid = semear_inventario(dados)
    from openpyxl import load_workbook
    wb = load_workbook(db.exportar_cadastros(dados, tmp_path / "c.xlsx"))
    wb["inv_leituras"].append([eid, 99999, "01 - SALA CCI", "2026-01-20 10:00:00", "Fulano", "", "", "", ""])   # bem inexistente
    wb["inv_leituras"].append([eid, 1001, "01 - SALA CCI", "x", "Fulano", "Ótimo", "", "", ""])                  # data e conservação
    wb["inv_salas"].append([77, "01 - SALA CCI"])                                                                # evento inexistente
    wb["inv_eventos"].append([2, "Outro aberto", None, "2026-02-01 00:00:00", None])                             # 2 abertos
    wb.save(tmp_path / "ruim.xlsx")
    with open(tmp_path / "ruim.xlsx", "rb") as f, pytest.raises(db.ImportacaoInvalida) as e:
        db.importar_cadastros(dados, f)
    msg = str(e.value)
    assert "99999" in msg and "conservação" in msg and "77" in msg and "aberto" in msg and "data" in msg
    assert inventario.resumo(dados, eid)["salas"] == 3                                                            # nada mudou


def test_titulo_do_xlsx_do_inventario_e_texto_literal(dados, tmp_path):
    from openpyxl import load_workbook
    eid = semear_inventario(dados)
    dados.execute("UPDATE inventario_eventos SET nome = '=1+1' WHERE id = ?", (eid,))
    dados.commit()
    wb = load_workbook(inventario.exportar_xlsx(dados, eid, tmp_path / "inv.xlsx"))
    assert wb["Bens"]["A1"].data_type == "s" and wb["Bens"]["A1"].value.startswith("=1+1")
    assert wb["Sobras"]["A1"].data_type == "s"


def test_encerrar_grava_snapshot_e_congela_o_evento(dados):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.ler(dados, eid, "01 - SALA CCI", 1003, "Fulano")          # BAIXADO lido: entra no snapshot por ter leitura
    inventario.encerrar_evento(dados, eid)
    snap = {r[0] for r in dados.execute("SELECT numero FROM inventario_bens_encerrados WHERE evento_id = ?", (eid,))}
    assert snap == {1001, 1002, 1003, 1004, 2001, 2002}
    inventario.encerrar_evento(dados, eid)                                 # idempotente: não regrava
    assert dados.execute("SELECT count(*) FROM inventario_bens_encerrados").fetchone()[0] == 6
    # o export do SPW do ano seguinte muda `bens`; o evento encerrado não muda
    dados.execute("UPDATE bens SET localizacao = '02 - SALA B', descricao = 'CADEIRA NOVA' WHERE numero = 1001")
    dados.execute("DELETE FROM bens WHERE numero = 2002")
    dados.execute("INSERT INTO bens VALUES (3001,'ATIVO','TV','LG','EQUIPAMENTOS','01 - SALA CCI','01/01/2027',1,1)")
    dados.commit()
    assert [(s["localizacao"], s["total"], s["localizados"]) for s in inventario.salas(dados, eid)] == \
        [("01 - SALA CCI", 2, 1), ("02 - SALA B", 2, 0), ("99 - SEM MAPA", 1, 0)]
    r = {x["numero"]: x for x in inventario.relatorio(dados, eid)}
    assert set(r) == {1001, 1002, 1003, 1004, 2001, 2002}
    assert r[1001]["descricao"] == "CADEIRA" and r[1001]["situacao_inv"] == "localizado" and r[1003]["situacao_bem"] == "BAIXADO"
    assert inventario.resumo(dados, eid)["bens"] == 5
    assert [b["numero"] for b in inventario.bens_da_sala(dados, eid, "01 - SALA CCI")["bens"]] == [1001, 1002]


def test_evento_encerrado_sem_snapshot_le_bens(dados):
    eid = semear_inventario(dados)
    dados.execute("UPDATE inventario_eventos SET encerrado_em = '2026-01-01 00:00:00' WHERE id = ?", (eid,))
    dados.commit()
    assert inventario.resumo(dados, eid)["bens"] == 5 and inventario._fonte_bens(dados, eid) == "bens"


def test_aba_inv_bens_encerrados_exporta_importa_e_valida(dados, tmp_path):
    from openpyxl import load_workbook
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.encerrar_evento(dados, eid)
    wb = load_workbook(db.exportar_cadastros(dados, tmp_path / "c.xlsx"))
    assert wb.sheetnames[-1] == "inv_bens_encerrados"
    linhas = list(wb["inv_bens_encerrados"].iter_rows(values_only=True))
    assert linhas[0] == tuple(inventario.ABAS["inv_bens_encerrados"])
    assert linhas[1] == (eid, 1001, "ATIVO", "CADEIRA", "GIRATÓRIA", "MÓVEIS", "01 - SALA CCI") and len(linhas) == 6
    dados.execute("DELETE FROM bens WHERE numero = 1004")          # o SPW mudou; o snapshot importado preserva 1004
    dados.commit()
    with open(tmp_path / "c.xlsx", "rb") as f:
        r = db.importar_cadastros(dados, f)
    assert r["inv_bens_encerrados"] == 5
    assert {x["numero"] for x in inventario.relatorio(dados, eid)} == {1001, 1002, 1004, 2001, 2002}
    # aba ausente (planilha exportada por versão anterior, com 5 abas inv_*): snapshot não é tocado
    wb.remove(wb["inv_bens_encerrados"])
    wb.save(tmp_path / "cinco.xlsx")
    with open(tmp_path / "cinco.xlsx", "rb") as f:
        r = db.importar_cadastros(dados, f)
    assert "inv_bens_encerrados" not in r and dados.execute("SELECT count(*) FROM inventario_bens_encerrados").fetchone()[0] == 5
    # validações: evento aberto, número inválido, repetido
    wb = load_workbook(tmp_path / "c.xlsx")
    wb["inv_eventos"].append([2, "Outro aberto", None, "2026-02-01 00:00:00", None])   # evento 2 existe mas não está encerrado
    wb["inv_bens_encerrados"].append([2, 1001, "", "", "", "", ""])
    wb["inv_bens_encerrados"].append([eid, "abc", "", "", "", "", ""])
    wb["inv_bens_encerrados"].append([eid, 1001, "", "", "", "", ""])                  # repete a linha 2 (eid, 1001)
    wb.save(tmp_path / "ruim.xlsx")
    with open(tmp_path / "ruim.xlsx", "rb") as f, pytest.raises(db.ImportacaoInvalida) as ex:
        db.importar_cadastros(dados, f)
    msg = str(ex.value)
    assert "não está encerrado" in msg and "número inválido" in msg and "repetido" in msg


def test_andar():
    assert inventario.andar("07 - COAD - SALA DE REUNIÃO") == "07"
    assert inventario.andar("02 - CCOM") == "02"
    assert inventario.andar("TERMOS INDIVIDUAIS") == inventario.ANDAR_SEM
    assert inventario.andar("") == inventario.ANDAR_SEM
    assert inventario.andar(" - X") == inventario.ANDAR_SEM


def test_ler_lote_e_desfazer_leituras(dados):
    eid = semear_inventario(dados)
    r = inventario.ler_lote(dados, eid, "01 - SALA CCI", [1001, 2001, 99999], "Fulano")
    assert r == {"lidos": 2, "nao_encontrados": [99999]}
    assert {b["numero"]: b["situacao_inv"] for b in inventario.bens_da_sala(dados, eid, "01 - SALA CCI")["bens"]} == {1001: "localizado", 1002: "pendente"}
    assert inventario.resumo(dados, eid)["divergentes"] == 1
    with pytest.raises(db.ErroDeNegocio):
        inventario.ler_lote(dados, eid, "01 - SALA CCI", [1002], "Ninguém")
    inventario.atualizar_leitura(dados, eid, 1001, foto_url="http://x/1001.webp")
    assert inventario.desfazer_leituras(dados, eid, [1001, 2001, 1004]) == ["http://x/1001.webp"]   # 1004 sem leitura: ignorado
    assert inventario.resumo(dados, eid)["lidos"] == 0 and inventario.resumo(dados, eid)["divergentes"] == 0
    assert inventario.desfazer_leituras(dados, eid, []) == []
    inventario.encerrar_evento(dados, eid)
    with pytest.raises(db.ErroDeNegocio):
        inventario.desfazer_leituras(dados, eid, [1002])
    with pytest.raises(db.ErroDeNegocio):
        inventario.ler_lote(dados, eid, "01 - SALA CCI", [1002], "Fulano")


def test_relatorio_filtros_busca_e_ordem(dados):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.ler(dados, eid, "02 - SALA B", 2001, "Beltrana")
    inventario.atualizar_leitura(dados, eid, 1001, conservacao="Ruim", quem_usa="José", foto_url="https://x/1001.webp")
    inventario.atualizar_leitura(dados, eid, 2001, observacao="tela quebrada")
    num = lambda **f: [x["numero"] for x in inventario.relatorio(dados, eid, **f)]
    assert num() == [1001, 1002, 2001, 2002, 1004]
    assert num(integrante="Fulano") == [1001] and num(integrante="Beltrana") == [2001]
    assert num(conservacao="Ruim") == [1001] and num(conservacao="-") == [1002, 2001, 2002, 1004]
    assert num(foto="com") == [1001] and num(foto="sem") == [1002, 2001, 2002, 1004]
    assert num(busca="jose") == [1001] and num(busca="QUEBRADA tela") == [2001]
    assert num(busca="cadeira giratoria") == [1001] and num(busca="20") == [2001, 2002] and num(busca="   ") == num()
    assert num(ordem="numero", dir="desc") == [2002, 2001, 1004, 1002, 1001]
    assert num(ordem="descricao") == [1004, 1001, 2002, 2001, 1002]            # ARMÁRIO, CADEIRA, IMPRESSORA, MONITOR, NOTEBOOK
    assert num(ordem="integrante") == [2001, 1001, 1002, 2002, 1004]           # vazios sempre por último
    assert num(ordem="integrante", dir="desc") == [1001, 2001, 1002, 2002, 1004]
    assert num(ordem="inexistente") == num()
    assert num(localizacao="02 - SALA B", situacao="pendente") == [2002]
    assert inventario.contar_fotos(inventario.relatorio(dados, eid)) == 1
    assert inventario.descrever_filtros({}) == "Todas as salas"
    assert inventario.descrever_filtros({"localizacao": "01 - SALA CCI", "situacao": "divergente", "integrante": "Fulano",
                                         "conservacao": "-", "foto": "com", "busca": " cadeira ", "ordem": "numero"}) == \
        'Sala 01 - SALA CCI · Situação Divergente · Integrante Fulano · Conservação Não informada · Com foto · Busca "cadeira"'


def test_painel_dados(dados):
    eid = semear_inventario(dados)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Fulano")
    inventario.ler(dados, eid, "01 - SALA CCI", 2001, "Beltrana")      # divergente (é da SALA B)
    inventario.ler(dados, eid, "01 - SALA CCI", 1002, "Fulano")
    inventario.atualizar_leitura(dados, eid, 1001, conservacao="Ruim")
    p = inventario.painel(dados, eid)
    assert p["resumo"]["lidos"] == 2
    assert [(x["chave"], x["rotulo"], x["quantidade"]) for x in p["situacao"]] == \
        [("localizado", "Localizado", 2), ("divergente", "Divergente", 1), ("pendente", "Não localizado", 3)]
    assert [(x["chave"], x["quantidade"]) for x in p["integrantes"]] == [("Fulano", 2), ("Beltrana", 1)]
    assert [(x["chave"], x["rotulo"], x["quantidade"]) for x in p["conservacao"]] == [("Ruim", "Ruim", 1), ("-", "Não informada", 2)]
    assert p["andares"] == [{"andar": "01", "total": 2, "localizados": 2, "pendentes": 0, "divergentes": 1, "salas": 1},
                            {"andar": "02", "total": 2, "localizados": 0, "pendentes": 2, "divergentes": 0, "salas": 1},
                            {"andar": "99", "total": 1, "localizados": 0, "pendentes": 1, "divergentes": 0, "salas": 1}]
    assert p["salas_do_andar"] == []
    assert inventario.painel(dados, eid, "02")["salas_do_andar"] == \
        [{"localizacao": "02 - SALA B", "total": 2, "localizados": 0, "pendentes": 2, "divergentes": 0}]
    with pytest.raises(db.ErroDeNegocio):
        inventario.painel(dados, 999)


def test_painel_sala_sem_andar_vai_por_ultimo(dados):
    semear(dados)
    dados.execute("INSERT INTO bens VALUES (5001,'ATIVO','QUADRO','','MÓVEIS','TERMOS INDIVIDUAIS','01/01/2020',1,1)")
    dados.commit()
    eid = inventario.abrir_evento(dados, "Inv", None, ["Fulano"])
    assert [a["andar"] for a in inventario.painel(dados, eid)["andares"]] == ["01", "99", inventario.ANDAR_SEM]
    assert [x["chave"] for x in inventario.painel(dados, eid)["integrantes"]] == [] and inventario.painel(dados, eid)["conservacao"] == []
