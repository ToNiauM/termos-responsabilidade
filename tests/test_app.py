import io
from urllib.parse import unquote

import pytest
from openpyxl import Workbook

import db
from tests.conftest import semear
from tests.test_db import CABECALHO


@pytest.fixture
def cliente(dados):
    semear(dados)
    from app import app
    app.config["TESTING"] = True
    with app.test_client() as c:
        yield c


def test_home_e_busca_de_bem(cliente):
    assert cliente.get("/").status_code == 200
    r = cliente.get("/bem?numero=1002")
    assert b"NOTEBOOK" in r.data and b"ANA SILVA" in r.data and b"JAQUELINE PORTELA" in r.data
    r = cliente.get("/bem?numero=9999", follow_redirects=True)
    assert "não encontrado".encode() in r.data


def test_termo_ccusto_documento_docx_planilha(cliente):
    r = cliente.post("/gerar", data={"ccusto": "CCI"})
    assert r.status_code == 302 and r.headers["Location"].endswith("/termo/ccusto/CCI")
    assert cliente.get("/termo/ccusto/CCI").status_code == 200
    doc = cliente.get("/termo/ccusto/CCI/documento").data.decode()
    assert "width:100%" in doc and "1001" in doc and "1002" not in doc  # 1002 está com ANA
    assert cliente.get("/termo/ccusto/CCI/docx").headers["Content-Disposition"].endswith('Termo_de_Responsabilidade_CCI.docx')
    assert cliente.get("/termo/ccusto/CCI/planilha").headers["Content-Disposition"].endswith('planilha_CCI.xlsx')


def test_termo_individual(cliente):
    r = cliente.post("/gerar-individual", data={"nome": "ANA SILVA"})
    assert r.status_code == 302
    doc = cliente.get("/termo/individual/ANA SILVA/documento").data.decode()
    assert "width:80%" in doc and "NOTEBOOK" in doc
    assert cliente.get("/termo/individual/ANA SILVA/docx").status_code == 200
    assert cliente.get("/termo/individual/NINGUEM").status_code == 404


def test_upload_importa_e_lista_sem_centro(cliente):
    wb = Workbook()
    ws = wb.active
    ws.append(CABECALHO)
    ws.append([1002, "ATIVO", "NOTEBOOK", "DELL", "EQ", "01 - SALA CCI", "x", 1, 1])
    ws.append([5000, "ATIVO", "TV", "", "EQ", "77 - NOVA SALA", "x", 1, 1])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    r = cliente.post("/upload", data={"arquivo": (buf, "export.xlsx")}, content_type="multipart/form-data",
                     follow_redirects=True)
    assert "2 bens importados (2 ativos)".encode() in r.data and b"77 - NOVA SALA" in r.data


def test_upload_invalido_mostra_erro(cliente):
    r = cliente.post("/upload", data={"arquivo": (io.BytesIO(b"nada"), "x.txt")}, content_type="multipart/form-data",
                     follow_redirects=True)
    assert b".xlsx" in r.data


def test_cadastro_responsaveis_editar_renomear_excluir(cliente):
    r = cliente.post("/cadastros/responsaveis/incluir", data={"ccustos": "decom", "tratamento": "Prezado",
                     "responsavel": "THIAGO", "email": "", "matricula": "481", "funcao": "gerente"}, follow_redirects=True)
    assert b"DECOM" in r.data
    r = cliente.get("/cadastros/responsaveis/CCI/editar")
    assert b"JAQUELINE PORTELA" in r.data and b"01 - SALA CCI" in r.data
    r = cliente.post("/cadastros/responsaveis/CCI/editar", data={"ccustos": "geserv", "tratamento": "Prezado",
                     "responsavel": "CARLOS", "email": "", "matricula": "7", "funcao": "gerente"}, follow_redirects=True)
    assert b"GESERV" in r.data and b"CARLOS" in r.data and b">CCI<" not in r.data
    assert cliente.get("/cadastros/responsaveis/CCI/editar").status_code == 404
    # excluir: pede confirmação; com bens sob guarda, bloqueia
    r = cliente.post("/cadastros/responsaveis/excluir", data={"ccustos": "GESERV"}, follow_redirects=True)
    assert b"sob guarda" in r.data
    r = cliente.post("/cadastros/responsaveis/excluir", data={"ccustos": "DECOM"}, follow_redirects=True)
    assert b"Confirmar exclus" in r.data
    r = cliente.post("/cadastros/responsaveis/excluir", data={"ccustos": "DECOM", "confirmar": "1"}, follow_redirects=True)
    assert b">DECOM<" not in r.data


def test_editar_responsavel_invalido_nao_renomeia(cliente, dados):
    r = cliente.post("/cadastros/responsaveis/CCI/editar", data={"ccustos": "GESERV", "responsavel": ""},
                     follow_redirects=True)
    assert b"obrigat" in r.data
    assert db.responsavel(dados, "CCI") is not None
    assert db.responsavel(dados, "GESERV") is None


def test_cadastro_localizacoes_mover(cliente):
    cliente.post("/cadastros/responsaveis/incluir", data={"ccustos": "PRES", "responsavel": "Y"})
    cliente.post("/cadastros/localizacoes/incluir", data={"localizacao": "99 - SEM MAPA", "ccustos": "CCI"})
    r = cliente.post("/cadastros/localizacoes/mover", data={"localizacoes": ["01 - SALA CCI", "99 - SEM MAPA"], "ccustos_destino": "PRES"},
                     follow_redirects=True)
    assert b"2 localiza" in r.data
    assert b"PRES" in cliente.get("/bem?numero=1001").data
    r = cliente.post("/cadastros/localizacoes/mover", data={"ccustos_destino": "PRES"}, follow_redirects=True)
    assert b"ao menos uma" in r.data


def test_cadastro_pessoas_editar_nome(cliente):
    r = cliente.get("/cadastros/pessoas/ANA SILVA/editar")
    assert b"ANA SILVA" in r.data
    r = cliente.post("/cadastros/pessoas/ANA SILVA/editar", data={"nome": "ana souza"}, follow_redirects=True)
    assert b"ANA SOUZA" in r.data and b"NOTEBOOK" in r.data
    assert cliente.get("/cadastros/pessoas/NINGUEM/editar").status_code == 404


def test_cadastro_localizacoes(cliente):
    r = cliente.get("/cadastros/localizacoes")
    assert b"99 - SEM MAPA" in r.data
    r = cliente.post("/cadastros/localizacoes/incluir", data={"localizacao": "99 - SEM MAPA", "ccustos": "CCI"}, follow_redirects=True)
    assert r.data.count(b"99 - SEM MAPA") >= 1 and b"pendente" not in r.data.lower()
    cliente.post("/cadastros/localizacoes/excluir", data={"localizacao": "99 - SEM MAPA"})
    assert b"99 - SEM MAPA" in cliente.get("/cadastros/localizacoes").data


def test_cadastro_pessoas_atribuir_com_confirmacao(cliente):
    cliente.post("/cadastros/pessoas/incluir", data={"nome": "bruno lima"})
    r = cliente.get("/cadastros/pessoas?nome=BRUNO LIMA")
    assert b"BRUNO LIMA" in r.data
    r = cliente.post("/cadastros/pessoas/atribuir", data={"nome": "BRUNO LIMA", "numero": "1002"}, follow_redirects=True)
    assert b"ANA SILVA" in r.data and b"confirmar" in r.data  # pede confirmação
    r = cliente.post("/cadastros/pessoas/atribuir", data={"nome": "BRUNO LIMA", "numero": "1002", "confirmar": "1002"},
                     follow_redirects=True)
    assert b"NOTEBOOK" in r.data
    r = cliente.post("/cadastros/pessoas/desatribuir", data={"nome": "BRUNO LIMA", "numero": "1002"}, follow_redirects=True)
    assert b"NOTEBOOK" not in r.data
    r = cliente.post("/cadastros/pessoas/atribuir", data={"nome": "BRUNO LIMA", "numero": "9999"}, follow_redirects=True)
    assert "não encontrado".encode() in r.data
    r = cliente.post("/cadastros/pessoas/excluir", data={"nome": "BRUNO LIMA"}, follow_redirects=True)
    assert b"confirmar" in r.data
    cliente.post("/cadastros/pessoas/excluir", data={"nome": "BRUNO LIMA", "confirmar": "1"})
    assert b"BRUNO LIMA" not in cliente.get("/cadastros/pessoas").data


def test_termo_devolucao_fluxo(cliente):
    r = cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "numero_bem": "1001"}, follow_redirects=True)
    assert b"CADEIRA" in r.data
    r = cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "numero_bem": "9999"}, follow_redirects=True)
    assert "não encontrado".encode() in r.data
    r = cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "gerar": "1"})
    assert unquote(r.headers["Location"]).endswith("/termo/devolucao/ANA SILVA")
    doc = cliente.get("/termo/devolucao/ANA SILVA/documento").data.decode()
    assert "TERMO DE DEVOLUÇÃO" in doc and "CADEIRA" in doc and "width:80%" in doc
    assert cliente.get("/termo/devolucao/ANA SILVA/docx").status_code == 200
    r = cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "remover": "1001"}, follow_redirects=True)
    assert b"CADEIRA" not in r.data


def test_termo_devolucao_normaliza_numero(cliente):
    cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "numero_bem": "1001"})
    r = cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "numero_bem": "01001"}, follow_redirects=True)
    assert r.data.count(b"CADEIRA") == 1
    r = cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "remover": "1001"}, follow_redirects=True)
    assert b"CADEIRA" not in r.data


def test_termo_devolucao_troca_de_pessoa_limpa_lista(cliente, dados):
    db.incluir_pessoa(dados, "BRUNO LIMA")
    cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "numero_bem": "1001"})
    r = cliente.post("/termo_devolucao", data={"nome": "BRUNO LIMA", "numero_bem": "1002"}, follow_redirects=True)
    assert b"CADEIRA" not in r.data and b"NOTEBOOK" in r.data


def test_textos_salvar_reflete_no_documento_e_restaurar(cliente):
    r = cliente.get("/textos")
    assert r.status_code == 200 and b"Compromissos" in r.data
    import textos
    dados_form = dict(textos.PADRAO, individual_abertura="TESTE {nome}.", orgao_nome="Órgão X")
    r = cliente.post("/textos", data=dados_form, follow_redirects=True)
    assert "Textos salvos".encode() in r.data and "Órgão X".encode() in r.data  # header usa orgao_nome
    doc = cliente.get("/termo/individual/ANA SILVA/documento").data.decode()
    assert "TESTE <b>ANA SILVA</b>." in doc
    r = cliente.post("/textos", data={"restaurar": "individual_abertura"}, follow_redirects=True)
    assert "Padrão restaurado".encode() in r.data
    doc = cliente.get("/termo/individual/ANA SILVA/documento").data.decode()
    assert "Pelo presente termo" in doc and "Órgão X".encode() in cliente.get("/").data


def test_textos_marcador_invalido_nao_grava_nada(cliente):
    import textos
    dados_form = dict(textos.PADRAO, individual_abertura="Eu {nomee}", cidade="Goiânia (GO)")
    r = cliente.post("/textos", data=dados_form, follow_redirects=True)
    assert b"nomee" in r.data
    doc = cliente.get("/termo/devolucao/ANA SILVA/documento").data.decode()
    assert "Goiânia" not in doc and "Brasília (DF)" in doc


def test_exportar_e_importar_cadastros(cliente):
    r = cliente.get("/cadastros/exportar")
    assert r.status_code == 200 and r.headers["Content-Disposition"].endswith("cadastros.xlsx")
    r = cliente.post("/importar-cadastros", data={"arquivo": (io.BytesIO(r.data), "cadastros.xlsx")},
                     content_type="multipart/form-data", follow_redirects=True)
    assert b"1 centro" in r.data and b"1 pessoa" in r.data
    r = cliente.post("/importar-cadastros", data={"arquivo": (io.BytesIO(b"nada"), "x.xlsx")},
                     content_type="multipart/form-data", follow_redirects=True)
    assert "inválido".encode() in r.data


def test_exportar_bens(cliente):
    r = cliente.get("/bens/exportar")
    assert r.status_code == 200 and r.headers["Content-Disposition"].endswith("bens.xlsx")
    assert b"Exportar bens" in cliente.get("/upload").data


def test_pesquisa_numero_vai_para_ficha_e_texto_lista(cliente):
    r = cliente.get("/pesquisa?q=1002")
    assert r.status_code == 302 and r.headers["Location"].endswith("/bem?numero=1002")
    r = cliente.get("/pesquisa?q=cci")
    assert r.status_code == 200
    assert b"JAQUELINE PORTELA" in r.data and b"CADEIRA" in r.data and b"/termo/ccusto/CCI" in r.data
    assert b"/pesquisa?ccusto=CCI" in r.data
    r = cliente.get("/pesquisa?q=ana")
    assert b"ANA SILVA" in r.data and b"/pesquisa?pessoa=ANA" in r.data
    r = cliente.get("/pesquisa?q=zzz")
    assert "Nada encontrado".encode() in r.data
    assert cliente.get("/pesquisa?q=9999").status_code == 200  # número inexistente cai na busca por texto


def test_pesquisa_filtra_bens_de_centro_e_pessoa(cliente):
    r = cliente.get("/pesquisa?ccusto=CCI")
    assert b"CADEIRA" in r.data and b"NOTEBOOK" not in r.data and b"/termo/ccusto/CCI" in r.data
    r = cliente.get("/pesquisa?pessoa=ANA SILVA")
    assert b"NOTEBOOK" in r.data and b"CADEIRA" not in r.data and b"/termo/individual/ANA" in r.data


def test_cabecalho_tem_campo_de_pesquisa(cliente):
    assert b'action="/pesquisa"' in cliente.get("/").data
