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


def test_cadastro_responsaveis_incluir_renomear_excluir(cliente):
    r = cliente.post("/cadastros/responsaveis/incluir", data={"ccustos": "decom", "tratamento": "Prezado",
                     "responsavel": "THIAGO", "email": "", "matricula": "481", "funcao": "gerente"}, follow_redirects=True)
    assert b"DECOM" in r.data
    r = cliente.post("/cadastros/responsaveis/renomear", data={"antigo": "CCI", "novo": "GESERV"}, follow_redirects=True)
    assert b"GESERV" in r.data and b">CCI<" not in r.data
    r = cliente.post("/cadastros/responsaveis/excluir", data={"ccustos": "GESERV"}, follow_redirects=True)
    assert "Remapeie".encode() in r.data


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
