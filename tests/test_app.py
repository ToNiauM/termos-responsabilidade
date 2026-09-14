import io

import pytest
from openpyxl import Workbook

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
