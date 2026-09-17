import io
from datetime import date
from urllib.parse import unquote

import pytest
from openpyxl import Workbook, load_workbook

import db
from tests.conftest import semear, confirmar_revisao
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
    cliente.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "T", "numero_sei": "1111", "vigente": "1"})
    r = cliente.post("/gerar", data={"ccusto": "CCI"})
    assert r.status_code == 302 and r.headers["Location"].endswith("/termo/ccusto/CCI")
    assert cliente.get("/termo/ccusto/CCI").status_code == 200
    doc = cliente.get("/termo/ccusto/CCI/documento").data.decode()
    assert "width:100%" in doc and "1001" in doc and "1002" not in doc  # 1002 está com ANA
    assert cliente.get("/termo/ccusto/CCI/docx").headers["Content-Disposition"].endswith('Termo_de_Responsabilidade_CCI.docx')
    assert cliente.get("/termo/ccusto/CCI/planilha").headers["Content-Disposition"].endswith('planilha_CCI.xlsx')


def test_termo_individual(cliente):
    cliente.post("/cadastros/processos/incluir", data={"tipo": "individual", "descricao": "T", "numero_sei": "1111", "vigente": "1"})
    r = cliente.post("/gerar-individual", data={"nome": "ANA SILVA"})
    assert r.status_code == 302
    doc = cliente.get("/termo/individual/ANA SILVA/documento").data.decode()
    assert "width:80%" in doc and "NOTEBOOK" in doc
    assert cliente.get("/termo/individual/ANA SILVA/docx").status_code == 200
    assert cliente.get("/termo/individual/NINGUEM").status_code == 404


def test_termo_sem_processo_vigente_nao_emite(cliente):
    r = cliente.get("/termo/ccusto/CCI")
    assert b"Cadastre um processo SEI vigente" in r.data and b'id="copiar"' not in r.data
    r = cliente.get("/termo/ccusto/CCI/docx", follow_redirects=True)
    assert b"Cadastre um processo SEI vigente" in r.data
    assert cliente.get("/termo/ccusto/CCI/documento").status_code == 200    # prévia continua
    r = cliente.post("/termo/ccusto/CCI/registrar")
    assert r.status_code == 409 and "erro" in r.get_json()


def test_termo_com_processo_registra_ao_baixar_e_ao_copiar(cliente):
    import db
    cliente.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "T", "numero_sei": "2222", "vigente": "1"})
    assert cliente.get("/termo/ccusto/CCI/planilha").status_code == 200
    assert db.termos_emitidos(db.conectar()) == []     # planilha não registra emissão
    r = cliente.get("/termo/ccusto/CCI")
    assert b'id="copiar"' in r.data and b"Nenhum termo registrado" in r.data
    assert cliente.get("/termo/ccusto/CCI/docx").status_code == 200
    r = cliente.get("/termo/ccusto/CCI")
    assert "Último termo registrado".encode() in r.data and b"Bens iguais aos de hoje" in r.data
    j = cliente.post("/termo/ccusto/CCI/registrar").get_json()
    assert j["id"] and j["emitido_em"][:4] == str(date.today().year)
    assert len(db.termos_emitidos(db.conectar())) == 1     # docx + registrar no mesmo dia = 1 registro


def test_termo_devolucao_registra_no_processo_de_devolucao(cliente):
    cliente.post("/cadastros/processos/incluir", data={"tipo": "devolucao", "descricao": "D", "numero_sei": "3333", "vigente": "1"})
    cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "numero_bem": "1001"})
    assert cliente.get("/termo/devolucao/ANA SILVA/docx").status_code == 200
    import db
    t = db.termos_emitidos(db.conectar(), tipo="devolucao")[0]
    assert t["chave"] == "ANA SILVA" and t["quantidade"] == 1


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


def test_upload_mostra_mudancas_e_detalhe(cliente, tmp_path):
    from tests.test_db import xlsx
    arq = xlsx(tmp_path, [
        [1001, "ATIVO", "CADEIRA", "GIRATÓRIA", "MÓVEIS", "02 - OUTRA", "31/12/1996", 75.94, 64.54],
        [1002, "ATIVO", "NOTEBOOK", "DELL", "EQUIPAMENTOS", "01 - SALA CCI", "06/12/2012", 3000, 1500],
    ])
    with open(arq, "rb") as f:
        r = cliente.post("/upload", data={"arquivo": (f, "export.xlsx")}, content_type="multipart/form-data", follow_redirects=True)
    assert b"2 bens importados" in r.data and b"1 movido" in r.data and b"2 removido" in r.data and b"export.xlsx" in r.data
    import db
    iid = db.importacoes(db.conectar())[0]["id"]
    r = cliente.get(f"/importacoes/{iid}")
    assert r.status_code == 200 and b"02 - OUTRA" in r.data and b"1004" in r.data and b"removido" in r.data
    r = cliente.get("/bem?numero=1001")
    assert b"02 - OUTRA" in r.data and b"movido" in r.data


def test_upload_invalido_mostra_erro(cliente):
    r = cliente.post("/upload", data={"arquivo": (io.BytesIO(b"nada"), "x.txt")}, content_type="multipart/form-data",
                     follow_redirects=True)
    assert b".xlsx" in r.data


def test_cadastro_responsaveis_editar_renomear_excluir(cliente):
    r = cliente.post("/cadastros/responsaveis/incluir", data={"ccustos": "decom",
                     "responsavel": "THIAGO", "email": "", "matricula": "481", "funcao": "gerente"}, follow_redirects=True)
    assert b"DECOM" in r.data
    r = cliente.get("/cadastros/responsaveis/CCI/editar")
    assert b"JAQUELINE PORTELA" in r.data and b"01 - SALA CCI" in r.data
    r = cliente.post("/cadastros/responsaveis/CCI/editar", data={"ccustos": "geserv",
                     "responsavel": "CARLOS", "email": "", "matricula": "7", "funcao": "gerente"}, follow_redirects=True)
    assert b"GESERV" in r.data and b"CARLOS" in r.data and b">CCI<" not in r.data
    assert cliente.get("/cadastros/responsaveis/CCI/editar").status_code == 404
    # excluir: pede confirmação; com bens sob guarda, bloqueia
    r = cliente.post("/cadastros/responsaveis/excluir", data={"ccustos": "GESERV"}, follow_redirects=True)
    assert b"sob guarda" in r.data
    r = cliente.post("/cadastros/responsaveis/excluir", data={"ccustos": "DECOM"}, follow_redirects=True)
    assert b"Confirmar exclus" in r.data
    r = confirmar_revisao(cliente, r, "/cadastros/responsaveis/excluir")
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
    assert b"PRES" not in cliente.get("/bem?numero=1001").data
    r = confirmar_revisao(cliente, r, "/cadastros/localizacoes/mover")
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
    r = cliente.post("/cadastros/localizacoes/excluir", data={"localizacao": "99 - SEM MAPA"})
    confirmar_revisao(cliente, r, "/cadastros/localizacoes/excluir")
    assert b"99 - SEM MAPA" in cliente.get("/cadastros/localizacoes").data


def test_cadastro_pessoas_atribuir_com_confirmacao(cliente):
    cliente.post("/cadastros/pessoas/incluir", data={"nome": "bruno lima"})
    r = cliente.get("/cadastros/pessoas?nome=BRUNO LIMA")
    assert b"BRUNO LIMA" in r.data
    r = cliente.post("/cadastros/pessoas/atribuir", data={"nome": "BRUNO LIMA", "numero": "1002"}, follow_redirects=True)
    assert b"ANA SILVA" in r.data and b"confirmar" in r.data  # pede confirmação
    r = confirmar_revisao(cliente, r, "/cadastros/pessoas/atribuir")
    assert b"NOTEBOOK" in r.data
    r = cliente.post("/cadastros/pessoas/desatribuir", data={"nome": "BRUNO LIMA", "numero": "1002"}, follow_redirects=True)
    r = confirmar_revisao(cliente, r, "/cadastros/pessoas/desatribuir")
    assert b"NOTEBOOK" not in r.data
    r = cliente.post("/cadastros/pessoas/atribuir", data={"nome": "BRUNO LIMA", "numero": "9999"}, follow_redirects=True)
    assert "não encontrado".encode() in r.data
    r = cliente.post("/cadastros/pessoas/excluir", data={"nome": "BRUNO LIMA"}, follow_redirects=True)
    assert b"Confirmar" in r.data
    confirmar_revisao(cliente, r, "/cadastros/pessoas/excluir")
    assert b"BRUNO LIMA" not in cliente.get("/cadastros/pessoas").data


def test_termo_devolucao_fluxo(cliente):
    cliente.post("/cadastros/processos/incluir", data={"tipo": "devolucao", "descricao": "T", "numero_sei": "1111", "vigente": "1"})
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
    dados_form = dict(textos.PADRAO, individual_abertura="TESTE {nome}.", orgao_nome="Órgão X", unidade_sigla="SN")
    r = cliente.post("/textos", data=dados_form, follow_redirects=True)
    assert "Textos salvos".encode() in r.data and "Órgão X".encode() in r.data  # header usa orgao_nome
    assert b'<div class="header-subtitle">SN</div>' in r.data  # e unidade_sigla no subtítulo
    doc = cliente.get("/termo/individual/ANA SILVA/documento").data.decode()
    assert "TESTE <b>Ana Silva</b>." in doc
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


def test_importar_cadastros_reimporta_inventario(cliente):
    eid = _abrir(cliente)
    cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    r = cliente.get("/cadastros/exportar")
    r = cliente.post("/importar-cadastros", data={"arquivo": (io.BytesIO(r.data), "cadastros.xlsx")},
                     content_type="multipart/form-data", follow_redirects=True)
    assert "Inventário: 1 evento".encode() in r.data


def test_exportar_bens(cliente):
    r = cliente.get("/bens/exportar")
    assert r.status_code == 200 and r.headers["Content-Disposition"].endswith("bens.xlsx")
    assert b"Baixar" in cliente.get("/upload").data and b"Carregar" in cliente.get("/upload").data


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


def test_cadastro_de_processos_sei(cliente):
    r = cliente.get("/cadastros/processos")
    assert r.status_code == 200 and b"Processos SEI" in r.data
    r = cliente.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "Termos 2026", "numero_sei": "2222", "vigente": "1"}, follow_redirects=True)
    assert b"Termos 2026" in r.data and b"2222" in r.data and b"vigente" in r.data
    r = cliente.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "", "numero_sei": "1"}, follow_redirects=True)
    assert "Descrição".encode() in r.data
    import db
    pid = db.processos(db.conectar())[0]["id"]
    r = cliente.post("/cadastros/processos/encerrar", data={"id": pid}, follow_redirects=True)
    r = confirmar_revisao(cliente, r, "/cadastros/processos/encerrar")
    assert b"encerrado" in r.data
    r = cliente.post("/cadastros/processos/vigente", data={"id": pid}, follow_redirects=True)
    r = confirmar_revisao(cliente, r, "/cadastros/processos/vigente")
    assert b"vigente" in r.data
    r = cliente.post("/cadastros/processos/excluir", data={"id": pid}, follow_redirects=True)
    r = confirmar_revisao(cliente, r, "/cadastros/processos/excluir")
    assert "exclu".encode() in r.data and b"Termos 2026" not in r.data


def test_termos_emitidos_lista_detalhe_e_documento_sei(cliente):
    cliente.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "T", "numero_sei": "2222", "vigente": "1"})
    cliente.get("/termo/ccusto/CCI/docx")
    r = cliente.get("/termos-emitidos")
    assert r.status_code == 200 and b"CCI" in r.data and b"2222" in r.data
    assert b"CCI" not in cliente.get("/termos-emitidos?tipo=individual").data.split(b"<tbody>")[1]
    assert b"CCI" in cliente.get("/termos-emitidos?chave=cc").data
    import db
    tid = db.termos_emitidos(db.conectar())[0]["id"]
    r = cliente.get(f"/termos-emitidos/{tid}")
    assert b"CADEIRA" in r.data and b"1001" in r.data and b'name="documento_sei"' in r.data
    r = cliente.post(f"/termos-emitidos/{tid}/documento", data={"documento_sei": "0459999"}, follow_redirects=True)
    assert b"0459999" in r.data and b"Informe documento e bloco" in r.data and b"mailto:" not in r.data
    # com documento + bloco e e-mail do responsável (CCI tem j@cfc.org.br): link mailto com assunto e corpo
    r = cliente.post(f"/termos-emitidos/{tid}/documento", data={"documento_sei": "0459999", "bloco_sei": "77"}, follow_redirects=True)
    html = r.data.decode()
    assert "mailto:j@cfc.org.br?subject=" in html and "bloco%20de%20assinatura%2077" in html
    assert "Prezado%28a%29%20Jaqueline%2C" in html and "Jaqueline%20Portela" not in html    # só o primeiro nome
    assert "E-mail não enviado" in html
    r = cliente.post(f"/termos-emitidos/{tid}/email", follow_redirects=True)
    assert "E-mail enviado em" in r.data.decode() and "Enviar e-mail novamente" in r.data.decode()
    assert b"0459999" in cliente.get("/termos-emitidos").data and b"77" in cliente.get("/termos-emitidos").data
    assert cliente.get("/termos-emitidos/999").status_code == 404
    assert b"Termos emitidos" in cliente.get("/").data    # menu


def test_painel_na_tela_inicial(cliente):
    r = cliente.get("/")
    assert r.status_code == 200
    assert b"dsgov-atalhos" in r.data and "Realizar inventário".encode() in r.data and b'href="/inventario"' in r.data   # carrossel de atalhos
    assert b"bens ativos" in r.data and b'data-grafico="g-centro"' in r.data and b'data-grafico="g-ano"' in r.data
    assert b"echarts.min.js" in r.data and b"echarts-dsgov.js" in r.data
    assert b"/recorte?situacao=ATIVO&amp;ccusto=CCI" in r.data or b"/recorte?situacao=ATIVO&ccusto=CCI" in r.data
    assert b"/termo/ccusto/CCI" in r.data and b"sem termo" in r.data
    assert b"Nenhuma" in r.data       # última importação: nenhuma
    assert b"sem centro nem pessoa" in r.data


def test_listas_mostram_situacao_do_termo(cliente):
    r = cliente.get("/centro-custos")
    assert b"sem termo" in r.data and b"/termo/ccusto/CCI" in r.data
    r = cliente.get("/termos-individuais")
    assert b"ANA SILVA" in r.data and b"sem termo" in r.data and b"/termo/individual/ANA" in r.data
    cliente.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "T", "numero_sei": "2", "vigente": "1"})
    cliente.get("/termo/ccusto/CCI/docx")
    r = cliente.get("/centro-custos")
    assert b"vigente" in r.data and b"sem doc./bloco" in r.data
    import db
    tid = db.termos_emitidos(db.conectar())[0]["id"]
    cliente.post(f"/termos-emitidos/{tid}/documento", data={"documento_sei": "1", "bloco_sei": "2"})
    assert "não enviado".encode() in cliente.get("/centro-custos").data
    cliente.post(f"/termos-emitidos/{tid}/email")
    r = cliente.get("/centro-custos")
    assert b"enviado " in r.data and f"/termos-emitidos/{tid}".encode() in r.data


def test_recorte_tela_filtros_termo_e_xlsx(cliente):
    r = cliente.get("/recorte")
    assert r.status_code == 200 and b"Bens ATIVO" in r.data and b"1001" in r.data and b"1003" not in r.data
    assert b'data-grafico="g-situacao"' not in r.data and b'data-grafico="g-centro"' in r.data
    r = cliente.get("/recorte?situacao=ATIVO&ccusto=CCI")
    assert b"centro de custo CCI" in r.data and b"/termo/ccusto/CCI" in r.data and b"sem termo" in r.data
    assert b'data-grafico="g-centro"' not in r.data and b"/recorte?situacao=ATIVO&amp;ccusto=CCI&amp;faixa=" in r.data
    r = cliente.get("/recorte?pessoa=ANA SILVA")
    assert b"/termo/individual/ANA" in r.data and b"NOTEBOOK" in r.data and b"CADEIRA" not in r.data
    r = cliente.get("/recorte?situacao=&valor_de=1.000,00&valor_ate=2000")
    assert b"NOTEBOOK" in r.data and b"CADEIRA" not in r.data and b"todas as situa" in r.data
    r = cliente.get("/recorte?situacao=&valor_de=1.000&valor_ate=1.600")     # ponto de milhar
    assert b"NOTEBOOK" in r.data and b"CADEIRA" not in r.data
    r = cliente.get("/recorte?situacao=&valor_de=1000.5&valor_ate=1600")     # ponto decimal
    assert b"NOTEBOOK" in r.data
    r = cliente.get("/recorte?situacao=&valor_de=1.000.000")
    assert b"NOTEBOOK" not in r.data and b"Nenhum bem" in r.data
    r = cliente.get("/recorte?valor_de=abc", follow_redirects=True)
    assert "Valor inválido".encode() in r.data
    r = cliente.get("/recorte/xlsx?ccusto=CCI")
    assert r.status_code == 200 and r.headers["Content-Disposition"].endswith("recorte.xlsx")
    assert b"Recorte" in cliente.get("/").data       # menu
    r = cliente.get("/recorte?situacao=")
    assert b"ccusto=CCI&amp;situacao=" in r.data or b"ccusto=CCI&situacao=" in r.data   # drill-down mantém "todas"
    assert b'href="/recorte/xlsx?situacao="' in r.data
    ws = load_workbook(io.BytesIO(cliente.get("/recorte/xlsx?situacao=").data)).active
    assert ws.max_row - 1 == 4     # 4 bens da semente (inclui 1003 BAIXADO); a tela mostra o mesmo total


def _abrir(cliente, integrante="Fulano"):
    cliente.post("/inventario/abrir", data={"nome": "Inv", "integrantes": "Fulano\nBeltrana", "escopo": "todas"})
    import db, inventario
    eid = inventario.evento_aberto(db.conectar())["id"]
    if integrante:
        cliente.post(f"/inventario/{eid}/integrante", data={"integrante": integrante})
    return eid


def test_inventario_eventos_abrir_e_encerrar(cliente):
    r = cliente.get("/inventario")
    assert r.status_code == 200 and b"Abrir evento" in r.data and b"Nenhum evento aberto" in r.data
    r = cliente.post("/inventario/abrir", data={"nome": "Inventário 2026", "descricao": "Portaria 1", "integrantes": "Fulano\nBeltrana", "escopo": "todas"}, follow_redirects=True)
    assert "Inventário 2026".encode() in r.data and b"01 - SALA CCI" in r.data and b"99 - SEM MAPA" in r.data
    assert "Inventário".encode() in cliente.get("/").data                          # menu
    r = cliente.post("/inventario/abrir", data={"nome": "Outro", "integrantes": "X", "escopo": "todas"}, follow_redirects=True)
    assert "já existe".encode() in r.data.lower() or "Já existe".encode() in r.data
    import db, inventario
    eid = inventario.evento_aberto(db.conectar())["id"]
    r = cliente.post(f"/inventario/{eid}/integrante", data={"integrante": "Fulano", "volta": f"/inventario/{eid}"}, follow_redirects=True)
    assert b"Fulano" in r.data
    r = cliente.post(f"/inventario/{eid}/integrante", data={"integrante": "Fulano", "volta": "//evil.example"})
    assert r.headers["Location"].startswith("/inventario/")
    r = cliente.post(f"/inventario/{eid}/integrante", data={"integrante": "Fulano", "volta": "/\\evil.example"})
    assert r.headers["Location"].startswith("/inventario/")
    r = cliente.post(f"/inventario/{eid}/encerrar", data={}, follow_redirects=True)
    assert b"Confirmar encerramento" in r.data
    r = cliente.post(f"/inventario/{eid}/encerrar", data={"confirmar": "1"}, follow_redirects=True)
    assert b"encerrado" in r.data
    assert cliente.get("/inventario/999").status_code == 404


def test_inventario_abrir_com_amostragem(cliente):
    r = cliente.post("/inventario/abrir", data={"nome": "Amostra", "integrantes": "A", "escopo": "escolher", "salas": ["99 - SEM MAPA"]}, follow_redirects=True)
    assert b"99 - SEM MAPA" in r.data and b"01 - SALA CCI" not in r.data.split(b"<tbody>")[1]


def test_inventario_sala_leitura_json(cliente):
    eid = _abrir(cliente, integrante=None)
    r = cliente.get(f"/inventario/{eid}/sala/01 - SALA CCI")
    assert r.status_code == 200 and b'id="leitura"' in r.data and b"html5-qrcode" in r.data and b"Escolha o integrante" in r.data
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    assert r.status_code == 409 and "integrante" in r.get_json()["erro"].lower()
    cliente.post(f"/inventario/{eid}/integrante", data={"integrante": "Fulano"})
    j = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "001001"}).get_json()
    assert j["situacao"] == "localizado" and j["numero"] == 1001 and j["descricao"] == "CADEIRA" and j["reler"] is False
    j = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1004"}).get_json()
    assert j["situacao"] == "divergente" and j["cadastrado_em"] == "99 - SEM MAPA"
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "99999"})
    assert r.status_code == 404 and r.get_json()["numero"] == 99999
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "abc"})
    assert r.status_code == 404
    for numero in ("ABC1001", "S/N-1001", "12.345", "1001x"):
        r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": numero})
        assert r.status_code == 404
    j = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"}).get_json()
    assert j["reler"] and j["leitura_anterior"]["integrante"] == "Fulano"
    r = cliente.get(f"/inventario/{eid}/sala/01 - SALA CCI")
    assert b"Localizado" in r.data and b"Divergente" in r.data and b"1004" in r.data
    assert r.data.index(b'data-numero="1002"') < r.data.index(b'data-numero="1001"')   # pendente antes de localizado
    r = cliente.post(f"/inventario/{eid}/leitura/1001", json={"conservacao": "Ruim", "quem_usa": "Ciclana"})
    assert r.status_code == 200 and r.get_json()["ok"]
    assert cliente.post(f"/inventario/{eid}/leitura/1001", json={"conservacao": "Péssimo"}).status_code == 409
    assert cliente.post(f"/inventario/{eid}/leitura/1002", json={"observacao": "x"}).status_code == 409   # não lido


def test_inventario_fotos_e_sobras(cliente, monkeypatch, tmp_path):
    import io
    from PIL import Image
    import fotos
    eid = _abrir(cliente)
    buf = io.BytesIO(); Image.new("RGB", (30, 20), (1, 2, 3)).save(buf, "PNG"); imagem = buf.getvalue()
    # fotos desativadas: sobra sem foto é aceita; foto de bem recusada
    for v in fotos.VARIAVEIS:
        monkeypatch.delenv(v, raising=False)
    r = cliente.get(f"/inventario/{eid}/sala/01 - SALA CCI")
    assert b"Fotos desativadas" in r.data
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/sobra", data={"descricao": "VENTILADOR", "observacao": "sem plaqueta"}, follow_redirects=True)
    assert b"VENTILADOR" in r.data
    cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    r = cliente.post(f"/inventario/{eid}/leitura/1001/foto", data={"foto": (io.BytesIO(imagem), "a.png")}, content_type="multipart/form-data")
    assert r.status_code == 409 and "desativadas" in r.get_json()["erro"]
    # fotos ativas: cliente falso
    for v in fotos.VARIAVEIS:
        monkeypatch.setenv(v, "x")
    monkeypatch.setenv("R2_PUBLIC_URL", "https://f.exemplo.org")
    enviados = []
    monkeypatch.setattr(fotos, "enviar", lambda nome, dados: enviados.append(nome) or f"https://f.exemplo.org/inventario/{nome}")
    apagados = []
    monkeypatch.setattr(fotos, "apagar", lambda url: apagados.append(url))
    r = cliente.post(f"/inventario/{eid}/leitura/1001/foto", data={"foto": (io.BytesIO(imagem), "a.png")}, content_type="multipart/form-data")
    assert r.status_code == 200 and r.get_json()["foto_url"].startswith("https://f.exemplo.org/inventario/INV") and enviados[-1].startswith(f"INV{eid}_BEM_1001_")
    r = cliente.post(f"/inventario/{eid}/leitura/1001/foto/excluir", follow_redirects=True)
    assert apagados and b"Foto removida" in r.data
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/sobra", data={"descricao": "CADEIRA VELHA", "observacao": "x"}, follow_redirects=True)
    assert "precisa de foto".encode() in r.data
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/sobra", data={"descricao": "CADEIRA VELHA", "observacao": "x", "foto": (io.BytesIO(imagem), "b.jpg")}, content_type="multipart/form-data", follow_redirects=True)
    assert b"CADEIRA VELHA" in r.data and enviados[-1].startswith(f"INV{eid}_SOBRA_")
    # falha no envio: sobra não fica registrada
    monkeypatch.setattr(fotos, "enviar", lambda nome, dados: (_ for _ in ()).throw(RuntimeError("bucket fora")))
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/sobra", data={"descricao": "MESA VELHA", "observacao": "x", "foto": (io.BytesIO(imagem), "c.jpg")}, content_type="multipart/form-data", follow_redirects=True)
    assert b"MESA VELHA" not in r.data and "não registrada".encode() in r.data
    import db, inventario
    sobras = inventario.bens_da_sala(db.conectar(), eid, "01 - SALA CCI")["sobras"]
    assert [s["descricao"] for s in sobras] == ["CADEIRA VELHA", "VENTILADOR"]
    r = cliente.post(f"/inventario/{eid}/sobra/{sobras[0]['id']}/excluir", follow_redirects=True)
    assert b"CADEIRA VELHA" not in r.data and len(apagados) == 2
    # evento encerrado: não mexe em foto no bucket nem aceita novo envio
    monkeypatch.setattr(fotos, "enviar", lambda nome, dados: enviados.append(nome) or f"https://f.exemplo.org/inventario/{nome}")
    cliente.post(f"/inventario/{eid}/encerrar", data={"confirmar": "1"})
    n_apagados, n_enviados = len(apagados), len(enviados)
    cliente.post(f"/inventario/{eid}/leitura/1001/foto/excluir", follow_redirects=True)
    assert len(apagados) == n_apagados
    r = cliente.post(f"/inventario/{eid}/leitura/1001/foto", data={"foto": (io.BytesIO(imagem), "d.png")}, content_type="multipart/form-data")
    assert r.status_code == 409 and len(enviados) == n_enviados


def test_inventario_relatorio_xlsx_e_card_do_painel(cliente):
    assert b"Nenhum invent" in cliente.get("/").data
    eid = _abrir(cliente)
    cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    r = cliente.get("/")
    assert b"Invent\xc3\xa1rio em andamento" in r.data and f"/inventario/{eid}".encode() in r.data
    r = cliente.get(f"/inventario/{eid}/relatorio")
    assert r.status_code == 200 and b"1001" in r.data and b"Localizado" in r.data and b"1002" in r.data
    assert b"1002" not in cliente.get(f"/inventario/{eid}/relatorio?situacao=localizado").data.split(b"<tbody>")[1]
    assert b"1004" not in cliente.get(f"/inventario/{eid}/relatorio?localizacao=01 - SALA CCI").data.split(b"<tbody>")[1]
    r = cliente.get(f"/inventario/{eid}/xlsx?localizacao=01 - SALA CCI")
    assert r.status_code == 200 and r.headers["Content-Disposition"].endswith(".xlsx")


def test_inventario_relatorio_filtros_ordem_modal_e_xlsx_com_fotos(cliente):
    import io
    from openpyxl import load_workbook
    eid = _abrir(cliente)
    cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    cliente.post(f"/inventario/{eid}/leitura/1001", json={"quem_usa": "José"})
    import db, inventario
    inventario.atualizar_leitura(db.conectar(), eid, 1001, foto_url="https://x/1001.webp")
    r = cliente.get(f"/inventario/{eid}/relatorio?integrante=Fulano&busca=jose&ordem=numero&dir=desc")
    corpo = r.data.split(b"<tbody>")[1]
    assert r.status_code == 200 and b">1001<" in corpo and b">1002<" not in corpo
    assert "Integrante Fulano".encode() in r.data and b"Busca &#34;jose&#34;" in r.data and b"fa-sort-down" in r.data
    assert b"ordem=numero&amp;dir=asc" in r.data or b"dir=asc&amp;ordem=numero" in r.data      # clique de novo inverte
    assert b'data-foto="https://x/1001.webp"' in r.data and b'id="scrim-foto"' in r.data and b'name="fotos"' in r.data
    assert b'name="ordem" value="numero"' in r.data and b"1 linha(s)" in r.data
    r = cliente.get(f"/inventario/{eid}/relatorio?foto=sem&conservacao=-")
    assert b">1001<" not in r.data.split(b"<tbody>")[1] and b"Sem foto" in r.data and "Não informada".encode() in r.data
    r_xss = cliente.get(f"/inventario/{eid}/relatorio?busca=<b>x</b>")
    assert b"<b>x</b>" not in r_xss.data
    r = cliente.get(f"/inventario/{eid}/xlsx?integrante=Fulano&fotos=1")
    assert r.status_code == 200 and r.headers["Content-Disposition"].endswith(".xlsx")
    ws = load_workbook(io.BytesIO(r.data))["Bens"]
    assert ws.cell(row=3, column=1).value == "Todas as salas · Integrante Fulano" and ws.max_row == 6
    assert ws.cell(row=6, column=13).value == '=_xlfn.IMAGE("https://x/1001.webp")'
    ws = load_workbook(io.BytesIO(cliente.get(f"/inventario/{eid}/xlsx").data))["Bens"]
    assert ws.cell(row=6, column=13).value == "https://x/1001.webp" and ws.max_row == 8      # 5 de cabeçalho + 1001, 1002, 1004


def test_inventario_painel(cliente):
    eid = _abrir(cliente)
    cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    r = cliente.get(f"/inventario/{eid}/painel")
    assert r.status_code == 200 and b'id="g-situacao"' in r.data and b'id="g-andares"' in r.data and b'id="g-salas"' not in r.data
    assert b"bens no escopo" in r.data and b"echarts-dsgov.js" in r.data and f"/inventario/{eid}/painel?andar=01".encode() in r.data
    assert b"1 (33.3%)" in r.data                                                       # localizados com % (1 de 3 bens ativos no escopo)
    r = cliente.get(f"/inventario/{eid}/painel?andar=01")
    assert b'id="g-salas"' in r.data and b"Salas do andar 01" in r.data and b"todos os andares" in r.data
    assert cliente.get("/inventario/999/painel").status_code == 404
    assert f"/inventario/{eid}/painel".encode() in cliente.get(f"/inventario/{eid}").data    # botão Painel
    cliente.post(f"/inventario/{eid}/encerrar", data={"confirmar": "1"})
    assert b"Evento encerrado" in cliente.get(f"/inventario/{eid}/painel").data


def test_pessoa_com_email_e_matricula(cliente):
    r = cliente.post("/cadastros/pessoas/incluir", data={"nome": "bruno lima", "email": "nao-e-email", "matricula": "1"})
    assert b"Informe um e-mail v" in r.data
    r = cliente.post("/cadastros/pessoas/incluir", data={"nome": "bruno lima", "email": "b@cfc.org.br", "matricula": "0012"},
                     follow_redirects=True)
    assert b"BRUNO LIMA" in r.data and b"b@cfc.org.br" in r.data      # tela da pessoa mostra o e-mail
    assert b"b@cfc.org.br" in cliente.get("/cadastros/pessoas").data   # lista também
    r = cliente.get("/cadastros/pessoas/BRUNO LIMA/editar")
    assert b'value="b@cfc.org.br"' in r.data and b'value="0012"' in r.data
    r = cliente.post("/cadastros/pessoas/BRUNO LIMA/editar", data={"nome": "bruno lima", "email": "", "matricula": "0012"},
                     follow_redirects=True)
    import db
    assert db.pessoa(db.conectar(), "BRUNO LIMA")["email"] is None
    assert b"tratamento" not in cliente.get("/cadastros/responsaveis/CCI/editar").data


def test_termo_individual_sem_email_avisa(cliente):
    cliente.post("/cadastros/processos/incluir", data={"tipo": "individual", "descricao": "T", "numero_sei": "3333", "vigente": "1"})
    cliente.get("/termo/individual/ANA SILVA/docx")
    import db
    tid = db.termos_emitidos(db.conectar(), tipo="individual")[0]["id"]
    cliente.post(f"/termos-emitidos/{tid}/documento", data={"documento_sei": "1", "bloco_sei": "2"})
    r = cliente.get(f"/termos-emitidos/{tid}")
    assert b"Sem e-mail cadastrado para ANA SILVA" in r.data and b"mailto:" not in r.data
    cliente.post("/cadastros/pessoas/ANA SILVA/editar", data={"nome": "ana silva", "email": "a@cfc.org.br"})
    r = cliente.get(f"/termos-emitidos/{tid}")
    assert b"mailto:a@cfc.org.br" in r.data and b"Prezado%28a%29%20Ana%2C" in r.data


def test_devolucao_situacao_por_pessoa_com_sinal_de_email(cliente):
    r = cliente.get("/termo_devolucao")
    assert "Situação dos termos de devolução".encode() in r.data and b"ANA SILVA" in r.data and b"Bens com ela" in r.data
    cliente.post("/cadastros/processos/incluir", data={"tipo": "devolucao", "descricao": "D", "numero_sei": "4444", "vigente": "1"})
    cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "numero_bem": "1002"})
    cliente.get("/termo/devolucao/ANA SILVA/docx")
    r = cliente.get("/termo_devolucao")
    assert b"sem doc./bloco" in r.data
    import db
    tid = db.termos_emitidos(db.conectar(), tipo="devolucao")[0]["id"]
    cliente.post(f"/termos-emitidos/{tid}/documento", data={"documento_sei": "1", "bloco_sei": "2"})
    cliente.post(f"/termos-emitidos/{tid}/email")
    r = cliente.get("/termo_devolucao")
    assert b"enviado " in r.data and f"/termos-emitidos/{tid}".encode() in r.data


def test_devolucao_sugere_bens_da_pessoa_sem_travar(cliente):
    # só escolher a pessoa (sem número) mostra os bens que estão com ela, sem erro
    r = cliente.post("/termo_devolucao", data={"nome": "ANA SILVA"}, follow_redirects=True)
    assert b"Bens com ANA SILVA" in r.data and b"1002" in r.data and b"Adicionar todos" in r.data and b"Erro." not in r.data
    # um bem que NÃO está com ela também pode entrar
    r = cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "numero_bem": "1001"}, follow_redirects=True)
    assert b"Bens a devolver" in r.data and b"1001" in r.data
    # adicionar todos: 1002 sai da sugestão e entra na lista
    r = cliente.post("/termo_devolucao", data={"nome": "ANA SILVA", "todos": "1"}, follow_redirects=True)
    assert b"Bens com ANA SILVA" not in r.data and r.data.count(b">1002<") == 1 and b">1001<" in r.data
    # sem pessoa e sem número: pede a pessoa
    cliente.post("/termo_devolucao", data={"limpar": "1"})
    with cliente.session_transaction() as s:
        s.pop("nome_devolucao", None)
    r = cliente.post("/termo_devolucao", data={"nome": ""}, follow_redirects=True)
    assert b"Escolha a pessoa que devolve" in r.data


def test_docx_tipo_invalido_da_404_e_head_nao_registra(cliente):
    assert cliente.get("/termo/xyz/CCI/docx").status_code == 404
    assert cliente.get("/termo/xyz/CCI").status_code == 404
    assert cliente.post("/termo/xyz/CCI/registrar").status_code == 404
    db.incluir_processo(db.conectar(), "ccusto", "Termos", "1111")
    assert cliente.head("/termo/ccusto/CCI/docx").status_code == 200
    assert db.termos_emitidos(db.conectar()) == []
    assert cliente.get("/termo/ccusto/CCI/docx").status_code == 200
    assert len(db.termos_emitidos(db.conectar())) == 1


def test_leitura_com_json_que_nao_e_objeto_da_400(cliente):
    eid = _abrir(cliente, integrante="Fulano")
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json=[1, 2])
    assert r.status_code == 400 and "objeto JSON" in r.get_json()["erro"]
    r = cliente.post(f"/inventario/{eid}/leitura/1001", json="texto")
    assert r.status_code == 400 and "objeto JSON" in r.get_json()["erro"]


def test_upload_acima_de_20mb_da_mensagem_e_nao_500(cliente):
    grande = io.BytesIO(b"x" * (20 * 1024 * 1024 + 1))
    r = cliente.post("/upload", data={"arquivo": (grande, "export.xlsx")}, content_type="multipart/form-data",
                     headers={"Referer": "http://localhost/upload"}, follow_redirects=True)
    assert r.status_code == 200 and "Arquivo muito grande".encode() in r.data


def test_form_sobra_nao_e_o_proprio_br_card(cliente):
    eid = _abrir(cliente, integrante="Fulano")
    html = cliente.get(f"/inventario/{eid}/sala/01 - SALA CCI").get_data(as_text=True)
    assert 'id="card-sobra"' in html and 'id="form-sobra"' in html
    assert 'class="br-card mt-3" id="form-sobra"' not in html   # BRCard trocaria o id do form
