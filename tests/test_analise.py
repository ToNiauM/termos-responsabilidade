"""Análise: separar valor não informado (NULL), valor zero e intervalos numéricos."""
import io
import re

import pytest
from openpyxl import load_workbook

import db
from tests.conftest import logar, semear


def test_zero_e_ausente_sao_conjuntos_distintos(dados):
    semear(dados)
    dados.execute('UPDATE bens SET valor_atual=NULL WHERE numero=1001')
    dados.execute('UPDATE bens SET valor_atual=0 WHERE numero=1002')
    dados.execute('UPDATE bens SET valor_atual=-10 WHERE numero=1004')
    r = db.recorte(dados, {'situacao': 'ATIVO'})
    assert (r['quantidade'], r['valor_nao_informado'], r['valor_zero']) == (3, 1, 1)
    n = db.recorte(dados, {'situacao': 'ATIVO', 'valor_status': 'nao_informado'})
    z = db.recorte(dados, {'situacao': 'ATIVO', 'valor_status': 'zero'})
    assert [b['numero'] for b in n['bens']] == [1001]
    assert [b['numero'] for b in z['bens']] == [1002]
    assert db.recorte(dados, {'valor_status': 'nao_informado', 'valor_de': '0'})['quantidade'] == 0
    assert {b['numero'] for b in db.recorte(dados, {'valor_ate': '0'})['bens']} == {1002, 1004}
    with pytest.raises(db.ErroDeNegocio):
        db.recorte(dados, {'valor_status': 'outro'})


def test_recorte_vazio_zera_as_seis_contagens(dados):
    semear(dados)
    r = db.recorte(dados, {'situacao': 'BAIXADO', 'ccusto': 'INEXISTENTE'})
    assert r['quantidade'] == 0
    assert r['valor_total'] == 0
    assert r['imoveis'] == 0
    assert r['valor_imoveis'] == 0
    assert r['sem_centro'] == 0
    assert r['valor_nao_informado'] == 0
    assert r['valor_zero'] == 0


def test_sem_centro_nem_pessoa_exige_ambas_as_ausencias(dados):
    semear(dados)
    # 1004 já está sem centro mapeado (localização "99 - SEM MAPA"); atribuir pessoa a ele deve
    # tirá-lo de sem_centro, pois falta só o centro (não os dois).
    dados.execute("INSERT INTO atribuicoes VALUES ('ANA SILVA', 1004)")
    # 5001 é um imóvel (TERRENOS) sem localização mapeada e sem pessoa: conta, pois faltam os dois.
    dados.execute("INSERT INTO bens VALUES (5001,'ATIVO','TERRENO','','TERRENOS','','01/01/2000',1,1000)")
    r = db.recorte(dados, {'situacao': 'ATIVO'})
    # 1001 (centro CCI, sem pessoa) e 1004 (sem centro, mas com pessoa) não contam; só 5001 conta.
    assert r['sem_centro'] == 1


def test_analise_http_seis_cards_e_situacao_do_valor(cliente, dados):
    """GET /analise (Tarefa 3 fase5b): os seis cards dsgov-kpi aparecem, o filtro 'Situação do valor'
    está no formulário e a tabela distingue 'Não informado' (NULL) de 'R$ 0,00' (moeda(0), valor zero)
    -- a troca de `or 0` por `is none` no template evita confundir os dois conjuntos."""
    dados.execute("INSERT INTO bens VALUES (9001,'ATIVO','SEM VALOR','','MÓVEIS','01 - SALA CCI','01/01/2020',NULL,NULL)")
    dados.execute("INSERT INTO bens VALUES (9002,'ATIVO','VALOR ZERO','','MÓVEIS','01 - SALA CCI','01/01/2020',0,0)")
    dados.commit()

    r = cliente.get('/analise')
    assert r.status_code == 200
    html = r.get_data(as_text=True)

    assert html.count('dsgov-kpi') == 6   # seis cards (com ou sem link de drill-down)
    kpi_hrefs = re.findall(r'<a class="br-card h-100 dsgov-kpi" href="([^"]*)">', html)
    assert kpi_hrefs and all(h.startswith('/analise?') for h in kpi_hrefs)   # ao menos um card com link

    assert 'name="valor_status"' in html
    assert 'value="nao_informado"' in html and 'value="zero"' in html
    assert 'Valor não informado' in html and 'Valor zero' in html

    linha_9001 = re.search(r'<tr>.*?9001.*?</tr>', html)
    linha_9002 = re.search(r'<tr>.*?9002.*?</tr>', html)
    assert linha_9001 and 'Não informado' in linha_9001.group()   # NULL
    assert linha_9002 and 'R$ 0,00' in linha_9002.group()         # zero, não "Não informado"

    descricao = cliente.get('/analise?valor_status=zero').get_data(as_text=True)
    assert 'valor zero' in descricao   # frase de painel.descrever para valor_status


def test_termo_completo_nao_e_limitado_pelo_filtro(cliente,dados):
    antes=dados.execute('SELECT count(*) FROM termos_emitidos').fetchone()[0]
    html=cliente.get('/analise?ccusto=CCI&classificacao=MÓVEIS').get_data(as_text=True)
    assert 'Abrir termo completo' in html
    assert 'Os filtros desta análise não limitam o termo' in html
    assert '/termo/ccusto/CCI' in html
    assert dados.execute('SELECT count(*) FROM termos_emitidos').fetchone()[0]==antes


def test_exportacao_integral_e_valores_preservados(dados):
    semear(dados)
    dados.executemany('''INSERT INTO bens VALUES (?,'ATIVO','BEM','','MÓVEIS',
      '99 - SEM MAPA','01/01/2020',1,?)''',
      [(n,None if n%2 else 0) for n in range(2000,3105)])
    f={'situacao':'ATIVO'}
    r=db.recorte(dados,f)
    assert r['truncado'] and len(r['bens'])==1000 and r['quantidade']==1108
    arq=io.BytesIO(); db.exportar_recorte(dados,f,arq)
    ws=load_workbook(io.BytesIO(arq.getvalue())).active
    assert ws.title=='analise' and ws.max_row-1==1108
    valores={row[0]:row[-1] for row in ws.iter_rows(min_row=2,values_only=True)}
    assert valores[2000]==0 and valores[2001] is None


def test_consulta_abre_analise_e_termo_sem_botoes_de_emissao(cliente, usuarios_exemplo, dados):
    """Consulta acessa a Análise e o termo completo (só leitura); os botões que registram emissão
    (Copiar para o SEI, Baixar .docx, Baixar planilha) são exclusivos de Operador/Admin -- com um
    processo SEI vigente cadastrado, para não confundir "sem processo" (que também esconde os botões,
    de qualquer função) com "sem permissão"."""
    db.incluir_processo(dados, 'ccusto', 'Termo de centro de custo', '1111', vigente=True)
    antes = dados.execute('SELECT count(*) FROM termos_emitidos').fetchone()[0]

    admin_html = cliente.get('/termo/ccusto/CCI').get_data(as_text=True)   # cliente já logado como admin
    assert 'Copiar para o SEI' in admin_html and 'Baixar .docx' in admin_html and 'Baixar planilha' in admin_html

    cliente.post('/sair'); logar(cliente, *usuarios_exemplo['operador'])
    op_html = cliente.get('/termo/ccusto/CCI').get_data(as_text=True)
    assert 'Copiar para o SEI' in op_html and 'Baixar .docx' in op_html and 'Baixar planilha' in op_html

    cliente.post('/sair'); logar(cliente, *usuarios_exemplo['consulta'])
    html = cliente.get('/analise?ccusto=CCI').get_data(as_text=True)
    assert 'Abrir termo completo' in html and '/termo/ccusto/CCI' in html
    r = cliente.get('/termo/ccusto/CCI')
    assert r.status_code == 200
    termo_html = r.get_data(as_text=True)
    assert 'Copiar para o SEI' not in termo_html
    assert 'Baixar .docx' not in termo_html
    assert 'Baixar planilha' not in termo_html
    assert dados.execute('SELECT count(*) FROM termos_emitidos').fetchone()[0] == antes   # a prévia não emite


def test_inventario_sozinho_recebe_403_na_analise_no_xlsx_e_nos_aliases(cliente, usuarios_exemplo):
    """Quem só tem a função Inventário não vê o acervo: a matriz nega antes de qualquer redirecionamento,
    inclusive nos aliases /recorte(/xlsx), que devem responder 403 e não 301."""
    cliente.post('/sair'); logar(cliente, *usuarios_exemplo['inventariante'])
    assert cliente.get('/analise').status_code == 403
    assert cliente.get('/analise/xlsx').status_code == 403
    assert cliente.get('/recorte').status_code == 403
    assert cliente.get('/recorte/xlsx').status_code == 403
