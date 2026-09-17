"""Análise: separar valor não informado (NULL), valor zero e intervalos numéricos."""
import re

import pytest
import db
from tests.conftest import semear


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
