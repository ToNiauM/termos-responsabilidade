"""Análise: separar valor não informado (NULL), valor zero e intervalos numéricos."""
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
