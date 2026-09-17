"""Cenários de uso e preservação dos vínculos no fluxo de manutenção."""
import re
from urllib.parse import parse_qs, urlsplit

import pytest

import db
from tests.conftest import semear, confirmar_revisao


@pytest.fixture
def cliente(dados):
    semear(dados)
    from app import app
    app.config['TESTING'] = True
    with app.test_client() as client:
        yield client


def test_busca_filtra_antes_de_paginar_e_normaliza_acentos(cliente, dados):
    for n in range(25):
        db.incluir_pessoa(dados, f'PESSOA {n:02}')
    db.incluir_pessoa(dados, 'ÁUREA LIMA')
    r = cliente.get('/cadastros/pessoas?q=aurea&por_pagina=10')
    assert 'ÁUREA LIMA' in r.text
    assert '1–1 de 1 registros' in r.text
    r = cliente.get('/cadastros/pessoas?q=pessoa&por_pagina=10&pagina=2')
    assert '11–20 de 25 registros' in r.text
    assert 'PESSOA 10' in r.text and 'PESSOA 00' not in r.text
    r = cliente.get('/cadastros/pessoas?q=%25')
    assert 'Nenhum registro encontrado' in r.text


def test_parametros_de_ordenacao_e_paginacao_invalidos(cliente):
    r = cliente.get('/cadastros/pessoas?pagina=-8&por_pagina=0&ordem=nome;DROP+TABLE+pessoas&direcao=qualquer')
    assert r.status_code == 200 and 'ANA SILVA' in r.text
    assert 'Página 1 de 1' in r.text
    assert cliente.get('/cadastros/pessoas?pagina=abc&por_pagina=x').status_code == 200


def test_edicao_mantem_campos_e_nao_altera_sigla_com_email_invalido(cliente, dados):
    r = cliente.post('/cadastros/responsaveis/CCI/editar', data={
        'ccustos': 'NOVO', 'responsavel': 'MARIA', 'email': 'invalido', 'matricula': '0007'})
    assert 'value="NOVO"' in r.text and 'value="0007"' in r.text
    assert 'id="erro-email"' in r.text and 'aria-invalid="true"' in r.text
    assert db.responsavel(dados, 'CCI') and not db.responsavel(dados, 'NOVO')


def test_renomear_centro_e_historico_sao_atomicos(cliente, dados):
    # Simula uma falha de escrita no histórico depois do UPDATE de sigla.
    dados.execute("CREATE TRIGGER falha_historico BEFORE UPDATE ON termos_emitidos BEGIN SELECT RAISE(ABORT, 'falha'); END")
    db.incluir_processo(dados, 'ccusto', 'T', '123')
    cliente.get('/termo/ccusto/CCI/docx')
    import sqlite3
    with pytest.raises(sqlite3.IntegrityError):
        db.salvar_centro(dados, 'CCI', {'ccustos': 'NOVO', 'responsavel': 'MARIA'})
    assert db.responsavel(dados, 'CCI') and not db.responsavel(dados, 'NOVO')
    assert db.localizacoes_mapeadas(dados)[0]['ccustos'] == 'CCI'


def test_edicao_retorna_a_busca_e_oferece_registro_fora_do_filtro(cliente, dados):
    r = cliente.post('/cadastros/responsaveis/CCI/editar', data={
        'ccustos': 'NOVO', 'responsavel': 'MARIA',
        'retorno': '/cadastros/responsaveis?q=CCI&por_pagina=10&pagina=1'})
    destino = urlsplit(r.location)
    assert parse_qs(destino.query)['q'] == ['CCI']
    assert destino.fragment.startswith('registro-')
    r = cliente.get(r.location)
    assert 'não aparece nesta página' in r.text and 'Acessar registro' in r.text
    assert db.responsavel(dados, 'NOVO')


def test_retorno_externo_e_ignorado(cliente):
    r = cliente.post('/cadastros/pessoas/incluir', data={'nome': 'JOAO', 'retorno': 'https://outro.test/cadastros/pessoas'})
    assert r.location.startswith('/cadastros/pessoas?')


def test_duplicidade_de_pessoa_preserva_nome(cliente):
    r = cliente.post('/cadastros/pessoas/incluir', data={'nome': 'ana silva'})
    assert 'Esta pessoa já está cadastrada' in r.text
    assert 'value="ana silva"' in r.text


def test_pendentes_respeitam_regra_de_guarda_individual(cliente, dados):
    db.atribuir(dados, 'ANA SILVA', 1004)
    r = cliente.get('/cadastros/localizacoes?situacao=sem_centro')
    assert '99 - SEM MAPA' not in r.text
    assert 'Nenhum registro encontrado' in r.text


def test_revisao_nao_muta_e_revalida_transferencia(cliente, dados):
    db.incluir_responsavel(dados, {'ccustos': 'GEX', 'responsavel': 'MARIA'})
    db.incluir_responsavel(dados, {'ccustos': 'OUTRO', 'responsavel': 'JOAO'})
    r = cliente.post('/cadastros/localizacoes/mover', data={'localizacoes': ['01 - SALA CCI'], 'ccustos_destino': 'GEX'})
    assert db.localizacoes_mapeadas(dados)[0]['ccustos'] == 'CCI'
    db.mover_localizacoes(dados, ['01 - SALA CCI'], 'OUTRO')
    r = confirmar_revisao(cliente, r, '/cadastros/localizacoes/mover')
    assert 'Os dados mudaram' in r.text and 'OUTRO' in r.text
    assert db.localizacoes_mapeadas(dados)[0]['ccustos'] == 'OUTRO'
    r = confirmar_revisao(cliente, r, '/cadastros/localizacoes/mover')
    assert 'movida(s) para GEX' in r.text
    assert db.localizacoes_mapeadas(dados)[0]['ccustos'] == 'GEX'
    assert db.pessoa_do_bem(dados, 1002) == 'ANA SILVA'


def test_token_nao_autoriza_outro_patrimonio(cliente, dados):
    db.incluir_pessoa(dados, 'BRUNO')
    r = cliente.post('/cadastros/pessoas/atribuir', data={'nome': 'BRUNO', 'numero': '1002'})
    token = re.search('name="revisao" value="([^"]+)"', r.text)[1]
    r = cliente.post('/cadastros/pessoas/atribuir', data={'nome': 'BRUNO', 'numero': '1001', 'revisao': token})
    assert 'Os dados mudaram' in r.text
    assert db.pessoa_do_bem(dados, 1001) is None
    assert db.pessoa_do_bem(dados, 1002) == 'ANA SILVA'


def test_atribuicao_sem_dono_tambem_pede_revisao(cliente, dados):
    r = cliente.post('/cadastros/pessoas/atribuir', data={'nome': 'ANA SILVA', 'numero': '1001'})
    assert 'CADEIRA' in r.text and 'CCI' in r.text
    assert db.pessoa_do_bem(dados, 1001) is None
    confirmar_revisao(cliente, r, '/cadastros/pessoas/atribuir')
    assert db.pessoa_do_bem(dados, 1001) == 'ANA SILVA'


def test_exclusao_pessoa_remove_so_vinculos_apos_revisao(cliente, dados):
    r = cliente.post('/cadastros/pessoas/excluir', data={'nome': 'ANA SILVA'})
    assert db.pessoa_do_bem(dados, 1002) == 'ANA SILVA'
    confirmar_revisao(cliente, r, '/cadastros/pessoas/excluir')
    assert db.buscar_bem(dados, 1002) and db.pessoa_do_bem(dados, 1002) is None


def test_substituicao_de_vigente_revisa_e_preserva_historico(cliente, dados):
    antigo = db.incluir_processo(dados, 'ccusto', 'Original', '111')
    cliente.get('/termo/ccusto/CCI/docx')
    r = cliente.post('/cadastros/processos/incluir', data={'tipo': 'ccusto', 'descricao': 'Novo', 'numero_sei': '222', 'vigente': '1'})
    assert '111' in r.text and '222' in r.text
    assert db.processo_vigente(dados, 'ccusto')['id'] == antigo
    confirmar_revisao(cliente, r, '/cadastros/processos/incluir')
    assert db.processo_vigente(dados, 'ccusto')['numero_sei'] == '222'
    assert dados.execute('SELECT processo_id FROM termos_emitidos').fetchone()[0] == antigo
    r = cliente.post('/cadastros/processos/excluir', data={'id': antigo})
    assert 'encerre-o em vez de excluir' in r.text and 'name="revisao"' not in r.text


def test_novos_formularios_e_conteudo_so_da_area_ativa(cliente):
    for aba in ('responsaveis', 'localizacoes', 'pessoas', 'processos'):
        r = cliente.get('/cadastros/' + aba + '/novo')
        assert r.status_code == 200
        assert r.text.count('<h1>') == 1
    r = cliente.get('/cadastros/responsaveis')
    assert 'cadastrar nova pessoa' not in r.text.lower()
    assert 'id="busca-cadastro"' in r.text and 'Novo centro de custo' in r.text


def test_destino_invalido_preserva_selecao_em_lote(cliente):
    r = cliente.post('/cadastros/localizacoes/mover', data={
        'localizacoes': ['01 - SALA CCI'], 'ccustos_destino': ''})
    assert 'id="erro-ccustos_destino"' in r.text
    assert 'name="localizacoes" value="01 - SALA CCI"' in r.text
    assert 'Revisar alteração' in r.text


def test_patrimonio_invalido_mantem_valor_para_correcao(cliente):
    r = cliente.post('/cadastros/pessoas/atribuir', data={'nome': 'ANA SILVA', 'numero': '9999'})
    assert 'value="9999"' in r.text and 'id="erro-numero"' in r.text
    assert 'não encontrado' in r.text


def test_localizacao_vinculada_fora_do_filtro_oferece_acesso(cliente):
    r = cliente.post('/cadastros/localizacoes/incluir', data={
        'localizacao': '99 - SEM MAPA', 'ccustos': 'CCI',
        'retorno': '/cadastros/localizacoes?situacao=sem_centro'}, follow_redirects=True)
    assert 'não aparece nesta página' in r.text and 'Acessar registro' in r.text


def test_cancelar_edicao_preserva_ancora_e_filtros(cliente):
    r = cliente.get('/cadastros/pessoas/ANA%20SILVA/editar?retorno=%2Fcadastros%2Fpessoas%3Fq%3Dana%23registro-0123456789abcdef')
    assert '/cadastros/pessoas?q=ana#registro-0123456789abcdef' in r.text
