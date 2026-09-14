import pytest

import db
import textos


def test_obter_devolve_padrao_e_sobreposicao(dados):
    t = textos.obter(dados)
    assert t == textos.PADRAO
    textos.salvar(dados, "individual_titulo", "TERMO X")
    assert textos.obter(dados)["individual_titulo"] == "TERMO X"
    assert textos.obter(dados)["orgao_sigla"] == "CFC"


def test_validar_recusa_marcador_desconhecido_e_chave_mal_formada():
    with pytest.raises(db.ErroDeNegocio) as e:
        textos.validar("individual_abertura", "Eu, {nomee}, declaro")
    assert "nomee" in str(e.value)
    with pytest.raises(db.ErroDeNegocio):
        textos.validar("individual_abertura", "Eu, {nome, declaro")
    with pytest.raises(db.ErroDeNegocio):
        textos.validar("chave_inexistente", "x")
    textos.validar("individual_abertura", "Eu, {nome}, do {orgao_sigla}")  # ok
    textos.validar("orgao_sigla", "")  # vazio permitido


def test_salvar_igual_ao_padrao_nao_deixa_linha(dados):
    textos.salvar(dados, "cidade", "Brasília (DF)")
    assert dados.execute("SELECT count(*) FROM textos").fetchone()[0] == 0
    textos.salvar(dados, "cidade", "Goiânia (GO)")
    assert dados.execute("SELECT count(*) FROM textos").fetchone()[0] == 1
    textos.restaurar(dados, "cidade")
    assert textos.obter(dados)["cidade"] == "Brasília (DF)"


def test_paragrafos_e_linhas():
    assert textos.paragrafos("a\nb\n\n\nc\n") == ["a b", "c"]
    assert textos.linhas("1) x\n\n2) y\n") == ["1) x", "2) y"]
    assert textos.paragrafos("") == []


def test_com_nome():
    assert textos.com_nome("Eu, {nome}, do {orgao_sigla}.", {"nome": "ANA", "orgao_sigla": "CFC"}) == ("Eu, ", "ANA", ", do CFC.")
    assert textos.com_nome("Sem marcador.", {"nome": "ANA"}) == ("Sem marcador.", None, "")


def test_padrao_cobre_todos_os_grupos_e_marcadores():
    chaves = {c for _, lista in textos.GRUPOS for c in lista}
    assert chaves == set(textos.PADRAO) == set(textos.MARCADORES)
    for chave, valor in textos.PADRAO.items():
        textos.validar(chave, valor)  # o padrão tem de passar na própria validação
