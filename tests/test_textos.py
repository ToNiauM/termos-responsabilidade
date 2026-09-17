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


@pytest.mark.parametrize("valor", ["{nome:d}", "{nome!r}", "{nome:{orgao_sigla}}", "{{nome}}"])
def test_validar_recusa_format_spec_e_chaves_duplas(valor):
    with pytest.raises(db.ErroDeNegocio):
        textos.validar("individual_abertura", valor)
    textos.validar("individual_abertura", "Eu, {nome}, do {orgao_sigla}")  # continua passando


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


def test_obter_ignora_valor_invalido_no_banco(dados):
    dados.execute("INSERT INTO textos VALUES ('individual_abertura', 'Eu {nomee}')")
    assert textos.obter(dados)["individual_abertura"] == textos.PADRAO["individual_abertura"]


def test_padrao_cobre_todos_os_grupos_e_marcadores():
    chaves = {c for _, lista in textos.GRUPOS for c in lista}
    assert chaves == set(textos.PADRAO) == set(textos.MARCADORES) == set(textos.ROTULOS)
    for chave, valor in textos.PADRAO.items():
        textos.validar(chave, valor)  # o padrão tem de passar na própria validação


def test_unidade_gestora_entra_como_marcador_em_todos_os_termos():
    t = dict(textos.PADRAO, unidade_nome="Setor Novo", unidade_sigla="SN")
    campos = textos.campos_gerais(t)
    assert campos == {"orgao_sigla": "CFC", "unidade_nome": "Setor Novo", "unidade_sigla": "SN"}
    for chave in ("individual_compromissos", "ccusto_paragrafos", "devolucao_abertura"):
        assert "{unidade_sigla}" in textos.PADRAO[chave]
        textos.validar(chave, "só a {unidade_nome} ({unidade_sigla})")
    assert "Gersev" not in "".join(v for k, v in textos.PADRAO.items() if k != "unidade_sigla")


def test_nome_proprio():
    assert textos.nome_proprio("ANTÔNIO RODRIGUES DE SOUSA JÚNIOR") == "Antônio Rodrigues de Sousa Júnior"
    assert textos.nome_proprio("MARIA DAS DORES E SILVA-LIMA") == "Maria das Dores e Silva-Lima"
    assert textos.nome_proprio("Elys Tevania Alves de Souza") == "Elys Tevania Alves de Souza"
    assert textos.nome_proprio("DE") == "De" and textos.nome_proprio(None) == ""


def test_com_nome_com_outro_marcador():
    assert textos.com_nome("Eu, {responsavel}, do {orgao_sigla}.", {"responsavel": "ANA", "orgao_sigla": "CFC"}, "responsavel") == ("Eu, ", "ANA", ", do CFC.")


def test_salvar_todos_e_tudo_ou_nada(dados):
    import sqlite3
    valores = dict(textos.PADRAO, orgao_nome="Conselho X", cidade="Goiânia (GO)")
    dados.execute("CREATE TRIGGER falha BEFORE INSERT ON textos WHEN (SELECT count(*) FROM textos) >= 1 "
                  "BEGIN SELECT RAISE(ABORT, 'falha simulada'); END")
    with pytest.raises(sqlite3.IntegrityError):
        textos.salvar_todos(dados, valores)
    assert dados.execute("SELECT count(*) FROM textos").fetchone()[0] == 0
    dados.execute("DROP TRIGGER falha")
    textos.salvar_todos(dados, valores)
    t = textos.obter(dados)
    assert t["orgao_nome"] == "Conselho X" and t["cidade"] == "Goiânia (GO)"
    assert dados.execute("SELECT count(*) FROM textos").fetchone()[0] == 2   # só o que difere do padrão
