"""Escopo do inventário por identidade (Fase 5A, tarefa 4): quem vê e quem confere um evento é decidido
pelos vínculos de `inventario_comissao_usuarios` (IDs), nunca pelo nome digitado na comissão."""
import sqlite3

import pytest

import comissoes
import db
import inventario
import usuarios
from tests.conftest import semear


def _integrantes(conn, eid):
    return inventario.evento(conn, eid)["integrantes"]


def _vinculos(conn, eid):
    return [tuple(r) for r in conn.execute("SELECT usuario_id, nome_na_comissao FROM inventario_comissao_usuarios "
                                           "WHERE evento_id=? ORDER BY usuario_id", (eid,))]


def test_escopo_por_id_e_nao_por_nome(dados):
    semear(dados)
    a = usuarios.criar(dados, 'ana1', 'Ana', 'Senha!234', ['inventariante'])
    b = usuarios.criar(dados, 'ana2', 'Ana', 'Senha!234', ['inventariante'])
    eid = inventario.abrir_evento(dados, 'Evento', '', ['Ana'])
    comissoes.definir(dados, eid, [a])
    assert comissoes.visivel(dados, usuarios.por_id(dados, a), eid)
    assert not comissoes.visivel(dados, usuarios.por_id(dados, b), eid)
    assert not comissoes.pode_conferir(dados, usuarios.por_id(dados, b), eid)


def test_definir_recusa_id_oculto_inativo_ou_sem_funcao(dados):
    semear(dados)
    a = usuarios.criar(dados, "ana", "Ana", "Senha!234", ["inventariante"])
    leitor = usuarios.criar(dados, "leitor", "Leitor", "Senha!234", ["consulta"])
    eid = inventario.abrir_evento(dados, "Evento", "", ["Ana"])
    comissoes.definir(dados, eid, [a])
    for ids in ([], [leitor], [a, leitor], [a + leitor + 10], ["Ana"], [None]):
        with pytest.raises(db.ErroDeNegocio):
            comissoes.definir(dados, eid, ids)
    usuarios.editar(dados, a, "Ana", ["inventariante"], ativo=0)
    with pytest.raises(db.ErroDeNegocio, match="ao menos um usuário ativo"):
        comissoes.definir(dados, eid, [a])
    assert _vinculos(dados, eid) == [(a, "Ana")]              # nada mudou nas recusas


def test_trocar_a_comissao_substitui_vinculos_e_preserva_leituras(dados):
    semear(dados)
    a = usuarios.criar(dados, "ana", "Ana", "Senha!234", ["inventariante"])
    b = usuarios.criar(dados, "beto", "Beto", "Senha!234", ["inventariante"])
    eid = comissoes.abrir(dados, "Evento", "", [a], None)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Ana")
    comissoes.definir(dados, eid, [b])
    assert _vinculos(dados, eid) == [(b, "Beto")] and _integrantes(dados, eid) == ["Beto"]
    assert not comissoes.pode_conferir(dados, usuarios.por_id(dados, a), eid)
    assert comissoes.pode_conferir(dados, usuarios.por_id(dados, b), eid)
    assert dados.execute("SELECT integrante FROM inventario_leituras WHERE evento_id=?", (eid,)).fetchone()[0] == "Ana"


def test_revogar_a_funcao_tira_a_conferencia_sem_apagar_o_vinculo(dados):
    semear(dados)
    a = usuarios.criar(dados, "ana", "Ana", "Senha!234", ["inventariante"])
    eid = comissoes.abrir(dados, "Evento", "", [a], None)
    usuarios.editar(dados, a, "Ana", ["consulta"], ativo=1)
    u = usuarios.por_id(dados, a)
    assert comissoes.membro(dados, u, eid) and not comissoes.pode_conferir(dados, u, eid)
    assert not comissoes.visivel(dados, u, eid)


def test_evento_encerrado_nao_muda_de_comissao(dados):
    semear(dados)
    a = usuarios.criar(dados, "ana", "Ana", "Senha!234", ["inventariante"])
    b = usuarios.criar(dados, "beto", "Beto", "Senha!234", ["inventariante"])
    eid = comissoes.abrir(dados, "Evento", "", [a], None)
    inventario.encerrar_evento(dados, eid)
    with pytest.raises(db.ErroDeNegocio, match="encerrado"):
        comissoes.definir(dados, eid, [b])
    assert _vinculos(dados, eid) == [(a, "Ana")] and _integrantes(dados, eid) == ["Ana"]


def test_renomear_conta_com_homonimo_so_mexe_no_evento_aberto(dados):
    semear(dados)
    a = usuarios.criar(dados, "ana1", "Ana", "Senha!234", ["inventariante"])
    b = usuarios.criar(dados, "ana2", "Ana", "Senha!234", ["inventariante"])
    antigo = comissoes.abrir(dados, "Evento antigo", "", [a], None)
    inventario.ler(dados, antigo, "01 - SALA CCI", 1001, "Ana")
    inventario.registrar_sobra(dados, antigo, "01 - SALA CCI", "VENT", "", "obs", "", "Ana", exigir_foto=False)
    inventario.encerrar_evento(dados, antigo)
    aberto = comissoes.abrir(dados, "Evento novo", "", [a], None)
    usuarios.editar(dados, a, "Ana Maria", ["inventariante"], ativo=1)
    comissoes.atualizar_nome(dados, a)
    assert _integrantes(dados, aberto) == ["Ana Maria"] and _vinculos(dados, aberto) == [(a, "Ana Maria")]
    assert _integrantes(dados, antigo) == ["Ana"] and _vinculos(dados, antigo) == [(a, "Ana")]
    assert dados.execute("SELECT integrante FROM inventario_leituras").fetchone()[0] == "Ana"
    assert dados.execute("SELECT integrante FROM inventario_sobras").fetchone()[0] == "Ana"
    assert comissoes.pode_conferir(dados, usuarios.por_id(dados, a), aberto)
    assert not comissoes.pode_conferir(dados, usuarios.por_id(dados, b), aberto)   # o homônimo continua de fora


def test_renomear_conta_mantem_o_integrante_legado_sem_vinculo(dados):
    """Comissão migrada: um nome ficou sem conta ligada e não pode sumir quando outro integrante é renomeado."""
    semear(dados)
    a = usuarios.criar(dados, "ana", "Ana", "Senha!234", ["inventariante"])
    eid = inventario.abrir_evento(dados, "Evento", "", ["Ana", "Zé Antigo"])
    dados.execute("INSERT INTO inventario_comissao_usuarios VALUES (?,?,?)", (eid, a, "Ana"))
    dados.commit()
    usuarios.editar(dados, a, "Ana Maria", ["inventariante"], ativo=1)
    comissoes.atualizar_nome(dados, a)
    assert _integrantes(dados, eid) == ["Ana Maria", "Zé Antigo"]
    assert _vinculos(dados, eid) == [(a, "Ana Maria")]


def test_abrir_desfaz_o_evento_quando_o_vinculo_falha(dados, monkeypatch):
    semear(dados)
    a = usuarios.criar(dados, "ana", "Ana", "Senha!234", ["inventariante"])
    monkeypatch.setattr(comissoes, "_usuarios_selecionados", lambda conn, ids: [{"id": a + 999, "nome": "Fantasma"}])
    with pytest.raises(sqlite3.IntegrityError):
        comissoes.abrir(dados, "Evento", "", [a], None)
    assert inventario.eventos(dados) == [] and inventario.evento_aberto(dados) is None
    assert not dados.execute("SELECT 1 FROM inventario_integrantes").fetchone()
    assert not dados.execute("SELECT 1 FROM inventario_salas").fetchone()


def test_excluir_evento_apaga_os_vinculos_em_cascata(dados):
    semear(dados)
    a = usuarios.criar(dados, "ana", "Ana", "Senha!234", ["inventariante"])
    eid = comissoes.abrir(dados, "Evento", "", [a], None)
    inventario.excluir_evento(dados, eid, "Evento")
    assert not dados.execute("SELECT 1 FROM inventario_comissao_usuarios").fetchone()


def test_apagar_usuario_apaga_os_vinculos_em_cascata(dados):
    semear(dados)
    a = usuarios.criar(dados, "ana", "Ana", "Senha!234", ["inventariante"])
    eid = comissoes.abrir(dados, "Evento", "", [a], None)
    dados.execute("DELETE FROM usuarios WHERE id=?", (a,))
    dados.commit()
    assert not dados.execute("SELECT 1 FROM inventario_comissao_usuarios").fetchone()
    assert _integrantes(dados, eid) == ["Ana"]               # o nome continua na comissão exibida (dado histórico)


def _linhas_de_importacao(eid, nome, aberto_em, integrantes):
    return {"inv_eventos": [(eid, nome, None, aberto_em, None)],
            "inv_integrantes": [(eid, n) for n in integrantes],
            "inv_salas": [(eid, "01 - SALA CCI")], "inv_leituras": [], "inv_sobras": [],
            "inv_bens_encerrados": [], "inv_fotos": []}


@pytest.mark.parametrize("nome, aberto, integrantes, mantem", [
    ("Evento", None, ["Ana"], True),                       # mesma identidade e mesmo nome na comissão
    ("Outro evento", None, ["Ana"], False),                # id reaproveitado por outro evento
    ("Evento", "2001-01-01 00:00:00", ["Ana"], False),     # mesma id e nome, mas outra abertura
    ("Evento", None, ["Zé"], False),                       # a pessoa saiu da comissão na planilha
])
def test_importacao_preserva_o_vinculo_so_com_identidade_intacta(dados, nome, aberto, integrantes, mantem):
    semear(dados)
    a = usuarios.criar(dados, "ana", "Ana", "Senha!234", ["inventariante"])
    eid = comissoes.abrir(dados, "Evento", "", [a], None)
    aberto_em = aberto or dados.execute("SELECT aberto_em FROM inventario_eventos WHERE id=?", (eid,)).fetchone()[0]
    inventario.substituir_tabelas(dados, _linhas_de_importacao(eid, nome, aberto_em, integrantes))
    dados.commit()
    assert bool(_vinculos(dados, eid)) is mantem
    assert comissoes.pode_conferir(dados, usuarios.por_id(dados, a), eid) is mantem
    assert _integrantes(dados, eid) == integrantes          # os nomes vêm sempre da planilha


def test_eventos_visiveis_por_funcao(dados):
    semear(dados)
    a = usuarios.criar(dados, "ana", "Ana", "Senha!234", ["inventariante"])
    b = usuarios.criar(dados, "beto", "Beto", "Senha!234", ["inventariante"])
    chefe = usuarios.criar(dados, "chefe", "Chefe", "Senha!234", ["consulta_inventarios"])
    e1 = comissoes.abrir(dados, "Evento 1", "", [a], None)
    inventario.encerrar_evento(dados, e1)
    e2 = comissoes.abrir(dados, "Evento 2", "", [b], None)
    assert [e["id"] for e in comissoes.eventos_visiveis(dados, usuarios.por_id(dados, a))] == [e1]
    assert [e["id"] for e in comissoes.eventos_visiveis(dados, usuarios.por_id(dados, b))] == [e2]
    assert {e["id"] for e in comissoes.eventos_visiveis(dados, usuarios.por_id(dados, chefe))} == {e1, e2}


def test_administrador_local_e_conferido_pelo_nome(dados):
    """Desktop: o administrador local não tem linha em usuarios, então vale o nome em inventario_integrantes."""
    semear(dados)
    eid = inventario.abrir_evento(dados, "Evento", "", ["Administrador local"])
    local = usuarios.USUARIO_LOCAL
    assert comissoes.membro(dados, local, eid) and comissoes.pode_conferir(dados, local, eid)
    outro = dict(local, nome="Outro qualquer")
    assert not comissoes.membro(dados, outro, eid)
    sem_admin = dict(local, funcoes=("consulta",))
    assert not comissoes.membro(dados, sem_admin, eid)


# ---------------------------------------------------------------- telas (contrato por ID)
def test_web_nao_autoriza_homonimo_pelo_nome_na_comissao(cliente, dados):
    """Duas contas chamadas 'Beltrana': só a vinculada ao evento lê e registra sobra."""
    from tests.conftest import SENHA_PADRAO, logar
    beltrana = usuarios.por_login(dados, "beltrana")["id"]
    usuarios.criar(dados, "beltrana2", "Beltrana", SENHA_PADRAO, ["inventariante"], trocar_senha=False)
    cliente.post("/inventario/abrir", data={"nome": "Inv", "usuarios": [beltrana], "escopo": "todas"})
    eid = inventario.evento_aberto(dados)["id"]
    assert _integrantes(dados, eid) == ["Beltrana"]
    cliente.post("/sair"); logar(cliente, "beltrana2", SENHA_PADRAO)
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    assert r.status_code == 409 and "comissão" in r.get_json()["erro"].lower()
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/sobra",
                     data={"descricao": "VENTILADOR", "observacao": "sem plaqueta"}, follow_redirects=True)
    assert "não faz parte da comissão".encode() in r.data
    assert not dados.execute("SELECT 1 FROM inventario_leituras").fetchone()
    assert not dados.execute("SELECT 1 FROM inventario_sobras").fetchone()
    cliente.post("/sair"); logar(cliente, "beltrana", SENHA_PADRAO)
    assert cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"}).status_code == 200


def test_tela_da_comissao_mostra_integrante_sem_conta_vinculada(cliente, dados):
    eid = inventario.abrir_evento(dados, "Inv", "", ["Zé Antigo"])
    r = cliente.get(f"/inventario/{eid}/comissao")
    assert "Integrantes anteriores sem conta vinculada: Zé Antigo".encode() in r.data
    assert b'name="usuarios"' in r.data and b'name="integrantes"' not in r.data
    beltrana = usuarios.por_login(dados, "beltrana")["id"]
    r = cliente.post(f"/inventario/{eid}/comissao", data={"usuarios": [beltrana]}, follow_redirects=True)
    assert "Comissão atualizada".encode() in r.data and _integrantes(dados, eid) == ["Beltrana"]
    assert b"sem conta vinculada" not in cliente.get(f"/inventario/{eid}/comissao").data
