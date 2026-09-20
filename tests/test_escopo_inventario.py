"""Escopo do inventário por identidade (Fase 5A, tarefa 4): quem vê e quem confere um evento é decidido
pelos vínculos de `inventario_comissao_usuarios` (IDs), nunca pelo nome digitado na comissão."""
import io
import sqlite3

import pytest

import comissoes
import db
import fotos
import inventario
import usuarios
from tests.conftest import ADMIN_LOGIN, ADMIN_SENHA, SENHA_PADRAO, logar, semear
from tests.test_permissoes import NEGADO


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


def test_definir_recusa_id_oculto_ou_inativo_mas_aceita_sem_funcao(dados):
    semear(dados)
    a = usuarios.criar(dados, "ana", "Ana", "Senha!234", ["inventariante"])
    leitor = usuarios.criar(dados, "leitor", "Leitor", "Senha!234", ["consulta"])
    eid = inventario.abrir_evento(dados, "Evento", "", ["Ana"])
    comissoes.definir(dados, eid, [a])
    for ids in ([], [a + leitor + 10], ["Ana"], [None]):
        with pytest.raises(db.ErroDeNegocio):
            comissoes.definir(dados, eid, ids)
    assert comissoes.definir(dados, eid, [a, leitor]) == ["Leitor"]   # qualquer ativo é aceito: leitor ganha a função
    usuarios.editar(dados, a, "Ana", ["inventariante"], ativo=0)
    with pytest.raises(db.ErroDeNegocio, match="ao menos um usuário ativo"):
        comissoes.definir(dados, eid, [a])
    assert _vinculos(dados, eid) == [(a, "Ana"), (leitor, "Leitor")]  # nada mudou na recusa


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
    return {"inv_eventos": [(eid, nome, None, aberto_em, None, None)],
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
    assert r.status_code == 403 and r.get_json()["erro"] == NEGADO       # o homônimo nem chega à view
    r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/sobra",
                     data={"descricao": "VENTILADOR", "observacao": "sem plaqueta"}, follow_redirects=True)
    assert r.status_code == 403 and NEGADO.encode() in r.data
    assert cliente.get(f"/inventario/{eid}").status_code == 403          # nem vê o evento alheio
    assert not dados.execute("SELECT 1 FROM inventario_leituras").fetchone()
    assert not dados.execute("SELECT 1 FROM inventario_sobras").fetchone()
    cliente.post("/sair"); logar(cliente, "beltrana", SENHA_PADRAO)
    assert cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"}).status_code == 200


def test_tela_de_eventos_nao_alude_a_evento_alheio(cliente, dados):
    """Inventariante fora de qualquer comissão: estado vazio, sem o nome nem os números do evento dos outros."""
    from tests.conftest import SENHA_PADRAO, logar
    fulano = dados.execute("SELECT id FROM usuarios WHERE login='admin'").fetchone()[0]
    cliente.post("/inventario/abrir", data={"nome": "Inventário Secreto", "usuarios": [fulano], "escopo": "todas"})
    eid = inventario.evento_aberto(dados)["id"]
    cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    cliente.post("/sair"); logar(cliente, "beltrana", SENHA_PADRAO)
    html = cliente.get("/inventario").get_data(as_text=True)
    assert "Nenhum inventário atribuído a você" in html
    assert "Inventário Secreto" not in html and "bens localizados" not in html and 'href="/administracao"' not in html
    assert f"/inventario/{eid}" not in html
    assert cliente.get(f"/inventario/{eid}").status_code == 403


def test_tela_da_comissao_mostra_integrante_sem_conta_vinculada(cliente, dados):
    eid = inventario.abrir_evento(dados, "Inv", "", ["Zé Antigo"])
    r = cliente.get(f"/inventario/{eid}/comissao")
    assert "Integrantes anteriores sem conta vinculada: Zé Antigo".encode() in r.data
    assert b'name="usuarios"' in r.data and b'name="integrantes"' not in r.data
    beltrana = usuarios.por_login(dados, "beltrana")["id"]
    r = cliente.post(f"/inventario/{eid}/comissao", data={"usuarios": [beltrana]}, follow_redirects=True)
    assert "Comissão atualizada".encode() in r.data and _integrantes(dados, eid) == ["Beltrana"]
    assert b"sem conta vinculada" not in cliente.get(f"/inventario/{eid}/comissao").data


# ---------------------------------------------------------------- tarefa 6: regressão de autorização
def test_revogacao_usa_banco_na_proxima_chamada(cliente, dados):
    """A revogação não depende de expirar sessão: a próxima requisição já lê a função atual do banco."""
    uid = usuarios.por_login(dados, 'beltrana')['id']
    usuarios.editar(dados, uid, 'Beltrana', ['inventariante', 'consulta'], True)
    cliente.post('/sair'); logar(cliente, 'beltrana', SENHA_PADRAO)
    assert cliente.get('/bem?numero=1001').status_code == 200
    usuarios.editar(dados, uid, 'Beltrana', ['inventariante'], True)
    assert cliente.get('/bem?numero=1001').status_code == 403


def _cenario_sem_vinculo(dados, cenario):
    """Monta um evento com a comissão 'Ana' + 'Carla' e devolve (eid, login) para cada motivo de negativa:
    só Consulta, inventariante de outra comissão, homônimo sem vínculo e função retirada depois do vínculo."""
    if cenario == "so_consulta":
        usuarios.criar(dados, "so_consulta", "Só Consulta", SENHA_PADRAO, ["consulta"], trocar_senha=False)
        ana = usuarios.criar(dados, "ana", "Ana", SENHA_PADRAO, ["inventariante"], trocar_senha=False)
        eid = comissoes.abrir(dados, "Evento", "", [ana], None)
        return eid, "so_consulta"
    if cenario == "outra_comissao":
        beto = usuarios.criar(dados, "beto", "Beto", SENHA_PADRAO, ["inventariante"], trocar_senha=False)
        antigo = comissoes.abrir(dados, "Antigo", "", [beto], None)
        inventario.encerrar_evento(dados, antigo)
        ana = usuarios.criar(dados, "ana", "Ana", SENHA_PADRAO, ["inventariante"], trocar_senha=False)
        eid = comissoes.abrir(dados, "Evento", "", [ana], None)
        return eid, "beto"
    if cenario == "homonimo":
        ana = usuarios.criar(dados, "ana", "Ana", SENHA_PADRAO, ["inventariante"], trocar_senha=False)
        usuarios.criar(dados, "ana2", "Ana", SENHA_PADRAO, ["inventariante"], trocar_senha=False)
        eid = comissoes.abrir(dados, "Evento", "", [ana], None)
        return eid, "ana2"
    if cenario == "funcao_retirada":
        ana = usuarios.criar(dados, "ana", "Ana", SENHA_PADRAO, ["inventariante"], trocar_senha=False)
        carla = usuarios.criar(dados, "carla", "Carla", SENHA_PADRAO, ["inventariante"], trocar_senha=False)
        eid = comissoes.abrir(dados, "Evento", "", [ana, carla], None)
        usuarios.editar(dados, carla, "Carla", ["consulta"], ativo=True)   # continua vinculada, mas sem a função
        return eid, "carla"
    raise ValueError(cenario)


_MUTACOES = [
    ("POST", "/inventario/{eid}/sala/01 - SALA CCI/ler", dict(json={"numero": "1001"})),
    ("POST", "/inventario/{eid}/leitura/1001", dict(json={"conservacao": "Bom"})),
    ("POST", "/inventario/{eid}/sala/01 - SALA CCI/lote", dict(data={"acao": "marcar", "numeros": ["1001"]})),
    ("POST", "/inventario/{eid}/leitura/1001/foto", "foto"),
    ("POST", "/inventario/{eid}/leitura/1001/foto/1/excluir", dict(data={})),
    ("POST", "/inventario/{eid}/sala/01 - SALA CCI/sobra", dict(data={"descricao": "VENTILADOR"})),
    ("POST", "/inventario/{eid}/sobra/1/excluir", dict(data={})),
]
_RELATORIOS = [
    ("GET", "/inventario/{eid}/relatorio", {}),
    ("GET", "/inventario/{eid}/painel", {}),
    ("GET", "/inventario/{eid}/xlsx", {}),
]


@pytest.mark.parametrize("cenario", ["so_consulta", "outra_comissao", "homonimo", "funcao_retirada"])
@pytest.mark.parametrize("metodo,rota,kw", _MUTACOES + _RELATORIOS,
                         ids=[r for _, r, _ in _MUTACOES + _RELATORIOS])
def test_mutacoes_e_relatorios_negam_quem_nao_tem_vinculo_valido(cliente, dados, monkeypatch, cenario, metodo, rota, kw):
    """Nenhuma mutação do inventário nem relatório aceita quem não tem vínculo válido com a comissão do
    evento (só Consulta, inventariante de outra comissão, homônimo sem vínculo, função retirada depois do
    vínculo): tudo cai em 403 antes da view, sem chamar o serviço nem o storage."""
    chamadas = []

    def _espiao(nome):
        def _fn(*a, **k):
            chamadas.append(nome)
            raise AssertionError(f"{nome} não deveria ser chamado")
        return _fn

    for nome in ("ler", "atualizar_leitura", "ler_lote", "desfazer_leituras", "adicionar_foto", "apagar_foto",
                 "registrar_sobra", "excluir_sobra", "relatorio", "painel", "exportar_xlsx"):
        monkeypatch.setattr(inventario, nome, _espiao(f"inventario.{nome}"))
    for nome in ("enviar", "apagar"):
        monkeypatch.setattr(fotos, nome, _espiao(f"fotos.{nome}"))

    if kw == "foto":     # BytesIO só pode ser lido uma vez: monta o arquivo fresco a cada chamada
        kw = dict(data={"foto": (io.BytesIO(b"fake"), "a.png")}, content_type="multipart/form-data")

    eid, login = _cenario_sem_vinculo(dados, cenario)
    cliente.post("/sair"); logar(cliente, login, SENHA_PADRAO)
    r = getattr(cliente, metodo.lower())(rota.format(eid=eid), **kw)
    assert r.status_code == 403
    assert chamadas == []


def test_consulta_de_inventarios_ve_evento_antigo_e_futuro_mas_so_le_com_vinculo(cliente, dados):
    """Consulta de inventários vê e baixa o .xlsx (GET e HEAD) de qualquer evento, encerrado antes da função
    ser concedida ou criado depois; não lê bens em nenhum. Somada a Inventário, só escreve no próprio evento."""
    ana = usuarios.por_login(dados, "beltrana")["id"]
    antigo = comissoes.abrir(dados, "Antigo", "", [ana], None)
    inventario.encerrar_evento(dados, antigo)
    chefe = usuarios.criar(dados, "chefe", "Chefe", SENHA_PADRAO, ["consulta_inventarios"], trocar_senha=False)
    novo = comissoes.abrir(dados, "Novo", "", [ana], None)          # criado depois de o chefe já ter a função
    cliente.post("/sair"); logar(cliente, "chefe", SENHA_PADRAO)
    for eid in (antigo, novo):
        assert cliente.get(f"/inventario/{eid}/xlsx").status_code == 200
        assert cliente.open(f"/inventario/{eid}/xlsx", method="HEAD").status_code == 200
        r = cliente.post(f"/inventario/{eid}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
        assert r.status_code == 403 and r.get_json()["erro"] == NEGADO
    cliente.post("/sair"); logar(cliente, ADMIN_LOGIN, ADMIN_SENHA)
    usuarios.editar(dados, chefe, "Chefe", ["consulta_inventarios", "inventariante"], ativo=True)
    cliente.post("/sair"); logar(cliente, "chefe", SENHA_PADRAO)
    assert cliente.post(f"/inventario/{antigo}/sala/01 - SALA CCI/ler", json={"numero": "1001"}).status_code == 403
    assert cliente.post(f"/inventario/{novo}/sala/01 - SALA CCI/ler", json={"numero": "1001"}).status_code == 403
    cliente.post("/sair"); logar(cliente, ADMIN_LOGIN, ADMIN_SENHA)
    comissoes.definir(dados, novo, [ana, chefe])
    cliente.post("/sair"); logar(cliente, "chefe", SENHA_PADRAO)
    r = cliente.post(f"/inventario/{novo}/sala/01 - SALA CCI/ler", json={"numero": "1001"})
    assert r.status_code == 200 and r.get_json()["situacao"] == "localizado"     # agora é da própria comissão


def test_home_sem_evento_atribuido_nao_mostra_cartao_alheio(cliente, dados):
    """Consulta + Inventário sem vínculo com o evento aberto: o cartão de inventário em andamento do Início
    não aparece (não vaza contagem nem nome de comissão alheia)."""
    fulano = usuarios.por_login(dados, ADMIN_LOGIN)["id"]
    mista = usuarios.criar(dados, "mista", "Mista", SENHA_PADRAO, ["consulta", "inventariante"], trocar_senha=False)
    comissoes.abrir(dados, "Inventário Alheio", "", [fulano], None)
    cliente.post("/sair"); logar(cliente, "mista", SENHA_PADRAO)
    html = cliente.get("/").get_data(as_text=True)
    assert "Inventário em andamento" not in html and "Inventário Alheio" not in html


def test_queda_de_funcao_nao_apaga_historico(cliente, dados):
    """Perder uma função (inventariante e operador) nunca apaga leituras, sobras, fotos ou termos emitidos:
    usuarios.editar só grava em usuarios e usuarios_funcoes."""
    ana = usuarios.criar(dados, "ana", "Ana", SENHA_PADRAO, ["inventariante", "operador"], trocar_senha=False)
    eid = comissoes.abrir(dados, "Evento", "", [ana], None)
    inventario.ler(dados, eid, "01 - SALA CCI", 1001, "Ana")
    inventario.registrar_sobra(dados, eid, "01 - SALA CCI", "VENT", "", "obs", "", "Ana", exigir_foto=False)
    inventario.adicionar_foto(dados, eid, 1001, lambda c: "https://x/a.webp")
    cliente.post("/cadastros/processos/incluir", data={"tipo": "ccusto", "descricao": "T", "numero_sei": "1111", "vigente": "1"})
    cliente.get("/termo/ccusto/CCI/docx")                              # gera e registra a emissão
    tabelas = ("inventario_leituras", "inventario_sobras", "inventario_fotos", "termos_emitidos")
    antes = {t: [tuple(r) for r in dados.execute(f"SELECT * FROM {t}")] for t in tabelas}
    usuarios.editar(dados, ana, "Ana", ["consulta"], ativo=True)       # perde inventariante e operador
    depois = {t: [tuple(r) for r in dados.execute(f"SELECT * FROM {t}")] for t in tabelas}
    assert antes == depois and all(antes[t] for t in tabelas)


def test_comissao_aceita_qualquer_ativo_e_concede_a_funcao_inventario(dados):
    semear(dados)
    admin = usuarios.criar(dados, "adm", "Admin", SENHA_PADRAO, ["admin"], trocar_senha=False)
    leitor = usuarios.criar(dados, "leitor", "Consulta Teste", SENHA_PADRAO, ["consulta"], trocar_senha=False)
    inativo = usuarios.criar(dados, "ina", "Inativo", SENHA_PADRAO, ["inventariante"], trocar_senha=False)
    usuarios.editar(dados, inativo, "Inativo", ["inventariante"], ativo=False)
    assert [u["login"] for u in usuarios.ativos_para_comissao(dados)] == ["adm", "leitor"]
    eid, concedidos = comissoes.criar(dados, "Inv", "", [admin, leitor], None, abrir=True)
    assert concedidos == ["Consulta Teste"]
    assert usuarios.por_id(dados, leitor)["funcoes"] == ("consulta", "inventariante")
    assert usuarios.por_id(dados, admin)["funcoes"] == ("admin",)                     # admin não precisa da função
    assert _vinculos(dados, eid) == [(admin, "Admin"), (leitor, "Consulta Teste")]
    assert inventario.evento_aberto(dados)["id"] == eid
    with pytest.raises(db.ErroDeNegocio, match="usuário ativo"):
        comissoes.definir(dados, eid, [inativo])
    assert comissoes.definir(dados, eid, [leitor]) == []                              # já tem a função: nada a conceder
    inventario.desligar_chave(dados, eid)
    assert comissoes.definir(dados, eid, [admin]) == [] and _integrantes(dados, eid) == ["Admin"]   # fechado aceita comissão
    inventario.encerrar_evento(dados, eid)
    with pytest.raises(db.ErroDeNegocio, match="encerrado"):
        comissoes.definir(dados, eid, [admin])


def test_criar_sem_abrir_nasce_fechado_e_conceder_e_atomico(dados):
    semear(dados)
    fulano = usuarios.criar(dados, "fulano2", "Fulano Dois", SENHA_PADRAO, ["consulta"], trocar_senha=False)
    eid, concedidos = comissoes.criar(dados, "Preparado", "", [fulano], None)
    assert inventario.evento(dados, eid)["estado"] == "fechado" and concedidos == ["Fulano Dois"]
    with pytest.raises(db.ErroDeNegocio):
        comissoes.criar(dados, "Preparado", "", [fulano], None)                        # nome repetido: nada gravado
    assert len(inventario.eventos(dados)) == 1
