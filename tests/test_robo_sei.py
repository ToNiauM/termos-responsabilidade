"""Orquestração do envio ao SEI sem navegador: um SEI falso registra as chamadas."""
from contextlib import contextmanager

import config
import db
import robo_sei
from tests.conftest import semear

ENV = {"SEI_USUARIO": "u", "SEI_SENHA": "s", "SEI_LOGIN_URL": "https://sei.cfc.org.br/sei/", "SEI_ORGAO": "CFC"}


class SEIFalso:
    def __init__(self, arvore=None, falhar_em=None, autenticado=True, titulo_ok=True, bloco_existe=True,
                 tipo_existe=True, unidades_ok=True, avisos=()):
        self.chamadas, self.arvore, self.avisos = [], dict(arvore or {}), list(avisos)
        self.falhar_em, self.autenticado, self.titulo_ok = falhar_em, autenticado, titulo_ok
        self.bloco_existe, self.tipo_existe, self.saiu = bloco_existe, tipo_existe, False
        self.unidades_ok = unidades_ok

    def _falha(self, passo):
        if self.falhar_em == passo:
            raise TimeoutError(f"Timeout 30000ms exceeded em {passo}")

    def login(self, env):
        self.chamadas.append(("login", env["SEI_USUARIO"]))
        return {"autenticado": self.autenticado, "unidade": "GELIC"}

    def trocar_unidade(self, sigla):
        self.chamadas.append(("trocar_unidade", sigla))
        if not self.unidades_ok:
            raise robo_sei.RoboErro(f"Unidade {sigla} não está disponível para este usuário no SEI.")
        return sigla

    def logout(self):
        self.saiu = True

    anotar = robo_sei.SEI.anotar                       # grava dados/sei/erro.txt como o robô real

    def abrir_processo(self, numero):
        self.chamadas.append(("abrir_processo", numero))
        if not self.titulo_ok:
            raise robo_sei.RoboErro(f"Processo {numero} não abriu no SEI; nada foi criado.")
        return {"titulo_confere": True, "id_procedimento": "555001"}

    def id_do_documento(self, numero):
        return {"1557099": "777099", "1557088": "777088"}.get(numero)

    def documento_na_arvore(self, rotulo):
        self.chamadas.append(("documento_na_arvore", rotulo))
        return self.arvore.get(rotulo)

    def incluir_documento(self, tipo_nome, nome_arvore, html, rotulo):
        self._falha("documento")
        if not self.tipo_existe:
            raise robo_sei.RoboErro(f"Tipo de documento '{tipo_nome}' não existe no SEI; corrija em Textos.")
        self.chamadas.append(("incluir_documento", tipo_nome, nome_arvore, len(html)))
        self.arvore[rotulo] = "1557099"
        return "1557099"

    def incluir_em_bloco(self, numero, nome_bloco):
        self._falha("bloco")
        if not self.bloco_existe:
            raise robo_sei.RoboErro(f"Bloco '{nome_bloco}' não existe no SEI; crie o bloco e clique em Incluir no bloco.")
        self.chamadas.append(("incluir_em_bloco", numero, nome_bloco))
        return "69766"


def _abrir(falso):
    @contextmanager
    def abrir(env):
        yield falso
    return abrir


def _pedido(conn):
    semear(conn)
    db.incluir_processo(conn, "ccusto", "T", "90796110000022.000059/2026-88")
    t = db.registrar_emissao(conn, "ccusto", "CCI", db.bens_do_centro(conn, "CCI"))
    t = db.preparar_envio_sei(conn, t["id"], agora="2026-09-20 10:00:00")
    pid = db.enfileirar_pedido(conn, "sei", termo_id=t["id"], html="<p>termo</p>", criado_por="admin")
    return db.pedido(conn, pid), t


def test_fluxo_completo_grava_documento_e_bloco(dados):
    p, t = _pedido(dados)
    falso = SEIFalso()
    r = robo_sei.enviar_termo(dados, p, abrir=_abrir(falso), env=ENV)
    assert r == {"passo": "concluido", "mensagem": "documento 1557099 no bloco 69766", "documento_sei": "1557099", "bloco_sei": "69766"}
    assert [c[0] for c in falso.chamadas] == ["login", "abrir_processo", "documento_na_arvore", "incluir_documento", "incluir_em_bloco"]
    assert falso.chamadas[1] == ("abrir_processo", "90796110000022.000059/2026-88")
    assert falso.chamadas[2] == ("documento_na_arvore", "Termo de Responsabilidade 01/2026 - CCI")
    assert falso.chamadas[3] == ("incluir_documento", "Termo de Responsabilidade", "01/2026 - CCI", len("<p>termo</p>"))
    assert falso.chamadas[4] == ("incluir_em_bloco", "1557099", "Termos CCI")
    assert falso.saiu
    t = db.termo_emitido(dados, t["id"])
    assert t["documento_sei"] == "1557099" and t["bloco_sei"] == "69766"
    assert t["id_documento"] == "777099" and t["id_procedimento"] == "555001"       # ids internos p/ hiperlinks
    assert db.pedido(dados, p["id"])["passo"] == "concluido"


def test_troca_de_unidade_apos_login_quando_diferente(dados):
    p, t = _pedido(dados)
    falso = SEIFalso()
    env = {**ENV, "SEI_UNIDADE": "GESERV"}
    r = robo_sei.enviar_termo(dados, p, abrir=_abrir(falso), env=env)
    assert r["passo"] == "concluido"
    assert [c[0] for c in falso.chamadas] == [
        "login", "trocar_unidade", "abrir_processo", "documento_na_arvore", "incluir_documento", "incluir_em_bloco"]
    assert falso.chamadas[1] == ("trocar_unidade", "GESERV")


def test_sem_troca_quando_sei_unidade_igual_a_do_login(dados):
    p, t = _pedido(dados)
    falso = SEIFalso()
    env = {**ENV, "SEI_UNIDADE": "GELIC"}                    # SEIFalso.login já devolve unidade "GELIC"
    r = robo_sei.enviar_termo(dados, p, abrir=_abrir(falso), env=env)
    assert r["passo"] == "concluido"
    assert "trocar_unidade" not in [c[0] for c in falso.chamadas]


def test_sem_sei_unidade_no_env_nao_troca(dados):
    p, t = _pedido(dados)
    falso = SEIFalso()
    r = robo_sei.enviar_termo(dados, p, abrir=_abrir(falso), env=ENV)
    assert r["passo"] == "concluido"
    assert "trocar_unidade" not in [c[0] for c in falso.chamadas]


def test_troca_de_unidade_recusada_nao_cria_documento(dados):
    p, t = _pedido(dados)
    falso = SEIFalso(unidades_ok=False)
    env = {**ENV, "SEI_UNIDADE": "GESERV"}
    r = robo_sei.enviar_termo(dados, p, abrir=_abrir(falso), env=env)
    assert r["passo"] == "erro"
    assert r["mensagem"] == "Unidade GESERV não está disponível para este usuário no SEI."
    assert "incluir_documento" not in [c[0] for c in falso.chamadas]
    assert db.termo_emitido(dados, t["id"])["documento_sei"] is None
    assert db.pedido(dados, p["id"])["passo"] == "erro"


def test_falha_no_bloco_deixa_documento_gravado(dados):
    p, t = _pedido(dados)
    falso = SEIFalso(bloco_existe=False)
    r = robo_sei.enviar_termo(dados, p, abrir=_abrir(falso), env=ENV)
    assert r["passo"] == "erro" and "Bloco 'Termos CCI' não existe" in r["mensagem"]
    t = db.termo_emitido(dados, t["id"])
    assert t["documento_sei"] == "1557099" and t["bloco_sei"] is None
    assert db.pedido(dados, p["id"])["passo"] == "erro" and falso.saiu


def test_retomada_com_documento_pula_a_criacao(dados):
    p, t = _pedido(dados)
    db.salvar_documento_sei(dados, t["id"], "1557099", "")
    falso = SEIFalso()
    r = robo_sei.enviar_termo(dados, p, abrir=_abrir(falso), env=ENV)
    assert r["passo"] == "concluido"
    assert [c[0] for c in falso.chamadas] == ["login", "abrir_processo", "incluir_em_bloco"]


def test_retomada_sem_documento_acha_na_arvore_e_nao_duplica(dados):
    p, t = _pedido(dados)
    falso = SEIFalso(arvore={"Termo de Responsabilidade 01/2026 - CCI": "1557088"})
    r = robo_sei.enviar_termo(dados, p, abrir=_abrir(falso), env=ENV)
    assert r["documento_sei"] == "1557088" and "incluir_documento" not in [c[0] for c in falso.chamadas]
    assert db.termo_emitido(dados, t["id"])["documento_sei"] == "1557088"


def test_processo_que_nao_abre_nao_cria_nada(dados):
    p, t = _pedido(dados)
    falso = SEIFalso(titulo_ok=False)
    r = robo_sei.enviar_termo(dados, p, abrir=_abrir(falso), env=ENV)
    assert r["passo"] == "erro" and "não abriu no SEI; nada foi criado" in r["mensagem"]
    assert "incluir_documento" not in [c[0] for c in falso.chamadas] and falso.saiu
    assert db.termo_emitido(dados, t["id"])["documento_sei"] is None


def test_login_recusado_tipo_inexistente_e_timeout(dados):
    p, _ = _pedido(dados)
    r = robo_sei.enviar_termo(dados, p, abrir=_abrir(SEIFalso(autenticado=False)), env=ENV)
    assert r["passo"] == "erro" and r["mensagem"] == "O SEI recusou seu usuário ou senha; atualize em Meus acessos."
    db.marcar_passo(dados, p["id"], "aguardando")
    r = robo_sei.enviar_termo(dados, p, abrir=_abrir(SEIFalso(tipo_existe=False)), env=ENV)
    assert "Tipo de documento 'Termo de Responsabilidade' não existe no SEI; corrija em Textos." == r["mensagem"]
    db.marcar_passo(dados, p["id"], "aguardando")
    r = robo_sei.enviar_termo(dados, p, abrir=_abrir(SEIFalso(falhar_em="bloco")), env=ENV)
    assert r["mensagem"] == "O SEI não respondeu a tempo ao incluir no bloco de assinatura."
    assert db.termo_emitido(dados, p["termo_id"])["documento_sei"] == "1557099"          # criado antes do timeout


def test_timeout_mostra_o_aviso_do_sei_e_grava_erro_bruto(dados, tmp_path, monkeypatch):
    monkeypatch.setattr(config, "pasta_dados", lambda: tmp_path)
    p, _ = _pedido(dados)
    sei = SEIFalso(falhar_em="documento", avisos=["Informe a Descrição do documento."])
    r = robo_sei.enviar_termo(dados, p, abrir=_abrir(sei), env=ENV)
    assert r["mensagem"] == "O SEI não respondeu a tempo ao criar o documento. O SEI avisou: «Informe a Descrição do documento.»"
    nota = (tmp_path / "sei" / "erro.txt").read_text(encoding="utf-8")
    assert "passo: documento" in nota and "Timeout 30000ms exceeded em documento" in nota and "Informe a Descrição" in nota


def test_env_ausente_vira_erro_legivel(dados):
    p, _ = _pedido(dados)
    r = robo_sei.enviar_termo(dados, p, abrir=_abrir(SEIFalso()))
    assert r["passo"] == "erro" and r["mensagem"] == "Credencial do SEI não informada ao robô."


def test_tipo_de_documento_vem_dos_textos_e_devolucao_usa_a_pessoa(dados):
    semear(dados)
    dados.execute("UPDATE pessoas SET unidade_sei = 'GECONT' WHERE nome = 'ANA SILVA'"); dados.commit()
    db.incluir_processo(dados, "devolucao", "D", "3333")
    t = db.registrar_emissao(dados, "devolucao", "ANA SILVA", db.bens_da_pessoa(dados, "ANA SILVA"))
    t = db.preparar_envio_sei(dados, t["id"], agora="2026-09-20 10:00:00")
    pid = db.enfileirar_pedido(dados, "sei", termo_id=t["id"], html="<p>d</p>")
    falso = SEIFalso()
    robo_sei.enviar_termo(dados, db.pedido(dados, pid), abrir=_abrir(falso), env=ENV)
    assert ("documento_na_arvore", "Termo de Devolução 01/2026 - ANA") in falso.chamadas   # primeiro nome da pessoa, não a unidade
    assert ("incluir_documento", "Termo de Devolução", "01/2026 - ANA", len("<p>d</p>")) in falso.chamadas
    assert ("incluir_em_bloco", "1557099", "Termos GECONT") in falso.chamadas                    # bloco continua pela unidade


def test_arvore_none_nao_derruba_documento_na_arvore(monkeypatch):
    """Logo após o SEI recarregar a ifrArvore (ex.: depois de fechar o editor), arvore() pode
    devolver None por uma fração de segundo; documento_na_arvore não pode explodir com TypeError."""
    sei = robo_sei.SEI.__new__(robo_sei.SEI)
    respostas = iter([None, {"anchors": [{"id": "1", "texto": "Termo de Responsabilidade 01/2026 - CCI (1557099)"}]}])
    monkeypatch.setattr(sei, "arvore", lambda: next(respostas))
    assert sei.documento_na_arvore("Termo de Responsabilidade 01/2026 - CCI") is None
    assert sei.documento_na_arvore("Termo de Responsabilidade 01/2026 - CCI") == "1557099"


def test_rotulo_repetido_nao_e_reaproveitado_e_novo_no_e_o_que_surgiu(monkeypatch):
    """Dois Antônios (unidades diferentes, ambos 01/2026) dão o mesmo rótulo no mesmo processo: o robô não pode
    reaproveitar por rótulo, e o documento recém-criado é o nó que não existia antes do Salvar."""
    sei = robo_sei.SEI.__new__(robo_sei.SEI)
    antes = [{"id": "1", "texto": "Termo de Responsabilidade 01/2026 - ANTÔNIO (1557101)"},
             {"id": "2", "texto": "Termo de Responsabilidade 01/2026 - ANTÔNIO (1557102)"}]
    monkeypatch.setattr(sei, "arvore", lambda: {"anchors": antes})
    assert sei.documento_na_arvore("Termo de Responsabilidade 01/2026 - ANTÔNIO") is None
    numeros = sei.numeros_na_arvore()
    assert numeros == {"1557101", "1557102"}
    depois = antes + [{"id": "3", "texto": "Termo de Responsabilidade 01/2026 - ANTÔNIO (1557103)"}]
    monkeypatch.setattr(sei, "arvore", lambda: {"anchors": depois})
    assert sei.documento_novo_na_arvore("Termo de Responsabilidade 01/2026 - ANTÔNIO", numeros) == "1557103"
    assert sei.documento_novo_na_arvore("Termo de Devolução 01/2026 - ANTÔNIO", numeros) is None   # rótulo diferente: não é o nosso


def test_rotulo_casa_com_o_separador_que_o_sei_usar(monkeypatch):
    """Tipo com campo Número (Termo de Devolução): o SEI monta o rótulo com o número e o nome na árvore,
    e o separador pode não ser " - "; o robô compara ignorando separador e espaços."""
    sei = robo_sei.SEI.__new__(robo_sei.SEI)
    monkeypatch.setattr(sei, "arvore", lambda: {"anchors": [
        {"id": "1", "texto": "Termo de Devolução 01/2026 ANTÔNIO (1557201)"},
        {"id": "2", "texto": "Termo de Devolução 02/2026 - MARIA (1557202)"},
        {"id": "3", "texto": "Termo de Devolução 02/2026 - MARIANA (1557203)"}]})
    assert sei.documento_na_arvore("Termo de Devolução 01/2026 - ANTÔNIO") == "1557201"
    assert sei.documento_na_arvore("Termo de Devolução 02/2026 - MARIA") == "1557202"
    assert sei.documento_na_arvore("Termo de Devolução 03/2026 - MARIA") is None


def test_campos_do_formulario_mantem_o_padrao_numero_traco_nome():
    """Tipo com campo Número (Termo de Devolução): o SEI monta "Tipo Número NomeNaÁrvore" com um espaço; para a árvore
    ficar "01/2026 - ANTÔNIO" como nos outros tipos, o traço vai junto do nome."""
    assert robo_sei.campos_do_formulario("01/2026 - ANTÔNIO", tem_numero=True) == ("01/2026", "- ANTÔNIO")
    assert robo_sei.campos_do_formulario("01/2026 - GELAI", tem_numero=False) == ("", "01/2026 - GELAI")


def test_id_do_documento_vem_do_anchor_da_arvore(monkeypatch):
    sei = robo_sei.SEI.__new__(robo_sei.SEI)
    monkeypatch.setattr(sei, "arvore", lambda: {"anchors": [
        {"id": "555001", "texto": "90796110000022.000059/2026-88"},
        {"id": "777099", "texto": "Termo de Responsabilidade 01/2026 - CCI (1557099)"}]})
    assert sei.id_do_documento("1557099") == "777099" and sei.id_do_documento("999") is None


def test_urls_do_sei():
    assert robo_sei.url_processo("555001") == "https://sei.cfc.org.br/sei/controlador.php?acao=procedimento_trabalhar&id_procedimento=555001"
    assert robo_sei.url_documento("777099") == "https://sei.cfc.org.br/sei/controlador.php?acao=documento_visualizar&id_documento=777099"


def test_modulo_importa_sem_playwright():
    import importlib
    importlib.reload(robo_sei)          # playwright é importado só dentro de abrir_sei
