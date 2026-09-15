"""Módulo de inventário: eventos (campanhas), salas, leituras por código de barras, sobras, relatório e
planilhas. Só dados: toda função recebe `conn` primeiro e não importa Flask (mesmo padrão de db.py).

Regras (spec 2026-09-15): `bens` é espelho do SPW e nunca muda aqui; "local sistema" = bens.localizacao,
"local inventário" = sala onde o bem foi lido; divergente = os dois diferem (calculado, nunca gravado)."""
import db
from db import ErroDeNegocio, _agora, _obrigatorio, _texto, _todos, _um

CONSERVACAO = ("Bom", "Regular", "Ruim", "Inservível")
ROTULO_SITUACAO = {"localizado": "Localizado", "divergente": "Divergente", "pendente": "Não localizado"}


class BemNaoEncontrado(ErroDeNegocio):
    """Número lido não existe em `bens`: a tela oferece registrar como sobra."""

    def __init__(self, numero):
        super().__init__(f"Bem {numero} não está na base.")
        self.numero = numero


# ---------------------------------------------------------------- eventos
def evento_aberto(conn) -> dict | None:
    return _um(conn, "SELECT * FROM inventario_eventos WHERE encerrado_em IS NULL")


def eventos(conn) -> list[dict]:
    return _todos(conn, "SELECT * FROM inventario_eventos ORDER BY (encerrado_em IS NULL) DESC, aberto_em DESC, id DESC")


def evento(conn, id: int) -> dict | None:
    e = _um(conn, "SELECT * FROM inventario_eventos WHERE id = ?", id)
    if e:
        e["integrantes"] = [r[0] for r in conn.execute(
            "SELECT nome FROM inventario_integrantes WHERE evento_id = ? ORDER BY nome", (id,))]
        e["resumo"] = resumo(conn, id)
    return e


def _evento_aberto_ou_erro(conn, id: int) -> dict:
    e = _um(conn, "SELECT * FROM inventario_eventos WHERE id = ?", id)
    if not e:
        raise ErroDeNegocio("Evento de inventário não encontrado.")
    if e["encerrado_em"]:
        raise ErroDeNegocio("Evento encerrado: não aceita alterações.")
    return e


def abrir_evento(conn, nome: str, descricao, integrantes: list, salas: list | None = None) -> int:
    """Um evento aberto por vez. salas=None → todas as localizações com bens ATIVO; lista → amostragem."""
    nome = _obrigatorio(nome, "Nome do evento")
    if evento_aberto(conn):
        raise ErroDeNegocio("Já existe um evento de inventário aberto; encerre-o antes de abrir outro.")
    nomes = sorted({" ".join(_texto(n).split()) for n in integrantes if _texto(n).strip()})
    if not nomes:
        raise ErroDeNegocio("Informe ao menos um integrante da comissão.")
    ativas = db.localizacoes_ativas(conn)
    escolhidas = ativas if salas is None else [s for s in ativas if s in set(salas)]
    if not escolhidas:
        raise ErroDeNegocio("Nenhuma sala com bens ativos no escopo do evento.")
    cur = conn.execute("INSERT INTO inventario_eventos (nome, descricao, aberto_em) VALUES (?,?,?)",
                       (nome, _texto(descricao) or None, _agora()))
    eid = cur.lastrowid
    conn.executemany("INSERT INTO inventario_integrantes VALUES (?,?)", [(eid, n) for n in nomes])
    conn.executemany("INSERT INTO inventario_salas (evento_id, localizacao) VALUES (?,?)", [(eid, s) for s in escolhidas])
    conn.commit()
    return eid


def encerrar_evento(conn, id: int) -> None:
    e = _um(conn, "SELECT * FROM inventario_eventos WHERE id = ?", id)
    if not e:
        raise ErroDeNegocio("Evento de inventário não encontrado.")
    if not e["encerrado_em"]:
        conn.execute("UPDATE inventario_eventos SET encerrado_em = ? WHERE id = ?", (_agora(), id))
        conn.commit()


# ---------------------------------------------------------------- salas
def salas(conn, evento_id: int) -> list[dict]:
    """Por sala do escopo: bens ativos (total), localizados aqui, divergentes lidos aqui, pendentes."""
    linhas = _todos(conn, """
        SELECT s.localizacao, l.ccustos,
          (SELECT count(*) FROM bens b WHERE b.localizacao = s.localizacao AND b.situacao = 'ATIVO') AS total,
          (SELECT count(*) FROM inventario_leituras r JOIN bens b ON b.numero = r.numero
             WHERE r.evento_id = s.evento_id AND r.localizacao = s.localizacao
               AND b.localizacao = s.localizacao AND b.situacao = 'ATIVO') AS localizados,
          (SELECT count(*) FROM inventario_leituras r JOIN bens b ON b.numero = r.numero
             WHERE r.evento_id = s.evento_id AND r.localizacao = s.localizacao AND b.localizacao <> s.localizacao
               AND b.situacao = 'ATIVO') AS divergentes
        FROM inventario_salas s LEFT JOIN localizacoes l ON l.localizacao = s.localizacao
        WHERE s.evento_id = ? ORDER BY s.localizacao""", evento_id)
    for s in linhas:
        s["pendentes"] = s["total"] - s["localizados"]
    return linhas


def _sala_ou_erro(conn, evento_id: int, localizacao: str) -> dict:
    s = _um(conn, "SELECT * FROM inventario_salas WHERE evento_id = ? AND localizacao = ?", evento_id, localizacao)
    if not s:
        raise ErroDeNegocio("Sala fora do escopo deste evento.")
    return s


def resumo(conn, evento_id: int) -> dict:
    """Progresso do evento pela quantidade de bens localizados (não há "concluir sala")."""
    ss = salas(conn, evento_id)
    bens = sum(s["total"] for s in ss)
    lidos = sum(s["localizados"] for s in ss)
    sobras = conn.execute("SELECT count(*) FROM inventario_sobras WHERE evento_id = ?", (evento_id,)).fetchone()[0]
    return {"salas": len(ss), "salas_iniciadas": sum(1 for s in ss if s["localizados"] or s["divergentes"]),
            "bens": bens, "lidos": lidos, "divergentes": sum(s["divergentes"] for s in ss),
            "pendentes": sum(s["pendentes"] for s in ss), "sobras": sobras,
            "pct_bens": round(100 * lidos / bens, 1) if bens else 0.0}


# ---------------------------------------------------------------- leituras
_LEITURA = "r.localizacao AS lido_em_sala, r.lido_em, r.integrante, r.conservacao, r.quem_usa, r.observacao, r.foto_url"


def ler(conn, evento_id: int, localizacao: str, numero: int, integrante: str) -> dict:
    """Registra (ou atualiza) a leitura do bem nesta sala. Qualquer bem cadastrado é aceito em qualquer sala;
    a divergência é só sinalizada (regra do sistema antigo)."""
    _evento_aberto_ou_erro(conn, evento_id)
    _sala_ou_erro(conn, evento_id, localizacao)
    if not conn.execute("SELECT 1 FROM inventario_integrantes WHERE evento_id = ? AND nome = ?", (evento_id, integrante)).fetchone():
        raise ErroDeNegocio("Escolha o integrante da comissão antes de ler.")
    bem = db.buscar_bem(conn, numero)
    if not bem:
        raise BemNaoEncontrado(numero)
    anterior = _um(conn, "SELECT * FROM inventario_leituras WHERE evento_id = ? AND numero = ?", evento_id, numero)
    agora = _agora()
    if anterior:
        conn.execute("UPDATE inventario_leituras SET localizacao = ?, lido_em = ?, integrante = ? WHERE id = ?",
                     (localizacao, agora, integrante, anterior["id"]))
    else:
        conn.execute("INSERT INTO inventario_leituras (evento_id, numero, localizacao, lido_em, integrante) VALUES (?,?,?,?,?)",
                     (evento_id, numero, localizacao, agora, integrante))
    conn.commit()
    return {"situacao": "localizado" if bem["localizacao"] == localizacao else "divergente", "bem": bem,
            "cadastrado_em": bem["localizacao"], "ativo": bem["situacao"] == "ATIVO", "reler": bool(anterior),
            "leitura_anterior": anterior, "lido_em": agora, "integrante": integrante}


def bens_da_sala(conn, evento_id: int, localizacao: str) -> dict:
    """bens: ativos cadastrados na sala (com a leitura do evento, se houver) e situacao_inv;
    trazidos: leituras feitas nesta sala de bens de outra sala ou não ativos; sobras: desta sala."""
    bens = _todos(conn, f"""
        SELECT b.*, {_LEITURA} FROM bens b
        LEFT JOIN inventario_leituras r ON r.numero = b.numero AND r.evento_id = ?
        WHERE b.localizacao = ? AND b.situacao = 'ATIVO' ORDER BY b.numero""", evento_id, localizacao)
    for b in bens:
        b["situacao_inv"] = "pendente" if not b["lido_em"] else ("localizado" if b["lido_em_sala"] == localizacao else "divergente")
    trazidos = _todos(conn, f"""
        SELECT b.*, {_LEITURA} FROM inventario_leituras r JOIN bens b ON b.numero = r.numero
        WHERE r.evento_id = ? AND r.localizacao = ? AND (b.localizacao <> ? OR b.situacao <> 'ATIVO')
        ORDER BY r.lido_em DESC""", evento_id, localizacao, localizacao)
    sobras = _todos(conn, "SELECT * FROM inventario_sobras WHERE evento_id = ? AND localizacao = ? ORDER BY id DESC",
                    evento_id, localizacao)
    return {"bens": bens, "trazidos": trazidos, "sobras": sobras}


def atualizar_leitura(conn, evento_id: int, numero: int, **campos) -> None:
    """Campos: conservacao, quem_usa, observacao, foto_url (só os presentes são gravados; '' vira NULL)."""
    _evento_aberto_ou_erro(conn, evento_id)
    permitidos = {"conservacao", "quem_usa", "observacao", "foto_url"}
    extra = set(campos) - permitidos
    if extra:
        raise ErroDeNegocio(f"Campo desconhecido: {', '.join(sorted(extra))}.")
    if "conservacao" in campos and campos["conservacao"] and campos["conservacao"] not in CONSERVACAO:
        raise ErroDeNegocio("Estado de conservação inválido.")
    if not _um(conn, "SELECT id FROM inventario_leituras WHERE evento_id = ? AND numero = ?", evento_id, numero):
        raise ErroDeNegocio("Leia o bem antes de preencher os dados.")
    for campo, valor in campos.items():
        conn.execute(f"UPDATE inventario_leituras SET {campo} = ? WHERE evento_id = ? AND numero = ?",
                     (_texto(valor) or None, evento_id, numero))
    conn.commit()


# ---------------------------------------------------------------- sobras
def registrar_sobra(conn, evento_id, localizacao, descricao, complemento, observacao, foto_url, integrante, exigir_foto=True) -> int:
    """Bem sem cadastro encontrado na sala. Foto obrigatória quando as fotos estão ativas (exigir_foto)."""
    _evento_aberto_ou_erro(conn, evento_id)
    _sala_ou_erro(conn, evento_id, localizacao)
    descricao = _obrigatorio(descricao, "Descrição")
    observacao = _obrigatorio(observacao, "Observação")
    integrante = _obrigatorio(integrante, "Integrante")
    if not conn.execute("SELECT 1 FROM inventario_integrantes WHERE evento_id = ? AND nome = ?", (evento_id, integrante)).fetchone():
        raise ErroDeNegocio("Escolha o integrante da comissão antes de ler.")
    if exigir_foto and not _texto(foto_url):
        raise ErroDeNegocio("A sobra precisa de foto.")
    cur = conn.execute("""INSERT INTO inventario_sobras (evento_id, localizacao, descricao, complemento, observacao, foto_url, integrante, criado_em)
                          VALUES (?,?,?,?,?,?,?,?)""",
                       (evento_id, localizacao, descricao, _texto(complemento) or None, observacao, _texto(foto_url), integrante, _agora()))
    conn.commit()
    return cur.lastrowid


def definir_foto_sobra(conn, sobra_id: int, foto_url: str) -> None:
    s = _um(conn, "SELECT evento_id FROM inventario_sobras WHERE id = ?", sobra_id)
    if not s:
        raise ErroDeNegocio("Sobra não encontrada.")
    _evento_aberto_ou_erro(conn, s["evento_id"])
    conn.execute("UPDATE inventario_sobras SET foto_url = ? WHERE id = ?", (_texto(foto_url), sobra_id))
    conn.commit()


def excluir_sobra(conn, evento_id: int, sobra_id: int) -> dict:
    """Só sobras podem ser apagadas (leituras de bens cadastrados, nunca). Devolve a sobra para apagar a foto."""
    _evento_aberto_ou_erro(conn, evento_id)
    s = _um(conn, "SELECT * FROM inventario_sobras WHERE evento_id = ? AND id = ?", evento_id, sobra_id)
    if not s:
        raise ErroDeNegocio("Sobra não encontrada.")
    conn.execute("DELETE FROM inventario_sobras WHERE id = ?", (sobra_id,))
    conn.commit()
    return s
