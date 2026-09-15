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
             WHERE r.evento_id = s.evento_id AND r.localizacao = s.localizacao AND b.localizacao <> s.localizacao) AS divergentes
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
