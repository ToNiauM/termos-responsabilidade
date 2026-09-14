"""Banco SQLite: esquema, importação de bens, consultas e cadastros.

Todas as funções recebem a conexão como primeiro argumento; quem abre e fecha é o chamador
(o Flask, por request; os testes, por fixture). Nenhuma função aqui usa Flask.
"""
import sqlite3
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

import config

ESQUEMA = """
CREATE TABLE IF NOT EXISTS bens (
  numero        INTEGER PRIMARY KEY,
  situacao      TEXT NOT NULL,
  descricao     TEXT NOT NULL,
  complemento   TEXT,
  classificacao TEXT,
  localizacao   TEXT,
  data_entrada  TEXT,
  valor_compra  REAL,
  valor_atual   REAL
);
CREATE TABLE IF NOT EXISTS responsaveis (
  ccustos     TEXT PRIMARY KEY,
  tratamento  TEXT,
  responsavel TEXT NOT NULL,
  email       TEXT,
  matricula   TEXT,
  funcao      TEXT
);
CREATE TABLE IF NOT EXISTS localizacoes (
  localizacao TEXT PRIMARY KEY,
  ccustos     TEXT NOT NULL REFERENCES responsaveis(ccustos) ON UPDATE CASCADE
);
CREATE TABLE IF NOT EXISTS pessoas (
  nome TEXT PRIMARY KEY
);
CREATE TABLE IF NOT EXISTS atribuicoes (
  nome   TEXT    NOT NULL REFERENCES pessoas(nome) ON UPDATE CASCADE ON DELETE CASCADE,
  numero INTEGER NOT NULL REFERENCES bens(numero) DEFERRABLE INITIALLY DEFERRED,
  PRIMARY KEY (nome, numero)
);
CREATE TABLE IF NOT EXISTS textos (
  chave TEXT PRIMARY KEY,
  valor TEXT NOT NULL
);
"""


class ErroDeNegocio(Exception):
    """Erro que vira mensagem para o usuário (flash), não traceback."""


def conectar(caminho: Path | None = None) -> sqlite3.Connection:
    conn = sqlite3.connect(str(caminho or config.caminho_db()))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def criar_esquema(conn: sqlite3.Connection) -> None:
    conn.executescript(ESQUEMA)


def inicializar() -> None:
    """Primeira execução do programa: pastas + esquema vazio."""
    config.preparar_pastas()
    conn = conectar()
    try:
        criar_esquema(conn)
    finally:
        conn.close()


COLUNAS_EXPORT = {
    "Número Bem": "numero",
    "Situação": "situacao",
    "Descrição": "descricao",
    "Complemento": "complemento",
    "Classificação Contábil": "classificacao",
    "Localização": "localizacao",
    "Data Entrada": "data_entrada",
    "Valor Compra": "valor_compra",
    "Valor Atual": "valor_atual",
}


class ImportacaoInvalida(ErroDeNegocio):
    pass


def _texto(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    return " ".join(str(v).split())


def _numero(v):
    try:
        return float(v) if v not in (None, "") else None
    except (TypeError, ValueError):
        return None


def _aba_do_export(wb):
    for ws in wb.worksheets:
        cabecalho = [_texto(c) for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())]
        if "Número Bem" in cabecalho:
            return ws, cabecalho
    raise ImportacaoInvalida("Nenhuma aba com a coluna 'Número Bem'. Envie o export do sistema de patrimônio.")


def importar_bens(conn: sqlite3.Connection, arquivo) -> dict:
    """Substitui a tabela `bens` pelo conteúdo do export. Tudo ou nada.

    Devolve {"total", "ativos", "sem_centro"}. Levanta ImportacaoInvalida (e não altera nada)
    se faltar coluna ou se algum bem atribuído a pessoa deixar de existir.
    """
    try:
        wb = load_workbook(arquivo, read_only=True, data_only=True)
    except Exception:
        raise ImportacaoInvalida("Arquivo inválido: envie o export do sistema em .xlsx.")

    try:
        ws, cabecalho = _aba_do_export(wb)
        faltando = [c for c in COLUNAS_EXPORT if c not in cabecalho]
        if faltando:
            raise ImportacaoInvalida("Colunas ausentes no export: " + ", ".join(faltando))
        indice = {campo: cabecalho.index(col) for col, campo in COLUNAS_EXPORT.items()}

        linhas = []
        for r in ws.iter_rows(min_row=2, values_only=True):
            num = _numero(r[indice["numero"]])
            if num is None:
                continue
            linhas.append((
                int(num), _texto(r[indice["situacao"]]), _texto(r[indice["descricao"]]),
                _texto(r[indice["complemento"]]), _texto(r[indice["classificacao"]]),
                _texto(r[indice["localizacao"]]), _texto(r[indice["data_entrada"]]),
                _numero(r[indice["valor_compra"]]), _numero(r[indice["valor_atual"]]),
            ))
    except ImportacaoInvalida:
        raise
    except Exception:
        raise ImportacaoInvalida("Não consegui ler a planilha; o arquivo pode estar corrompido.")
    finally:
        wb.close()

    try:
        conn.execute("DELETE FROM bens")
        conn.executemany("INSERT INTO bens VALUES (?,?,?,?,?,?,?,?,?)", linhas)
        orfaos = [str(r[0]) for r in conn.execute(
            "SELECT numero FROM atribuicoes WHERE numero NOT IN (SELECT numero FROM bens) ORDER BY numero")]
        if orfaos:
            raise ImportacaoInvalida(
                "O export não traz bens que estão atribuídos a pessoas: " + ", ".join(orfaos)
                + ". Remova a atribuição na aba Pessoas ou use um export completo.")
        conn.commit()
    except sqlite3.IntegrityError:
        conn.rollback()
        raise ImportacaoInvalida("O export tem número de bem repetido; corrija a planilha e envie de novo.")
    except Exception:
        conn.rollback()
        raise

    total = conn.execute("SELECT count(*) FROM bens").fetchone()[0]
    ativos = conn.execute("SELECT count(*) FROM bens WHERE situacao='ATIVO'").fetchone()[0]
    return {"total": total, "ativos": ativos, "sem_centro": localizacoes_sem_centro(conn)}


def localizacoes_sem_centro(conn: sqlite3.Connection) -> list[str]:
    """Localizações de bens ATIVOS que não têm centro de custo mapeado."""
    return [r[0] for r in conn.execute(
        "SELECT DISTINCT localizacao FROM bens WHERE situacao='ATIVO' AND localizacao <> '' "
        "AND localizacao NOT IN (SELECT localizacao FROM localizacoes) ORDER BY localizacao")]


def _todos(conn, sql, *args) -> list[dict]:
    return [dict(r) for r in conn.execute(sql, args)]


def _um(conn, sql, *args) -> dict | None:
    r = conn.execute(sql, args).fetchone()
    return dict(r) if r else None


def centros(conn) -> list[dict]:
    return _todos(conn, "SELECT * FROM responsaveis ORDER BY ccustos")


def responsavel(conn, ccustos: str) -> dict | None:
    return _um(conn, "SELECT * FROM responsaveis WHERE ccustos = ?", ccustos)


def bens_do_centro(conn, ccustos: str) -> list[dict]:
    """Bens ATIVOS nas localizações do centro, excluindo os atribuídos a pessoas (setor OU pessoa)."""
    return _todos(conn, """
        SELECT b.* FROM bens b JOIN localizacoes l ON l.localizacao = b.localizacao
        WHERE l.ccustos = ? AND b.situacao = 'ATIVO'
          AND b.numero NOT IN (SELECT numero FROM atribuicoes)
        ORDER BY b.numero""", ccustos)


def pessoas(conn) -> list[str]:
    return [r[0] for r in conn.execute("SELECT nome FROM pessoas ORDER BY nome")]


def bens_da_pessoa(conn, nome: str) -> list[dict]:
    return _todos(conn, """
        SELECT b.* FROM atribuicoes a JOIN bens b ON b.numero = a.numero
        WHERE a.nome = ? ORDER BY b.numero""", nome)


def buscar_bem(conn, numero: int) -> dict | None:
    return _um(conn, "SELECT * FROM bens WHERE numero = ?", numero)


def pessoa_do_bem(conn, numero: int) -> str | None:
    r = conn.execute("SELECT nome FROM atribuicoes WHERE numero = ?", (numero,)).fetchone()
    return r[0] if r else None


def ficha_do_bem(conn, numero: int) -> dict | None:
    """Bem + centro de custo/responsável do setor (pela localização) + pessoa (pela atribuição)."""
    bem = buscar_bem(conn, numero)
    if not bem:
        return None
    setor = _um(conn, """
        SELECT r.ccustos, r.responsavel FROM localizacoes l JOIN responsaveis r ON r.ccustos = l.ccustos
        WHERE l.localizacao = ?""", bem["localizacao"]) or {"ccustos": None, "responsavel": None}
    return {**bem, **setor, "pessoa": pessoa_do_bem(conn, numero)}


def localizacoes_mapeadas(conn) -> list[dict]:
    return _todos(conn, "SELECT localizacao, ccustos FROM localizacoes ORDER BY localizacao")


class CentroEmUso(ErroDeNegocio):
    pass


class BemNaoEncontrado(ErroDeNegocio):
    pass


class JaAtribuido(ErroDeNegocio):
    def __init__(self, numero, pessoa):
        super().__init__(f"O bem {numero} está com {pessoa}.")
        self.pessoa = pessoa


def _obrigatorio(valor, rotulo) -> str:
    v = " ".join(str(valor or "").split())
    if not v:
        raise ErroDeNegocio(f"{rotulo} é obrigatório.")
    return v


def incluir_responsavel(conn, dados: dict) -> None:
    sigla = _obrigatorio(dados.get("ccustos"), "Centro de custo").upper()
    nome = _obrigatorio(dados.get("responsavel"), "Responsável")
    if responsavel(conn, sigla):
        raise ErroDeNegocio(f"O centro de custo {sigla} já existe.")
    conn.execute("INSERT INTO responsaveis VALUES (?,?,?,?,?,?)", (
        sigla, _texto(dados.get("tratamento")), nome, _texto(dados.get("email")),
        _texto(dados.get("matricula")), _texto(dados.get("funcao"))))
    conn.commit()


def checar_exclusao_centro(conn, ccustos: str) -> None:
    """Centro só pode ser excluído sem bens ativos sob guarda (bens em localizações dele, não atribuídos)."""
    if not responsavel(conn, ccustos):
        raise ErroDeNegocio(f"Centro de custo {ccustos} não encontrado.")
    n = len(bens_do_centro(conn, ccustos))
    if n:
        bem_str = "bem" if n == 1 else "bens"
        ativo_str = "ativo" if n == 1 else "ativos"
        raise CentroEmUso(f"{ccustos} tem {n} {bem_str} {ativo_str} sob guarda. "
                          "Mova as localizações para outro centro antes de excluir.")


def excluir_responsavel(conn, ccustos: str) -> None:
    checar_exclusao_centro(conn, ccustos)
    conn.execute("DELETE FROM localizacoes WHERE ccustos = ?", (ccustos,))  # voltam a "pendentes"
    conn.execute("DELETE FROM responsaveis WHERE ccustos = ?", (ccustos,))
    conn.commit()


def atualizar_responsavel(conn, ccustos: str, dados: dict) -> None:
    if not responsavel(conn, ccustos):
        raise ErroDeNegocio(f"Centro de custo {ccustos} não encontrado.")
    nome = _obrigatorio(dados.get("responsavel"), "Responsável")
    conn.execute("UPDATE responsaveis SET tratamento=?, responsavel=?, email=?, matricula=?, funcao=? WHERE ccustos=?", (
        _texto(dados.get("tratamento")), nome, _texto(dados.get("email")),
        _texto(dados.get("matricula")), _texto(dados.get("funcao")), ccustos))
    conn.commit()


def renomear_centro(conn, antigo: str, novo: str) -> None:
    novo = _obrigatorio(novo, "Nova sigla").upper()
    if responsavel(conn, novo):
        raise ErroDeNegocio(f"Já existe um centro de custo {novo}.")
    if not responsavel(conn, antigo):
        raise ErroDeNegocio(f"Centro de custo {antigo} não encontrado.")
    conn.execute("UPDATE responsaveis SET ccustos = ? WHERE ccustos = ?", (novo, antigo))  # cascateia
    conn.commit()


def incluir_localizacao(conn, localizacao: str, ccustos: str) -> None:
    loc = _obrigatorio(localizacao, "Localização")
    if not responsavel(conn, ccustos):
        raise ErroDeNegocio(f"Centro de custo {ccustos} não cadastrado.")
    conn.execute("INSERT OR REPLACE INTO localizacoes VALUES (?, ?)", (loc, ccustos))
    conn.commit()


def excluir_localizacao(conn, localizacao: str) -> None:
    conn.execute("DELETE FROM localizacoes WHERE localizacao = ?", (localizacao,))
    conn.commit()


def mover_localizacoes(conn, localizacoes: list[str], ccustos: str) -> int:
    """De-Para: muda o centro de custo de uma ou várias localizações de uma vez."""
    if not localizacoes:
        raise ErroDeNegocio("Selecione ao menos uma localização.")
    if not responsavel(conn, ccustos):
        raise ErroDeNegocio(f"Centro de custo {ccustos} não cadastrado.")
    marcadores = ",".join("?" * len(localizacoes))
    cur = conn.execute(f"UPDATE localizacoes SET ccustos = ? WHERE localizacao IN ({marcadores})", (ccustos, *localizacoes))
    conn.commit()
    return cur.rowcount


def incluir_pessoa(conn, nome: str) -> str:
    nome = _obrigatorio(nome, "Nome").upper()
    conn.execute("INSERT OR IGNORE INTO pessoas VALUES (?)", (nome,))
    conn.commit()
    return nome


def excluir_pessoa(conn, nome: str) -> None:
    conn.execute("DELETE FROM pessoas WHERE nome = ?", (nome,))  # atribuições vão junto (cascade)
    conn.commit()


def renomear_pessoa(conn, antigo: str, novo: str) -> str:
    novo = _obrigatorio(novo, "Nome").upper()
    if antigo not in pessoas(conn):
        raise ErroDeNegocio(f"Pessoa {antigo} não encontrada.")
    if novo == antigo:
        return novo
    if novo in pessoas(conn):
        raise ErroDeNegocio(f"Já existe uma pessoa chamada {novo}.")
    conn.execute("UPDATE pessoas SET nome = ? WHERE nome = ?", (novo, antigo))  # cascateia em atribuicoes
    conn.commit()
    return novo


def atribuir(conn, nome: str, numero: int, confirmar: bool = False) -> None:
    """Coloca o bem sob responsabilidade da pessoa. Se já está com outra, exige confirmar=True."""
    if not buscar_bem(conn, numero):
        raise BemNaoEncontrado(f"Bem {numero} não encontrado na base.")
    atual = pessoa_do_bem(conn, numero)
    if atual == nome:
        return
    if atual and not confirmar:
        raise JaAtribuido(numero, atual)
    conn.execute("DELETE FROM atribuicoes WHERE numero = ?", (numero,))
    conn.execute("INSERT INTO atribuicoes VALUES (?, ?)", (nome, numero))
    conn.commit()


def desatribuir(conn, nome: str, numero: int) -> None:
    conn.execute("DELETE FROM atribuicoes WHERE nome = ? AND numero = ?", (nome, numero))
    conn.commit()


CADASTROS = {
    "responsaveis": ["ccustos", "tratamento", "responsavel", "email", "matricula", "funcao"],
    "localizacoes": ["localizacao", "ccustos"],
    "pessoas": ["nome"],
    "atribuicoes": ["nome", "numero"],
}


def exportar_cadastros(conn, destino: Path) -> Path:
    """Planilha com as 4 tabelas de cadastro, no formato do banco (para backup e edição em massa)."""
    wb = Workbook()
    wb.remove(wb.active)
    for tabela, colunas in CADASTROS.items():
        ws = wb.create_sheet(tabela)
        ws.append(colunas)
        for linha in conn.execute(f"SELECT {', '.join(colunas)} FROM {tabela} ORDER BY {colunas[0]}"):
            ws.append(list(linha))
    wb.save(str(destino))
    return Path(destino)


def _ler_aba_cadastro(wb, tabela: str, problemas: list) -> list[dict]:
    colunas = CADASTROS[tabela]
    if tabela not in wb.sheetnames:
        problemas.append(f"aba '{tabela}' não encontrada")
        return []
    ws = wb[tabela]
    it = ws.iter_rows(values_only=True)
    cabecalho = [_texto(c) for c in next(it, ())]
    faltando = [c for c in colunas if c not in cabecalho]
    if faltando:
        problemas.append(f"aba '{tabela}': coluna(s) ausente(s): {', '.join(faltando)}")
        return []
    idx = [cabecalho.index(c) for c in colunas]
    linhas = []
    for n, r in enumerate(it, start=2):
        if r is None or all(v is None or _texto(v) == "" for v in r):
            continue
        linhas.append({"_linha": n, **{c: r[i] if i < len(r) else None for c, i in zip(colunas, idx)}})
    return linhas


def importar_cadastros(conn, arquivo) -> dict:
    """Substitui responsaveis, localizacoes, pessoas e atribuicoes pelo conteúdo da planilha. Tudo ou nada."""
    try:
        wb = load_workbook(arquivo, read_only=True, data_only=True)
    except Exception:
        raise ImportacaoInvalida("Arquivo inválido: envie a planilha de cadastros em .xlsx.")
    problemas: list[str] = []
    try:
        brutos = {t: _ler_aba_cadastro(wb, t, problemas) for t in CADASTROS}
    except ImportacaoInvalida:
        raise
    except Exception:
        raise ImportacaoInvalida("Não consegui ler a planilha; o arquivo pode estar corrompido.")
    finally:
        wb.close()
    if problemas:
        raise ImportacaoInvalida("Planilha de cadastros: " + "; ".join(problemas))

    responsaveis, siglas = [], set()
    for r in brutos["responsaveis"]:
        sigla, nome = _texto(r["ccustos"]).upper(), _texto(r["responsavel"])
        if not sigla:
            problemas.append(f"responsaveis linha {r['_linha']}: sigla vazia")
        elif sigla in siglas:
            problemas.append(f"responsaveis linha {r['_linha']}: sigla {sigla} repetida")
        elif not nome:
            problemas.append(f"responsaveis linha {r['_linha']}: responsável vazio")
        else:
            siglas.add(sigla)
            responsaveis.append((sigla, _texto(r["tratamento"]), nome, _texto(r["email"]), _texto(r["matricula"]), _texto(r["funcao"])))

    localizacoes, locs = [], set()
    for r in brutos["localizacoes"]:
        loc, sigla = _texto(r["localizacao"]), _texto(r["ccustos"]).upper()
        if not loc:
            problemas.append(f"localizacoes linha {r['_linha']}: localização vazia")
        elif loc in locs:
            problemas.append(f"localizacoes linha {r['_linha']}: localização {loc} repetida")
        elif sigla not in siglas:
            problemas.append(f"localizacoes linha {r['_linha']}: centro {sigla or '(vazio)'} não está na aba responsaveis")
        else:
            locs.add(loc)
            localizacoes.append((loc, sigla))

    nomes = set()
    for r in brutos["pessoas"]:
        nome = _texto(r["nome"]).upper()
        if not nome:
            problemas.append(f"pessoas linha {r['_linha']}: nome vazio")
        elif nome in nomes:
            problemas.append(f"pessoas linha {r['_linha']}: nome {nome} repetido")
        else:
            nomes.add(nome)

    atribuicoes, numeros = [], set()
    for r in brutos["atribuicoes"]:
        nome, num = _texto(r["nome"]).upper(), _numero(r["numero"])
        if nome not in nomes:
            problemas.append(f"atribuicoes linha {r['_linha']}: {nome or '(vazio)'} não está na aba pessoas")
        elif num is None or num != int(num):
            problemas.append(f"atribuicoes linha {r['_linha']}: número inválido")
        elif not buscar_bem(conn, int(num)):
            problemas.append(f"atribuicoes linha {r['_linha']}: bem {int(num)} não existe na base")
        elif int(num) in numeros:
            problemas.append(f"atribuicoes linha {r['_linha']}: bem {int(num)} repetido (um bem, uma pessoa)")
        else:
            numeros.add(int(num))
            atribuicoes.append((nome, int(num)))

    if problemas:
        extra = f" (+{len(problemas) - 20})" if len(problemas) > 20 else ""
        raise ImportacaoInvalida("Planilha de cadastros: " + "; ".join(problemas[:20]) + extra)

    try:
        for t in ("atribuicoes", "pessoas", "localizacoes", "responsaveis"):
            conn.execute(f"DELETE FROM {t}")
        conn.executemany("INSERT INTO responsaveis VALUES (?,?,?,?,?,?)", responsaveis)
        conn.executemany("INSERT INTO localizacoes VALUES (?,?)", localizacoes)
        conn.executemany("INSERT INTO pessoas VALUES (?)", [(n,) for n in sorted(nomes)])
        conn.executemany("INSERT INTO atribuicoes VALUES (?,?)", atribuicoes)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    return {"responsaveis": len(responsaveis), "localizacoes": len(localizacoes), "pessoas": len(nomes),
            "atribuicoes": len(atribuicoes), "sem_centro": localizacoes_sem_centro(conn)}
