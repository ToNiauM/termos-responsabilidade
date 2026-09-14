"""Textos dos termos: padrão em código, alterações na tabela `textos`. Única fonte dos dizeres.

Marcadores ({nome}, {ccustos}...) são substituídos na hora de gerar o documento. Só o que difere do
padrão fica no banco; restaurar = apagar a linha.
"""
import string

import db

PADRAO = {
    # gerais
    "orgao_nome": "Conselho Federal de Contabilidade",
    "orgao_sigla": "CFC",
    "cidade": "Brasília (DF)",
    "assinatura_eletronica": "Assinado eletronicamente via SEI",
    "recebedor_nome": "ANTÔNIO RODRIGUES DE SOUSA JÚNIOR",
    "recebedor_cargo": "Supervisor de Patrimônio",
    # termo individual
    "individual_titulo": "TERMO DE RESPONSABILIDADE",
    "individual_abertura": "Pelo presente termo, eu, {nome}, declaro que o(s) equipamento(s) abaixo discriminado(s) se encontra(m) sob a minha guarda e responsabilidade.",
    "individual_compromissos_intro": "Comprometo-me a:",
    "individual_compromissos": "\n".join([
        "1) zelar pela guarda, uso adequado e conservação do(s) bem(ns), utilizando-o(s) exclusivamente para fins profissionais do {orgao_sigla};",
        "2) informar imediatamente ao Setor de Patrimônio qualquer dano, inutilização, perda ou roubo, apresentando boletim de ocorrência quando necessário;",
        "3) ressarcir o {orgao_sigla} por danos ou perdas decorrentes de negligência do responsável, após decisão da Câmara de Assuntos Administrativos (CAD) e homologação pelo Plenário do {orgao_sigla}, em conformidade com o Manual de Gestão Patrimonial do {orgao_sigla};",
        "4) devolver o(s) equipamento(s) e acessórios ao término do vínculo, mediante solicitação ou em caso de substituição, em condições compatíveis com o uso; e",
        "5) fornecer informações sobre o(s) bem(ns) sempre que solicitado, especialmente durante o inventário patrimonial.",
    ]),
    "individual_ciencia": "Declaro estar ciente das responsabilidades mencionadas acima e assumo total responsabilidade pelos bens listados.",
    # termo por centro de custo
    "ccusto_titulo": "Termo de Responsabilidade - {ccustos}",
    "ccusto_paragrafos": "\n\n".join([
        "Pelo presente termo, eu, {responsavel}, matrícula n.º {matricula}, {funcao} do(a) {ccustos} do {orgao_sigla}, declaro que os bens patrimoniais abaixo discriminados se encontram na localização sob a minha guarda e responsabilidade.",
        "Assumo TOTAL responsabilidade pelos referidos bens, comprometendo-me a informar o Setor de Patrimônio quanto a qualquer alteração e/ou irregularidade, bem como zelar pela guarda e bom uso do patrimônio público.",
        "Em caso de extravio ou dano a bem sob a minha responsabilidade, comprometo-me a ressarcir o {orgao_sigla} dos prejuízos causados.",
        "Observações:",
        "Em caso de perda ou roubo do bem, o responsável deverá registrar boletim de ocorrência policial e apresentar ao Setor de Patrimônio;",
        "Ao final do mandato, função ou designação, o responsável deverá devolver o bem, se for o caso.",
        "No caso de movimentação e transferência de bens entre as unidades administrativas, o Setor de Patrimônio utilizará o Termo de Transferência disponível no SEI, que será apensado a processo específico até a emissão de um novo termo atualizado.",
    ]),
    "ccusto_assinatura": "{responsavel}\n{funcao} do(a) {ccustos} do {orgao_sigla}",
    # termo de devolução
    "devolucao_titulo": "TERMO DE DEVOLUÇÃO",
    "devolucao_abertura": "Pelo presente termo, eu, {nome}, declaro que devolvi ao Setor de Patrimônio o(s) bem(ns) abaixo discriminado(s), que se encontrava(m) sob minha guarda e responsabilidade:",
    "devolucao_data": "{cidade}, {data}",
    "devolucao_recebimento": "Declaro que recebi o(s) bem(ns) acima especificado(s):",
}

_GERAIS = set()
_INDIVIDUAL = {"nome", "orgao_sigla"}
_CCUSTO = {"responsavel", "matricula", "funcao", "ccustos", "orgao_sigla"}
_DEVOLUCAO = {"nome", "orgao_sigla", "cidade", "data"}
MARCADORES = {
    "orgao_nome": _GERAIS, "orgao_sigla": _GERAIS, "cidade": _GERAIS, "assinatura_eletronica": _GERAIS,
    "recebedor_nome": _GERAIS, "recebedor_cargo": _GERAIS,
    "individual_titulo": _INDIVIDUAL, "individual_abertura": _INDIVIDUAL, "individual_compromissos_intro": _INDIVIDUAL,
    "individual_compromissos": _INDIVIDUAL, "individual_ciencia": _INDIVIDUAL,
    "ccusto_titulo": _CCUSTO, "ccusto_paragrafos": _CCUSTO, "ccusto_assinatura": _CCUSTO,
    "devolucao_titulo": _DEVOLUCAO, "devolucao_abertura": _DEVOLUCAO, "devolucao_data": _DEVOLUCAO,
    "devolucao_recebimento": _DEVOLUCAO,
}

# Ordem e agrupamento da tela Textos.
GRUPOS = [
    ("Gerais", ["orgao_nome", "orgao_sigla", "cidade", "assinatura_eletronica", "recebedor_nome", "recebedor_cargo"]),
    ("Termo individual", ["individual_titulo", "individual_abertura", "individual_compromissos_intro",
                          "individual_compromissos", "individual_ciencia"]),
    ("Termo por centro de custo", ["ccusto_titulo", "ccusto_paragrafos", "ccusto_assinatura"]),
    ("Termo de devolução", ["devolucao_titulo", "devolucao_abertura", "devolucao_data", "devolucao_recebimento"]),
]
# Blocos que viram br-textarea; os demais são br-input.
TEXTAREA = {"individual_abertura", "individual_compromissos", "individual_ciencia", "ccusto_paragrafos",
            "ccusto_assinatura", "devolucao_abertura", "devolucao_recebimento"}
ROTULOS = {
    "orgao_nome": "Nome do órgão", "orgao_sigla": "Sigla do órgão", "cidade": "Cidade (data do termo)",
    "assinatura_eletronica": "Texto da assinatura eletrônica", "recebedor_nome": "Quem recebe a devolução — nome",
    "recebedor_cargo": "Quem recebe a devolução — cargo",
    "individual_titulo": "Título", "individual_abertura": "Abertura", "individual_compromissos_intro": "Introdução dos compromissos",
    "individual_compromissos": "Compromissos (uma linha por item)", "individual_ciencia": "Declaração de ciência",
    "ccusto_titulo": "Título", "ccusto_paragrafos": "Parágrafos (linha em branco separa parágrafos)",
    "ccusto_assinatura": "Assinatura (uma linha por linha)",
    "devolucao_titulo": "Título", "devolucao_abertura": "Abertura", "devolucao_data": "Linha da data",
    "devolucao_recebimento": "Declaração de recebimento",
}


def validar(chave: str, valor: str) -> None:
    if chave not in PADRAO:
        raise db.ErroDeNegocio(f"Texto '{chave}' não existe.")
    try:
        campos = {c for _, c, _, _ in string.Formatter().parse(valor) if c is not None}
    except ValueError:
        raise db.ErroDeNegocio(f"{ROTULOS[chave]}: chave {{ sem fechar ou mal formada.")
    estranhos = sorted(campos - MARCADORES[chave])
    if estranhos:
        raise db.ErroDeNegocio(f"{ROTULOS[chave]}: marcador {{{estranhos[0]}}} não existe neste bloco.")


def obter(conn) -> dict:
    t = dict(PADRAO)
    for chave, valor in conn.execute("SELECT chave, valor FROM textos"):
        if chave in t:
            t[chave] = valor
    return t


def salvar(conn, chave: str, valor: str) -> None:
    validar(chave, valor)
    if valor == PADRAO[chave]:
        restaurar(conn, chave)
        return
    conn.execute("INSERT OR REPLACE INTO textos VALUES (?, ?)", (chave, valor))
    conn.commit()


def restaurar(conn, chave: str) -> None:
    conn.execute("DELETE FROM textos WHERE chave = ?", (chave,))
    conn.commit()


def paragrafos(texto: str) -> list[str]:
    """Blocos separados por linha em branco; linhas consecutivas viram um parágrafo só."""
    saida, atual = [], []
    for linha in texto.splitlines() + [""]:
        if linha.strip():
            atual.append(linha.strip())
        elif atual:
            saida.append(" ".join(atual))
            atual = []
    return saida


def linhas(texto: str) -> list[str]:
    return [l.strip() for l in texto.splitlines() if l.strip()]


def com_nome(texto: str, campos: dict) -> tuple[str, str | None, str]:
    """Divide em (antes, nome, depois) para o nome sair em negrito; sem {nome}, devolve (texto, None, '')."""
    antes, marcador, depois = texto.partition("{nome}")
    if not marcador:
        return antes.format_map(campos), None, ""
    return antes.format_map(campos), str(campos.get("nome", "")), depois.format_map(campos)
