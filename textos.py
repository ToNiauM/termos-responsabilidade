"""Textos dos termos: padrão em código, alterações na tabela `textos`. Única fonte dos dizeres.

Marcadores ({nome}, {ccustos}...) são substituídos na hora de gerar o documento. Só o que difere do
padrão fica no banco; restaurar = apagar a linha.

Fundamentos citados nos dizeres (o Manual de Gestão Patrimonial do CFC é a norma interna; as demais são
aplicadas ao CFC, autarquia federal, no que couber):
- Lei n.º 4.320/1964, arts. 94 a 96: registro analítico dos bens e dos agentes responsáveis pela guarda;
  inventário analítico de cada unidade.
- IN SEDAP n.º 205/1988: item 7.11 (carga do bem mediante Termo de Responsabilidade), item 8 (inventário)
  e item 10 (responsabilidade e indenização por desaparecimento ou dano, por dolo ou culpa).
- Decreto n.º 12.785/2025 (revogou o 9.373/2018 e o 10.340/2020): circularidade, cessão, transferência e
  desfazimento de bens móveis, só pela área de patrimônio.
- Código Civil, arts. 186 e 927: reparação do dano causado por ato ilícito, doloso ou culposo.
- Rito interno de ressarcimento: apuração com contraditório e ampla defesa, deliberação da CAM-GAFI e
  homologação pelo Plenário (Manual de Gestão Patrimonial do CFC).
A unidade gestora do patrimônio (Gersev) entra nos dizeres pelos marcadores {unidade_nome} e {unidade_sigla}.
"""
import string

import db

PADRAO = {
    # gerais
    "orgao_nome": "Conselho Federal de Contabilidade",
    "orgao_sigla": "CFC",
    "cidade": "Brasília (DF)",
    "unidade_nome": "Gerência de Serviços Administrativos",
    "unidade_sigla": "Gersev",
    "assinatura_eletronica": "Assinado eletronicamente via SEI",
    "recebedor_nome": "Bruno de Araujo Gomes",
    "recebedor_cargo": "Gerente de Serviços Administrativos",
    # termo individual
    "individual_titulo": "TERMO DE RESPONSABILIDADE",
    "individual_abertura": "Pelo presente termo, eu, {nome}, declaro ter recebido do {orgao_sigla} o(s) bem(ns) patrimonial(is) abaixo discriminado(s), conferido(s) neste ato quanto à identificação patrimonial e ao estado de conservação, e que ele(s) se encontra(m) sob minha guarda, uso e responsabilidade, na forma dos arts. 94 a 96 da Lei n.º 4.320/1964, dos itens 7.11 e 10 da Instrução Normativa SEDAP n.º 205/1988 e do Manual de Gestão Patrimonial do {orgao_sigla}.",
    "individual_compromissos_intro": "Comprometo-me a:",
    "individual_compromissos": "\n".join([
        "1) zelar pela guarda, conservação e uso adequado do(s) bem(ns), empregando-o(s) exclusivamente nas atividades institucionais do {orgao_sigla} e mantendo íntegra e legível a plaqueta de identificação patrimonial;",
        "2) não ceder, emprestar, remover para outro local ou unidade, nem dar qualquer destinação ao(s) bem(ns) sem prévia autorização da {unidade_nome} ({unidade_sigla}); autorizada a movimentação, alteração ou substituição, a equipe da {unidade_sigla} emitirá novo Termo de Responsabilidade e o disponibilizará no SEI para assinatura, cabendo-lhe também a destinação dos bens inservíveis, na forma do Decreto n.º 12.785/2025;",
        "3) comunicar imediatamente à {unidade_sigla}, por escrito, qualquer dano, defeito, inutilização, extravio, furto ou roubo, registrando boletim de ocorrência policial nos casos de furto, roubo ou extravio e encaminhando-o à {unidade_sigla} com o relato dos fatos;",
        "4) responder pelo desaparecimento do(s) bem(ns) e pelo dano que, por dolo ou culpa, lhe(s) causar, ressarcindo o {orgao_sigla} do prejuízo apurado em procedimento próprio, assegurados o contraditório e a ampla defesa, após deliberação da Câmara de Gestão Administrativa-Financeira (CAM-GAFI) e homologação pelo Plenário do {orgao_sigla}, conforme o Manual de Gestão Patrimonial do {orgao_sigla}, o item 10 da IN SEDAP n.º 205/1988 e os arts. 186 e 927 do Código Civil;",
        "5) apresentar o(s) bem(ns) e prestar as informações solicitadas pela {unidade_sigla} ou pela comissão de inventário sempre que requisitado, em especial no inventário anual (Lei n.º 4.320/1964, art. 96; IN SEDAP n.º 205/1988, item 8); e",
        "6) devolver o(s) bem(ns) e seus acessórios à {unidade_sigla}, mediante Termo de Devolução, ao término do vínculo com o {orgao_sigla}, em caso de afastamento prolongado, mudança de lotação ou de função, substituição do equipamento ou sempre que solicitado, em condições compatíveis com o uso regular.",
    ]),
    "individual_ciencia": "Declaro estar ciente de que este termo constitui a carga patrimonial do(s) bem(ns) em meu nome (IN SEDAP n.º 205/1988, item 7.11), de que a responsabilidade aqui assumida permanece até a assinatura de novo Termo de Responsabilidade que o substitua ou do Termo de Devolução, e de que o descumprimento das obrigações acima me sujeita às medidas administrativas e civis cabíveis, sem prejuízo de outras previstas em lei.",
    # termo por centro de custo
    "ccusto_titulo": "Termo de Responsabilidade - {ccustos}",
    "ccusto_paragrafos": "\n\n".join([
        "Pelo presente termo, eu, {responsavel}, matrícula n.º {matricula}, {funcao} do(a) {ccustos} do {orgao_sigla}, declaro que os bens patrimoniais abaixo discriminados se encontram nas localizações indicadas, sob minha guarda e responsabilidade, na condição de agente responsável pela guarda e administração a que se refere o art. 94 da Lei n.º 4.320/1964, na forma do item 7.11 da Instrução Normativa SEDAP n.º 205/1988 e do Manual de Gestão Patrimonial do {orgao_sigla}.",
        "Comprometo-me a zelar pela guarda, conservação e uso adequado dos bens, empregando-os exclusivamente nas atividades institucionais do {orgao_sigla}; a manter íntegras e legíveis as plaquetas de identificação patrimonial; a orientar os colaboradores lotados na unidade quanto ao uso correto dos bens; e a comunicar à {unidade_nome} ({unidade_sigla}), por escrito, toda alteração ou irregularidade, em especial dano, inutilização, extravio, furto, roubo, bem sem plaqueta ou bem não relacionado neste termo.",
        "Nenhum bem será cedido, emprestado, removido para outra unidade ou localização, nem receberá qualquer destinação sem prévia autorização da {unidade_sigla}. Autorizada a movimentação, alteração ou substituição de bens, a equipe da {unidade_sigla} emitirá novo Termo de Responsabilidade atualizado, que substituirá este, e o disponibilizará no SEI para assinatura. Os bens sem uso na unidade serão devolvidos à {unidade_sigla} para reaproveitamento ou desfazimento, na forma do Decreto n.º 12.785/2025.",
        "Em caso de extravio ou dano a bem sob minha responsabilidade, decorrente de dolo ou culpa, comprometo-me a ressarcir o {orgao_sigla} do prejuízo apurado em procedimento próprio, assegurados o contraditório e a ampla defesa, após deliberação da Câmara de Gestão Administrativa-Financeira (CAM-GAFI) e homologação pelo Plenário do {orgao_sigla}, conforme o Manual de Gestão Patrimonial do {orgao_sigla}, o item 10 da IN SEDAP n.º 205/1988 e os arts. 186 e 927 do Código Civil.",
        "Observações:",
        "Em caso de furto, roubo ou extravio, o responsável registrará boletim de ocorrência policial e o encaminhará à {unidade_sigla}, com o relato dos fatos, para a instauração do procedimento de apuração.",
        "Os bens e as informações a eles relativas serão apresentados à {unidade_sigla} e à comissão de inventário sempre que solicitados, em especial no inventário anual (Lei n.º 4.320/1964, art. 96; IN SEDAP n.º 205/1988, item 8).",
        "Ao final do mandato, da função ou da designação, ou na mudança de lotação, o responsável apresentará os bens à {unidade_sigla} para conferência e transferência da carga ao sucessor, permanecendo responsável por eles até a assinatura do novo Termo de Responsabilidade emitido pela {unidade_sigla}.",
    ]),
    "ccusto_assinatura": "{responsavel}\n{funcao} do(a) {ccustos}",
    # termo de devolução
    "devolucao_titulo": "TERMO DE DEVOLUÇÃO",
    "devolucao_abertura": "Pelo presente termo, eu, {nome}, declaro que devolvo à {unidade_nome} ({unidade_sigla}) do {orgao_sigla} o(s) bem(ns) patrimonial(is) abaixo discriminado(s), com seus acessórios, que se encontrava(m) sob minha guarda e responsabilidade, ficando desonerado(a) da respectiva carga patrimonial a partir do recebimento atestado abaixo:",
    "devolucao_data": "{cidade}, {data}",
    "devolucao_recebimento": "Atesto o recebimento do(s) bem(ns) acima especificado(s), conferido(s) quanto à identificação patrimonial e ao estado de conservação, para fins de baixa da carga do(a) responsável e atualização dos registros patrimoniais do {orgao_sigla}.",
    # e-mail pedindo a assinatura no SEI (abre no programa de e-mail de quem envia)
    "email_assunto": "{termo} para assinatura no SEI - processo {processo}",
    "email_corpo": "\n".join([
        "Prezado(a) {primeiro_nome},",
        "",
        "O {termo} foi inserido no processo SEI {processo}, documento {documento}, bloco de assinatura {bloco}, e aguarda a sua assinatura.",
        "",
        "Em caso de dúvida sobre os bens relacionados, fale com a {unidade_nome} ({unidade_sigla}).",
        "",
        "Atenciosamente,",
        "{unidade_sigla}",
    ]),
}

_GERAIS = set()
_UNIDADE = {"orgao_sigla", "unidade_nome", "unidade_sigla"}
_INDIVIDUAL = {"nome"} | _UNIDADE
_CCUSTO = {"responsavel", "matricula", "funcao", "ccustos"} | _UNIDADE
_DEVOLUCAO = {"nome", "cidade", "data"} | _UNIDADE
_EMAIL = {"nome", "primeiro_nome", "termo", "processo", "documento", "bloco"} | _UNIDADE
MARCADORES = {
    "orgao_nome": _GERAIS, "orgao_sigla": _GERAIS, "cidade": _GERAIS, "unidade_nome": _GERAIS, "unidade_sigla": _GERAIS,
    "assinatura_eletronica": _GERAIS,
    "recebedor_nome": _GERAIS, "recebedor_cargo": _GERAIS,
    "individual_titulo": _INDIVIDUAL, "individual_abertura": _INDIVIDUAL, "individual_compromissos_intro": _INDIVIDUAL,
    "individual_compromissos": _INDIVIDUAL, "individual_ciencia": _INDIVIDUAL,
    "ccusto_titulo": _CCUSTO, "ccusto_paragrafos": _CCUSTO, "ccusto_assinatura": _CCUSTO,
    "devolucao_titulo": _DEVOLUCAO, "devolucao_abertura": _DEVOLUCAO, "devolucao_data": _DEVOLUCAO,
    "devolucao_recebimento": _DEVOLUCAO,
    "email_assunto": _EMAIL, "email_corpo": _EMAIL,
}

# Ordem e agrupamento da tela Textos.
GRUPOS = [
    ("Gerais", ["orgao_nome", "orgao_sigla", "unidade_nome", "unidade_sigla", "cidade", "assinatura_eletronica",
                "recebedor_nome", "recebedor_cargo"]),
    ("Termo individual", ["individual_titulo", "individual_abertura", "individual_compromissos_intro",
                          "individual_compromissos", "individual_ciencia"]),
    ("Termo por centro de custo", ["ccusto_titulo", "ccusto_paragrafos", "ccusto_assinatura"]),
    ("Termo de devolução", ["devolucao_titulo", "devolucao_abertura", "devolucao_data", "devolucao_recebimento"]),
    ("E-mail de assinatura no SEI", ["email_assunto", "email_corpo"]),
]
# Blocos que viram br-textarea; os demais são br-input.
TEXTAREA = {"individual_abertura", "individual_compromissos", "individual_ciencia", "ccusto_paragrafos",
            "ccusto_assinatura", "devolucao_abertura", "devolucao_recebimento", "email_corpo"}
ROTULOS = {
    "orgao_nome": "Nome do órgão", "orgao_sigla": "Sigla do órgão", "unidade_nome": "Unidade gestora do patrimônio — nome",
    "unidade_sigla": "Unidade gestora do patrimônio — sigla (também no cabeçalho do sistema)", "cidade": "Cidade (data do termo)",
    "assinatura_eletronica": "Texto da assinatura eletrônica", "recebedor_nome": "Quem recebe a devolução — nome",
    "recebedor_cargo": "Quem recebe a devolução — cargo",
    "individual_titulo": "Título", "individual_abertura": "Abertura", "individual_compromissos_intro": "Introdução dos compromissos",
    "individual_compromissos": "Compromissos (uma linha por item)", "individual_ciencia": "Declaração de ciência",
    "ccusto_titulo": "Título", "ccusto_paragrafos": "Parágrafos (linha em branco separa parágrafos)",
    "ccusto_assinatura": "Assinatura (uma linha por linha)",
    "devolucao_titulo": "Título", "devolucao_abertura": "Abertura", "devolucao_data": "Linha da data",
    "devolucao_recebimento": "Declaração de recebimento",
    "email_assunto": "Assunto", "email_corpo": "Corpo (enviado pelo seu programa de e-mail)",
}


def validar(chave: str, valor: str) -> None:
    if chave not in PADRAO:
        raise db.ErroDeNegocio(f"Texto '{chave}' não existe.")
    try:
        partes = list(string.Formatter().parse(valor))
    except ValueError:
        raise db.ErroDeNegocio(f"{ROTULOS[chave]}: chave {{ sem fechar ou mal formada.")
    campos = {c for _, c, _, _ in partes if c is not None}
    for _, campo, format_spec, conversion in partes:
        if campo is not None and (format_spec or conversion is not None):
            raise db.ErroDeNegocio(f"{ROTULOS[chave]}: use só {{marcador}}, sem ':' ou '!'.")
    estranhos = sorted(campos - MARCADORES[chave])
    if estranhos:
        raise db.ErroDeNegocio(f"{ROTULOS[chave]}: marcador {{{estranhos[0]}}} não existe neste bloco.")
    try:
        com_nome(valor, {m: "x" for m in MARCADORES[chave]})
    except (ValueError, KeyError, IndexError):
        raise db.ErroDeNegocio(f"{ROTULOS[chave]}: chaves {{ }} mal formadas.")


def obter(conn) -> dict:
    t = dict(PADRAO)
    for chave, valor in conn.execute("SELECT chave, valor FROM textos"):
        if chave not in t:
            continue
        try:
            validar(chave, valor)
        except db.ErroDeNegocio:
            continue  # texto inválido gravado por fora da tela: mantém o padrão
        t[chave] = valor
    return t


def salvar(conn, chave: str, valor: str) -> None:
    _gravar(conn, chave, valor)
    conn.commit()


def restaurar(conn, chave: str) -> None:
    conn.execute("DELETE FROM textos WHERE chave = ?", (chave,))
    conn.commit()


def salvar_todos(conn, valores: dict) -> None:
    """Tela Textos: valida tudo antes de gravar qualquer coisa e commita uma vez. Tudo ou nada."""
    for chave, valor in valores.items():
        validar(chave, valor)
    try:
        for chave, valor in valores.items():
            _gravar(conn, chave, valor)
        conn.commit()
    except Exception:
        conn.rollback()
        raise


def _gravar(conn, chave: str, valor: str) -> None:
    """Sem commit. Igual ao padrão = apaga a linha (só o que difere fica no banco)."""
    validar(chave, valor)
    if valor == PADRAO[chave]:
        conn.execute("DELETE FROM textos WHERE chave = ?", (chave,))
    else:
        conn.execute("INSERT OR REPLACE INTO textos VALUES (?, ?)", (chave, valor))


def campos_gerais(t: dict) -> dict:
    """Marcadores que existem em todos os blocos: siglas do órgão e da unidade gestora."""
    return {k: t[k] for k in ("orgao_sigla", "unidade_nome", "unidade_sigla")}


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


def com_nome(texto: str, campos: dict, marcador: str = "nome") -> tuple[str, str | None, str]:
    """Divide em (antes, nome, depois) para o nome sair em negrito; sem o marcador, devolve (texto, None, '')."""
    antes, achou, depois = texto.partition("{" + marcador + "}")
    if not achou:
        return antes.format_map(campos), None, ""
    return antes.format_map(campos), str(campos.get(marcador, "")), depois.format_map(campos)


_PARTICULAS = {"de", "da", "do", "das", "dos", "e"}


def nome_proprio(nome) -> str:
    """'ANTÔNIO DE SOUSA JÚNIOR' -> 'Antônio de Sousa Júnior'. O banco guarda pessoas em caixa alta (é a chave);
    nos documentos o nome sai assim."""
    saida = []
    for i, palavra in enumerate(str(nome or "").split()):
        baixa = palavra.lower()
        saida.append(baixa if i and baixa in _PARTICULAS else "-".join(p.capitalize() for p in baixa.split("-")))
    return " ".join(saida)
