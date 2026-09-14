"""Corpo HTML dos três termos, no padrão do /gelic: funções pequenas que montam o HTML com html.escape.

As tabelas levam style= inline de propósito: é o que a área de transferência carrega para o SEI.
É o único lugar do projeto com inline style.
"""
import html
from datetime import date
from pathlib import Path

import config

TABELA = "border-collapse:collapse;width:{largura};margin:8pt auto;font-size:10.5pt"
TH = "border:1px solid #000;padding:3pt 5pt;text-align:center;background:#e6e6e6;font-weight:bold"
TD = "border:1px solid #000;padding:3pt 5pt;text-align:{alinhamento}"

MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
         'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']


def formatar_moeda(valor) -> str:
    return f"R$ {valor or 0:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")


def esc(s) -> str:
    return html.escape("" if s is None else str(s))


def tabela(colunas: list[str], linhas: list[list], largura: str, total: float, colunas_total: int) -> str:
    """colunas_total: quantas colunas o rótulo TOTAL ocupa (mescladas); o valor vai na última."""
    cab = "".join(f'<th style="{TH}">{esc(c)}</th>' for c in colunas)
    corpo = ""
    for linha in linhas:
        celulas = ""
        for i, v in enumerate(linha):
            alinhamento = "right" if i == len(linha) - 1 else "center" if i == 0 else "left"
            celulas += f'<td style="{TD.format(alinhamento=alinhamento)}">{esc(v)}</td>'
        corpo += f"<tr>{celulas}</tr>"
    rodape = (f'<tr><td colspan="{colunas_total}" style="{TD.format(alinhamento="center")};font-weight:bold">TOTAL</td>'
              f'<td style="{TD.format(alinhamento="right")};font-weight:bold">{formatar_moeda(total)}</td></tr>')
    return (f'<table style="{TABELA.format(largura=largura)}"><thead><tr>{cab}</tr></thead>'
            f'<tbody>{corpo}{rodape}</tbody></table>')


def _total(bens) -> float:
    return sum(b["valor_atual"] or 0 for b in bens)


COMPROMISSOS_INDIVIDUAL = [
    "1) zelar pela guarda, uso adequado e conservação do(s) bem(ns), utilizando-o(s) exclusivamente para fins profissionais do CFC;",
    "2) informar imediatamente ao Setor de Patrimônio qualquer dano, inutilização, perda ou roubo, apresentando boletim de ocorrência quando necessário;",
    "3) ressarcir o CFC por danos ou perdas decorrentes de negligência do responsável, após decisão da Câmara de Assuntos Administrativos (CAD) e homologação pelo Plenário do CFC, em conformidade com o Manual de Gestão Patrimonial do CFC;",
    "4) devolver o(s) equipamento(s) e acessórios ao término do vínculo, mediante solicitação ou em caso de substituição, em condições compatíveis com o uso; e",
    "5) fornecer informações sobre o(s) bem(ns) sempre que solicitado, especialmente durante o inventário patrimonial.",
]


def corpo_individual(nome: str, bens: list[dict]) -> str:
    linhas = [[b["numero"], b["descricao"], b["complemento"], formatar_moeda(b["valor_atual"])] for b in bens]
    return (
        "<h1>TERMO DE RESPONSABILIDADE</h1>"
        f"<p class=\"semrecuo\">Pelo presente termo, eu, <b>{esc(nome)}</b>, declaro que o(s) equipamento(s) abaixo "
        "discriminado(s) se encontra(m) sob a minha guarda e responsabilidade.</p>"
        + tabela(["Patrimônio", "Descrição", "Complemento", "Valor Atual"], linhas, "80%", _total(bens), 3)
        + "<p class=\"semrecuo\">Comprometo-me a:</p>"
        + "".join(f"<p class=\"semrecuo\">{esc(c)}</p>" for c in COMPROMISSOS_INDIVIDUAL)
        + "<p class=\"semrecuo\">Declaro estar ciente das responsabilidades mencionadas acima e assumo total "
        "responsabilidade pelos bens listados.</p>"
        f"<p class=\"assinatura\"><b>{esc(nome)}</b><br>Assinado eletronicamente via SEI</p>"
    )


PARAGRAFOS_CCUSTO = [
    "Pelo presente termo, eu, {responsavel}, matrícula n.º {matricula}, {funcao} do(a) {ccustos} do CFC, declaro que os bens patrimoniais abaixo discriminados se encontram na localização sob a minha guarda e responsabilidade.",
    "Assumo TOTAL responsabilidade pelos referidos bens, comprometendo-me a informar o Setor de Patrimônio quanto a qualquer alteração e/ou irregularidade, bem como zelar pela guarda e bom uso do patrimônio público.",
    "Em caso de extravio ou dano a bem sob a minha responsabilidade, comprometo-me a ressarcir o CFC dos prejuízos causados.",
    "Observações:",
    "Em caso de perda ou roubo do bem, o responsável deverá registrar boletim de ocorrência policial e apresentar ao Setor de Patrimônio;",
    "Ao final do mandato, função ou designação, o responsável deverá devolver o bem, se for o caso.",
    "No caso de movimentação e transferência de bens entre as unidades administrativas, o Setor de Patrimônio utilizará o Termo de Transferência disponível no SEI, que será apensado a processo específico até a emissão de um novo termo atualizado.",
]


def corpo_ccusto(ccustos: str, responsavel: dict, bens: list[dict]) -> str:
    campos = {k: esc(responsavel.get(k)) for k in ("responsavel", "matricula", "funcao")}
    campos["ccustos"] = esc(ccustos)
    linhas = [[b["numero"], b["descricao"], b["complemento"], b["localizacao"], formatar_moeda(b["valor_atual"])]
              for b in sorted(bens, key=lambda b: b["numero"])]
    return (
        f"<h1>Termo de Responsabilidade - {esc(ccustos)}</h1>"
        + "".join(f"<p>{p.format(**campos)}</p>" for p in PARAGRAFOS_CCUSTO)
        + tabela(["Número Bem", "Descrição", "Complemento", "Localização", "Valor Atual"], linhas, "100%", _total(bens), 4)
        + f"<p class=\"assinatura\">{campos['responsavel']}<br>{campos['funcao']} do(a) {campos['ccustos']} do CFC</p>"
    )


def corpo_devolucao(nome: str, bens: list[dict], hoje: date | None = None) -> str:
    hoje = hoje or date.today()
    linhas = [[b["numero"], b["descricao"], b["complemento"], formatar_moeda(b["valor_atual"])] for b in bens]
    return (
        "<h1>TERMO DE DEVOLUÇÃO</h1>"
        f"<p class=\"semrecuo\">Pelo presente termo, eu, <b>{esc(nome)}</b>, declaro que devolvi ao Setor de Patrimônio "
        "o(s) bem(ns) abaixo discriminado(s), que se encontrava(m) sob minha guarda e responsabilidade:</p>"
        + tabela(["Patrimônio", "Descrição", "Complemento", "Valor Atual"], linhas, "80%", _total(bens), 3)
        + f"<p class=\"direita\">Brasília (DF), {hoje.day} de {MESES[hoje.month - 1]} de {hoje.year}</p>"
        f"<p class=\"assinatura\"><b>{esc(nome)}</b><br>Assinado eletronicamente via SEI</p>"
        "<p class=\"semrecuo\">Declaro que recebi o(s) bem(ns) acima especificado(s):</p>"
        "<p class=\"assinatura\"><b>ANTÔNIO RODRIGUES DE SOUSA JÚNIOR</b><br>Supervisor de Patrimônio<br>"
        "Assinado eletronicamente via SEI</p>"
    )


def documento(titulo: str, corpo: str) -> str:
    base = (config.pasta_recursos() / "templates" / "termo_base.html").read_text(encoding="utf8")
    return base.replace("{{titulo}}", esc(titulo)).replace("{{corpo}}", corpo)
