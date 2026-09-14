"""Corpo HTML dos três termos, no padrão do /gelic: funções pequenas que montam o HTML com html.escape.

As tabelas levam style= inline de propósito: é o que a área de transferência carrega para o SEI.
É o único lugar do projeto com inline style.
"""
import html
from datetime import date

import config
import textos as textos_mod

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


def data_por_extenso(hoje: date) -> str:
    return f"{hoje.day} de {MESES[hoje.month - 1]} de {hoje.year}"


def _p(texto: str, classe: str = "semrecuo") -> str:
    return f'<p class="{classe}">{esc(texto)}</p>'


def _abertura(texto: str, campos: dict) -> str:
    antes, nome, depois = textos_mod.com_nome(texto, campos)
    meio = f"<b>{esc(nome)}</b>" if nome is not None else ""
    return f'<p class="semrecuo">{esc(antes)}{meio}{esc(depois)}</p>'


def corpo_individual(nome: str, bens: list[dict], textos: dict | None = None) -> str:
    t = textos or textos_mod.PADRAO
    campos = {"nome": nome, "orgao_sigla": t["orgao_sigla"]}
    linhas = [[b["numero"], b["descricao"], b["complemento"], formatar_moeda(b["valor_atual"])] for b in bens]
    return (
        f"<h1>{esc(t['individual_titulo'].format_map(campos))}</h1>"
        + _abertura(t["individual_abertura"], campos)
        + tabela(["Patrimônio", "Descrição", "Complemento", "Valor Atual"], linhas, "80%", _total(bens), 3)
        + _p(t["individual_compromissos_intro"].format_map(campos))
        + "".join(_p(c.format_map(campos)) for c in textos_mod.linhas(t["individual_compromissos"]))
        + _p(t["individual_ciencia"].format_map(campos))
        + f'<p class="assinatura"><b>{esc(nome)}</b><br>{esc(t["assinatura_eletronica"])}</p>'
    )


def corpo_ccusto(ccustos: str, responsavel: dict, bens: list[dict], textos: dict | None = None) -> str:
    t = textos or textos_mod.PADRAO
    campos = {k: responsavel.get(k) or "" for k in ("responsavel", "matricula", "funcao")}
    campos.update(ccustos=ccustos, orgao_sigla=t["orgao_sigla"])
    linhas = [[b["numero"], b["descricao"], b["complemento"], b["localizacao"], formatar_moeda(b["valor_atual"])]
              for b in sorted(bens, key=lambda b: b["numero"])]
    assinatura = "<br>".join(esc(l.format_map(campos)) for l in textos_mod.linhas(t["ccusto_assinatura"]))
    return (
        f"<h1>{esc(t['ccusto_titulo'].format_map(campos))}</h1>"
        + "".join(f"<p>{esc(p.format_map(campos))}</p>" for p in textos_mod.paragrafos(t["ccusto_paragrafos"]))
        + tabela(["Número Bem", "Descrição", "Complemento", "Localização", "Valor Atual"], linhas, "100%", _total(bens), 4)
        + f'<p class="assinatura">{assinatura}</p>'
    )


def corpo_devolucao(nome: str, bens: list[dict], hoje: date | None = None, textos: dict | None = None) -> str:
    t = textos or textos_mod.PADRAO
    hoje = hoje or date.today()
    campos = {"nome": nome, "orgao_sigla": t["orgao_sigla"], "cidade": t["cidade"], "data": data_por_extenso(hoje)}
    linhas = [[b["numero"], b["descricao"], b["complemento"], formatar_moeda(b["valor_atual"])] for b in bens]
    return (
        f"<h1>{esc(t['devolucao_titulo'].format_map(campos))}</h1>"
        + _abertura(t["devolucao_abertura"], campos)
        + tabela(["Patrimônio", "Descrição", "Complemento", "Valor Atual"], linhas, "80%", _total(bens), 3)
        + _p(t["devolucao_data"].format_map(campos), "direita")
        + f'<p class="assinatura"><b>{esc(nome)}</b><br>{esc(t["assinatura_eletronica"])}</p>'
        + _p(t["devolucao_recebimento"].format_map(campos))
        + f'<p class="assinatura"><b>{esc(t["recebedor_nome"])}</b><br>{esc(t["recebedor_cargo"])}<br>'
        f'{esc(t["assinatura_eletronica"])}</p>'
    )


def documento(titulo: str, corpo: str) -> str:
    base = (config.pasta_recursos() / "templates" / "termo_base.html").read_text(encoding="utf8")
    return base.replace("{{corpo}}", corpo).replace("{{titulo}}", esc(titulo))
