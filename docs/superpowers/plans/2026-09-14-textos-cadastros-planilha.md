# Textos editáveis, cadastros completos e planilha de cadastros — Plano de implementação

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Tirar os textos dos termos do código (editáveis no app, com marcadores), completar a edição de cadastros com um De-Para de verdade, e permitir exportar/importar os cadastros em planilha.

**Architecture:** `textos.py` é a única fonte dos dizeres (padrão em código, alterações na tabela `textos`); geradores `.docx` e `termos_html.py` recebem o dicionário e param de ter texto próprio. `db.py` ganha as operações de cadastro que faltavam e a exportação/importação das 4 tabelas de cadastro (tudo ou nada, como `importar_bens`). `app.py` ganha a tela Textos, páginas de edição e as rotas de mover/exportar/importar.

**Tech Stack:** Python 3.14, Flask 3.1, sqlite3, openpyxl, python-docx, pytest — nada novo. DSGov 3.7.0 já vendorizado.

**Spec:** `docs/superpowers/specs/2026-09-14-textos-cadastros-planilha-design.md`

## Global Constraints

- Simplicidade: adaptar o que existe; nenhuma dependência nova; nada de CRUD genérico.
- Nenhum `<style>`, hex ou `<select>` nativo nos templates DSGov; um único `br-button primary` visível por tela; um `h1` por página. Exceção de `style=` inline continua restrita a `termos_html.py`/`termo_base.html`.
- Marcadores permitidos por bloco exatamente como na spec §3.2; `{nome}` nos blocos `*_abertura` sai em negrito; linha em branco separa parágrafos; uma linha por compromisso.
- Formatação Word (fontes, tabela, larguras 80 %/100 %, TOTAL, data por extenso) **não muda** — só a origem do texto.
- Excluir centro de custo: bloqueado se tiver bens ATIVOS sob guarda (`db.bens_do_centro` não vazio); senão apaga o centro e seus mapeamentos.
- Importar cadastros substitui as 4 tabelas inteiras, tudo ou nada; a planilha é a verdade.
- Código, comentários, mensagens e commits em português. Rodar com `.venv/bin/python` / `.venv/bin/pytest`; saída de testes limpa (sem warnings).
- Nunca iniciar servidor nas portas 5000/5001 nesta VPS; nunca rodar `importar_planilhas.py`.
- Suíte atual: 41 testes verdes na `main` (`08b14e2`).

---

## Estrutura de arquivos

| Arquivo | Responsabilidade |
|---|---|
| `textos.py` (novo) | `PADRAO`, `MARCADORES`, `GRUPOS`, `TEXTAREA`, `obter`, `validar`, `salvar`, `restaurar`, `paragrafos`, `linhas`, `com_nome` |
| `db.py` | + tabela `textos` no `ESQUEMA`; `atualizar_responsavel`, `checar_exclusao_centro`, `excluir_responsavel` (regra nova), `mover_localizacoes`, `renomear_pessoa`, `exportar_cadastros`, `importar_cadastros` |
| `termos_html.py` | `corpo_*` recebem `textos` (opcional, default `PADRAO`) |
| `Script_Termo_Individual.py`, `Termo_de_Responsabilidade.py`, `termo_devolucao.py` | idem; sem texto próprio |
| `app.py` | rotas `/textos`, editar responsável/pessoa, mover, exportar, importar-cadastros; `DSGOV` lê `orgao_nome` |
| `templates/textos.html`, `editar_responsavel.html`, `editar_pessoa.html` (novos) | telas |
| `templates/cadastros.html`, `upload.html`, `_macros.html`, `base.html` | ajustes |
| `tests/test_textos.py` (novo); `test_db.py`, `test_app.py`, `test_docx.py`, `test_termos_html.py` | ampliados |

---

### Task 1: `textos.py` — padrão, marcadores, tabela e API

**Files:**
- Create: `textos.py`, `tests/test_textos.py`
- Modify: `db.py` (`ESQUEMA`, acrescentar a tabela `textos`)

**Interfaces:**
- Produces: `textos.PADRAO: dict[str, str]`; `textos.MARCADORES: dict[str, set[str]]`; `textos.GRUPOS: list[tuple[str, list[str]]]`; `textos.TEXTAREA: set[str]`; `textos.obter(conn) -> dict`; `textos.validar(chave, valor) -> None` (levanta `db.ErroDeNegocio`); `textos.salvar(conn, chave, valor) -> None`; `textos.restaurar(conn, chave) -> None`; `textos.paragrafos(texto) -> list[str]`; `textos.linhas(texto) -> list[str]`; `textos.com_nome(texto, campos) -> tuple[str, str | None, str]` (antes, nome-ou-None, depois — já formatados).

- [ ] **Step 1: Testes que falham**

`tests/test_textos.py`:

```python
import pytest

import db
import textos


def test_obter_devolve_padrao_e_sobreposicao(dados):
    t = textos.obter(dados)
    assert t == textos.PADRAO
    textos.salvar(dados, "individual_titulo", "TERMO X")
    assert textos.obter(dados)["individual_titulo"] == "TERMO X"
    assert textos.obter(dados)["orgao_sigla"] == "CFC"


def test_validar_recusa_marcador_desconhecido_e_chave_mal_formada():
    with pytest.raises(db.ErroDeNegocio) as e:
        textos.validar("individual_abertura", "Eu, {nomee}, declaro")
    assert "nomee" in str(e.value)
    with pytest.raises(db.ErroDeNegocio):
        textos.validar("individual_abertura", "Eu, {nome, declaro")
    with pytest.raises(db.ErroDeNegocio):
        textos.validar("chave_inexistente", "x")
    textos.validar("individual_abertura", "Eu, {nome}, do {orgao_sigla}")  # ok
    textos.validar("orgao_sigla", "")  # vazio permitido


def test_salvar_igual_ao_padrao_nao_deixa_linha(dados):
    textos.salvar(dados, "cidade", "Brasília (DF)")
    assert dados.execute("SELECT count(*) FROM textos").fetchone()[0] == 0
    textos.salvar(dados, "cidade", "Goiânia (GO)")
    assert dados.execute("SELECT count(*) FROM textos").fetchone()[0] == 1
    textos.restaurar(dados, "cidade")
    assert textos.obter(dados)["cidade"] == "Brasília (DF)"


def test_paragrafos_e_linhas():
    assert textos.paragrafos("a\nb\n\n\nc\n") == ["a b", "c"]
    assert textos.linhas("1) x\n\n2) y\n") == ["1) x", "2) y"]
    assert textos.paragrafos("") == []


def test_com_nome():
    assert textos.com_nome("Eu, {nome}, do {orgao_sigla}.", {"nome": "ANA", "orgao_sigla": "CFC"}) == ("Eu, ", "ANA", ", do CFC.")
    assert textos.com_nome("Sem marcador.", {"nome": "ANA"}) == ("Sem marcador.", None, "")


def test_padrao_cobre_todos_os_grupos_e_marcadores():
    chaves = {c for _, lista in textos.GRUPOS for c in lista}
    assert chaves == set(textos.PADRAO) == set(textos.MARCADORES)
    for chave, valor in textos.PADRAO.items():
        textos.validar(chave, valor)  # o padrão tem de passar na própria validação
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_textos.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'textos'`

- [ ] **Step 3: Tabela no esquema**

Em `db.py`, acrescentar ao final de `ESQUEMA` (antes das aspas de fechamento):

```sql
CREATE TABLE IF NOT EXISTS textos (
  chave TEXT PRIMARY KEY,
  valor TEXT NOT NULL
);
```

- [ ] **Step 4: `textos.py`**

```python
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
```

- [ ] **Step 5: Rodar e ver passar**

Run: `.venv/bin/pytest tests/test_textos.py tests/test_db.py -q`
Expected: todos passam (6 novos + os existentes).

- [ ] **Step 6: Commit**

```bash
git add textos.py db.py tests/test_textos.py
git commit -m "textos: padrão dos dizeres, marcadores por bloco e tabela textos"
```

---

### Task 2: `termos_html.py` lê os textos

**Files:**
- Modify: `termos_html.py`, `tests/test_termos_html.py`

**Interfaces:**
- Consumes: `textos.PADRAO`, `textos.paragrafos`, `textos.linhas`, `textos.com_nome`.
- Produces: `corpo_individual(nome, bens, textos=None)`, `corpo_ccusto(ccustos, responsavel, bens, textos=None)`, `corpo_devolucao(nome, bens, hoje=None, textos=None)`; `textos=None` usa `PADRAO`. `data_por_extenso(hoje) -> str` ("14 de setembro de 2026").

- [ ] **Step 1: Testes que falham**

Acrescentar a `tests/test_termos_html.py`:

```python
import textos


def test_individual_com_texto_alterado_e_nome_em_negrito():
    t = dict(textos.PADRAO, individual_abertura="TESTE {nome} do {orgao_sigla}.",
             individual_compromissos="a\nb\nc", orgao_sigla="XYZ")
    html = th.corpo_individual("ANA SILVA", bens(), textos=t)
    assert "TESTE <b>ANA SILVA</b> do XYZ." in html
    assert html.count("<p class=\"semrecuo\">a</p>") == 1 and "<p class=\"semrecuo\">c</p>" in html


def test_ccusto_com_dois_paragrafos_e_sigla():
    t = dict(textos.PADRAO, ccusto_paragrafos="Primeiro {ccustos}.\n\nSegundo do {orgao_sigla}.", orgao_sigla="XYZ")
    resp = {"ccustos": "CCI", "responsavel": "JAQUELINE", "matricula": "46", "funcao": "coordenadora"}
    html = th.corpo_ccusto("CCI", resp, bens(), textos=t)
    assert "<p>Primeiro CCI.</p><p>Segundo do XYZ.</p>" in html
    assert "coordenadora do(a) CCI do XYZ" in html


def test_devolucao_usa_cidade_e_recebedor_dos_textos():
    t = dict(textos.PADRAO, cidade="Goiânia (GO)", recebedor_nome="FULANO", recebedor_cargo="Chefe")
    html = th.corpo_devolucao("ANA SILVA", bens(), hoje=date(2026, 9, 14), textos=t)
    assert "Goiânia (GO), 14 de setembro de 2026" in html and "<b>FULANO</b><br>Chefe<br>" in html


def test_escapa_texto_vindo_do_banco():
    t = dict(textos.PADRAO, individual_ciencia="<script>x</script>")
    assert "&lt;script&gt;" in th.corpo_individual("A", [], textos=t)
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_termos_html.py -v`
Expected: FAIL — `TypeError: ... unexpected keyword argument 'textos'`

- [ ] **Step 3: Reescrever as funções de corpo**

Em `termos_html.py`: `import textos as textos_mod` (evita conflito com o parâmetro). Apagar `COMPROMISSOS_INDIVIDUAL` e `PARAGRAFOS_CCUSTO`. Substituir as três funções de corpo por:

```python
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
```

- [ ] **Step 4: Rodar e ver passar**

Run: `.venv/bin/pytest tests/test_termos_html.py tests/test_app.py -q`
Expected: todos passam (os 4 testes antigos continuam valendo com o padrão).

- [ ] **Step 5: Commit**

```bash
git add termos_html.py tests/test_termos_html.py
git commit -m "termos_html: corpo dos termos vem de textos.py; nome em negrito por marcador"
```

---

### Task 3: Geradores `.docx` leem os textos

**Files:**
- Modify: `Script_Termo_Individual.py`, `Termo_de_Responsabilidade.py`, `termo_devolucao.py`, `tests/test_docx.py`

**Interfaces:**
- Consumes: `textos.PADRAO`, `textos.paragrafos`, `textos.linhas`, `textos.com_nome`, `termos_html.data_por_extenso`.
- Produces: `criar_termo_responsabilidade(nome, bens, destino, textos=None)`, `gerar_termo_centro(ccustos, responsavel, bens, destino, textos=None)`, `gerar_termo_devolucao(nome, bens, destino, textos=None)`.

- [ ] **Step 1: Testes que falham**

Acrescentar a `tests/test_docx.py`:

```python
import textos


def texto(caminho):
    return "\n".join(p.text for p in Document(caminho).paragraphs)


def test_individual_com_textos_alterados(dados, tmp_path):
    from Script_Termo_Individual import criar_termo_responsabilidade
    t = dict(textos.PADRAO, individual_abertura="TESTE {nome} do {orgao_sigla}.", individual_compromissos="a\nb\nc",
             orgao_sigla="XYZ", assinatura_eletronica="Assinado via X")
    destino = criar_termo_responsabilidade("ANA SILVA", bens(), tmp_path / "t.docx", textos=t)
    doc = Document(destino)
    abertura = next(p for p in doc.paragraphs if p.text.startswith("TESTE"))
    assert [r.text for r in abertura.runs] == ["TESTE ", "ANA SILVA", " do XYZ."] and abertura.runs[1].bold
    corpo = texto(destino)
    assert "\na\n" in corpo and "\nc\n" in corpo and corpo.rstrip().endswith("Assinado via X")


def test_centro_com_paragrafos_e_assinatura_dos_textos(dados, tmp_path):
    from Termo_de_Responsabilidade import gerar_termo_centro
    t = dict(textos.PADRAO, ccusto_paragrafos="Primeiro {ccustos}.\n\nSegundo do {orgao_sigla}.",
             ccusto_assinatura="{responsavel}\nChefe", orgao_sigla="XYZ")
    resp = {"ccustos": "CCI", "responsavel": "JAQUELINE", "matricula": "46", "funcao": "coordenadora"}
    destino = gerar_termo_centro("CCI", resp, bens(), tmp_path / "c.docx", textos=t)
    corpo = texto(destino)
    assert "Primeiro CCI.\nSegundo do XYZ." in corpo
    assert Document(destino).paragraphs[-1].text == "JAQUELINE\nChefe"


def test_devolucao_com_recebedor_e_cidade_dos_textos(dados, tmp_path):
    from termo_devolucao import gerar_termo_devolucao
    t = dict(textos.PADRAO, cidade="Goiânia (GO)", recebedor_nome="FULANO", recebedor_cargo="Chefe")
    destino = gerar_termo_devolucao("ANA SILVA", bens(), tmp_path / "d.docx", textos=t)
    corpo = texto(destino)
    assert "Goiânia (GO), " in corpo and "\nFULANO\nChefe\n" in corpo
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_docx.py -v`
Expected: FAIL — `TypeError: ... 'textos'`

- [ ] **Step 3: `Script_Termo_Individual.py`**

Acrescentar `import textos as textos_mod` e mudar a assinatura para `def criar_termo_responsabilidade(nome, bens, destino, textos=None):` com `t = textos or textos_mod.PADRAO` e `campos = {"nome": nome, "orgao_sigla": t["orgao_sigla"]}` logo após `doc = Document(...)`. Trocar os trechos de texto (o restante do arquivo — tabela, larguras, fontes — fica igual):

```python
    run = p.add_run(t["individual_titulo"].format_map(campos))
```

```python
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    antes, nome_negrito, depois = textos_mod.com_nome(t["individual_abertura"], campos)
    p.add_run(antes)
    if nome_negrito is not None:
        p.add_run(nome_negrito).bold = True
        p.add_run(depois)
```

O bloco `novo_paragrafo = """..."""` e o laço `for paragraph in novo_paragrafo...` saem; no lugar (mantendo os parágrafos vazios entre itens, como o documento atual):

```python
    doc.add_paragraph("")
    p = doc.add_paragraph(t["individual_compromissos_intro"].format_map(campos))
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for item in textos_mod.linhas(t["individual_compromissos"]):
        doc.add_paragraph("")
        p = doc.add_paragraph(item.format_map(campos))
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
```

```python
    doc.add_paragraph(t["individual_ciencia"].format_map(campos))
```

```python
    p_assinado.add_run(t["assinatura_eletronica"])
```

- [ ] **Step 4: `Termo_de_Responsabilidade.py`**

Apagar `texto_padrao`. `import textos as textos_mod`. Assinatura `def gerar_termo_centro(ccustos, responsavel, bens, destino, textos=None):`; após `documento = Document(...)`:

```python
    t = textos or textos_mod.PADRAO
    campos = {k: responsavel.get(k) or "" for k in ("responsavel", "matricula", "funcao")}
    campos.update(ccustos=ccustos, orgao_sigla=t["orgao_sigla"])
    cabecalho_run = cabecalho.add_run(t["ccusto_titulo"].format_map(campos))
```

Laço dos parágrafos:

```python
    for paragraph in textos_mod.paragrafos(t["ccusto_paragrafos"]):
        paragrafo = documento.add_paragraph()
        paragrafo.paragraph_format.first_line_indent = Inches(0.59)
        paragrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragrafo.add_run(paragraph.format_map(campos))
```

Assinatura:

```python
    paragrafo_assinatura.add_run("\n".join(l.format_map(campos) for l in textos_mod.linhas(t["ccusto_assinatura"])))
```

- [ ] **Step 5: `termo_devolucao.py`**

`import textos as textos_mod`, `from termos_html import data_por_extenso`; remover `from datetime import datetime` se ficar sem uso (a data agora vem de `date.today()`: `from datetime import date`). Assinatura `def gerar_termo_devolucao(nome, bens, destino, textos=None):`; após `doc = Document(...)`:

```python
    t = textos or textos_mod.PADRAO
    campos = {"nome": nome, "orgao_sigla": t["orgao_sigla"], "cidade": t["cidade"], "data": data_por_extenso(date.today())}
```

Trocas: título `run = p.add_run(t["devolucao_titulo"].format_map(campos))`; introdução →

```python
    antes, nome_negrito, depois = textos_mod.com_nome(t["devolucao_abertura"], campos)
    p.add_run(antes).bold = False
    if nome_negrito is not None:
        p.add_run(nome_negrito).bold = True
        p.add_run(depois)
```

Data: apagar `hoje`, `meses`, `data_formatada`; `p_data = doc.add_paragraph(t["devolucao_data"].format_map(campos))`. Assinaturas: `p2 = doc.add_paragraph(t["assinatura_eletronica"])`, `p3 = doc.add_paragraph(t["devolucao_recebimento"].format_map(campos))`, `p4.add_run(t["recebedor_nome"]).bold = True`, `p5 = doc.add_paragraph(t["recebedor_cargo"])`, `p6 = doc.add_paragraph(t["assinatura_eletronica"])`.

- [ ] **Step 6: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam. `grep -n "CFC\|Brasília\|ANTÔNIO" Script_Termo_Individual.py Termo_de_Responsabilidade.py termo_devolucao.py termos_html.py` → nenhum resultado.

- [ ] **Step 7: Commit**

```bash
git add Script_Termo_Individual.py Termo_de_Responsabilidade.py termo_devolucao.py tests/test_docx.py
git commit -m "Geradores .docx leem os textos de textos.py; sem texto próprio"
```

---

### Task 4: Tela "Textos" e cabeçalho com o nome do órgão

**Files:**
- Modify: `app.py`, `templates/base.html` (nada — o menu vem de `MENU`), `tests/test_app.py`
- Create: `templates/textos.html`

**Interfaces:**
- Consumes: `textos.obter/salvar/restaurar/GRUPOS/TEXTAREA/ROTULOS/MARCADORES/PADRAO`.
- Produces: rotas `GET /textos` (`textos_tela`), `POST /textos` (`textos_salvar`). Rotas de termo passam `textos=textos.obter(conn)` aos geradores. `DSGOV["ORGAO"]` vem de `orgao_nome`.

- [ ] **Step 1: Testes que falham**

Acrescentar a `tests/test_app.py`:

```python
def test_textos_salvar_reflete_no_documento_e_restaurar(cliente):
    r = cliente.get("/textos")
    assert r.status_code == 200 and b"Compromissos" in r.data
    import textos
    dados_form = dict(textos.PADRAO, individual_abertura="TESTE {nome}.", orgao_nome="Órgão X")
    r = cliente.post("/textos", data=dados_form, follow_redirects=True)
    assert "Textos salvos".encode() in r.data and "Órgão X".encode() in r.data  # header usa orgao_nome
    doc = cliente.get("/termo/individual/ANA SILVA/documento").data.decode()
    assert "TESTE <b>ANA SILVA</b>." in doc
    r = cliente.post("/textos", data={"restaurar": "individual_abertura"}, follow_redirects=True)
    assert "Padrão restaurado".encode() in r.data
    doc = cliente.get("/termo/individual/ANA SILVA/documento").data.decode()
    assert "Pelo presente termo" in doc and "Órgão X".encode() in cliente.get("/").data


def test_textos_marcador_invalido_nao_grava_nada(cliente):
    import textos
    dados_form = dict(textos.PADRAO, individual_abertura="Eu {nomee}", cidade="Goiânia (GO)")
    r = cliente.post("/textos", data=dados_form, follow_redirects=True)
    assert b"nomee" in r.data
    doc = cliente.get("/termo/devolucao/ANA SILVA/documento").data.decode()
    assert "Goiânia" not in doc and "Brasília (DF)" in doc
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_app.py -v -k textos`
Expected: FAIL — 404 em `/textos`

- [ ] **Step 3: `app.py`**

`import textos` no topo. Substituir `DSGOV` e o `context_processor`:

```python
DSGOV_FIXO = {"SISTEMA": "Termos de Responsabilidade", "SUBTITULO": "Setor de Patrimônio"}


@app.context_processor
def contexto_dsgov():
    dsgov = dict(DSGOV_FIXO, ORGAO=textos.obter(obter_conn())["orgao_nome"])
    return {"DSGOV": dsgov, "MENU": [
        ("Início", "fa-home", url_for("home")),
        ("Termo por centro de custo", "fa-building", url_for("centro_custos")),
        ("Termo individual", "fa-user-check", url_for("termos_individuais")),
        ("Termo de devolução", "fa-box-open", url_for("termo_devolucao")),
        ("Cadastros", "fa-address-book", url_for("cadastros", aba="responsaveis")),
        ("Textos", "fa-file-signature", url_for("textos_tela")),
        ("Atualizar base", "fa-upload", url_for("upload")),
    ]}
```

Em `_bens_do_termo`, obter `t = textos.obter(conn)` no início e passar `textos=t` nas três chamadas `termos_html.corpo_*`; usar `t["ccusto_titulo"].format_map({...})`? Não — o título da página (breadcrumb) continua `f"Termo de Responsabilidade - {chave}"`; só o documento usa os textos. Em `termo_docx`, `t = textos.obter(conn)` e `textos=t` nas três chamadas dos geradores.

Rotas novas (antes de `# ---- cadastros`):

```python
# ---------------------------------------------------------------- textos do termo
@app.route("/textos")
def textos_tela():
    return render_template("textos.html", valores=textos.obter(obter_conn()), grupos=textos.GRUPOS,
                           textarea=textos.TEXTAREA, rotulos=textos.ROTULOS, marcadores=textos.MARCADORES,
                           padrao=textos.PADRAO, trilha=[("Textos", None)])


@app.route("/textos", methods=["POST"])
def textos_salvar():
    conn = obter_conn()
    chave = request.form.get("restaurar")
    if chave:
        textos.restaurar(conn, chave)
        flash(f"Padrão restaurado: {textos.ROTULOS.get(chave, chave)}.", "success")
        return redirect(url_for("textos_tela"))
    novos = {c: request.form.get(c, "").replace("\r\n", "\n") for c in textos.PADRAO}
    for c, v in novos.items():
        textos.validar(c, v)          # tudo validado antes de gravar qualquer coisa
    for c, v in novos.items():
        textos.salvar(conn, c, v)
    flash("Textos salvos.", "success")
    return redirect(url_for("textos_tela"))
```

(`textos.validar` levanta `ErroDeNegocio` → o `errorhandler` faz flash e redireciona ao referrer.)

- [ ] **Step 4: `templates/textos.html`**

```html
{% extends "base.html" %}
{% block titulo %}Textos{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Textos dos termos</h1></div>
<form method="post" action="{{ url_for('textos_salvar') }}" class="col-md-10">
  <p class="text-gray-70">Os marcadores entre chaves são substituídos ao gerar o termo. Linha em branco separa parágrafos; nos compromissos, uma linha por item. Vale para o documento colado no SEI e para o .docx.</p>
  {% for titulo, chaves in grupos %}
  <div class="br-card mb-4">
    <div class="card-header"><div class="text-weight-semi-bold text-up-01">{{ titulo }}</div></div>
    <div class="card-content">
      {% for chave in chaves %}
      <div class="mb-3">
        {% if chave in textarea %}
        <div class="br-textarea">
          <label for="t-{{ chave }}">{{ rotulos[chave] }}</label>
          <textarea id="t-{{ chave }}" name="{{ chave }}" rows="{{ 8 if chave in ('ccusto_paragrafos', 'individual_compromissos') else 3 }}">{{ valores[chave] }}</textarea>
        </div>
        {% else %}
        <div class="br-input">
          <label for="t-{{ chave }}">{{ rotulos[chave] }}</label>
          <input id="t-{{ chave }}" name="{{ chave }}" type="text" value="{{ valores[chave] }}"/>
        </div>
        {% endif %}
        <div class="d-flex align-items-center mt-1">
          <span class="text-down-01 text-gray-70">{% if marcadores[chave] %}Marcadores: {% for m in marcadores[chave]|sort %}{{ '{' ~ m ~ '}' }}{% if not loop.last %}, {% endif %}{% endfor %}{% else %}Sem marcadores.{% endif %}</span>
          {% if valores[chave] != padrao[chave] %}
          <button class="br-button small ml-auto" type="submit" name="restaurar" value="{{ chave }}"><i class="fas fa-undo mr-1" aria-hidden="true"></i>Restaurar padrão</button>
          {% endif %}
        </div>
      </div>
      {% endfor %}
    </div>
  </div>
  {% endfor %}
  <button class="br-button primary" type="submit"><i class="fas fa-save mr-1" aria-hidden="true"></i>Salvar</button>
</form>
{% endblock %}
```

- [ ] **Step 5: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam.

- [ ] **Step 6: Commit**

```bash
git add app.py templates/textos.html tests/test_app.py
git commit -m "Tela Textos: edição dos dizeres com marcadores e restaurar padrão"
```

---

### Task 5: `db.py` — edição de responsável, regra nova de exclusão, mover localizações, renomear pessoa

**Files:**
- Modify: `db.py`, `tests/test_db.py`

**Interfaces:**
- Produces: `atualizar_responsavel(conn, ccustos, dados: dict) -> None`; `checar_exclusao_centro(conn, ccustos) -> None` (levanta `CentroEmUso` se houver bens ativos sob guarda; `ErroDeNegocio` se o centro não existir); `excluir_responsavel(conn, ccustos)` (usa `checar_exclusao_centro`, apaga mapeamentos e centro); `mover_localizacoes(conn, localizacoes: list[str], ccustos) -> int`; `renomear_pessoa(conn, antigo, novo) -> str`.

- [ ] **Step 1: Testes que falham**

Acrescentar a `tests/test_db.py` e **ajustar** `test_excluir_centro_em_uso_falha`:

```python
def test_excluir_centro_em_uso_falha(dados):
    semear(dados)  # CCI tem o bem 1001 ativo em "01 - SALA CCI"
    with pytest.raises(db.CentroEmUso) as e:
        db.excluir_responsavel(dados, "CCI")
    assert "1 bem" in str(e.value)
    db.desatribuir(dados, "ANA SILVA", 1002)  # agora 1001 e 1002 respondem pelo setor
    with pytest.raises(db.CentroEmUso) as e:
        db.excluir_responsavel(dados, "CCI")
    assert "2 bens" in str(e.value)


def test_excluir_centro_sem_bens_apaga_mapeamentos(dados):
    semear(dados)
    db.incluir_responsavel(dados, {"ccustos": "VAZIO", "responsavel": "X"})
    db.incluir_localizacao(dados, "99 - SEM MAPA", "VAZIO")   # 1004 é ATIVO aqui...
    with pytest.raises(db.CentroEmUso):
        db.excluir_responsavel(dados, "VAZIO")
    dados.execute("UPDATE bens SET situacao='BAIXADO' WHERE numero=1004")
    dados.commit()
    db.excluir_responsavel(dados, "VAZIO")                       # ...sem bens ativos: some, e a sala volta a pendente
    assert db.responsavel(dados, "VAZIO") is None
    assert db.localizacoes_mapeadas(dados) == [{"localizacao": "01 - SALA CCI", "ccustos": "CCI"}]
    with pytest.raises(db.ErroDeNegocio):
        db.excluir_responsavel(dados, "NAO_EXISTE")


def test_atualizar_responsavel(dados):
    semear(dados)
    db.atualizar_responsavel(dados, "CCI", {"tratamento": "Prezado", "responsavel": " carlos ", "email": "c@cfc",
                                           "matricula": "99", "funcao": "gerente"})
    r = db.responsavel(dados, "CCI")
    assert (r["responsavel"], r["funcao"], r["matricula"]) == ("carlos", "gerente", "99")
    with pytest.raises(db.ErroDeNegocio):
        db.atualizar_responsavel(dados, "CCI", {"responsavel": ""})
    with pytest.raises(db.ErroDeNegocio):
        db.atualizar_responsavel(dados, "NAO_EXISTE", {"responsavel": "X"})


def test_mover_localizacoes(dados):
    semear(dados)
    db.incluir_responsavel(dados, {"ccustos": "PRES", "responsavel": "Y"})
    db.incluir_localizacao(dados, "99 - SEM MAPA", "CCI")
    n = db.mover_localizacoes(dados, ["01 - SALA CCI", "99 - SEM MAPA"], "PRES")
    assert n == 2 and {l["ccustos"] for l in db.localizacoes_mapeadas(dados)} == {"PRES"}
    assert [b["numero"] for b in db.bens_do_centro(dados, "PRES")] == [1001, 1004]
    with pytest.raises(db.ErroDeNegocio):
        db.mover_localizacoes(dados, [], "PRES")
    with pytest.raises(db.ErroDeNegocio):
        db.mover_localizacoes(dados, ["01 - SALA CCI"], "NAO_EXISTE")


def test_renomear_pessoa_mantem_atribuicoes(dados):
    semear(dados)
    assert db.renomear_pessoa(dados, "ANA SILVA", " ana  souza ") == "ANA SOUZA"
    assert db.pessoas(dados) == ["ANA SOUZA"] and db.pessoa_do_bem(dados, 1002) == "ANA SOUZA"
    db.incluir_pessoa(dados, "BRUNO")
    with pytest.raises(db.ErroDeNegocio):
        db.renomear_pessoa(dados, "ANA SOUZA", "bruno")
    with pytest.raises(db.ErroDeNegocio):
        db.renomear_pessoa(dados, "NINGUEM", "X")
    assert db.renomear_pessoa(dados, "ANA SOUZA", "ANA SOUZA") == "ANA SOUZA"
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_db.py -v -k "excluir_centro or atualizar or mover or renomear_pessoa"`
Expected: FAIL (`AttributeError` / mensagem antiga "localização(ões) mapeada(s)")

- [ ] **Step 3: Implementar em `db.py`**

Substituir `excluir_responsavel` e acrescentar as demais (junto das operações de cadastro):

```python
def checar_exclusao_centro(conn, ccustos: str) -> None:
    """Centro só pode ser excluído sem bens ativos sob guarda (bens em localizações dele, não atribuídos)."""
    if not responsavel(conn, ccustos):
        raise ErroDeNegocio(f"Centro de custo {ccustos} não encontrado.")
    n = len(bens_do_centro(conn, ccustos))
    if n:
        raise CentroEmUso(f"{ccustos} tem {n} bem{'ns' if n > 1 else ''} ativo{'s' if n > 1 else ''} sob guarda. "
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
```

Nota: `CentroEmUso` já existe em `db.py`; a mensagem mudou — `tests/test_app.py::test_cadastro_responsaveis_incluir_renomear_excluir` asserta `"Remapeie"`; esse teste é reescrito na Task 6 (deixe-o falhar agora só se falhar; rode `tests/test_db.py`).

- [ ] **Step 4: Rodar e ver passar**

Run: `.venv/bin/pytest tests/test_db.py -q`
Expected: todos passam (5 novos/alterados + os existentes).

- [ ] **Step 5: Commit**

```bash
git add db.py tests/test_db.py
git commit -m "db: editar responsável, exclusão só sem bens sob guarda, mover localizações, renomear pessoa"
```

---

### Task 6: Telas de cadastro — editar responsável, mover localizações, editar pessoa

**Files:**
- Modify: `app.py`, `templates/cadastros.html`, `templates/_macros.html`, `tests/test_app.py`
- Create: `templates/editar_responsavel.html`, `templates/editar_pessoa.html`

**Interfaces:**
- Consumes: Task 5.
- Produces: rotas `GET|POST /cadastros/responsaveis/<ccustos>/editar` (`responsaveis_editar`), `POST /cadastros/responsaveis/excluir` (agora com confirmação em dois POSTs; `?excluir=<sigla>` na volta), `POST /cadastros/localizacoes/mover` (`localizacoes_mover`), `GET|POST /cadastros/pessoas/<nome>/editar` (`pessoas_editar`). Rota `responsaveis_renomear` **removida**. Macros `cabecalho_tabela(titulo, id, selecao=False)`, `th_selecao(id)`, `td_selecao(id, indice, name, value)`.

- [ ] **Step 1: Testes que falham**

Em `tests/test_app.py`, **substituir** `test_cadastro_responsaveis_incluir_renomear_excluir` por:

```python
def test_cadastro_responsaveis_editar_renomear_excluir(cliente):
    r = cliente.post("/cadastros/responsaveis/incluir", data={"ccustos": "decom", "tratamento": "Prezado",
                     "responsavel": "THIAGO", "email": "", "matricula": "481", "funcao": "gerente"}, follow_redirects=True)
    assert b"DECOM" in r.data
    r = cliente.get("/cadastros/responsaveis/CCI/editar")
    assert b"JAQUELINE PORTELA" in r.data and b"01 - SALA CCI" in r.data
    r = cliente.post("/cadastros/responsaveis/CCI/editar", data={"ccustos": "geserv", "tratamento": "Prezado",
                     "responsavel": "CARLOS", "email": "", "matricula": "7", "funcao": "gerente"}, follow_redirects=True)
    assert b"GESERV" in r.data and b"CARLOS" in r.data and b">CCI<" not in r.data
    assert cliente.get("/cadastros/responsaveis/CCI/editar").status_code == 404
    # excluir: pede confirmação; com bens sob guarda, bloqueia
    r = cliente.post("/cadastros/responsaveis/excluir", data={"ccustos": "GESERV"}, follow_redirects=True)
    assert b"sob guarda" in r.data
    r = cliente.post("/cadastros/responsaveis/excluir", data={"ccustos": "DECOM"}, follow_redirects=True)
    assert b"Confirmar exclus" in r.data
    r = cliente.post("/cadastros/responsaveis/excluir", data={"ccustos": "DECOM", "confirmar": "1"}, follow_redirects=True)
    assert b">DECOM<" not in r.data


def test_cadastro_localizacoes_mover(cliente):
    cliente.post("/cadastros/responsaveis/incluir", data={"ccustos": "PRES", "responsavel": "Y"})
    cliente.post("/cadastros/localizacoes/incluir", data={"localizacao": "99 - SEM MAPA", "ccustos": "CCI"})
    r = cliente.post("/cadastros/localizacoes/mover", data={"localizacoes": ["01 - SALA CCI", "99 - SEM MAPA"], "ccustos_destino": "PRES"},
                     follow_redirects=True)
    assert b"2 localiza" in r.data
    assert b"PRES" in cliente.get("/bem?numero=1001").data
    r = cliente.post("/cadastros/localizacoes/mover", data={"ccustos_destino": "PRES"}, follow_redirects=True)
    assert b"ao menos uma" in r.data


def test_cadastro_pessoas_editar_nome(cliente):
    r = cliente.get("/cadastros/pessoas/ANA SILVA/editar")
    assert b"ANA SILVA" in r.data
    r = cliente.post("/cadastros/pessoas/ANA SILVA/editar", data={"nome": "ana souza"}, follow_redirects=True)
    assert b"ANA SOUZA" in r.data and b"NOTEBOOK" in r.data
    assert cliente.get("/cadastros/pessoas/NINGUEM/editar").status_code == 404
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_app.py -v -k "editar or mover"`
Expected: FAIL (404)

- [ ] **Step 3: Macros** (`templates/_macros.html`)

Alterar a abertura de `cabecalho_tabela` e acrescentar duas macros:

```html
{% macro cabecalho_tabela(titulo, id, selecao=False) %}
{# abre um br-table com busca (e seleção de linhas, se selecao); quem chama fecha com </table></div> #}
<div class="br-table" data-search="data-search"{% if selecao %} data-selection="data-selection"{% endif %}>
  <div class="table-header">
    ... (top-bar e search-bar como estão) ...
    {% if selecao %}
    <div class="selected-bar">
      <div class="info"><span class="count">0</span><span class="text">item selecionado</span></div>
    </div>
    {% endif %}
  </div>
  <table>
    <caption class="sr-only">{{ titulo }}</caption>
{% endmacro %}

{% macro th_selecao(id) %}
<th class="column-checkbox" scope="col"><div class="br-checkbox hidden-label"><input id="check-all-{{ id }}" name="check-all-{{ id }}" type="checkbox" aria-label="Selecionar tudo" data-parent="check-{{ id }}"/><label for="check-all-{{ id }}">Selecionar todas as linhas</label></div></th>
{% endmacro %}

{% macro td_selecao(id, indice, name, value) %}
<td class="column-checkbox"><div class="br-checkbox hidden-label"><input id="check-{{ id }}-{{ indice }}" name="{{ name }}" type="checkbox" value="{{ value }}" data-child="check-{{ id }}"/><label for="check-{{ id }}-{{ indice }}">Selecionar linha</label></div></td>
{% endmacro %}
```

- [ ] **Step 4: Rotas em `app.py`**

Remover `responsaveis_renomear`. Substituir `responsaveis_excluir` e acrescentar:

```python
@app.route("/cadastros/responsaveis/<ccustos>/editar", methods=["GET", "POST"])
def responsaveis_editar(ccustos):
    conn = obter_conn()
    c = db.responsavel(conn, ccustos) or abort(404)
    if request.method == "POST":
        nova = " ".join(request.form.get("ccustos", "").split()).upper()
        if nova and nova != ccustos:
            db.renomear_centro(conn, ccustos, nova)
            ccustos = nova
        db.atualizar_responsavel(conn, ccustos, request.form)
        flash(f"Centro de custo {ccustos} atualizado.", "success")
        return _volta("responsaveis")
    locais = [l["localizacao"] for l in db.localizacoes_mapeadas(conn) if l["ccustos"] == ccustos]
    return render_template("editar_responsavel.html", c=c, locais=locais,
                           trilha=[("Cadastros", url_for("cadastros", aba="responsaveis")), (f"Editar {ccustos}", None)])


@app.route("/cadastros/responsaveis/excluir", methods=["POST"])
def responsaveis_excluir():
    conn, sigla = obter_conn(), request.form["ccustos"]
    db.checar_exclusao_centro(conn, sigla)   # bloqueia já aqui se houver bens sob guarda
    if not request.form.get("confirmar"):
        n = sum(1 for l in db.localizacoes_mapeadas(conn) if l["ccustos"] == sigla)
        flash(f"Excluir {sigla} remove também o mapeamento de {n} localização(ões), que voltam a pendentes. "
              "Clique em confirmar para prosseguir.", "warning")
        return _volta("responsaveis", excluir=sigla)
    db.excluir_responsavel(conn, sigla)
    flash(f"Centro de custo {sigla} excluído.", "success")
    return _volta("responsaveis")


@app.route("/cadastros/localizacoes/mover", methods=["POST"])
def localizacoes_mover():
    destino = request.form.get("ccustos_destino", "")
    n = db.mover_localizacoes(obter_conn(), request.form.getlist("localizacoes"), destino)
    flash(f"{n} localização(ões) movida(s) para {destino}.", "success")
    return _volta("localizacoes")


@app.route("/cadastros/pessoas/<nome>/editar", methods=["GET", "POST"])
def pessoas_editar(nome):
    conn = obter_conn()
    if nome not in db.pessoas(conn):
        abort(404)
    if request.method == "POST":
        novo = db.renomear_pessoa(conn, nome, request.form.get("nome", ""))
        flash(f"Pessoa renomeada para {novo}.", "success")
        return _volta("pessoas", nome=novo)
    return render_template("editar_pessoa.html", nome=nome,
                           trilha=[("Cadastros", url_for("cadastros", aba="pessoas")), (f"Editar {nome}", None)])
```

Em `cadastros()`, passar também `excluir=request.args.get("excluir")` ao template.

- [ ] **Step 5: Templates**

`templates/editar_responsavel.html`:

```html
{% extends "base.html" %}
{% block titulo %}Editar {{ c.ccustos }}{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Editar centro de custo {{ c.ccustos }}</h1></div>
<form method="post" class="col-md-8">
  {% for campo, rotulo in [('ccustos', 'Sigla do centro de custo (mudar a sigla renomeia e mantém as localizações)'), ('responsavel', 'Responsável'), ('tratamento', 'Tratamento (Prezado/Prezada)'), ('funcao', 'Função'), ('matricula', 'Matrícula'), ('email', 'E-mail')] %}
  <div class="br-input mb-3"><label for="e-{{ campo }}">{{ rotulo }}</label><input id="e-{{ campo }}" name="{{ campo }}" type="text" value="{{ c[campo] or '' }}"{% if campo in ('ccustos', 'responsavel') %} required{% endif %}/></div>
  {% endfor %}
  <div class="mb-4">
    <p class="mb-1 text-weight-semi-bold">Localizações deste centro ({{ locais|length }})</p>
    {% if locais %}<p class="text-gray-70">{{ locais|join(' · ') }}</p>{% else %}<p class="text-gray-70">Nenhuma.</p>{% endif %}
    <a href="{{ url_for('cadastros', aba='localizacoes') }}">Mover localizações</a>
  </div>
  <div class="dsgov-acoes-formulario">
    <a class="br-button" href="{{ url_for('cadastros', aba='responsaveis') }}">Cancelar</a>
    <button class="br-button primary" type="submit"><i class="fas fa-save mr-1" aria-hidden="true"></i>Salvar</button>
  </div>
</form>
{% endblock %}
```

`templates/editar_pessoa.html`:

```html
{% extends "base.html" %}
{% block titulo %}Editar {{ nome }}{% endblock %}
{% block conteudo %}
<div class="d-flex align-items-center mb-4"><h1 class="mb-0">Editar pessoa</h1></div>
<form method="post" class="col-md-8">
  <div class="br-input mb-3"><label for="e-nome">Nome (os bens atribuídos acompanham)</label><input id="e-nome" name="nome" type="text" value="{{ nome }}" required/></div>
  <div class="dsgov-acoes-formulario">
    <a class="br-button" href="{{ url_for('cadastros', aba='pessoas', nome=nome) }}">Cancelar</a>
    <button class="br-button primary" type="submit"><i class="fas fa-save mr-1" aria-hidden="true"></i>Salvar</button>
  </div>
</form>
{% endblock %}
```

`templates/cadastros.html` — três mudanças:

1. Importar as macros novas: `{% from "_macros.html" import select, cabecalho_tabela, th_selecao, td_selecao %}`.
2. Aba Responsáveis, célula de ações (substitui o form de renomear):

```html
          <td class="dsgov-acoes">
            <a class="br-button circle small" href="{{ url_for('responsaveis_editar', ccustos=c.ccustos) }}" aria-label="Editar {{ c.ccustos }}"><i class="fas fa-pen" aria-hidden="true"></i></a>
            <form method="post" action="{{ url_for('responsaveis_excluir') }}" class="d-inline">
              <input type="hidden" name="ccustos" value="{{ c.ccustos }}"/>
              {% if excluir == c.ccustos %}<input type="hidden" name="confirmar" value="1"/>
              <button class="br-button secondary small" type="submit"><i class="fas fa-check mr-1" aria-hidden="true"></i>Confirmar exclusão</button>
              {% else %}<button class="br-button circle small" type="submit" aria-label="Excluir {{ c.ccustos }}"><i class="fas fa-trash" aria-hidden="true"></i></button>{% endif %}
            </form>
          </td>
```

3. Aba Localizações: envolver a tabela num form de mover, com seleção:

```html
      <form method="post" action="{{ url_for('localizacoes_mover') }}">
      {{ cabecalho_tabela('Localizações mapeadas', 'loc', selecao=True) }}
        <thead><tr>{{ th_selecao('loc') }}<th scope="col">Localização</th><th scope="col">Centro de custo</th><th scope="col" class="dsgov-acoes">Ações</th></tr></thead>
        <tbody>
        {% for l in mapeadas %}
        <tr>{{ td_selecao('loc', loop.index, 'localizacoes', l.localizacao) }}<td>{{ l.localizacao }}</td><td>{{ l.ccustos }}</td>
          <td class="dsgov-acoes"><button class="br-button circle small" type="submit" form="excluir-loc-{{ loop.index }}" aria-label="Remover mapeamento de {{ l.localizacao }}"><i class="fas fa-trash" aria-hidden="true"></i></button></td></tr>
        {% endfor %}
        </tbody>
      </table></div>
      <div class="row mt-3">
        <div class="col-md-5">{{ select('ccustos_destino', 'Mover selecionadas para', centros|map(attribute='ccustos')|list) }}</div>
        <div class="col-md-3 d-flex align-items-end"><button class="br-button secondary" type="submit"><i class="fas fa-exchange-alt mr-1" aria-hidden="true"></i>Mover</button></div>
      </div>
      </form>
      {% for l in mapeadas %}
      <form method="post" action="{{ url_for('localizacoes_excluir') }}" id="excluir-loc-{{ loop.index }}"><input type="hidden" name="localizacao" value="{{ l.localizacao }}"/></form>
      {% endfor %}
```

(O botão de excluir usa `form="..."` para não ficar aninhado no form de mover — HTML não permite form dentro de form. O `select` de mover chama-se `ccustos_destino` porque o form "Mapear" acima já usa `ccustos` e os ids gerados pela macro colidiriam.)

4. Aba Pessoas, ao lado do `h2`: `<a class="br-button circle small ml-2" href="{{ url_for('pessoas_editar', nome=nome) }}" aria-label="Editar nome de {{ nome }}"><i class="fas fa-pen" aria-hidden="true"></i></a>`.

- [ ] **Step 6: Rodar e ver passar**

Run: `.venv/bin/pytest -q`
Expected: todos passam.

- [ ] **Step 7: Verificação manual**

`TERMOS_PORTA=5002 .venv/bin/python main.py` (cai no navegador nesta VPS; use `curl` ou abra por túnel): na aba Localizações, marcar duas linhas, escolher destino, Mover → flash com "2 localização(ões)". Editar um centro mudando a sigla → linhas da aba Localizações mostram a sigla nova. Parar com Ctrl+C.

- [ ] **Step 8: Commit**

```bash
git add app.py templates tests/test_app.py
git commit -m "Cadastros: editar responsável e pessoa, mover localizações, exclusão com confirmação"
```

---

### Task 7: `db.py` — exportar e importar cadastros

**Files:**
- Modify: `db.py`, `tests/test_db.py`

**Interfaces:**
- Produces: `exportar_cadastros(conn, destino: Path) -> Path`; `importar_cadastros(conn, arquivo) -> dict` com `responsaveis`, `localizacoes`, `pessoas`, `atribuicoes` (contagens) e `sem_centro`; levanta `ImportacaoInvalida` com a lista de problemas (até 20) sem alterar nada.

- [ ] **Step 1: Testes que falham**

Acrescentar a `tests/test_db.py`:

```python
from openpyxl import load_workbook


def test_exportar_cadastros_quatro_abas(dados, tmp_path):
    semear(dados)
    arq = db.exportar_cadastros(dados, tmp_path / "c.xlsx")
    wb = load_workbook(arq)
    assert wb.sheetnames == ["responsaveis", "localizacoes", "pessoas", "atribuicoes"]
    assert [c.value for c in wb["responsaveis"][1]] == ["ccustos", "tratamento", "responsavel", "email", "matricula", "funcao"]
    assert [c.value for c in wb["atribuicoes"][2]] == ["ANA SILVA", 1002]
    assert wb["localizacoes"].max_row == 2 and wb["pessoas"].max_row == 2


def test_importar_a_propria_exportacao_e_idempotente(dados, tmp_path):
    semear(dados)
    arq = db.exportar_cadastros(dados, tmp_path / "c.xlsx")
    resumo = db.importar_cadastros(dados, arq)
    assert resumo == {"responsaveis": 1, "localizacoes": 1, "pessoas": 1, "atribuicoes": 1, "sem_centro": ["99 - SEM MAPA"]}
    assert db.ficha_do_bem(dados, 1002)["pessoa"] == "ANA SILVA"


def cadastros_xlsx(tmp_path, **abas):
    wb = Workbook()
    wb.remove(wb.active)
    cabecalhos = {"responsaveis": ["ccustos", "tratamento", "responsavel", "email", "matricula", "funcao"],
                  "localizacoes": ["localizacao", "ccustos"], "pessoas": ["nome"], "atribuicoes": ["nome", "numero"]}
    for aba, cab in cabecalhos.items():
        ws = wb.create_sheet(aba)
        ws.append(cab)
        for linha in abas.get(aba, []):
            ws.append(linha)
    caminho = tmp_path / "cad.xlsx"
    wb.save(caminho)
    return caminho


def test_importar_cadastros_substitui_e_normaliza(dados, tmp_path):
    semear(dados)
    arq = cadastros_xlsx(tmp_path,
                         responsaveis=[["geserv ", "Prezado", " carlos", "c@cfc", 7, "gerente"], ["pres", "", "MARIA", "", "", ""]],
                         localizacoes=[["01 - SALA CCI", "GESERV"], ["99 - SEM MAPA", "pres"]],
                         pessoas=[[" bruno lima "]], atribuicoes=[["bruno lima", "1001"]])
    resumo = db.importar_cadastros(dados, arq)
    assert resumo["responsaveis"] == 2 and resumo["atribuicoes"] == 1 and resumo["sem_centro"] == []
    assert db.responsavel(dados, "CCI") is None and db.responsavel(dados, "GESERV")["matricula"] == "7"
    f = db.ficha_do_bem(dados, 1001)
    assert f["ccustos"] == "GESERV" and f["pessoa"] == "BRUNO LIMA"
    assert db.pessoas(dados) == ["BRUNO LIMA"]


@pytest.mark.parametrize("abas, trecho", [
    ({"localizacoes": [["01 - SALA CCI", "NAOEXISTE"]]}, "NAOEXISTE"),
    ({"atribuicoes": [["ANA SILVA", 1001]]}, "pessoas"),                     # pessoa fora da aba pessoas
    ({"pessoas": [["ANA"]], "atribuicoes": [["ANA", 9999]]}, "9999"),
    ({"pessoas": [["ANA"], ["BRUNO"]], "atribuicoes": [["ANA", 1001], ["BRUNO", 1001]]}, "repetido"),
    ({"responsaveis": [["", "", "X", "", "", ""]]}, "sigla"),
    ({"responsaveis": [["A", "", "", "", "", ""]]}, "responsável"),
    ({"responsaveis": [["A", "", "X", "", "", ""], ["a", "", "Y", "", "", ""]]}, "repetid"),
])
def test_importar_cadastros_invalidos_nao_alteram_nada(dados, tmp_path, abas, trecho):
    semear(dados)
    with pytest.raises(db.ImportacaoInvalida) as e:
        db.importar_cadastros(dados, cadastros_xlsx(tmp_path, **abas))
    assert trecho in str(e.value)
    assert db.centros(dados)[0]["ccustos"] == "CCI" and db.pessoas(dados) == ["ANA SILVA"]


def test_importar_cadastros_sem_aba_ou_coluna(dados, tmp_path):
    semear(dados)
    wb = Workbook()
    wb.active.title = "responsaveis"
    wb.active.append(["ccustos", "responsavel"])
    arq = tmp_path / "x.xlsx"
    wb.save(arq)
    with pytest.raises(db.ImportacaoInvalida) as e:
        db.importar_cadastros(dados, arq)
    assert "localizacoes" in str(e.value) or "coluna" in str(e.value)
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_db.py -v -k cadastros`
Expected: FAIL — `AttributeError: exportar_cadastros`

- [ ] **Step 3: Implementar em `db.py`** (ao final do arquivo; `from openpyxl import Workbook, load_workbook` no topo)

```python
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
    brutos = {t: _ler_aba_cadastro(wb, t, problemas) for t in CADASTROS}
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
        elif num is None:
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
```

- [ ] **Step 4: Rodar e ver passar**

Run: `.venv/bin/pytest tests/test_db.py -q`
Expected: todos passam (12 novos + os existentes).

- [ ] **Step 5: Commit**

```bash
git add db.py tests/test_db.py
git commit -m "db: exportar e importar cadastros em planilha (tudo ou nada, com validação)"
```

---

### Task 8: Rotas de exportar/importar cadastros e README

**Files:**
- Modify: `app.py`, `templates/cadastros.html`, `templates/upload.html`, `tests/test_app.py`, `README.md`

**Interfaces:**
- Consumes: Task 7.
- Produces: `GET /cadastros/exportar` (`cadastros_exportar`), `POST /importar-cadastros` (`importar_cadastros`).

- [ ] **Step 1: Testes que falham**

Acrescentar a `tests/test_app.py`:

```python
def test_exportar_e_importar_cadastros(cliente):
    r = cliente.get("/cadastros/exportar")
    assert r.status_code == 200 and r.headers["Content-Disposition"].endswith("cadastros.xlsx")
    r = cliente.post("/importar-cadastros", data={"arquivo": (io.BytesIO(r.data), "cadastros.xlsx")},
                     content_type="multipart/form-data", follow_redirects=True)
    assert b"1 centro" in r.data and b"1 pessoa" in r.data
    r = cliente.post("/importar-cadastros", data={"arquivo": (io.BytesIO(b"nada"), "x.xlsx")},
                     content_type="multipart/form-data", follow_redirects=True)
    assert "inválido".encode() in r.data
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `.venv/bin/pytest tests/test_app.py -v -k exportar`
Expected: FAIL (404)

- [ ] **Step 3: Rotas**

```python
@app.route("/cadastros/exportar")
def cadastros_exportar():
    destino = db.exportar_cadastros(obter_conn(), config.pasta_saida() / "cadastros.xlsx")
    return send_file(destino, as_attachment=True, download_name="cadastros.xlsx")


@app.route("/importar-cadastros", methods=["POST"])
def importar_cadastros():
    arquivo = request.files.get("arquivo")
    if not arquivo or not arquivo.filename.lower().endswith(".xlsx"):
        flash("Envie a planilha de cadastros em .xlsx.", "error")
        return redirect(url_for("upload"))
    r = db.importar_cadastros(obter_conn(), arquivo.stream)
    flash(f"Cadastros importados: {r['responsaveis']} centro(s) de custo, {r['localizacoes']} localização(ões), "
          f"{r['pessoas']} pessoa(s), {r['atribuicoes']} atribuição(ões).", "success")
    return redirect(url_for("upload"))
```

- [ ] **Step 4: Templates**

`cadastros.html`, cabeçalho da página:

```html
<div class="d-flex align-items-center mb-4">
  <h1 class="mb-0">Cadastros</h1>
  <div class="ml-auto"><a class="br-button" href="{{ url_for('cadastros_exportar') }}"><i class="fas fa-file-excel mr-1" aria-hidden="true"></i>Exportar cadastros</a></div>
</div>
```

`upload.html`, após o primeiro formulário:

```html
<form method="post" action="{{ url_for('importar_cadastros') }}" enctype="multipart/form-data" class="col-md-8 mb-4">
  <p class="text-gray-70">Importar cadastros: a planilha exportada em <a href="{{ url_for('cadastros_exportar') }}">Cadastros → Exportar</a>, editada. Substitui responsáveis, localizações, pessoas e atribuições inteiros; o que não estiver na planilha some. Bens não mudam.</p>
  <div class="br-upload mb-3">
    <label class="upload-label" for="arquivo-cadastros"><span>Planilha de cadastros .xlsx</span></label>
    <input class="upload-input" id="arquivo-cadastros" name="arquivo" type="file" accept=".xlsx" required/>
    <div class="upload-list"></div>
  </div>
  <button class="br-button secondary" type="submit"><i class="fas fa-upload mr-1" aria-hidden="true"></i>Importar cadastros</button>
</form>
```

- [ ] **Step 5: README** — na seção "Uso", acrescentar dois itens:

```
5. **Textos**: os dizeres dos termos (abertura, compromissos, parágrafos, quem recebe a devolução,
   cidade, sigla do órgão) são editáveis no menu Textos, com marcadores como `{nome}` e `{ccustos}`;
   "Restaurar padrão" volta ao texto original.
6. **Planilha de cadastros**: Cadastros → *Exportar cadastros* gera `cadastros.xlsx` (4 abas). Edite no
   Excel e importe em *Atualizar base → Importar cadastros* — substitui as 4 tabelas inteiras.
```

E em "Arquivos": `| textos.py | textos padrão dos termos e marcadores |`.

- [ ] **Step 6: Rodar tudo**

Run: `.venv/bin/pytest -q`
Expected: todos passam.

- [ ] **Step 7: Commit**

```bash
git add app.py templates/cadastros.html templates/upload.html tests/test_app.py README.md
git commit -m "Exportar e importar cadastros em planilha; README"
```

---

## Self-review

**Spec coverage**
- §3.1 tabela/API → Task 1. §3.2 chaves/valores/marcadores → Task 1 (`PADRAO`, `MARCADORES`). §3.3 renderização (negrito, multilinha, geradores sem texto) → Tasks 2–3. §3.4 tela Textos → Task 4.
- §4.1 editar/excluir responsável → Tasks 5–6. §4.2 mover localizações com seleção → Tasks 5–6 (macros). §4.3 renomear pessoa → Tasks 5–6.
- §5.1–5.3 exportar/importar → Tasks 7–8. §6 estrutura → coincide. §7 testes → distribuídos.

**Placeholders**: nenhum; os trechos "(como estão)" nas macros referem-se a código existente que não muda.

**Consistência de nomes**: `textos.obter/validar/salvar/restaurar/paragrafos/linhas/com_nome/PADRAO/MARCADORES/GRUPOS/TEXTAREA/ROTULOS` (T1) usados em T2–T4; `data_por_extenso` (T2) usado em T3; `db.atualizar_responsavel/checar_exclusao_centro/excluir_responsavel/mover_localizacoes/renomear_pessoa` (T5) usados em T6; `db.exportar_cadastros/importar_cadastros/CADASTROS` (T7) usados em T8; macros `cabecalho_tabela(selecao=)`, `th_selecao`, `td_selecao` (T6). O `select` de mover usa `ccustos_destino` (rota e teste, ver nota na T6).
