"""Aparência por usuário na casca Tabler: claro/escuro, cor principal, tom dos cinzas e cantos, gravados na conta."""

import json

import pytest

import usuarios
from tests.test_tabler import tabler  # noqa: F401 — fixture


def _csrf(cliente):
    with cliente.session_transaction() as s:
        return s.setdefault("csrf", "token-teste")


def test_padrao_quando_nunca_escolheu():
    assert usuarios.aparencia({"aparencia": None}) == usuarios.APARENCIA_PADRAO
    assert usuarios.aparencia({"aparencia": "{estragado"}) == usuarios.APARENCIA_PADRAO
    assert usuarios.aparencia(None) == usuarios.APARENCIA_PADRAO


def test_grava_na_conta_e_responde_json(cliente, dados):
    resposta = cliente.post("/aparencia", data={"cor": "green", "base": "stone", "cantos": "2", "tema": "escuro",
                                                "csrf": _csrf(cliente)}, headers={"Accept": "application/json"})
    assert resposta.status_code == 200
    assert resposta.get_json() == {"tema": "escuro", "cor": "green", "base": "stone", "cantos": "2"}
    linha = dados.execute("SELECT aparencia FROM usuarios WHERE aparencia IS NOT NULL").fetchone()
    assert json.loads(linha[0])["cor"] == "green"


def test_valor_fora_da_lista_nao_grava(cliente):
    resposta = cliente.post("/aparencia", data={"cor": "#ff0000", "csrf": _csrf(cliente)},
                            headers={"Accept": "application/json"})
    assert resposta.status_code == 400


def test_envio_sem_js_volta_para_a_pagina(cliente):
    resposta = cliente.post("/aparencia", data={"cor": "red", "csrf": _csrf(cliente)},
                            headers={"Referer": "http://localhost/pesquisa"})
    assert resposta.status_code == 302 and resposta.headers["Location"].endswith("/pesquisa")
    resposta = cliente.post("/aparencia", data={"cor": "red", "csrf": _csrf(cliente)},
                            headers={"Referer": "https://outro-site.com/x"})
    assert "outro-site" not in resposta.headers["Location"]


def test_pagina_abre_com_a_aparencia_da_conta(tabler):  # noqa: F811
    cliente = tabler
    cliente.post("/aparencia", data={"cor": "purple", "tema": "escuro", "csrf": _csrf(cliente)},
                 headers={"Accept": "application/json"})
    html = cliente.get("/").get_data(as_text=True)
    assert '"cor": "purple"' in html and '"tema": "escuro"' in html
    assert "tabler-themes.min.css" in html and 'id="painel-aparencia"' in html
    assert "style=" not in html.split('id="painel-aparencia"')[1].split("</form>")[0]
