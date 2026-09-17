"""Cards ECharts do painel do evento de inventário."""


def _dados():
    return {"resumo": {},
            "situacao": [{"chave": "localizado", "rotulo": "Localizado", "quantidade": 2},
                         {"chave": "divergente", "rotulo": "Divergente", "quantidade": 1},
                         {"chave": "pendente", "rotulo": "Não localizado", "quantidade": 3}],
            "integrantes": [{"chave": f"I{i}", "rotulo": f"I{i}", "quantidade": 10 - i} for i in range(7)],
            "conservacao": [{"chave": "Bom", "rotulo": "Bom", "quantidade": 3}, {"chave": "-", "rotulo": "Não informada", "quantidade": 1}],
            "andares": [{"andar": "02", "total": 4, "localizados": 1, "pendentes": 3, "divergentes": 0, "salas": 2},
                        {"andar": "Sem andar", "total": 1, "localizados": 0, "pendentes": 1, "divergentes": 0, "salas": 1}],
            "salas_do_andar": []}


def test_cards_tipos_urls_e_tabelas(dados):
    from app import app
    import painel_inventario
    with app.test_request_context():
        c = {x["id"]: x for x in painel_inventario.cards(_dados(), 7)}
    assert list(c) == ["g-situacao", "g-integrantes", "g-conservacao", "g-andares"]
    s = c["g-situacao"]["opcoes"]["series"][0]
    assert s["type"] == "pie" and s["data"][0] == {"name": "Localizado", "value": 2, "itemStyle": {"color": "#168821"}, "url": "/inventario/7/relatorio?situacao=localizado"}
    assert c["g-situacao"]["opcoes"]["graphic"][0]["style"]["text"] == "6\nbens"
    i = c["g-integrantes"]["opcoes"]
    assert i["yAxis"]["type"] == "category" and i["series"][0]["data"][0] == {"value": 10, "url": "/inventario/7/relatorio?integrante=I0"}   # 7 itens → barras
    assert c["g-integrantes"]["tabela"]["colunas"] == ["Integrante", "Leituras"] and c["g-integrantes"]["tabela"]["linhas"][0][0]["url"] == "/inventario/7/relatorio?integrante=I0"
    k = c["g-conservacao"]["opcoes"]["series"][0]
    assert k["type"] == "pie" and k["data"][1]["url"] == "/inventario/7/relatorio?conservacao=-"
    a = c["g-andares"]["opcoes"]
    assert a["xAxis"]["data"] == ["02", "Sem andar"] and [x["name"] for x in a["series"]] == ["Localizados", "Pendentes"]
    assert a["series"][0]["stack"] == "total" and a["series"][0]["itemStyle"] == {"color": "#168821"}
    assert a["series"][1]["data"][0] == {"value": 3, "url": "/inventario/7/painel?andar=02"}
    assert c["g-andares"]["tabela"]["colunas"] == ["Andar", "Bens", "Localizados", "Divergentes", "Pendentes"]
    assert c["g-andares"]["tabela"]["linhas"][0] == [{"valor": "02", "url": "/inventario/7/painel?andar=02"}, {"valor": 4}, {"valor": 1}, {"valor": 0}, {"valor": 3}]
    assert c["g-andares"]["col"] is None and c["g-integrantes"]["altura"] is None


def test_card_salas_do_andar(dados):
    from app import app
    import painel_inventario
    d = _dados()
    d["salas_do_andar"] = [{"localizacao": f"02 - SALA {i}", "total": 1, "localizados": 0, "pendentes": 1, "divergentes": 0} for i in range(11)]
    with app.test_request_context():
        c = {x["id"]: x for x in painel_inventario.cards(d, 7, "02")}
        assert "g-salas" not in {x["id"] for x in painel_inventario.cards(d, 7)}       # sem andar escolhido, sem card
    s = c["g-salas"]
    assert s["titulo"] == "Salas do andar 02" and s["col"] == "col-12" and s["opcoes"]["xAxis"]["axisLabel"]["rotate"] == 45
    assert s["opcoes"]["xAxis"]["data"][0] == "SALA 0"
    assert s["opcoes"]["series"][1]["data"][0]["url"] == "/inventario/7/sala/02%20-%20SALA%200"
    assert s["tabela"]["linhas"][0][0] == {"valor": "02 - SALA 0", "url": "/inventario/7/sala/02%20-%20SALA%200"}
    d["salas_do_andar"] = [{"localizacao": "TERMOS INDIVIDUAIS", "total": 1, "localizados": 0, "pendentes": 1, "divergentes": 0}]
    with app.test_request_context():
        s = next(x for x in painel_inventario.cards(d, 7, "Sem andar") if x["id"] == "g-salas")
    assert s["opcoes"]["xAxis"]["data"] == ["TERMOS INDIVIDUAIS"] and s["col"] is None and "rotate" not in str(s["opcoes"]["xAxis"])
