import itertools
import pytest
import permissoes


def test_funcoes_somam_sem_conceder_outras():
    f = {"inventariante", "consulta"}
    assert permissoes.permitido(f, "bem")
    assert permissoes.permitido(f, "inventario.ler", "POST")
    assert not permissoes.permitido(f, "inventario.xlsx")
    assert not permissoes.permitido(f, "termo_docx")
    assert not permissoes.permitido({"operador"}, "inventario.ler", "POST")
    assert permissoes.permitido({"consulta_inventarios"}, "inventario.xlsx")
    assert not permissoes.permitido({"consulta_inventarios"}, "bem")
    assert not permissoes.permitido({"admin"}, "rota_inexistente")
    assert not permissoes.permitido({"consulta"}, "termo_devolucao", "POST")
    assert permissoes.permitido({"admin"}, "termo_docx", "HEAD")
    assert not permissoes.permitido({"admin"}, "bem", "DELETE")


def _gerar_todos_subconjuntos():
    """Gera todos os 32 subconjuntos de FUNCOES (2^5)."""
    for r in range(len(permissoes.FUNCOES) + 1):
        for combo in itertools.combinations(permissoes.FUNCOES, r):
            yield frozenset(combo)


@pytest.mark.parametrize("subset", list(_gerar_todos_subconjuntos()))
@pytest.mark.parametrize("endpoint,metodo", list(permissoes.PERMISSOES.keys()))
def test_uniao_de_funcoes(subset, endpoint, metodo):
    """Para cada subconjunto, o resultado é a união das funções isoladas."""
    esperado = any(
        permissoes.permitido({funcao}, endpoint, metodo)
        for funcao in subset
    )
    assert permissoes.permitido(subset, endpoint, metodo) == esperado


def test_conjunto_vazio_retorna_false():
    """Conjunto vazio não deve ter acesso a nada."""
    for endpoint, metodo in permissoes.PERMISSOES.keys():
        assert not permissoes.permitido(frozenset(), endpoint, metodo)


def test_funcao_desconhecida_retorna_false():
    """Função desconhecida não deve ter acesso a nada."""
    funcao_desconhecida = {"admin", "x"}
    for endpoint, metodo in permissoes.PERMISSOES.keys():
        assert not permissoes.permitido(funcao_desconhecida, endpoint, metodo)
