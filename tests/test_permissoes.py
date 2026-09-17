"""Matriz de permissões: unidade (permitido) e cobertura de todas as rotas do app."""
import pytest

import usuarios


def test_permitido_por_perfil():
    assert usuarios.permitido("consulta", "home") and usuarios.permitido("inventariante", "bem")
    assert usuarios.permitido("consulta", "termo") and usuarios.permitido("consulta", "termo_documento")
    assert not usuarios.permitido("inventariante", "termo") and not usuarios.permitido("inventariante", "centro_custos")
    assert not usuarios.permitido("consulta", "termo_docx") and not usuarios.permitido("consulta", "termo_registrar", "POST")
    assert usuarios.permitido("consulta", "termo_devolucao") and not usuarios.permitido("consulta", "termo_devolucao", "POST")
    assert usuarios.permitido("operador", "termo_devolucao", "POST")
    assert usuarios.permitido("operador", "responsaveis_editar", "POST") and not usuarios.permitido("operador", "responsaveis_excluir", "POST")
    assert usuarios.permitido("operador", "upload", "POST") and not usuarios.permitido("operador", "importar_cadastros", "POST")
    assert not usuarios.permitido("inventariante", "cadastros") and not usuarios.permitido("inventariante", "textos_tela")
    assert usuarios.permitido("inventariante", "inventario.ler", "POST") and not usuarios.permitido("consulta", "inventario.ler", "POST")
    assert not usuarios.permitido("operador", "inventario.abrir", "POST") and usuarios.permitido("admin", "inventario.excluir", "POST")
    assert not usuarios.permitido("operador", "usuarios.lista") and usuarios.permitido("admin", "usuarios.lista")
    for perfil in usuarios.PERFIS:
        assert usuarios.permitido(perfil, "usuarios.senha", "POST") and usuarios.permitido(perfil, "usuarios.sair", "POST")
    assert usuarios.permitido("admin", "termo_docx", "HEAD")                 # HEAD conta como GET
    assert not usuarios.permitido("admin", "rota_que_nao_existe")            # fora da matriz = negado a todos
    assert not usuarios.permitido("chefe", "home")


def test_toda_rota_do_app_esta_na_matriz():
    from app import app
    faltam = []
    for regra in app.url_map.iter_rules():
        if regra.endpoint == "static":
            continue
        for metodo in regra.methods - {"HEAD", "OPTIONS"}:
            chave = regra.endpoint if metodo == "GET" else f"{regra.endpoint}:{metodo}"
            if chave not in usuarios.PERMISSOES and regra.endpoint not in usuarios.PERMISSOES:
                faltam.append(f"{regra.endpoint} [{metodo}]")
    assert not faltam, "Rotas sem regra em usuarios.PERMISSOES: " + ", ".join(sorted(faltam))
