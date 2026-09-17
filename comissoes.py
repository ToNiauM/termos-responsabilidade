"""Comissão de inventário por identidade de usuário. Só dados e regras: toda função recebe `conn` primeiro
e não importa Flask (padrão de db.py e inventario.py).

Quem pode ver e quem pode conferir um evento é decidido por `inventario_comissao_usuarios` (evento_id +
usuario_id). O nome em `inventario_integrantes` continua existindo — é o que aparece nas telas, nas leituras
e na planilha de intercâmbio —, mas nunca autoriza nada na web: homônimos, renomeações e contas recriadas
não mudam o escopo de ninguém.

Exceção do modo desktop: o "Administrador local" não tem linha em `usuarios` (id None), então é conferido
pelo nome em `inventario_integrantes`, como sempre foi (lá só existe um usuário, o administrador da máquina).

Transações: `definir` e `atualizar_nome` abrem transação própria e exigem conexão sem transação pendente;
`abrir` envolve `inventario.abrir_evento(..., confirmar=False)` para que evento e vínculos caiam juntos.
"""
import db
import permissoes


def membro(conn, u, eid) -> bool:
    """O usuário faz parte da comissão deste evento (por vínculo; desktop, pelo nome do administrador local)."""
    if u["id"] is None:
        return "admin" in u["funcoes"] and bool(conn.execute(
            "SELECT 1 FROM inventario_integrantes WHERE evento_id=? AND nome=?", (eid, u["nome"])).fetchone())
    return bool(conn.execute("SELECT 1 FROM inventario_comissao_usuarios WHERE evento_id=? AND usuario_id=?",
                             (eid, u["id"])).fetchone())


def visivel(conn, u, eid) -> bool:
    """Admin e consulta de inventários veem todos os eventos; inventariante só vê os eventos de que participa."""
    f = set(u["funcoes"])
    return bool(f & {"admin", "consulta_inventarios"}) or ("inventariante" in f and membro(conn, u, eid))


def pode_conferir(conn, u, eid) -> bool:
    """Ler bens, tirar fotos e registrar sobras: precisa de função de conferência E do vínculo com o evento."""
    return bool(set(u["funcoes"]) & permissoes.CONFERENCIA) and membro(conn, u, eid)


def eventos_visiveis(conn, u) -> list[dict]:
    import inventario
    return [e for e in inventario.eventos(conn) if visivel(conn, u, e["id"])]


def _usuarios_selecionados(conn, ids) -> list[dict]:
    """IDs vindos do formulário → usuários elegíveis. ID oculto, inativo ou sem função de inventário é recusado."""
    import usuarios
    try:
        ids = sorted({int(i) for i in ids})
    except (TypeError, ValueError):
        raise db.ErroDeNegocio("Selecione integrantes válidos.")
    elegiveis = {u["id"]: u for u in usuarios.elegiveis_comissao(conn)}
    if not ids or not set(ids) <= set(elegiveis):
        raise db.ErroDeNegocio("Selecione ao menos um usuário ativo com função de inventário.")
    return [elegiveis[i] for i in ids]


def definir(conn, eid, ids) -> None:
    """Substitui a comissão do evento aberto: vínculos e nomes exibidos. Leituras já feitas não mudam."""
    import inventario
    with conn:
        conn.execute("BEGIN IMMEDIATE")
        inventario._evento_aberto_ou_erro(conn, eid)
        selecionados = _usuarios_selecionados(conn, ids)
        conn.execute("DELETE FROM inventario_comissao_usuarios WHERE evento_id=?", (eid,))
        conn.executemany("INSERT INTO inventario_comissao_usuarios VALUES (?,?,?)",
                         [(eid, u["id"], u["nome"]) for u in selecionados])
        conn.execute("DELETE FROM inventario_integrantes WHERE evento_id=?", (eid,))
        conn.executemany("INSERT INTO inventario_integrantes VALUES (?,?)",
                         [(eid, n) for n in sorted({u["nome"] for u in selecionados})])


def abrir(conn, nome, descricao, ids, salas) -> int:
    """Abre o evento e grava os vínculos na mesma transação (falha no vínculo desfaz o evento)."""
    import inventario
    selecionados = _usuarios_selecionados(conn, ids)
    with conn:
        eid = inventario.abrir_evento(conn, nome, descricao, [u["nome"] for u in selecionados],
                                      salas, confirmar=False)
        conn.executemany("INSERT INTO inventario_comissao_usuarios VALUES (?,?,?)",
                         [(eid, u["id"], u["nome"]) for u in selecionados])
    return eid


def atualizar_nome(conn, uid) -> None:
    """Usuário renomeado: atualiza o nome dele na comissão dos eventos ABERTOS (vínculo e nome exibido).
    Histórico não muda: leituras, sobras e eventos encerrados mantêm o nome de quando foram registrados.
    Nomes legados (integrantes sem conta vinculada) continuam na lista do evento."""
    u = conn.execute("SELECT nome FROM usuarios WHERE id=?", (uid,)).fetchone()
    if not u:
        return
    eventos = [r[0] for r in conn.execute("""SELECT c.evento_id FROM inventario_comissao_usuarios c
      JOIN inventario_eventos e ON e.id=c.evento_id WHERE c.usuario_id=? AND e.encerrado_em IS NULL""", (uid,))]
    with conn:
        for eid in eventos:
            vinculados = {r[0] for r in conn.execute(
                "SELECT nome_na_comissao FROM inventario_comissao_usuarios WHERE evento_id=?", (eid,))}
            nomes = {r[0] for r in conn.execute(
                "SELECT nome FROM inventario_integrantes WHERE evento_id=?", (eid,))}
            legado = nomes - vinculados
            conn.execute("UPDATE inventario_comissao_usuarios SET nome_na_comissao=? WHERE evento_id=? AND usuario_id=?",
                         (u[0], eid, uid))
            atuais = {r[0] for r in conn.execute(
                "SELECT nome_na_comissao FROM inventario_comissao_usuarios WHERE evento_id=?", (eid,))}
            conn.execute("DELETE FROM inventario_integrantes WHERE evento_id=?", (eid,))
            conn.executemany("INSERT INTO inventario_integrantes VALUES (?,?)",
                             [(eid, n) for n in sorted(legado | atuais)])


def vinculos(conn, eid) -> list[dict]:
    """Comissão do evento com identidade: [{usuario_id, nome_na_comissao}] em ordem de nome."""
    return [{"usuario_id": r[0], "nome_na_comissao": r[1]} for r in conn.execute(
        "SELECT usuario_id, nome_na_comissao FROM inventario_comissao_usuarios WHERE evento_id=? ORDER BY nome_na_comissao",
        (eid,))]
