FUNCOES = ("admin", "operador", "consulta", "inventariante", "consulta_inventarios")
ROTULOS = dict(zip(FUNCOES, ("Administrador", "Operador", "Consulta", "Inventário", "Consulta de inventários")))
TODAS = frozenset(FUNCOES)
ACERVO = frozenset({"admin", "operador", "consulta"})
GESTAO = frozenset({"admin", "operador"})
ADMIN = frozenset({"admin"})
CONFERENCIA = frozenset({"admin", "inventariante"})
EVENTOS = frozenset({"admin", "inventariante", "consulta_inventarios"})
RELATORIOS = frozenset({"admin", "consulta_inventarios"})
PERMISSOES = {}


def _registrar(funcoes, linhas):
    for linha in linhas.strip().splitlines():
        endpoint, *metodos = linha.split()
        for metodo in metodos:
            chave = (endpoint, metodo)
            if chave in PERMISSOES:
                raise ValueError(f"Permissão repetida: {chave}")
            PERMISSOES[chave] = funcoes


_registrar(TODAS, """
home GET
ajuda GET
usuarios.login GET POST
usuarios.sair POST
usuarios.senha GET POST
usuarios.aparencia POST
usuarios.acessos GET
usuarios.salvar_acesso_sei POST
usuarios.apagar_acesso_sei POST
usuarios.salvar_acesso_spw POST
usuarios.apagar_acesso_spw POST
""")
_registrar(ACERVO, """
bem GET
pesquisa GET
recorte GET
recorte_xlsx GET
analise GET
analise_xlsx GET
centro_custos GET
termos_individuais GET
termo GET
termo_documento GET
termo_devolucao GET
termos_emitidos_tela GET
termo_emitido_tela GET
""")
_registrar(GESTAO, """
gerar POST
gerar_individual POST
termo_docx GET
termo_planilha GET
termo_registrar POST
termo_devolucao POST
termo_emitido_documento POST
termo_emitido_email POST
termo_enviar_sei POST
termo_emitido_enviar_sei POST
cadastros GET
cadastro_novo GET
responsaveis_incluir POST
responsaveis_editar GET POST
pessoas_incluir POST
pessoas_editar GET POST
pessoas_atribuir POST
pessoas_desatribuir POST
localizacoes_incluir POST
localizacoes_alterar GET
localizacoes_mover POST
processos_incluir POST
processos_vigente POST
processos_encerrar POST
cadastros_exportar GET
textos_tela GET
textos_salvar POST
upload GET POST
bens_exportar GET
importacao_tela GET
""")
_registrar(ADMIN, """
responsaveis_excluir POST
pessoas_excluir POST
localizacoes_excluir POST
processos_excluir POST
importar_cadastros POST
admin.tela GET
inventario.abrir POST
inventario.abrir_chave POST
inventario.encerrar POST
inventario.fechar POST
inventario.comissao GET POST
inventario.excluir GET POST
usuarios.lista GET
usuarios.novo GET
usuarios.incluir POST
usuarios.editar GET POST
usuarios.nova_senha POST
usuarios.apagar_acessos POST
base_atualizar_spw POST
""")
_registrar(EVENTOS, """
inventario.eventos_tela GET
inventario.evento_tela GET
inventario.sala_tela GET
""")
_registrar(RELATORIOS, """
inventario.relatorio_tela GET
inventario.painel_tela GET
inventario.xlsx GET
""")
_registrar(CONFERENCIA, """
inventario.ler POST
inventario.atualizar_leitura POST
inventario.lote POST
inventario.foto_leitura POST
inventario.foto_excluir POST
inventario.sobra POST
inventario.sobra_excluir POST
""")


def permitido(funcoes, endpoint, metodo="GET"):
    if isinstance(funcoes, str):
        return False
    concedidas = frozenset(funcoes or ())
    if not concedidas <= TODAS:
        return False
    metodo = "GET" if metodo == "HEAD" else metodo
    return bool(concedidas & PERMISSOES.get((endpoint, metodo), frozenset()))
