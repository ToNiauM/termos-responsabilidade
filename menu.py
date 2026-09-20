"""Navegação e ajuda: a árvore do menu e as seções do guia, ambas recortadas pelas funções do usuário.

Sem banco: `montar` recebe o evento aberto já autorizado (ou None) e só depende da matriz de permissões.
"""
from flask import url_for

from permissoes import permitido

CADASTROS = [('Centros de custo', 'responsaveis'), ('Localizações', 'localizacoes'),
             ('Pessoas', 'pessoas'), ('Processos SEI', 'processos')]
TERMOS = [('Termo por centro de custo', 'centro_custos'), ('Termo individual', 'termos_individuais'),
          ('Termo de devolução', 'termo_devolucao'), ('Termos emitidos', 'termos_emitidos_tela')]
# Telas de cadastro sem item próprio: acendem a aba de onde vieram
CADASTRO_ENDPOINTS = {
    'responsaveis_editar': 'responsaveis', 'responsaveis_incluir': 'responsaveis', 'responsaveis_excluir': 'responsaveis',
    'pessoas_editar': 'pessoas', 'pessoas_incluir': 'pessoas', 'pessoas_excluir': 'pessoas',
    'pessoas_atribuir': 'pessoas', 'pessoas_desatribuir': 'pessoas',
    'localizacoes_alterar': 'localizacoes', 'localizacoes_incluir': 'localizacoes',
    'localizacoes_mover': 'localizacoes', 'localizacoes_excluir': 'localizacoes',
    'processos_incluir': 'processos', 'processos_vigente': 'processos',
    'processos_encerrar': 'processos', 'processos_excluir': 'processos',
}


def destino_atual(endpoint, args):
    """Endpoint do menu que representa a tela aberta (telas derivadas acendem o item de origem)."""
    args = dict(args or {})
    if endpoint in CADASTRO_ENDPOINTS:
        return 'cadastros', {'aba': CADASTRO_ENDPOINTS[endpoint]}
    if endpoint == 'cadastro_novo':
        return 'cadastros', {'aba': args.get('aba')}
    if endpoint == 'termo':
        return {'ccusto': 'centro_custos', 'individual': 'termos_individuais',
                'devolucao': 'termo_devolucao'}.get(args.get('tipo'), 'centro_custos'), {}
    if endpoint in {'termo_emitido_tela', 'termo_emitido_documento', 'termo_emitido_email'}:
        return 'termos_emitidos_tela', {}
    if endpoint in {'gerar', 'gerar_individual'}:
        return ('centro_custos' if endpoint == 'gerar' else 'termos_individuais'), {}
    if endpoint in {'usuarios.novo', 'usuarios.incluir', 'usuarios.editar', 'usuarios.nova_senha', 'usuarios.apagar_acessos'}:
        return 'usuarios.lista', {}
    if endpoint in {'inventario.sala_tela', 'inventario.comissao', 'inventario.excluir'}:
        return 'inventario.evento_tela', {'id': args.get('id')}
    if endpoint == 'importacao_tela':
        return 'upload', {}
    return endpoint, args


def montar(funcoes, evento_aberto, endpoint_atual, argumentos=None, login_ativo=True):
    """Árvore do menu: itens e grupos permitidos, com o item da tela atual marcado."""
    ep, args = destino_atual(endpoint_atual, argumentos)

    def item(rotulo, icone, endpoint, kw=None):
        kw = kw or {}   # todo nó tem a mesma forma: a folha traz id=None e filhos vazios; o grupo, url=None

        return dict(id=None, rotulo=rotulo, icone=icone, endpoint=endpoint, url=url_for(endpoint, **kw),
                    ativo=ep == endpoint and all(args.get(k) == v for k, v in kw.items()),
                    filhos=[], aberto=False)

    def grupo(id, rotulo, icone, filhos, aberto):
        filhos = [f for f in filhos if permitido(funcoes, f['endpoint'])]
        return dict(id=id, rotulo=rotulo, icone=icone, filhos=filhos, aberto=aberto, ativo=False, url=None)

    itens = []
    if permitido(funcoes, 'analise'):
        itens.append(item('Início', 'fa-home', 'home'))
    g = grupo('termos', 'Termos de Responsabilidade', 'fa-file-signature',
              [item(r, '', e) for r, e in TERMOS], ep in {e for _, e in TERMOS})
    if g['filhos']:
        itens.append(g)
    if permitido(funcoes, 'analise'):
        itens.append(item('Análise', 'fa-chart-bar', 'analise'))
    filhos = [item('Eventos', '', 'inventario.eventos_tela')]
    if evento_aberto:
        for r, e in [(evento_aberto['nome'], 'inventario.evento_tela'), ('Painel', 'inventario.painel_tela'),
                     ('Relatório', 'inventario.relatorio_tela')]:
            filhos.append(item(r, '', e, {'id': evento_aberto['id']}))
    g = grupo('inventario', 'Inventário', 'fa-clipboard-check', filhos,
              (endpoint_atual or '').startswith('inventario.'))
    if g['filhos']:
        itens.append(g)
    g = grupo('cadastros', 'Cadastros', 'fa-address-book',
              [item(r, '', 'cadastros', {'aba': aba}) for r, aba in CADASTROS], ep == 'cadastros')
    if g['filhos']:
        itens.append(g)
    for r, i, e in [('Textos', 'fa-pen-nib', 'textos_tela'), ('Atualizar base', 'fa-upload', 'upload'),
                    ('Usuários', 'fa-users', 'usuarios.lista'), ('Ajuda', 'fa-question-circle', 'ajuda')]:
        if permitido(funcoes, e) and (e != 'usuarios.lista' or login_ativo):
            itens.append(item(r, i, e))
    return itens


# Seção do guia: (id, título, endpoint que a autoriza, método)
SECOES = [
    ('inicio', 'Início', 'analise', 'GET'), ('pesquisa', 'Pesquisa', 'pesquisa', 'GET'),
    ('termos', 'Termos de Responsabilidade', 'termo', 'GET'), ('analise', 'Análise', 'analise', 'GET'),
    ('inventario', 'Conferência do inventário', 'inventario.ler', 'POST'),
    ('consulta-inventarios', 'Consulta de inventários', 'inventario.relatorio_tela', 'GET'),
    ('cadastros', 'Cadastros', 'cadastros', 'GET'), ('textos', 'Textos', 'textos_tela', 'GET'),
    ('atualizar-base', 'Atualizar base', 'upload', 'GET'), ('usuarios', 'Usuários', 'usuarios.lista', 'GET'),
    ('conta', 'Sua conta', 'usuarios.senha', 'GET'), ('perguntas', 'Perguntas frequentes', 'ajuda', 'GET'),
]
# Tela → seção do guia (telas fora daqui não ganham atalho contextual)
AJUDA = {
    'home': 'inicio', 'pesquisa': 'pesquisa', 'bem': 'pesquisa', 'analise': 'analise',
    'centro_custos': 'termos', 'termos_individuais': 'termos', 'termo_devolucao': 'termos',
    'termos_emitidos_tela': 'termos',
    'cadastros': 'cadastros', 'textos_tela': 'textos', 'textos_salvar': 'textos', 'upload': 'atualizar-base',
    'usuarios.lista': 'usuarios', 'usuarios.senha': 'conta', 'usuarios.acessos': 'conta',
    'inventario.painel_tela': 'consulta-inventarios', 'inventario.relatorio_tela': 'consulta-inventarios',
}


def secoes_ajuda(funcoes, login_ativo):
    """Seções do guia que este usuário pode ver, na ordem de SECOES."""
    return [dict(id=id, titulo=titulo) for id, titulo, ep, m in SECOES
            if permitido(funcoes, ep, m) and (id not in {'conta', 'usuarios'} or login_ativo)]


def ancora_ajuda(funcoes, endpoint, argumentos, login_ativo):
    """Seção do guia correspondente à tela atual, ou None (a própria Ajuda, login, erro e documentos)."""
    if endpoint in {None, 'ajuda', 'usuarios.login'}:
        return None
    ep, _ = destino_atual(endpoint, argumentos)
    if (ep or '').startswith('inventario.') and ep not in AJUDA:
        ancora = 'inventario' if permitido(funcoes, 'inventario.ler', 'POST') else 'consulta-inventarios'
    else:
        ancora = AJUDA.get(ep)
    return ancora if ancora in {s['id'] for s in secoes_ajuda(funcoes, login_ativo)} else None
