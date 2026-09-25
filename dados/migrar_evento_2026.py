"""Migra o inventário 2026 do sistema anterior (planilha INVENTARIO_SEPAT, aba BASE, ou o relatório dele
convertido por relatorio_para_base.py) para o evento "2026" deste sistema.

Usa o migrar_inventario.py do próprio sistema, só que:
  - mantém o evento já criado no sistema (id 1, nome "2026"), em vez de criar "Inventário 2026";
  - aberto_em = primeira leitura da planilha (18/08/2026), como o script original;
  - renomeia integrantes para bater com os usuários cadastrados (--renomear "A=B");
  - traduz salas lidas com o nome antigo para o nome atual do SPW (--depara csv "antiga;atual");
  - por padrão PRESERVA o que já foi feito neste sistema no mesmo evento: leitura mais recente vence (as feitas
    aqui depois da campanha substituem a migrada do mesmo bem), fotos e sobras acompanham, integrantes se somam;
    com --descartar-feitas-aqui as leituras, fotos e sobras já existentes no evento são apagadas (eram testes do
    sistema) e só a comissão (integrantes e vínculos) é mantida — foi assim na carga de 25/09/2026;
  - exporta a planilha de cadastros direto do banco (não precisa baixar pela tela);
  - sem --gravar só gera a planilha (ensaio); com --gravar carrega pela mesma validação tudo-ou-nada da tela
    (db.importar_cadastros) e repõe os vínculos da comissão, que a carga descartaria porque aberto_em muda.

    cd /opt/web/sistema-inventario && .venv/bin/python dados/migrar_evento_2026.py dados/relatorio_como_BASE_2026-09-22.xlsx \
        dados/cadastros_com_inventario_2026.xlsx --depara dados/depara_localizacoes_2026.csv \
        --renomear "Denise Cristiane=Denise Cristiane Silva" [--pular-inexistentes] [--descartar-feitas-aqui] [--gravar]
"""
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, "/opt/web/sistema-inventario")
import db                      # noqa: E402
import migrar_inventario as mi  # noqa: E402

EVENTO_ID = 1   # evento "2026" já criado na tela Administração › Inventários


def preservar_existentes(conn, abas: dict, descartar: bool = False) -> dict:
    """Leituras, sobras, fotos e integrantes que o evento já tem neste sistema entram na planilha gerada.
    Leitura do mesmo bem: fica a mais recente (a migrada perde a foto se for substituída, porque não tinha).
    descartar=True: só os integrantes (a comissão) são mantidos; leituras, fotos e sobras existentes somem."""
    eid = EVENTO_ID
    novas = [] if descartar else conn.execute("SELECT numero, localizacao, lido_em, integrante, conservacao, quem_usa, observacao, fotos_seq "
                         "FROM inventario_leituras WHERE evento_id = ?", (eid,)).fetchall()
    pos = {l[1]: i for i, l in enumerate(abas["inv_leituras"])}
    substituidas, acrescentadas, mantidas_migradas = 0, 0, 0
    for n in novas:
        linha = [eid, n["numero"], n["localizacao"], n["lido_em"], n["integrante"], n["conservacao"], n["quem_usa"], n["observacao"], n["fotos_seq"]]
        if n["numero"] in pos:
            if n["lido_em"] >= abas["inv_leituras"][pos[n["numero"]]][3]:
                abas["inv_leituras"][pos[n["numero"]]] = linha; substituidas += 1
            else:
                mantidas_migradas += 1
        else:
            pos[n["numero"]] = len(abas["inv_leituras"]); abas["inv_leituras"].append(linha); acrescentadas += 1
    vence = {l[1]: l[3] for l in abas["inv_leituras"]}
    fotos = [] if descartar else conn.execute("SELECT numero, nfoto, url, criado_em FROM inventario_fotos WHERE evento_id = ?", (eid,)).fetchall()
    abas["inv_fotos"] = [f for f in abas["inv_fotos"] if f[1] in vence] + \
                        [[eid, f["numero"], f["nfoto"], f["url"], f["criado_em"]] for f in fotos if f["numero"] in vence]
    sobras = [] if descartar else conn.execute("SELECT localizacao, descricao, complemento, observacao, foto_url, integrante, criado_em "
                                               "FROM inventario_sobras WHERE evento_id = ?", (eid,)).fetchall()
    abas["inv_sobras"] += [[eid, *s] for s in sobras]
    nomes = {i[1] for i in abas["inv_integrantes"]} | {l[4] for l in abas["inv_leituras"]} | {s[6] for s in abas["inv_sobras"]}
    nomes |= {r[0] for r in conn.execute("SELECT nome FROM inventario_integrantes WHERE evento_id = ?", (eid,))}
    abas["inv_integrantes"] = [[eid, n] for n in sorted(nomes)]
    abas["inv_salas"] = [[eid, s] for s in sorted({s[1] for s in abas["inv_salas"]} | {l[2] for l in abas["inv_leituras"]} | {s[1] for s in abas["inv_sobras"]})]
    return {"existentes": len(novas), "substituidas": substituidas, "acrescentadas": acrescentadas,
            "mantidas_migradas": mantidas_migradas, "fotos": len(fotos), "sobras": len(sobras)}


def main(argv):
    pular, gravar, descartar = "--pular-inexistentes" in argv, "--gravar" in argv, "--descartar-feitas-aqui" in argv
    renomear, depara = {}, {}
    args = []
    it = iter(argv)
    for a in it:
        if a == "--renomear":
            de, para = next(it).split("=", 1); renomear[de.strip()] = para.strip()
        elif a == "--depara":                          # csv "antiga;atual" (nomes de localização do SPW)
            for ln in Path(next(it)).read_text(encoding="utf-8").splitlines()[1:]:
                if ";" in ln:
                    de, para = ln.split(";", 1)
                    if para.strip() and para.strip() != "?":
                        depara[de.strip().upper()] = para.strip()
        elif not a.startswith("--"):
            args.append(a)
    if len(args) < 2:
        raise SystemExit(__doc__)
    antiga, saida = Path(args[0]), Path(args[1])

    conn = db.conectar()
    ev = conn.execute("SELECT id, nome, descricao FROM inventario_eventos WHERE id = ?", (EVENTO_ID,)).fetchone()
    if not ev:
        raise SystemExit(f"Evento id {EVENTO_ID} não existe no banco; crie-o na tela Administração › Inventários.")
    mi.EVENTO_NOME, mi.EVENTO_DESCRICAO = ev[1], ev[2] or mi.EVENTO_DESCRICAO

    linhas = mi.ler_base(antiga)
    for r in linhas:                                   # nomes como estão nos usuários do sistema
        u = " ".join(db._texto(r.get("usuario")).split())
        if u in renomear:
            r["usuario"] = renomear[u]
        loc = db._texto(r.get("local_verificado")).strip().upper()
        if loc in depara:                              # sala lida com o nome antigo -> nome atual do SPW
            r["local_verificado"] = depara[loc]
    abas, r = mi.transformar(conn, linhas, pular)
    print(f"Planilha antiga: {r['leituras']} leituras ({r['divergentes']} divergentes) | sobras: {r['sobras']} | salas: {r['salas']}")
    p = preservar_existentes(conn, abas, descartar)
    hoje = set(db.localizacoes_ativas(conn))
    fantasmas = sorted({l[2] for l in abas["inv_leituras"]} - hoje)
    if fantasmas:
        print(f"ATENÇÃO: {len(fantasmas)} localização(ões) lida(s) não existem hoje no sistema: {', '.join(fantasmas)}")

    cadastros = saida.with_name("cadastros_atual.xlsx")
    db.exportar_cadastros(conn, cadastros)             # 4 abas de cadastro (+ inv_* atuais, que serão substituídas)
    mi.gravar(cadastros, saida, abas)
    leituras = abas["inv_leituras"]
    divergentes = sum(1 for l in leituras if db.buscar_bem(conn, l[1])["localizacao"] != l[2])

    print(f"Gerado: {saida}")
    print(f"  evento: {mi.EVENTO_NOME} (id {EVENTO_ID}, aberto), aberto_em {abas['inv_eventos'][0][3]}")
    print(f"  leituras: {len(leituras)} ({divergentes} divergentes) | sobras: {len(abas['inv_sobras'])} | salas: {len(abas['inv_salas'])} | fotos: {len(abas['inv_fotos'])}")
    for nome, n in sorted(Counter(l[4] for l in leituras).items()):
        print(f"    {nome}: {n}")
    if descartar:
        n_desc = conn.execute("SELECT count(*) FROM inventario_leituras WHERE evento_id = ?", (EVENTO_ID,)).fetchone()[0]
        print(f"  DESCARTADAS (--descartar-feitas-aqui): {n_desc} leitura(s) já feitas neste sistema, com fotos e sobras")
    print(f"  já feitas neste sistema: {p['existentes']} leitura(s) ({p['substituidas']} substituíram a migrada do mesmo bem, "
          f"{p['acrescentadas']} de bens ainda não lidos, {p['mantidas_migradas']} mais antigas que a migrada), "
          f"{p['fotos']} foto(s), {p['sobras']} sobra(s)")
    if r["inexistentes"]:
        print(f"  DEIXADAS DE FORA: {len(r['inexistentes'])} leitura(s) de bens fora da base: "
              f"{', '.join(map(str, r['inexistentes'][:20]))}{' ...' if len(r['inexistentes']) > 20 else ''}")
    for pr in r["problemas"]:
        print(f"  aviso: {pr}")

    if not gravar:
        conn.close()
        print("Ensaio: nada gravado. Confira a planilha e rode de novo com --gravar (ou carregue em Atualizar base → Cadastros).")
        return
    vinculos = conn.execute("SELECT evento_id, usuario_id, nome_na_comissao FROM inventario_comissao_usuarios WHERE evento_id = ?", (EVENTO_ID,)).fetchall()
    res = db.importar_cadastros(conn, saida)           # tudo ou nada; apaga e regrava as tabelas inv_*
    nomes = {i[1] for i in abas["inv_integrantes"]}
    repostos = [tuple(v) for v in vinculos if v[2] in nomes]
    conn.executemany("INSERT OR IGNORE INTO inventario_comissao_usuarios VALUES (?,?,?)", repostos)
    conn.commit()
    print(f"GRAVADO no banco: {res}")
    print(f"  vínculos da comissão repostos: {len(repostos)} de {len(vinculos)}")
    conn.close()


if __name__ == "__main__":
    main(sys.argv[1:])
