"""Confere, no banco deste sistema, os números da campanha que o TCC apresenta (seção "Conferência sistema × TCC"
em numeros-da-pesquisa.md). Só lê. Rodar na raiz do projeto:  .venv/bin/python docs/tcc/conferir_sistema.py"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
import db  # noqa: E402

c = db.conectar()
q = lambda s: c.execute(s).fetchone()   # noqa: E731
B = "inventario_bens_encerrados"        # snapshot congelado no encerramento do evento
ativos = q(f"select count(*) from {B} where situacao='ATIVO'")[0]
lidos_ativos = q(f"select count(*) from inventario_leituras l join {B} b on b.numero=l.numero where b.situacao='ATIVO'")[0]
print("Evento:", tuple(q("select nome, aberto_em, encerrado_em from inventario_eventos")))
print("Base hoje: registros", q("select count(*) from bens")[0], "| ativos", q("select count(*) from bens where situacao='ATIVO'")[0],
      "| localizações com ativos", q("select count(distinct localizacao) from bens where situacao='ATIVO'")[0])
print("Salas do evento:", q("select count(*) from inventario_salas")[0], "| ativos no snapshot:", ativos)
print("Leituras:", q("select count(*) from inventario_leituras")[0], "| sobras:", q("select count(*) from inventario_sobras")[0])
print("Ativos lidos:", lidos_ativos, f"| cobertura {100 * lidos_ativos / ativos:.1f}%", "| pendentes:", ativos - lidos_ativos)
print("Baixados/doados lidos:", q(f"select count(*) from inventario_leituras l join {B} b on b.numero=l.numero where b.situacao<>'ATIVO'")[0])
print("Divergentes: todas", q(f"select count(*) from inventario_leituras l join {B} b on b.numero=l.numero where l.localizacao<>b.localizacao")[0],
      "| só ativos", q(f"select count(*) from inventario_leituras l join {B} b on b.numero=l.numero where l.localizacao<>b.localizacao and b.situacao='ATIVO'")[0])
print("Período:", tuple(q("select min(lido_em), max(lido_em) from inventario_leituras")),
      "| dias de campo:", q("select count(distinct substr(lido_em,1,10)) from inventario_leituras")[0],
      "| integrantes que leram:", q("select count(distinct integrante) from inventario_leituras")[0])
print("Conservação:", [tuple(r) for r in c.execute("select conservacao, count(*) from inventario_leituras group by 1")])
