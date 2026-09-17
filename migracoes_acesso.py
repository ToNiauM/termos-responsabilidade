"""Migração única do esquema de acesso: perfil único vira funções (usuarios_funcoes) e a comissão
de inventário ganha identidade de usuário (inventario_comissao_usuarios), quando for possível ligar
sem ambiguidade o nome digitado na comissão a uma conta ativa com função de admin ou inventariante.

`migrar` é idempotente: repetir não reconcede funções removidas nem religa comissão desfeita
(marcador em `migracoes_acesso`), mas as tabelas em si são sempre criadas (IF NOT EXISTS), inclusive
num banco que nasce direto no esquema novo.
"""
import sqlite3


def migrar(conn: sqlite3.Connection) -> None:
    with conn:
        conn.execute('BEGIN IMMEDIATE')
        conn.execute("CREATE TABLE IF NOT EXISTS migracoes_acesso (nome TEXT PRIMARY KEY)")
        conn.execute("""CREATE TABLE IF NOT EXISTS usuarios_funcoes (
          usuario_id INTEGER NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
          funcao TEXT NOT NULL CHECK(funcao IN
            ('admin','operador','consulta','inventariante','consulta_inventarios')),
          PRIMARY KEY(usuario_id, funcao))""")
        conn.execute("""CREATE TABLE IF NOT EXISTS inventario_comissao_usuarios (
          evento_id INTEGER NOT NULL REFERENCES inventario_eventos(id) ON DELETE CASCADE,
          usuario_id INTEGER NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
          nome_na_comissao TEXT NOT NULL,
          PRIMARY KEY(evento_id, usuario_id))""")
        if conn.execute("SELECT 1 FROM migracoes_acesso WHERE nome='fase5_funcoes'").fetchone():
            return
        colunas = {r[1] for r in conn.execute('PRAGMA table_info(usuarios)')}
        if 'perfil' in colunas:
            conn.execute("INSERT INTO usuarios_funcoes SELECT id, perfil FROM usuarios")
            conn.execute('ALTER TABLE usuarios DROP COLUMN perfil')
        conn.execute("""INSERT INTO inventario_comissao_usuarios
          SELECT i.evento_id, u.id, i.nome
          FROM inventario_integrantes i JOIN usuarios u ON u.nome=i.nome
          WHERE u.ativo=1
            AND (SELECT count(*) FROM usuarios x WHERE x.nome=i.nome)=1
            AND EXISTS(SELECT 1 FROM usuarios_funcoes f WHERE f.usuario_id=u.id
                       AND f.funcao IN ('admin','inventariante'))""")
        conn.execute("INSERT INTO migracoes_acesso VALUES ('fase5_funcoes')")
