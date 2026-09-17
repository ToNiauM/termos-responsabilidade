import sqlite3

import pytest

from migracoes_acesso import migrar


def _banco_legado():
    c = sqlite3.connect(":memory:")
    c.executescript("""
      PRAGMA foreign_keys=ON;
      CREATE TABLE usuarios(id INTEGER PRIMARY KEY, nome TEXT, ativo INTEGER,
        perfil TEXT CHECK(perfil IN ('admin','operador','consulta','inventariante')));
      CREATE TABLE inventario_eventos(id INTEGER PRIMARY KEY);
      CREATE TABLE inventario_integrantes(evento_id INTEGER, nome TEXT);
      INSERT INTO usuarios VALUES(1,'Ana',1,'inventariante'),(2,'Ana',1,'inventariante'),
        (3,'Beto',1,'inventariante'),(4,'Cris',1,'operador');
      INSERT INTO inventario_eventos VALUES(1);
      INSERT INTO inventario_integrantes VALUES(1,'Ana'),(1,'Beto'),(1,'Cris');
    """)
    return c


def test_migracao_nao_reconcede_funcao_nem_homonimo():
    c = _banco_legado()
    migrar(c)
    assert c.execute("SELECT usuario_id FROM inventario_comissao_usuarios").fetchall() == [(3,)]
    assert c.execute("SELECT funcao FROM usuarios_funcoes WHERE usuario_id=4").fetchall() == [('operador',)]
    c.execute("DELETE FROM usuarios_funcoes WHERE usuario_id=3")
    c.execute("DELETE FROM inventario_comissao_usuarios WHERE usuario_id=3")
    c.commit()
    migrar(c)
    assert not c.execute("SELECT 1 FROM usuarios_funcoes WHERE usuario_id=3").fetchone()
    assert not c.execute("SELECT 1 FROM inventario_comissao_usuarios").fetchone()
    assert 'perfil' not in [r[1] for r in c.execute('PRAGMA table_info(usuarios)')]


def test_migracao_banco_novo_cria_tabelas_sem_dados_legados():
    """Banco sem coluna perfil (já nasceu no esquema novo): as tabelas são criadas mas nada é
    inserido a partir de perfil, e a migração fica marcada como concluída."""
    c = sqlite3.connect(":memory:")
    c.executescript("""
      PRAGMA foreign_keys=ON;
      CREATE TABLE usuarios(id INTEGER PRIMARY KEY, nome TEXT, ativo INTEGER);
      CREATE TABLE inventario_eventos(id INTEGER PRIMARY KEY);
      CREATE TABLE inventario_integrantes(evento_id INTEGER, nome TEXT);
    """)
    migrar(c)
    nomes = {r[0] for r in c.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    assert {"usuarios_funcoes", "inventario_comissao_usuarios", "migracoes_acesso"} <= nomes
    assert c.execute("SELECT count(*) FROM usuarios_funcoes").fetchone()[0] == 0
    assert c.execute("SELECT 1 FROM migracoes_acesso WHERE nome='fase5_funcoes'").fetchone()


def test_migracao_preserva_ids_hash_email_ativo():
    c = sqlite3.connect(":memory:")
    c.executescript("""
      PRAGMA foreign_keys=ON;
      CREATE TABLE usuarios(id INTEGER PRIMARY KEY, login TEXT, email TEXT, nome TEXT,
        senha_hash TEXT, perfil TEXT CHECK(perfil IN ('admin','operador','consulta','inventariante')),
        ativo INTEGER, trocar_senha INTEGER, falhas INTEGER, bloqueado_ate TEXT,
        criado_em TEXT, ultimo_acesso TEXT);
      CREATE TABLE inventario_eventos(id INTEGER PRIMARY KEY);
      CREATE TABLE inventario_integrantes(evento_id INTEGER, nome TEXT);
      INSERT INTO usuarios VALUES
        (7,'fulano','fulano@x.com','Fulano','hash-fictício-123','admin',1,0,0,NULL,'2026-01-01','2026-02-01'),
        (9,'ciclana',NULL,'Ciclana','hash-fictício-456','consulta',0,1,2,'2026-03-01','2026-01-01',NULL);
    """)
    migrar(c)
    linha = c.execute(
        "SELECT id, login, email, nome, senha_hash, ativo, trocar_senha, falhas, bloqueado_ate, "
        "criado_em, ultimo_acesso FROM usuarios WHERE id=7"
    ).fetchone()
    assert linha == (7, 'fulano', 'fulano@x.com', 'Fulano', 'hash-fictício-123', 1, 0, 0, None,
                      '2026-01-01', '2026-02-01')
    outra = c.execute("SELECT id, login, email, ativo FROM usuarios WHERE id=9").fetchone()
    assert outra == (9, 'ciclana', None, 0)
    assert c.execute("SELECT funcao FROM usuarios_funcoes WHERE usuario_id=7").fetchall() == [('admin',)]
    assert c.execute("SELECT funcao FROM usuarios_funcoes WHERE usuario_id=9").fetchall() == [('consulta',)]


class _ConexaoComFalha(sqlite3.Connection):
    """Conexão que injeta uma falha no INSERT da comissão, para testar o rollback da migração."""
    falha_ativa = False

    def execute(self, sql, *args, **kwargs):
        if self.falha_ativa and "INSERT INTO inventario_comissao_usuarios" in sql:
            raise sqlite3.OperationalError("falha injetada")
        return super().execute(sql, *args, **kwargs)


def test_migracao_rollback_por_erro_injetado_preserva_perfil_e_dados():
    c = sqlite3.connect(":memory:", factory=_ConexaoComFalha)
    c.executescript("""
      PRAGMA foreign_keys=ON;
      CREATE TABLE usuarios(id INTEGER PRIMARY KEY, nome TEXT, ativo INTEGER,
        perfil TEXT CHECK(perfil IN ('admin','operador','consulta','inventariante')));
      CREATE TABLE inventario_eventos(id INTEGER PRIMARY KEY);
      CREATE TABLE inventario_integrantes(evento_id INTEGER, nome TEXT);
      INSERT INTO usuarios VALUES(1,'Ana',1,'inventariante'),(2,'Ana',1,'inventariante'),
        (3,'Beto',1,'inventariante'),(4,'Cris',1,'operador');
      INSERT INTO inventario_eventos VALUES(1);
      INSERT INTO inventario_integrantes VALUES(1,'Ana'),(1,'Beto'),(1,'Cris');
    """)

    c.falha_ativa = True
    with pytest.raises(sqlite3.OperationalError):
        migrar(c)
    c.falha_ativa = False

    assert 'perfil' in [r[1] for r in c.execute('PRAGMA table_info(usuarios)')]
    assert c.execute("SELECT id, nome, ativo, perfil FROM usuarios ORDER BY id").fetchall() == [
        (1, 'Ana', 1, 'inventariante'), (2, 'Ana', 1, 'inventariante'),
        (3, 'Beto', 1, 'inventariante'), (4, 'Cris', 1, 'operador'),
    ]
    # BEGIN IMMEDIATE engloba até os CREATE TABLE: o rollback desfaz a transação inteira, então nem
    # as tabelas novas (nem o marcador) sobrevivem à falha injetada — nada de sucesso parcial.
    nomes = {r[0] for r in c.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    assert not nomes & {"migracoes_acesso", "usuarios_funcoes", "inventario_comissao_usuarios"}
