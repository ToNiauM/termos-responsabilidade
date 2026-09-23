"""Apaga eventos de inventário inteiros, fotos do bucket incluídas — a parte que roda DENTRO do container `web`
(lá estão o banco em /app/dados e as credenciais R2_*). Chamado por scripts/zerar_inventario.sh, que lista,
pergunta e faz a cópia do banco; este arquivo não pergunta nada.

    python - [--copia PASTA] [--manter-fotos] ID [ID ...]   < scripts/zerar_inventario.py

Por evento: (1) com --copia, baixa TODAS as fotos (bens e sobras) para PASTA/<id>-<pasta do evento>/; se alguma
não baixar, para sem apagar nada; (2) apaga no bucket e depois no banco com inventario.excluir_evento — a mesma
função do "Excluir evento" da tela: se o bucket falhar, o banco não muda e dá para rodar de novo (apagar no
bucket um objeto que já não existe não é erro)."""
import argparse
import os
import sys

sys.path.insert(0, os.getcwd())      # no container o código fica na pasta de trabalho (/app)

import db          # noqa: E402
import fotos       # noqa: E402
import inventario  # noqa: E402


def chave_da_url(url):
    """Mesma regra de fotos.apagar: URL nova = base pública + chave; antiga tem "inventario/" no meio."""
    base = fotos._url("")
    if url.startswith(base):
        return url[len(base):] or None
    pos = url.find(fotos._PREFIXO_ANTIGO)
    return url[pos:] if pos >= 0 else None


def copiar_fotos(conn, evento, destino):
    """Baixa as fotos do evento para destino/. Devolve (baixadas, ausentes). Qualquer outro erro PROPAGA."""
    from botocore.exceptions import ClientError
    pasta = os.path.join(destino, f"{evento['id']}-{fotos.pasta(evento['nome'], evento['id'])}")
    os.makedirs(pasta, exist_ok=True)
    cliente, bucket = fotos._cliente(), os.environ["R2_BUCKET_NAME"]
    baixadas, ausentes = 0, []
    for url in inventario.urls_das_fotos(conn, evento["id"]):
        chave = chave_da_url(url)
        if not chave:
            ausentes.append(url)
            continue
        try:
            corpo = cliente.get_object(Bucket=bucket, Key=chave)["Body"].read()
        except ClientError as e:
            if e.response.get("Error", {}).get("Code") in ("NoSuchKey", "404"):
                ausentes.append(url)       # já não está no bucket: nada a copiar nem a apagar
                continue
            raise
        with open(os.path.join(pasta, chave.replace("/", "__")), "wb") as f:
            f.write(corpo)
        baixadas += 1
    return baixadas, ausentes


def apagar_no_bucket(url):
    """Como app_inventario._apagar_no_bucket: falha vira ErroDeNegocio e o banco não muda."""
    try:
        fotos.apagar(url)
    except Exception as e:
        raise db.ErroDeNegocio(f"não foi possível apagar a foto {url} no bucket ({e}); o banco não foi alterado")


def main(argv=None):
    p = argparse.ArgumentParser()
    p.add_argument("ids", nargs="+", type=int)
    p.add_argument("--copia")
    p.add_argument("--manter-fotos", action="store_true")
    p.add_argument("--dono", help="UID:GID do usuário do servidor, dono da cópia das fotos (o container roda como root)")
    a = p.parse_args(argv)
    if not a.manter_fotos and not fotos.configurado():
        print("ERRO: bucket de fotos não configurado neste ambiente; use --manter-fotos ou rode no container web")
        return 1
    conn = db.conectar()
    try:
        for evento_id in a.ids:
            e = conn.execute("SELECT * FROM inventario_eventos WHERE id = ?", (evento_id,)).fetchone()
            if not e:
                print(f"evento {evento_id}: não existe (já apagado?) — pulado")
                continue
            e = dict(e)
            n = len(inventario.urls_das_fotos(conn, evento_id))
            if a.copia and not a.manter_fotos and n:
                baixadas, ausentes = copiar_fotos(conn, e, a.copia)
                print(f"evento {evento_id} ({e['nome']}): {baixadas} foto(s) copiada(s)"
                      + (f", {len(ausentes)} já ausente(s) no bucket" if ausentes else ""))
            inventario.excluir_evento(conn, evento_id, e["nome"], apagar=None if a.manter_fotos else apagar_no_bucket)
            print(f"evento {evento_id} ({e['nome']}): apagado"
                  + ("" if a.manter_fotos else f", {n} foto(s) removida(s) do bucket"))
    except db.ErroDeNegocio as erro:
        print(f"ERRO: {erro}")
        return 1
    finally:
        conn.close()
        if a.copia and a.dono and os.path.isdir(a.copia):
            uid, gid = (int(x) for x in a.dono.split(":"))
            for raiz, pastas, arquivos in os.walk(a.copia):
                for nome in [raiz] + [os.path.join(raiz, x) for x in pastas + arquivos]:
                    os.chown(nome, uid, gid)
    return 0


if __name__ == "__main__":
    sys.exit(main())
