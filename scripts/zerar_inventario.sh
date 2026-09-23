#!/bin/bash
# Apaga eventos de inventário INTEIROS: o evento e tudo o que é dele (salas do escopo, comissão, leituras, sobras,
# bens encerrados) E AS FOTOS NO BUCKET R2. Mesmo efeito do "Excluir evento" da tela, pelo terminal e para vários
# de uma vez. Não toca em bens, pessoas, cadastros, termos, usuários nem em outros eventos.
#
# Segurança, nesta ordem: (1) lista e pede confirmação; (2) cópia do banco em dados/; (3) cópia de TODAS as fotos
# dos eventos em /opt/backups/termos/fotos-inventario-<data>/ — se alguma não baixar, para sem apagar nada;
# (4) apaga no bucket e depois no banco (se o bucket falhar, o banco não muda; pode rodar de novo).
# A parte que fala com o bucket roda dentro do container `web` (scripts/zerar_inventario.py).
#
#   ./zerar_inventario.sh                  # lista e pergunta qual(is) apagar (ex.: 3  ou  1,3  ou  todos)
#   ./zerar_inventario.sh 3                # apaga o evento 3
#   ./zerar_inventario.sh --ALL            # apaga TODOS os eventos (também --all / --todos)
#   opções: --sim (sem confirmação) · --manter-fotos (só o banco; fotos ficam no bucket)
#           --sem-copia (não baixa a cópia das fotos antes de apagar)
set -euo pipefail
cd "$(dirname "$0")/.."      # raiz do projeto
DB=dados/termos.db
SQLITE=/usr/bin/sqlite3
BACKUPS=/opt/backups/termos

[ -f "$DB" ] || { echo "banco $DB não encontrado"; exit 1; }

SIM=0; TODOS=0; MANTER=0; COPIA=1; ESCOLHA=""
for arg in "$@"; do
    case "$arg" in
        --sim) SIM=1 ;;
        --ALL|--all|--todos) TODOS=1 ;;
        --manter-fotos) MANTER=1 ;;
        --sem-copia) COPIA=0 ;;
        -*) echo "opção desconhecida: $arg"; exit 1 ;;
        *) ESCOLHA="$ESCOLHA $arg" ;;
    esac
done

TOTAL=$($SQLITE "$DB" "SELECT count(*) FROM inventario_eventos")
[ "$TOTAL" = "0" ] && { echo "nenhum evento de inventário no banco"; exit 0; }

echo "Eventos de inventário:"
$SQLITE -header -column "$DB" "
    SELECT e.id, e.nome,
           CASE WHEN e.encerrado_em IS NOT NULL THEN 'encerrado'
                WHEN e.suspenso_em IS NOT NULL THEN 'fechado' ELSE 'ABERTO' END AS estado,
           substr(e.aberto_em, 1, 10) AS inicio,
           (SELECT count(*) FROM inventario_salas s WHERE s.evento_id = e.id) AS salas,
           (SELECT count(*) FROM inventario_leituras l WHERE l.evento_id = e.id) AS leituras,
           (SELECT count(*) FROM inventario_fotos f WHERE f.evento_id = e.id)
             + (SELECT count(*) FROM inventario_sobras o WHERE o.evento_id = e.id AND coalesce(o.foto_url, '') <> '') AS fotos,
           (SELECT count(*) FROM inventario_sobras o WHERE o.evento_id = e.id) AS sobras
    FROM inventario_eventos e ORDER BY e.id"
echo

if [ "$TODOS" = "0" ] && [ -z "${ESCOLHA// /}" ]; then
    read -r -p "Qual evento apagar? (número, vários separados por vírgula, ou 'todos'; Enter cancela) " ESCOLHA
    [ -n "${ESCOLHA// /}" ] || { echo "cancelado"; exit 1; }
fi
if [ "$TODOS" = "1" ] || [ "$(echo "$ESCOLHA" | tr -d ' ' | tr 'A-Z' 'a-z')" = "todos" ]; then
    IDS=$($SQLITE "$DB" "SELECT id FROM inventario_eventos ORDER BY id")
else
    IDS=""
    for id in $(echo "$ESCOLHA" | tr ',' ' '); do
        [[ "$id" =~ ^[0-9]+$ ]] || { echo "número inválido: $id"; exit 1; }
        [ "$($SQLITE "$DB" "SELECT count(*) FROM inventario_eventos WHERE id = $id")" = "1" ] || { echo "evento $id não existe"; exit 1; }
        IDS="$IDS $id"
    done
fi
IDS=$(echo $IDS)                 # normaliza espaços
LISTA=$(echo "$IDS" | tr ' ' ',')

FOTOS=$($SQLITE "$DB" "SELECT (SELECT count(*) FROM inventario_fotos WHERE evento_id IN ($LISTA))
                              + (SELECT count(*) FROM inventario_sobras WHERE evento_id IN ($LISTA) AND coalesce(foto_url, '') <> '')")
echo "Vai APAGAR:"
$SQLITE "$DB" "SELECT '  evento ' || id || ' - ' || nome || CASE WHEN encerrado_em IS NULL AND suspenso_em IS NULL THEN '   <-- ABERTO (em uso?)' ELSE '' END
               FROM inventario_eventos WHERE id IN ($LISTA) ORDER BY id"
echo "  leituras: $($SQLITE "$DB" "SELECT count(*) FROM inventario_leituras WHERE evento_id IN ($LISTA)")"
if [ "$MANTER" = "1" ]; then
    echo "  fotos: $FOTOS ficam no bucket (--manter-fotos)"
else
    echo "  fotos: $FOTOS, APAGADAS DO BUCKET$([ "$COPIA" = "1" ] && [ "$FOTOS" != "0" ] && echo " (antes, cópia em $BACKUPS/)")"
fi
echo

if [ "$SIM" != "1" ]; then
    read -r -p "Confirma? Não há desfazer além das cópias. (digite APAGAR) " resposta
    [ "$resposta" = "APAGAR" ] || { echo "cancelado"; exit 1; }
fi

DATA=$(date +%F-%H%M%S)
COPIA_DB=dados/termos-antes-de-zerar-inventario-$DATA.db
$SQLITE "$DB" ".backup '$COPIA_DB'"
echo "cópia do banco em $COPIA_DB"

ARGS="$IDS"
[ "$MANTER" = "1" ] && ARGS="--manter-fotos $ARGS"
PASTA_TMP=""
if [ "$MANTER" = "0" ] && [ "$COPIA" = "1" ] && [ "$FOTOS" != "0" ]; then
    PASTA_TMP=fotos-inventario-$DATA            # dentro de dados/ (o container enxerga), depois vai para $BACKUPS
    ARGS="--copia /app/dados/$PASTA_TMP --dono $(id -u):$(id -g) $ARGS"
fi

# O Python roda no container web: lá estão /app/dados (mesmo banco) e as credenciais do bucket.
set +e
docker compose exec -T web python - $ARGS < scripts/zerar_inventario.py
RESULTADO=$?
set -e

if [ -n "$PASTA_TMP" ] && [ -d "dados/$PASTA_TMP" ]; then
    mv "dados/$PASTA_TMP" "$BACKUPS/" && echo "cópia das fotos em $BACKUPS/$PASTA_TMP ($(find "$BACKUPS/$PASTA_TMP" -type f | wc -l) arquivo(s))"
fi
[ "$RESULTADO" = "0" ] || { echo "parou com erro (veja acima); o que não foi apagado continua no banco"; exit "$RESULTADO"; }
echo "pronto. Eventos de inventário agora: $($SQLITE "$DB" "SELECT count(*) FROM inventario_eventos")"
