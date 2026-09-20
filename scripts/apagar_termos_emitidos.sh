#!/bin/bash
# Apaga TODOS os termos emitidos (e os bens de cada termo e os pedidos de emissão no SEI ligados a eles)
# do banco dados/termos.db — para limpar os testes antes de entrar em produção. A numeração dos termos
# (01/2026, 02/2026…) volta a começar do 01, pois é calculada a partir dos termos existentes.
# Não toca em bens, pessoas, cadastros, processos do SEI, inventário, usuários nem nos documentos já
# criados no SEI (esses continuam lá: exclua-os no próprio SEI, se for o caso).
#
#   ./apagar_termos_emitidos.sh          # mostra o que vai apagar e pede confirmação
#   ./apagar_termos_emitidos.sh --sim    # apaga sem perguntar
set -euo pipefail
cd "$(dirname "$0")/.."      # raiz do projeto
DB=dados/termos.db
SQLITE=/usr/bin/sqlite3

[ -f "$DB" ] || { echo "banco $DB não encontrado"; exit 1; }

ATIVOS=$($SQLITE "$DB" "SELECT count(*) FROM robo_pedidos WHERE tipo='sei' AND passo NOT IN ('concluido','erro')")
if [ "$ATIVOS" != "0" ]; then
    echo "há $ATIVOS pedido(s) de emissão no SEI em andamento; espere terminar (ou docker compose stop robo) e rode de novo"
    exit 1
fi

echo "Termos emitidos no banco:"
$SQLITE -header -column "$DB" "SELECT id, tipo, chave, numero_termo AS numero, unidade_sei AS unidade, documento_sei AS doc_sei, emitido_em FROM termos_emitidos ORDER BY id"
TERMOS=$($SQLITE "$DB" "SELECT count(*) FROM termos_emitidos")
PEDIDOS=$($SQLITE "$DB" "SELECT count(*) FROM robo_pedidos WHERE tipo='sei'")
echo
echo "Vai apagar: $TERMOS termo(s) emitido(s), seus bens e $PEDIDOS pedido(s) de emissão no SEI."
[ "$TERMOS" = "0" ] && [ "$PEDIDOS" = "0" ] && { echo "nada a apagar"; exit 0; }

if [ "${1:-}" != "--sim" ]; then
    read -r -p "Confirma? (digite APAGAR) " resposta
    [ "$resposta" = "APAGAR" ] || { echo "cancelado"; exit 1; }
fi

COPIA=dados/termos-antes-de-apagar-termos-$(date +%F-%H%M).db
$SQLITE "$DB" ".backup '$COPIA'"
echo "cópia do banco em $COPIA"

$SQLITE "$DB" <<'SQL'
BEGIN;
DELETE FROM robo_pedidos WHERE tipo = 'sei';
DELETE FROM termos_emitidos_bens;
DELETE FROM termos_emitidos;
COMMIT;
SQL
echo "apagados. Termos emitidos agora: $($SQLITE "$DB" "SELECT count(*) FROM termos_emitidos")"
