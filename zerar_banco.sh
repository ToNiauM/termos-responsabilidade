#!/bin/bash
# Zera o HISTÓRICO de dados/termos.db para começar a produção do zero, preservando os cadastros e a carga atual:
#   mantém: bens e atribuicoes (o bem 14359 continua com quem está), pessoas, responsaveis (centros de custo),
#           localizacoes, processos_sei, textos, usuarios / usuarios_funcoes (e seus acessos ao SEI/SPW),
#           migracoes_acesso — e o inventário, salvo com --inventario
#   apaga:  importacoes / importacoes_mudancas / robo_execucoes (histórico das cargas: o próximo upload ou robô
#           do SPW compara com a base atual, não com uma execução anterior), termos_emitidos / termos_emitidos_bens /
#           robo_pedidos (emissões e fila); com --inventario, também inventario_* (eventos, leituras, fotos,
#           sobras, comissão — as fotos no bucket R2 ficam lá)
#
#   ./zerar_banco.sh                     # mostra os totais e pede confirmação
#   ./zerar_banco.sh --sim               # zera sem perguntar
#   ./zerar_banco.sh --inventario        # inclui o inventário (pode combinar com --sim)
set -euo pipefail
cd "$(dirname "$0")"
DB=dados/termos.db
SQLITE=/usr/bin/sqlite3

[ -f "$DB" ] || { echo "banco $DB não encontrado"; exit 1; }

ATIVOS=$($SQLITE "$DB" "SELECT count(*) FROM robo_pedidos WHERE passo NOT IN ('concluido','erro')")
if [ "$ATIVOS" != "0" ]; then
    echo "há $ATIVOS pedido(s) do robô em andamento; espere terminar (ou docker compose stop robo) e rode de novo"
    exit 1
fi

SIM=0; INVENTARIO=0
for arg in "$@"; do
    case "$arg" in
        --sim) SIM=1 ;;
        --inventario) INVENTARIO=1 ;;
        *) echo "opção desconhecida: $arg"; exit 1 ;;
    esac
done

APAGAR="importacoes_mudancas robo_execucoes importacoes termos_emitidos_bens robo_pedidos termos_emitidos"
MANTER="bens atribuicoes pessoas responsaveis localizacoes processos_sei textos usuarios usuarios_funcoes migracoes_acesso"
INV="inventario_fotos inventario_leituras inventario_sobras inventario_salas inventario_integrantes inventario_comissao_usuarios
inventario_bens_encerrados inventario_eventos"
if [ "$INVENTARIO" = "1" ]; then APAGAR="$APAGAR $INV"; else MANTER="$MANTER $INV"; fi

echo "Vai APAGAR:"
for t in $APAGAR; do printf "  %-30s %s\n" "$t" "$($SQLITE "$DB" "SELECT count(*) FROM $t")"; done
echo "Vai MANTER:"
for t in $MANTER; do printf "  %-30s %s\n" "$t" "$($SQLITE "$DB" "SELECT count(*) FROM $t")"; done
echo

if [ "$SIM" != "1" ]; then
    read -r -p "Confirma? (digite ZERAR) " resposta
    [ "$resposta" = "ZERAR" ] || { echo "cancelado"; exit 1; }
fi

COPIA=dados/termos-antes-de-zerar-$(date +%F-%H%M).db
$SQLITE "$DB" ".backup '$COPIA'"
echo "cópia do banco em $COPIA"

{
    echo "BEGIN;"
    for t in $APAGAR; do echo "DELETE FROM $t;"; done      # a ordem acima respeita as chaves estrangeiras
    echo "COMMIT;"
    echo "VACUUM;"
} | $SQLITE "$DB"
echo "zerado. bens: $($SQLITE "$DB" 'SELECT count(*) FROM bens'); atribuições: $($SQLITE "$DB" 'SELECT count(*) FROM atribuicoes'); termos emitidos: $($SQLITE "$DB" 'SELECT count(*) FROM termos_emitidos'); importações: $($SQLITE "$DB" 'SELECT count(*) FROM importacoes')"
