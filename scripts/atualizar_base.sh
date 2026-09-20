#!/bin/bash
# Atualiza a base de bens a partir do SPW, na hora (fora do cron).
# Uso: ./atualizar_base.sh          -> roda o robô contra dados/termos.db e mostra o resultado
#      ./atualizar_base.sh --teste  -> roda contra uma CÓPIA em dados/robo-teste, sem tocar na base real
# O robô roda dentro do container `robo` (docker compose exec); o script só entra nele.
# A mesma linha vai para dados/robo_spw.log. Sai com 0 (importado/sem mudança) ou 1 (erro).
set -uo pipefail
cd "$(dirname "$(readlink -f "$0")")/.."

if ! docker compose ps --status running --services 2>/dev/null | grep -qx robo; then
  echo "O container robo não está rodando. Suba com: docker compose up -d" >&2
  exit 1
fi

if [ "${1:-}" = "--teste" ]; then
  mkdir -p dados/robo-teste && cp dados/termos.db dados/robo-teste/termos.db
  echo "Teste numa cópia: dados/robo-teste/termos.db (a base real não muda)"
  docker compose exec -T -e TERMOS_DADOS=/app/dados/robo-teste robo python importar_spw.py
  exit $?
fi

# Mesmo lock do trabalhador (atender_pedidos.py, agora dentro do container): nunca dois
# processos no banco e no navegador ao mesmo tempo. dados/robo.lock é o mesmo arquivo dos
# dois lados (volume ./dados:/app/dados), então o flock vale para host e container.
exec 9>dados/robo.lock
if ! flock -w 900 9; then
  echo "Outra atualização ou emissão está em andamento há mais de 15 min; tente de novo." >&2
  exit 1
fi

echo "Atualizando a base a partir do SPW..."
docker compose exec -T robo python importar_spw.py | tee -a dados/robo_spw.log
codigo=${PIPESTATUS[0]}
if [ "$codigo" -ne 0 ]; then
  echo "Falhou. Veja dados/spw/erro.png e a tabela de execuções em Atualizar base." >&2
fi
exit "$codigo"
