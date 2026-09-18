#!/bin/bash
# Atualiza a base de bens a partir do SPW, na hora (fora do cron).
# Uso: ./atualizar_base.sh          -> roda o robô contra dados/termos.db e mostra o resultado
#      ./atualizar_base.sh --teste  -> roda contra uma CÓPIA em /tmp/robo-teste, sem tocar na base real
# A mesma linha vai para dados/robo_spw.log. Sai com 0 (importado/sem mudança) ou 1 (erro).
set -uo pipefail
cd "$(dirname "$(readlink -f "$0")")"

if [ ! -x .venv-robo/bin/python ]; then
  echo "Falta o venv do robô. Instale: python3 -m venv .venv-robo && .venv-robo/bin/pip install -r requirements-robo.txt && .venv-robo/bin/playwright install --with-deps chromium" >&2
  exit 1
fi

if [ "${1:-}" = "--teste" ]; then
  mkdir -p /tmp/robo-teste && cp dados/termos.db /tmp/robo-teste/termos.db
  echo "Teste numa cópia: /tmp/robo-teste/termos.db (a base real não muda)"
  TERMOS_DADOS=/tmp/robo-teste .venv-robo/bin/python importar_spw.py
  exit $?
fi

echo "Atualizando a base a partir do SPW..."
.venv-robo/bin/python importar_spw.py | tee -a dados/robo_spw.log
codigo=${PIPESTATUS[0]}
if [ "$codigo" -ne 0 ]; then
  echo "Falhou. Veja dados/spw/erro.png e a tabela de execuções em Atualizar base." >&2
fi
exit "$codigo"
