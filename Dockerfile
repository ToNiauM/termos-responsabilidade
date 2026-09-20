# Servir o sistema na web (patrimonio.sistemascfc.org). O programa de desktop (main.py/build.bat) não usa este arquivo.
#
# Dois alvos:
#   web  -> a imagem do site (Flask + waitress), como sempre foi.
#   robo -> FROM web, com Playwright e xlrd por cima, para atender a fila robo_pedidos
#           (emissão no SEI e atualização com o SPW). O site nunca ganha Playwright;
#           só o alvo robo (compose.yml usa build.target para escolher cada um).
FROM python:3.12-slim AS web

# TERMOS_LOGIN=1: a imagem sempre exige login; o compose pode sobrescrever se algum dia precisar do modo desktop.
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 TERMOS_DADOS=/app/dados TERMOS_LOGIN=1
WORKDIR /app

# Só o que o servidor precisa: pywebview e pyinstaller são do desktop.
RUN pip install --no-cache-dir flask openpyxl python-docx waitress Pillow boto3

COPY . .

EXPOSE 8000
# db.inicializar() cria a pasta de dados, copia o timbrado e aplica o esquema; depois o waitress serve o app.
CMD ["sh", "-c", "python -c 'import db; db.inicializar()' && waitress-serve --listen=0.0.0.0:8000 --threads=4 app:app"]

# Alvo do trabalhador da fila (atender_pedidos.py): herda tudo do alvo web e
# acrescenta Playwright (SEI e SPW) e xlrd (planilha exportada pelo SPW).
FROM web AS robo

RUN pip install --no-cache-dir playwright xlrd
RUN playwright install --with-deps chromium

CMD ["python", "atender_pedidos.py"]
