# Servir o sistema na web (patrimonio.sistemascfc.org). O programa de desktop (main.py/build.bat) não usa este arquivo.
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 TERMOS_DADOS=/app/dados
WORKDIR /app

# Só o que o servidor precisa: pywebview e pyinstaller são do desktop.
RUN pip install --no-cache-dir flask openpyxl python-docx waitress Pillow boto3

COPY . .

EXPOSE 8000
# db.inicializar() cria a pasta de dados, copia o timbrado e aplica o esquema; depois o waitress serve o app.
CMD ["sh", "-c", "python -c 'import db; db.inicializar()' && waitress-serve --listen=0.0.0.0:8000 --threads=4 app:app"]
