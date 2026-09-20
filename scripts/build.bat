@echo off
REM Gera dist\TermosCFC\TermosCFC.exe. Rode A PARTIR DA RAIZ do projeto (scripts\build.bat), dentro da venv: python -m venv .venv && .venv\Scripts\activate && pip install -r requirements.txt
pyinstaller --noconfirm --clean --onedir --windowed --name TermosCFC ^
  --add-data "templates;templates" --add-data "static;static" --add-data "timbrado.docx;." ^
  main.py
if errorlevel 1 exit /b 1
if not exist dist\TermosCFC\dados mkdir dist\TermosCFC\dados
echo.
echo Pronto: dist\TermosCFC\TermosCFC.exe  (copie termos.db para dist\TermosCFC\dados\ se quiser levar os dados)
