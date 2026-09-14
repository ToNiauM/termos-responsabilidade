"""Programa de desktop: sobe o Flask numa thread e abre a janela. Fechar a janela encerra o programa.

Sem WebView2/WebKit disponível, abre o navegador padrão e fica servindo até Ctrl+C.
Porta ocupada = o programa já está aberto: abre o navegador nela e sai.
"""
import os
import socket
import threading
import webbrowser

import db
from app import app

PORTA = int(os.environ.get("TERMOS_PORTA", "5000"))
URL = f"http://127.0.0.1:{PORTA}"


def porta_livre() -> bool:
    with socket.socket() as s:
        return s.connect_ex(("127.0.0.1", PORTA)) != 0


def servidor():
    app.run(host="127.0.0.1", port=PORTA, debug=False, use_reloader=False)


def main():
    if not porta_livre():
        webbrowser.open(URL)
        return
    db.inicializar()
    threading.Thread(target=servidor, daemon=True).start()
    try:
        import webview
        webview.create_window("Termos de Responsabilidade – CFC", URL, width=1100, height=750)
        webview.start()
    except Exception:
        webbrowser.open(URL)
        threading.Event().wait()  # mantém o servidor vivo até Ctrl+C


if __name__ == "__main__":
    main()
