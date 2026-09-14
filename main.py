"""Programa de desktop: sobe o Flask numa thread e abre a janela. Fechar a janela encerra o programa.

Sem WebView2/WebKit disponível, abre o navegador padrão e fica servindo até Ctrl+C.
Porta ocupada = o programa já está aberto: abre o navegador nela e sai.
"""
import os
import socket
import sys
import threading
import webbrowser

import db
from app import app

PORTA = int(os.environ.get("TERMOS_PORTA", "5000"))
URL = f"http://127.0.0.1:{PORTA}"


def avisar(texto: str) -> None:
    """Mostra uma caixa de mensagem no Windows; no terminal, imprime."""
    if sys.platform == "win32":
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, texto, "Termos de Responsabilidade – CFC", 0x40)
    else:
        print(texto)


def porta_livre() -> bool:
    with socket.socket() as s:
        return s.connect_ex(("127.0.0.1", PORTA)) != 0


def servidor():
    app.run(host="127.0.0.1", port=PORTA, debug=False, use_reloader=False)


def main():
    if not porta_livre():
        avisar(f"A porta {PORTA} já está em uso — provavelmente o programa já está aberto. "
               "Vou abrir o navegador nele.")
        webbrowser.open(URL)
        return
    db.inicializar()
    threading.Thread(target=servidor, daemon=True).start()
    try:
        import webview
        webview.create_window("Termos de Responsabilidade – CFC", URL, width=1100, height=750)
        webview.start()
    except Exception:
        # Sem WebView2/WebKit: o programa continua rodando no navegador. No Windows, a caixa de
        # mensagem bloqueia até o usuário clicar OK — é o que fecha o programa nesse caso, já que
        # não há janela para fechar (o processo roda sem console sob --windowed).
        webbrowser.open(URL)
        if sys.platform == "win32":
            avisar("Não foi possível abrir a janela do programa (WebView2 ausente). Ele está rodando "
                   f"no navegador em {URL}. Clique em OK para encerrar o programa quando terminar.")
            return
        avisar("Não foi possível abrir a janela do programa (WebView2/WebKit ausente). Ele está rodando "
               f"no navegador em {URL}. Encerre com Ctrl+C quando terminar.")
        threading.Event().wait()  # mantém o servidor vivo até Ctrl+C


if __name__ == "__main__":
    main()
