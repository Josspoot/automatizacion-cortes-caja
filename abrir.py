#!/usr/bin/env python3
"""
Abre UNA ventana de Chromium (perfil guardado) con puerto de depuracion
y la deja abierta. El bot se conecta a esta ventana via CDP, asi NO se
cierra al terminar cada dia y no tienes que iniciar sesion de nuevo.

Lanza el Chromium directamente (no via Playwright) para que el puerto 9222
si quede a la escucha.

Uso:
    python abrir.py        # abre la ventana y la mantiene viva
"""
import subprocess
import time
import urllib.request
from pathlib import Path
from playwright.sync_api import sync_playwright

import config

BASE = Path(__file__).parent
PERFIL = config.PERFIL_NAVEGADOR
URL_MOV = config.URL_MOV
PUERTO = 9222


def chromium_path():
    with sync_playwright() as p:
        return p.chromium.executable_path


def puerto_listo():
    try:
        urllib.request.urlopen(f"http://127.0.0.1:{PUERTO}/json/version", timeout=1)
        return True
    except Exception:
        return False


def main():
    exe = chromium_path()
    proc = subprocess.Popen([
        exe,
        f"--remote-debugging-port={PUERTO}",
        f"--user-data-dir={PERFIL}",
        "--no-first-run",
        "--no-default-browser-check",
        "--start-maximized",          # abre al tamaño de tu pantalla
        URL_MOV,
    ])
    # esperar a que el puerto este listo
    for _ in range(30):
        if puerto_listo():
            break
        time.sleep(0.5)

    print("=" * 60)
    if puerto_listo():
        print("  Ventana abierta y conectada (puerto 9222) ✅")
    else:
        print("  (!) El puerto 9222 no respondio; revisa la ventana.")
    print("  - Inicia sesion si hace falta.")
    print("  - DEJA esta ventana abierta; el bot se conecta aqui.")
    print("  - Para terminar: cierra la ventana del navegador.")
    print("=" * 60)

    # mantener vivo mientras el navegador siga abierto
    try:
        proc.wait()
    except KeyboardInterrupt:
        proc.terminate()
    print("  Ventana cerrada. Fin.")


if __name__ == "__main__":
    main()
