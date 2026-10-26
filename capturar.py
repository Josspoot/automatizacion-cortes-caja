#!/usr/bin/env python3
"""Captura un corte en LAS DOS plataformas de un jalon:
  1) WEB   el ERP web (bot.py --auto)
  2) EXCEL el libro .xlsm con macros (bot_excel.py --real)

Uso:
    python capturar.py 12          # busca "CC 12-MM-AAAA.pdf" (mes actual);
                                   # si no esta en cortes_pdf/ lo copia de ~/Documents
    python capturar.py 10 11 12    # varios dias seguidos
    python capturar.py "cortes_pdf/CC 12-07-2026.pdf"    # ruta directa
    --sin-web / --sin-excel        # saltar una plataforma

El Excel tiene candado antiduplicados: si el dia ya esta en el libro, se salta
solo (no hace falta pensarlo al re-correr).
"""
import argparse
import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path

BASE = Path(__file__).parent
PY = BASE / ".venv" / "bin" / "python"
DOCS = Path.home() / "Documents"


def encontrar_pdf(arg):
    p = Path(arg)
    if p.suffix.lower() == ".pdf":
        p = p if p.is_absolute() else BASE / p
        return p if p.exists() else None
    try:
        dia = int(arg)
    except ValueError:
        return None
    hoy = date.today()
    nombre = f"CC {dia:02d}-{hoy.month:02d}-{hoy.year}.pdf"
    destino = BASE / "cortes_pdf" / nombre
    if destino.exists():
        return destino
    if (DOCS / nombre).exists():
        shutil.copy(DOCS / nombre, destino)
        print(f"  (copié {nombre} de Documents a cortes_pdf/)")
        return destino
    # otro mes: el mas reciente que exista con ese dia
    cands = sorted((BASE / "cortes_pdf").glob(f"CC {dia:02d}-*.pdf"))
    return cands[-1] if cands else None


def correr(script, pdf, extra):
    r = subprocess.run([str(PY), str(BASE / script), str(pdf)] + extra, cwd=BASE)
    return r.returncode


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("dias", nargs="+", help="numero(s) de dia del corte o ruta(s) a PDF")
    ap.add_argument("--sin-web", action="store_true")
    ap.add_argument("--sin-excel", action="store_true")
    args = ap.parse_args()

    resumen = []
    for d in args.dias:
        pdf = encontrar_pdf(d)
        if not pdf:
            print(f"✗ No encontré el PDF para '{d}' (ni en cortes_pdf/ ni en Documents)")
            resumen.append((str(d), "SIN PDF", "SIN PDF"))
            continue

        print(f"\n{'='*22} {pdf.name} {'='*22}")
        web = excel = "no pedido"
        if not args.sin_web:
            print("\n--- 1/2 WEB (bot.py --auto) ---")
            web = "OK" if correr("bot.py", pdf, ["--auto"]) == 0 else "REVISAR"
        if not args.sin_excel:
            print("\n--- 2/2 EXCEL (bot_excel.py --real) ---")
            rc = correr("bot_excel.py", pdf, ["--real"])
            excel = "OK" if rc == 0 else ("ya estaba" if rc == 3 else "REVISAR")
        resumen.append((pdf.name, web, excel))

    print(f"\n{'='*18} RESUMEN {'='*18}")
    for nombre, web, excel in resumen:
        print(f"  {nombre:24}  web: {web:10}  excel: {excel}")


if __name__ == "__main__":
    main()
