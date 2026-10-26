#!/usr/bin/env python3
"""Configuracion centralizada: lee variables de entorno (.env) para que el
repo no tenga nada especifico de un negocio en particular.

Copia .env.example a .env y llena tus propios valores antes de usar el bot.
"""
import os
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

BASE = Path(__file__).parent

# URL base del ERP / plataforma web donde se capturan los movimientos
URL_BASE = os.environ.get("CORTES_URL_BASE", "https://tu-erp.ejemplo.com")
URL_MOV = f"{URL_BASE}/movimientos"

# Carpeta donde Playwright guarda la sesion de navegador ya logueada
# (se crea sola la primera vez que inicias sesion manualmente)
PERFIL_NAVEGADOR = Path(os.environ.get("CORTES_PERFIL_NAVEGADOR", str(BASE / ".perfil_navegador")))

# Ruta al archivo Excel/VBA donde tambien se registra el doble-registro
XLSM_PATH = Path(os.environ.get(
    "CORTES_XLSM_PATH",
    str(Path.home() / "Documents" / "Corte de Caja.xlsm"),
))

# Carpeta donde caen los PDF de "corte de caja" a procesar
CORTES_PDF_DIR = Path(os.environ.get("CORTES_PDF_DIR", str(BASE / "cortes_pdf")))

# Archivos de conocimiento / registro (se generan o personalizan localmente,
# nunca se suben al repo con datos reales)
CONOCIMIENTO_PATH = Path(os.environ.get("CORTES_CONOCIMIENTO_PATH", str(BASE / "conocimiento.json")))
CONOCIMIENTO_EXCEL_PATH = Path(os.environ.get("CORTES_CONOCIMIENTO_EXCEL_PATH", str(BASE / "conocimiento_excel.json")))
CATEGORIAS_PATH = Path(os.environ.get("CORTES_CATEGORIAS_PATH", str(BASE / "categorias.json")))
REGISTRO_PATH = Path(os.environ.get("CORTES_REGISTRO_PATH", str(BASE / "registro.json")))
REGISTRO_REPORTE_PATH = Path(os.environ.get("CORTES_REGISTRO_REPORTE_PATH", str(BASE / "registro_reporte.json")))
PENDIENTES_PATH = Path(os.environ.get("CORTES_PENDIENTES_PATH", str(BASE / "PENDIENTES.txt")))
