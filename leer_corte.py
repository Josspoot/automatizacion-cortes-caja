#!/usr/bin/env python3
"""
Lector de Corte Caja Diario (Elektra Moto) -> lista lista para capturar.

Lee el PDF, reconstruye la tabla de movimientos por coordenadas de palabras,
aplica las reglas de mapeo (cliente / proveedor / categoria / forma de pago)
y separa los movimientos en: CAPTURAR, NO REGISTRAR y PENDIENTES (avisar).

Uso:
    python leer_corte.py "cortes_pdf/CC 10-06-2026.pdf"
    python leer_corte.py            # procesa todos los PDF de cortes_pdf/
"""
import sys
import re
import glob
import json
import unicodedata
from pathlib import Path
from dataclasses import dataclass

import pdfplumber

sys.path.insert(0, str(Path(__file__).parent))
import config

# ---------------------------------------------------------------------------
# CEREBRO: conocimiento.json (clientes / proveedores / categorias aprendidas)
# ---------------------------------------------------------------------------
CONOC_PATH = Path(__file__).parent / "conocimiento.json"


def cargar_conocimiento():
    with open(CONOC_PATH, encoding="utf-8") as f:
        return json.load(f)


CONOC = cargar_conocimiento()

# ---------------------------------------------------------------------------
# BASE DE EJEMPLOS: movimientos ya capturados en la plataforma (ejemplos.json,
# se genera con leer_historico.py). Sirve para clasificar por PARECIDO.
# ---------------------------------------------------------------------------
EJEMPLOS_PATH = Path(__file__).parent / "ejemplos.json"
_STOP = {"DE", "DEL", "LA", "EL", "LOS", "LAS", "Y", "CON", "SIN", "PARA",
         "POR", "REM", "SER", "UNIDADES", "PAGO", "COMPRA"}


def _tokens(texto):
    """Palabras significativas del concepto (sin acentos, numeros ni folios)."""
    t = unicodedata.normalize("NFD", texto.upper())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    palabras = re.findall(r"[A-ZÑ]{3,}", t)
    return {p for p in palabras if p not in _STOP}


def cargar_ejemplos():
    try:
        ejemplos = json.loads(EJEMPLOS_PATH.read_text(encoding="utf-8"))
    except Exception:
        return []
    for e in ejemplos:
        e["_toks"] = _tokens(e.get("concepto", ""))
    return [e for e in ejemplos if e["_toks"]]


EJEMPLOS = cargar_ejemplos()


def buscar_parecido(concepto, tipo):
    """Ejemplo mas parecido al concepto (mismo tipo). Devuelve (ejemplo, score)
    o (None, 0). Score = proporcion de palabras compartidas (0..1)."""
    toks = _tokens(concepto)
    seg = concepto.split(" - ")[0].strip().upper()
    mejor, mejor_score = None, 0.0
    for e in EJEMPLOS:
        if e.get("tipo") != tipo or not e.get("cat1"):
            continue
        inter = toks & e["_toks"]
        if not inter:
            continue
        score = len(inter) / max(1, min(len(toks), len(e["_toks"])))
        # bono si el primer segmento coincide (mismo origen)
        if e["concepto"].split(" - ")[0].strip().upper() == seg:
            score += 0.25
        if score > mejor_score:
            mejor, mejor_score = e, score
    # umbral: minimo 60% de coincidencia y 2+ palabras compartidas
    if mejor and mejor_score >= 0.6 and len(toks & mejor["_toks"]) >= 2:
        return mejor, min(mejor_score, 1.0)
    return None, 0.0

# ---------------------------------------------------------------------------
# Limites de columnas (coordenada x0 de cada palabra). Calibrados con los PDF.
# ---------------------------------------------------------------------------
COL_NUM      = (0,   100)   # N°
COL_FECHA    = (100, 155)   # FECHA
COL_BANCO    = (155, 192)   # BANCO (EFECTIVO / AZTECA)
COL_CONCEPTO = (192, 352)   # CONCEPTO
COL_CAJA     = (352, 378)   # codigo de caja (001/002)
COL_SUCURSAL = (378, 425)   # SUCURSAL
COL_MONTO    = (425, 999)   # INGRESO / GASTO
INGRESO_GASTO_CORTE = 462   # x0 del "$": menor -> ingreso, mayor -> gasto

# ---------------------------------------------------------------------------
# REGLAS DE MAPEO  (editable: aqui se afinan clientes / proveedores / categorias)
# ---------------------------------------------------------------------------
# Conceptos que NO se registran (con su motivo)
NO_REGISTRAR = {
    "DEPOSITO EN EFECTIVO": "Deposito/ventas en efectivo: no se registra",
}
# Conceptos que son solo referencia de cuadre (saldo inicial)
REFERENCIA = ["INICIO DIA"]
# Conceptos VARIABLES: su proveedor/categoria depende del contenido -> manual
VARIABLES = {k for k in CONOC.get("variables", {}) if not k.startswith("_")}


@dataclass
class Movimiento:
    n: str
    fecha: str
    banco: str            # EFECTIVO / AZTECA (texto)
    concepto: str
    sucursal: str
    importe: float
    es_ingreso: bool
    caja: str = ""        # codigo de caja: 001=Efectivo, 002=Azteca
    # resultado del mapeo:
    accion: str = ""            # CAPTURAR / NO_REGISTRAR / REFERENCIA / PENDIENTE
    tipo: str = ""              # INGRESO / GASTO / DEPOSITO
    contacto: str = ""          # cliente o proveedor
    cat1: str = ""              # Categoria Principal (Nivel 1)
    cat2: str = ""              # Categoria (Nivel 2)
    cat3: str = ""              # Subcategoria (Nivel 3)
    forma_pago: str = ""        # Efectivo / Transferencia
    cuenta: str = ""            # Banco Azteca (si transferencia)
    nota: str = ""              # motivo si es pendiente
    sugerido: bool = False      # True si vino del matcher de ejemplos (verificar)

    @property
    def segmento(self) -> str:
        return self.concepto.split(" - ")[0].strip().upper()


# ---------------------------------------------------------------------------
# EXTRACCION
# ---------------------------------------------------------------------------
def _col(x0):
    if COL_NUM[0]      <= x0 < COL_NUM[1]:      return "num"
    if COL_FECHA[0]    <= x0 < COL_FECHA[1]:    return "fecha"
    if COL_BANCO[0]    <= x0 < COL_BANCO[1]:    return "banco"
    if COL_CONCEPTO[0] <= x0 < COL_CONCEPTO[1]: return "concepto"
    if COL_CAJA[0]     <= x0 < COL_CAJA[1]:     return "caja"
    if COL_SUCURSAL[0] <= x0 < COL_SUCURSAL[1]: return "sucursal"
    return "monto"


def leer_pdf(ruta):
    pdf = pdfplumber.open(ruta)
    page = pdf.pages[0]
    words = page.extract_words(use_text_flow=True, keep_blank_chars=False)

    # Agrupar palabras por renglon (coordenada 'top', tolerancia 3px)
    filas = {}
    for w in words:
        if w["x0"] > 530 or not (128 < w["top"] < 296):
            continue
        clave = None
        for t in filas:
            if abs(t - w["top"]) <= 3:
                clave = t
                break
        if clave is None:
            clave = w["top"]
            filas[clave] = []
        filas[clave].append(w)

    movimientos = []
    for top in sorted(filas):
        cols = {"num": [], "fecha": [], "banco": [], "concepto": [],
                "caja": [], "sucursal": [], "monto": []}
        for w in sorted(filas[top], key=lambda w: w["x0"]):
            cols[_col(w["x0"])].append(w)

        num = " ".join(w["text"] for w in cols["num"]).strip()
        fecha = " ".join(w["text"] for w in cols["fecha"]).strip()
        if not re.match(r"^\d+/$", num) or not fecha:
            continue  # renglon vacio

        banco = " ".join(w["text"] for w in cols["banco"]).strip()
        caja = " ".join(w["text"] for w in cols["caja"]).strip()
        concepto = " ".join(w["text"] for w in cols["concepto"]).strip()
        sucursal = " ".join(w["text"] for w in cols["sucursal"]).strip()

        # "INICIO DIA EFECTIVO/BANCO" no tiene columna BANCO; va en concepto
        if not concepto and banco:
            concepto, banco = banco, ""

        if not cols["monto"]:
            continue
        primer_monto_x = min(w["x0"] for w in cols["monto"])
        es_ingreso = primer_monto_x < INGRESO_GASTO_CORTE
        texto_monto = " ".join(w["text"] for w in cols["monto"])
        m = re.search(r"([\d,]+\.\d{2})", texto_monto)
        if not m:
            continue
        importe = float(m.group(1).replace(",", ""))

        movimientos.append(Movimiento(
            n=num, fecha=fecha, banco=banco, concepto=concepto,
            sucursal=sucursal, importe=importe, es_ingreso=es_ingreso, caja=caja,
        ))

    # Fecha del documento y total inicio dia (referencia de cuadre)
    fecha_doc = movimientos[0].fecha if movimientos else "?"
    total_inicio = _buscar_total_inicio(page)
    pdf.close()
    return movimientos, fecha_doc, total_inicio


def _buscar_total_inicio(page):
    for t in page.extract_tables():
        for fila in t:
            celdas = [c for c in fila if c]
            if celdas and "TOTAL INICIO DIA" in str(celdas[0]).upper():
                m = re.search(r"([\d,]+\.\d{2})", str(celdas[-1]))
                if m:
                    return float(m.group(1).replace(",", ""))
    return None


# ---------------------------------------------------------------------------
# REGLAS
# ---------------------------------------------------------------------------
def _detectar_nomina(concepto):
    """Si el concepto menciona nomina, deduce CUAL (por el texto, no por el
    primer segmento: 'DISPERSIÓN - NOMINA COMERCIAL - ROMARIO' -> Comercial)."""
    c = concepto.upper()
    if "NOMINA" not in c:
        return None
    if "COMERCIAL" in c:
        return "Nomina Comercial"
    if "ADMIN" in c:
        return "Nomina Administrativa"
    if "CESEM" in c or "PRODUCTIVA" in c or "TALLER" in c:
        return "Nomina CESEM Taller"
    return "Nomina Administrativa"   # dispersion de nomina sin tipo -> admin


def _sugerir_por_parecido(mov: Movimiento):
    """Ultimo recurso antes de PENDIENTE: buscar en la base de ejemplos un
    movimiento ya capturado que se PAREZCA. Si hay, se sugiere (supervisado)."""
    ejemplo, score = buscar_parecido(mov.concepto, mov.tipo)
    if not ejemplo:
        return False
    mov.accion = "CAPTURAR"
    mov.sugerido = True
    mov.contacto = ejemplo.get("contacto") or ""
    mov.cat1 = ejemplo.get("cat1") or ""
    mov.cat2 = ejemplo.get("cat2") or ""
    mov.cat3 = ejemplo.get("cat3") or ""
    mov.nota = (f"SUGERIDO ({score:.0%} parecido a: "
                f"{ejemplo['concepto'][:45]}) - VERIFICA")
    return True


def aplicar_reglas(mov: Movimiento):
    seg = mov.segmento

    # Forma de pago: criterio principal = codigo de caja (002=Azteca, 001=Efectivo);
    # respaldo = texto de la columna BANCO.
    if "002" in mov.caja:
        es_azteca = True
    elif "001" in mov.caja:
        es_azteca = False
    else:
        es_azteca = mov.banco.upper() == "AZTECA"   # respaldo
    if es_azteca:
        mov.forma_pago, mov.cuenta = "Transferencia", "Banco Azteca"
    else:
        mov.forma_pago = "Efectivo"

    # 1. Referencia de cuadre (saldo inicial) -> no se captura
    if any(seg.startswith(r) for r in REFERENCIA):
        mov.accion, mov.tipo = "REFERENCIA", "—"
        return

    # 2. No registrar (ventas en efectivo / traspasos entre cuentas)
    for prefijo, motivo in NO_REGISTRAR.items():
        if seg.startswith(prefijo):
            mov.accion, mov.tipo = "NO_REGISTRAR", "—"
            mov.nota = motivo
            return
    # 2b. Traspasos entre cuentas: la RESERVA (aguinaldo/ahorro) SI se captura
    #     como provision; los traspasos ESPEJO (entrada+salida, p.ej. efectivo
    #     para nomina / gastos varios) NO se registran en ningun lado.
    conc = mov.concepto.upper()
    if seg.startswith("TRASPASO ENTRE CUENTAS"):
        if "AGUINALDO" in conc or "RESERVA" in conc:
            mov.accion, mov.tipo = "CAPTURAR", "GASTO"
            mov.contacto = "Provisiones"
            mov.cat1 = "Provisiones y Reservas(Ahorro/Emergencia)"
            mov.cat2 = "Aguinaldos"
            return
        mov.accion, mov.tipo = "NO_REGISTRAR", "—"
        mov.nota = "Traspaso espejo entre cuentas (entrada y salida se eliminan)"
        return
    if seg == "INGRESO EFECTIVO" and ("TRASPASO" in conc or "GASTOS VARIOS" in conc):
        mov.accion, mov.tipo = "NO_REGISTRAR", "—"
        mov.nota = "Traspaso espejo entre cuentas (entrada y salida se eliminan)"
        return

    # 3. Especiales (reembolso -> deposito en efectivo)
    if seg in CONOC.get("depositos", {}):
        mov.accion, mov.tipo = "CAPTURAR", "DEPOSITO"
        mov.contacto = CONOC["depositos"][seg].get("tipo", "DEPOSITO EN EFECTIVO")
        mov.forma_pago = "Efectivo"
        return

    # 4. Ingresos
    if mov.es_ingreso:
        mov.tipo = "INGRESO"
        info = CONOC.get("ingresos", {}).get(seg)
        if info:
            mov.accion = "CAPTURAR"
            mov.contacto = info.get("cliente", "")
            mov.cat1 = info.get("cat1") or ""
            mov.cat2 = info.get("cat2") or ""
            mov.cat3 = info.get("cat3") or ""
        elif not _sugerir_por_parecido(mov):
            mov.accion = "PENDIENTE"
            mov.nota = f"Ingreso sin cliente definido para '{seg}'"
        return

    # 5. Gastos
    mov.tipo = "GASTO"
    # 5a. Conceptos VARIABLES (OXXO, OTROS, REEMBOLSOS, DISPERSIÓN...): el destino
    #     depende del contenido. Si menciona nomina, se deduce cual; si no, manual.
    if seg in VARIABLES:
        nomina = _detectar_nomina(mov.concepto)
        if nomina:
            mov.accion = "CAPTURAR"
            mov.contacto = nomina
            mov.cat1 = "Gastos Fijos de Operacion - Personal"
            mov.cat2 = "Nominas- Sueldos y salarios"
            mov.cat3 = nomina
        elif not _sugerir_por_parecido(mov):
            mov.accion = "PENDIENTE"
            mov.nota = (f"'{seg}' varia segun contenido: "
                        f"elige proveedor/categoria a mano")
        return
    info = CONOC.get("gastos", {}).get(seg)
    if info:
        mov.accion = "CAPTURAR"
        mov.contacto = info.get("proveedor", "")
        mov.cat1 = info.get("cat1") or ""
        mov.cat2 = info.get("cat2") or ""
        mov.cat3 = info.get("cat3") or ""
    elif not _sugerir_por_parecido(mov):
        mov.accion = "PENDIENTE"
        mov.nota = f"Gasto '{seg}': proveedor/categoria por definir"


# ---------------------------------------------------------------------------
# SALIDA
# ---------------------------------------------------------------------------
def procesar(ruta):
    movimientos, fecha_doc, total_inicio = leer_pdf(ruta)
    for mov in movimientos:
        aplicar_reglas(mov)

    capturar  = [m for m in movimientos if m.accion == "CAPTURAR"]
    pendientes = [m for m in movimientos if m.accion == "PENDIENTE"]
    no_reg    = [m for m in movimientos if m.accion == "NO_REGISTRAR"]
    referencia = [m for m in movimientos if m.accion == "REFERENCIA"]

    print("=" * 78)
    print(f"  CORTE {fecha_doc}   ({ruta.split('/')[-1]})")
    if total_inicio is not None:
        print(f"  TOTAL INICIO DIA (debe cuadrar con Caja Disponible): ${total_inicio:,.2f}")
    print("=" * 78)

    print(f"\n  A CAPTURAR ({len(capturar)}):")
    print(f"  {'#':<4}{'TIPO':<9}{'FORMA':<14}{'CONTACTO/CATEGORIA':<42}{'IMPORTE':>12}")
    print("  " + "-" * 76)
    for m in capturar:
        destino = m.contacto if m.tipo != "GASTO" else f"{m.contacto}  [{m.cat1}]"
        forma = m.forma_pago + (" (Azteca)" if m.cuenta else "")
        print(f"  {m.n:<4}{m.tipo:<9}{forma:<14}{destino[:40]:<42}{m.importe:>12,.2f}")
        if m.sugerido:
            print(f"        ~ {m.nota}")
    print("  " + "-" * 76)
    tot_ing = sum(m.importe for m in capturar if m.tipo == "INGRESO")
    tot_gas = sum(m.importe for m in capturar if m.tipo == "GASTO")
    tot_dep = sum(m.importe for m in capturar if m.tipo == "DEPOSITO")
    print(f"  Ingresos: ${tot_ing:,.2f}   Gastos: ${tot_gas:,.2f}   Depositos: ${tot_dep:,.2f}")

    if no_reg:
        print(f"\n  NO SE REGISTRAN ({len(no_reg)}):")
        for m in no_reg:
            print(f"    {m.n:<4}{m.concepto[:55]:<57}{m.importe:>12,.2f}")

    if pendientes:
        print(f"\n  *** PENDIENTES / AVISAR ({len(pendientes)}): ***")
        for m in pendientes:
            print(f"    {m.n:<4}{m.concepto[:50]:<52}{m.importe:>12,.2f}")
            print(f"        -> {m.nota}")

    return {"fecha": fecha_doc, "ruta": ruta, "total_inicio": total_inicio,
            "capturar": capturar, "pendientes": pendientes,
            "no_registrar": no_reg, "referencia": referencia}


def main():
    rutas = sys.argv[1:] or sorted(glob.glob(
        str(config.CORTES_PDF_DIR / "*.pdf")))
    if not rutas:
        print("No se encontraron PDF en cortes_pdf/")
        return
    todos_pendientes = []
    for ruta in rutas:
        res = procesar(ruta)
        for m in res["pendientes"]:
            todos_pendientes.append((res["fecha"], m))
        print()

    if todos_pendientes:
        # Guardar bloc de pendientes para avisar
        salida = str(config.PENDIENTES_PATH)
        with open(salida, "w", encoding="utf-8") as f:
            f.write("PENDIENTES POR DEFINIR (proveedor / categoria / cliente)\n")
            f.write("=" * 60 + "\n\n")
            for fecha, m in todos_pendientes:
                f.write(f"[{fecha}] renglon {m.n}\n")
                f.write(f"  Concepto : {m.concepto}\n")
                f.write(f"  Banco    : {m.banco}    Importe: ${m.importe:,.2f}\n")
                f.write(f"  Motivo   : {m.nota}\n\n")
        print(f"** {len(todos_pendientes)} pendientes guardados en {salida} **")


if __name__ == "__main__":
    main()
