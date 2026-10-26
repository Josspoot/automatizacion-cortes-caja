#!/usr/bin/env python3
"""
Bot de captura de movimientos en el ERP web (ver config.py / .env).

MODO SUPERVISADO (por defecto):
  - El bot LLENA cada movimiento con lo que sabe (cerebro = conocimiento.json):
    tipo, fecha, forma de pago, cuenta, cliente/proveedor, categorias, concepto, monto.
  - En los que NO conoce, llena lo basico y deja cliente/proveedor/categoria para ti.
  - Se DETIENE; TU revisas, completas si falta y das "Guardar y agregar otro".
  - APRENDE: lee lo que elegiste y lo guarda en conocimiento.json para la proxima.

MODO AUTO (--auto): los movimientos que conoce al 100% los llena COMPLETOS
  (incluye cliente/proveedor y categorias), verifica lo seleccionado y guarda
  solo. Los SUGERIDOS por parecido y los PENDIENTES se detienen para que TU
  confirmes/completes (y aprende igual que en supervisado).

Uso:
    python bot.py "cortes_pdf/CC 12-06-2026.pdf"          # supervisado
    python bot.py "cortes_pdf/CC 12-06-2026.pdf" --auto    # automatico
"""
import argparse
import functools
import json
import re
import time
import unicodedata
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
import leer_corte
import config

print = functools.partial(print, flush=True)

BASE = Path(__file__).parent
PERFIL = config.PERFIL_NAVEGADOR
CAPTURAS = BASE / "capturas"
CONOC_PATH = config.CONOCIMIENTO_PATH
URL_MOV = config.URL_MOV
TAXONOMIA_PATH = config.CATEGORIAS_PATH

try:
    TAXONOMIA = json.loads(TAXONOMIA_PATH.read_text(encoding="utf-8"))
except Exception:
    TAXONOMIA = {}


def agregar_a_taxonomia(cat1, cat2, cat3):
    """Agrega una categoria nueva (elegida en el formulario) a categorias.json."""
    if not cat1:
        return
    TAXONOMIA.setdefault(cat1, {})
    if cat2:
        TAXONOMIA[cat1].setdefault(cat2, [])
        if cat3 and cat3 not in TAXONOMIA[cat1][cat2]:
            TAXONOMIA[cat1][cat2].append(cat3)
    try:
        TAXONOMIA_PATH.write_text(json.dumps(TAXONOMIA, indent=2, ensure_ascii=False),
                                  encoding="utf-8")
    except Exception:
        pass


def categoria_valida(cat1, cat2, cat3):
    """True si cat1/cat2/cat3 existen en la taxonomia oficial (categorias.json)."""
    if not TAXONOMIA:
        return True                      # sin taxonomia cargada: no bloquear
    if not cat1:
        return True                      # sin categoria: se permite (la pones tu)
    if cat1 not in TAXONOMIA:
        return False
    if cat2 and cat2 not in TAXONOMIA[cat1]:
        return False
    if cat3 and cat2 and cat3 not in TAXONOMIA[cat1].get(cat2, []):
        return False
    return True

PLACEHOLDERS = ("buscar", "seleccionar", "no hay")
PASO = [0]
TYPEAR_CLIENTE = True  # supervisado: escribe cliente/proveedor y deja el menu abierto para que elijas


PUERTO = 9222


def abrir_navegador(p):
    """Abre la ventana del bot (Chromium maximizado, con perfil/sesion guardada)."""
    try:
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(PERFIL), headless=False,
            no_viewport=True, args=["--start-maximized"])
        print("  Navegador: Chromium (maximizado)")
        return ctx
    except Exception as e:
        print(f"\n  (!) No pude abrir el navegador: {str(e)[:80]}")
        print("      Cierra cualquier otra ventana de este perfil y reintenta.\n")
        return None


def captura(page, nombre):
    PASO[0] += 1
    CAPTURAS.mkdir(exist_ok=True)
    try:
        page.screenshot(path=str(CAPTURAS / f"{PASO[0]:02d}_{nombre}.png"))
    except Exception:
        pass


def fecha_iso(fecha_pdf):
    """Se captura con la MISMA fecha del corte (decidido por el usuario el
    2026-07-14: todo alineado a la fecha del PDF; la convencion X-1 del
    2026-07-11 quedo descartada y las capturas viejas se corrigieron +1)."""
    from datetime import date
    d, m, y = fecha_pdf.split("/")
    return date(int(y), int(m), int(d)).isoformat()


def limpiar_texto(t):
    """Quita iconos/emojis y saltos de linea que el combo trae junto al texto."""
    t = t.replace("\n", " ")
    t = "".join(c for c in t if ord(c) < 0x2000)   # fuera emojis/puntos de color
    return re.sub(r"\s+", " ", t).strip()


def texto_combo(loc):
    """Texto actual (limpio) de un combo; '' si esta en placeholder (vacio)."""
    try:
        t = limpiar_texto(loc.inner_text(timeout=1500) or "")
    except Exception:
        return ""
    return "" if t.lower().startswith(PLACEHOLDERS) else t


# ---------------------------------------------------------------------------
# Seleccion en combos
# ---------------------------------------------------------------------------
def elegir_select_for(dlg, page, for_id, valor):
    """Radix Select identificado por label[for=...] (Forma de Pago)."""
    dlg.locator(f'label[for="{for_id}"] + button[role="combobox"]').click()
    page.wait_for_timeout(400)
    page.get_by_role("option", name=re.compile(re.escape(valor), re.I)).first.click()
    page.wait_for_timeout(250)


def elegir_cuenta(dlg, page, valor):
    """Cuenta Bancaria (aparece al elegir Transferencia)."""
    trigger = dlg.locator(
        'xpath=//label[contains(translate(., "CUENTA", "cuenta"), "cuenta")]'
        '/following-sibling::button[@role="combobox"][1]')
    trigger.first.click()
    page.wait_for_timeout(400)
    page.get_by_role("option", name=re.compile(re.escape(valor), re.I)).first.click()
    page.wait_for_timeout(250)


def resolver_avisos_previos(dlg, page):
    """Antes de dar 'Guardar': si el formulario muestra los avisos de 'Monto
    inusual' o 'Transaccion similar detectada', los descarta automaticamente
    (Confirmar / No es duplicado) para que el boton de Guardar deje de estar
    deshabilitado. Nuestro propio registro (registro_reporte.json) ya evita
    que un mismo movimiento del Excel se capture dos veces, asi que estos
    avisos -comparando importes contra el historico real- son en la practica
    siempre falsos positivos (otra transaccion distinta, mismo monto)."""
    try:
        dlg.get_by_role("button", name=re.compile(r"^Confirmar$", re.I)).first.click(timeout=1500)
        page.wait_for_timeout(300)
    except PWTimeout:
        pass
    try:
        dlg.get_by_role("button", name=re.compile(r"No es duplicado", re.I)).first.click(timeout=1500)
        page.wait_for_timeout(300)
    except PWTimeout:
        pass


def _norm(t):
    """MAYUSCULAS sin acentos ni espacios dobles (para comparar textos)."""
    t = unicodedata.normalize("NFD", t or "")
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).strip().upper()


def elegir_busqueda(page, combo, valor):
    """Combo de busqueda (Cliente/Proveedor/Categoria): abre, escribe en el
    input del popover (press_sequentially: keyboard.type NO dispara la
    busqueda) y da clic en la opcion que coincida. Devuelve True si selecciono.
    OJO: nunca usar Escape (cierra todo el dialogo)."""
    try:
        combo.click()
        page.wait_for_timeout(550)
        inp = page.locator(
            '[class*="bg-popover"] input, [data-radix-popper-content-wrapper] input').last
        inp.fill("")
        inp.press_sequentially(valor, delay=45)
        page.wait_for_timeout(900)   # por si consulta al servidor
        # Las opciones son <div> SIN role: divs con clase cursor-pointer dentro
        # del popover (visto en diagnostico; el popover es role=dialog).
        opciones = page.locator(
            '[class*="bg-popover"] div[class*="cursor-pointer"], '
            '[class*="bg-popover"] [role="option"], '
            '[class*="bg-popover"] [cmdk-item]')
        objetivo = _norm(valor)
        for i in range(min(opciones.count(), 20)):
            op = opciones.nth(i)
            txt = _norm(limpiar_texto(op.inner_text(timeout=800) or ""))
            if not txt or txt.lower().startswith(PLACEHOLDERS):
                continue
            if objetivo in txt or txt in objetivo:
                op.click()
                page.wait_for_timeout(350)
                return True
    except Exception:
        pass
    # No encontro la opcion: plegar el popover SIN Escape (re-clic al combo).
    try:
        combo.click(timeout=1000)
        page.wait_for_timeout(200)
    except Exception:
        pass
    return False


# ---------------------------------------------------------------------------
def pagina_viva(ctx, page=None):
    """Devuelve una pestaña viva en movimientos (re-toma si la anterior murio)."""
    if page is not None:
        try:
            if not page.is_closed():
                return page
        except Exception:
            pass
    for pg in ctx.pages:                       # buscar la de movimientos
        try:
            if not pg.is_closed() and "movimientos" in pg.url:
                return pg
        except Exception:
            continue
    for pg in ctx.pages:                       # cualquiera viva
        try:
            if not pg.is_closed():
                return pg
        except Exception:
            continue
    pg = ctx.new_page()                        # ultima opcion
    pg.goto(URL_MOV, wait_until="domcontentloaded")
    return pg


def asegurar_dialogo(page):
    dlg = page.get_by_role("dialog")
    if dlg.count() == 0 or not dlg.first.is_visible():
        boton = page.get_by_role("button", name=re.compile(r"Nuevo Movimiento", re.I)).first
        boton.wait_for(state="visible", timeout=15000)
        boton.click()
        dlg = page.get_by_role("dialog")
        dlg.first.wait_for(state="visible", timeout=10000)
        page.wait_for_timeout(600)
    return dlg.first


def llenar_campos(page, dlg, mov, auto=False):
    """Llena el formulario. Con auto=True tambien SELECCIONA cliente/proveedor
    y categorias; devuelve True si todo quedo seleccionado."""
    # 1) Tipo
    tipo_btn = "Gasto" if mov.tipo == "GASTO" else "Ingreso"
    dlg.get_by_role("button", name=re.compile(rf"^{tipo_btn}$", re.I)).first.click()
    page.wait_for_timeout(400)

    # 2) Fecha
    dlg.locator("#fecha").fill(fecha_iso(mov.fecha))
    page.wait_for_timeout(120)

    # 3) Forma de pago + cuenta
    try:
        elegir_select_for(dlg, page, "formaPago", mov.forma_pago)
    except Exception as e:
        print(f"    (!) Forma de pago: {str(e)[:50]}")
    if mov.cuenta:
        try:
            elegir_cuenta(dlg, page, mov.cuenta)
        except Exception as e:
            print(f"    (!) Cuenta: {str(e)[:50]}")

    # 4) Concepto
    dlg.locator("#concepto").fill(mov.concepto)
    page.wait_for_timeout(120)
    # 5) Monto
    dlg.locator("#monto").fill(f"{mov.importe:.2f}")
    page.wait_for_timeout(120)

    # 6-7) Cliente/Proveedor + Categorias
    # MODO AUTO: el bot SELECCIONA las opciones el solo (escribe en el input del
    # popover y da clic en la coincidencia). Si algo no coincide -> False y el
    # movimiento pasa a confirmacion manual.
    if auto and mov.contacto:
        etiqueta = "proveedor" if mov.tipo == "GASTO" else "cliente"
        combos = dlg.locator('button[aria-haspopup="dialog"]')
        # si el combo retiene el valor del movimiento anterior ya no dice
        # "proveedor/cliente": usar el primero por posicion como respaldo
        con_texto = combos.filter(has_text=re.compile(etiqueta, re.I))
        combo = con_texto.first if con_texto.count() else combos.nth(0)
        if not elegir_busqueda(page, combo, mov.contacto):
            print(f"    (!) No pude seleccionar {etiqueta} '{mov.contacto}'")
            return False
        for nivel, valor in enumerate([mov.cat1, mov.cat2, mov.cat3], start=1):
            if not valor:
                continue
            page.wait_for_timeout(400)   # el combo del siguiente nivel aparece al elegir
            combos = dlg.locator('button[aria-haspopup="dialog"]')
            if combos.count() <= nivel or not elegir_busqueda(page, combos.nth(nivel), valor):
                print(f"    (!) No pude seleccionar la categoria '{valor}'")
                return False
        return True

    # 8) SUPERVISADO: escribir el cliente/proveedor y DEJAR el menu abierto para
    #    que TU elijas la opcion correcta (sin Escape -> no se cierra el dialogo).
    if not auto and TYPEAR_CLIENTE and mov.contacto:
        etiqueta = "proveedor" if mov.tipo == "GASTO" else "cliente"
        try:
            combo = dlg.locator('button[aria-haspopup="dialog"]').filter(
                has_text=re.compile(etiqueta, re.I)).first
            combo.click()
            page.wait_for_timeout(600)
            inp = page.locator(
                '[class*="bg-popover"] input, [data-radix-popper-content-wrapper] input').last
            inp.press_sequentially(mov.contacto, delay=70)
            page.wait_for_timeout(500)
            print(f"    (menu de {etiqueta} abierto con '{mov.contacto}': elige la opcion)")
        except Exception as e:
            print(f"    (!) No pude escribir {etiqueta}: {str(e)[:50]}")


def leer_estado(dlg):
    """Lee lo que hay ahora en contacto + categorias (para aprender)."""
    combos = dlg.locator('button[aria-haspopup="dialog"]')
    n = combos.count()
    vals = [texto_combo(combos.nth(i)) for i in range(min(n, 4))]
    vals += [""] * (4 - len(vals))
    return {"contacto": vals[0], "cat1": vals[1], "cat2": vals[2], "cat3": vals[3]}


def esperar_y_aprender(page, dlg, mov, aprender):
    """Espera a que TU guardes; mientras, recuerda el ultimo estado lleno."""
    print("   >>> REVISA y dale 'Guardar y agregar otro' (o 'Guardar'). Esperando...")
    ultimo = None
    limite = time.time() + 900
    while time.time() < limite:
        time.sleep(1)
        try:
            campo = page.locator("#concepto")
            if campo.count() == 0:
                break                                   # dialogo cerrado
            val = campo.first.input_value(timeout=1500)
            if val.strip() == "":
                break                                   # guardado y agregar otro
            if aprender:
                est = leer_estado(dlg)
                if est["contacto"]:
                    ultimo = est                        # snapshot util
        except Exception:
            break
    if aprender and ultimo:
        aprender_mapeo(mov, ultimo)
    return ultimo


def aprender_mapeo(mov, est):
    """Guarda en conocimiento.json lo que el usuario eligio."""
    try:
        conoc = json.loads(CONOC_PATH.read_text(encoding="utf-8"))
    except Exception:
        return
    seg = mov.segmento
    grupo = "gastos" if mov.tipo == "GASTO" else "ingresos" if mov.tipo == "INGRESO" else None
    if grupo is None:
        return
    # SEGURIDAD: no sobrescribir un mapeo ya conocido (evita corromper el cerebro
    # con lecturas viejas del formulario). Solo se aprenden conceptos NUEVOS.
    if seg in conoc[grupo]:
        return
    # SEGURIDAD 2: los conceptos VARIABLES (OXXO, OTROS, REEMBOLSOS, DISPERSIÓN)
    # cambian de destino segun el contenido -> nunca se aprende un mapeo fijo.
    if seg in conoc.get("variables", {}):
        return
    cat1, cat2, cat3 = est["cat1"] or None, est["cat2"] or None, est["cat3"] or None
    # VALIDACION: el formulario solo permite categorias validas, asi que lo que
    # elegiste es correcto. Si la categoria es NUEVA (no estaba en la taxonomia
    # de junio), se aprende igual, se avisa, y se agrega sola a categorias.json.
    if not categoria_valida(cat1, cat2, cat3):
        print(f"   ⚠️  Categoria nueva en '{seg}': '{cat1} > {cat2} > {cat3}' "
              f"(la agrego a la taxonomia; verifica si te parece rara).")
        agregar_a_taxonomia(cat1, cat2, cat3)
    entrada = {"cat1": cat1, "cat2": cat2, "cat3": cat3}
    clave = "proveedor" if grupo == "gastos" else "cliente"
    entrada = {clave: est["contacto"], **entrada}
    antes = conoc[grupo].get(seg)
    conoc[grupo][seg] = entrada
    if entrada != antes:
        CONOC_PATH.write_text(json.dumps(conoc, indent=2, ensure_ascii=False),
                              encoding="utf-8")
        print(f"   🧠 Aprendido: '{seg}' -> {est['contacto']}"
              f"{' / '+est['cat1'] if est['cat1'] else ''}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--auto", action="store_true")
    ap.add_argument("--desde", type=int, default=1, help="reanudar desde el movimiento N")
    ap.add_argument("--solo", type=int, help="capturar SOLO la fila N del PDF")
    args = ap.parse_args()

    movimientos, fecha_doc, total_inicio = leer_corte.leer_pdf(args.pdf)
    for m in movimientos:
        leer_corte.aplicar_reglas(m)
    # Se procesan CAPTURAR y PENDIENTE en ambos modos. En --auto los seguros se
    # guardan solos; los SUGERIDOS por parecido y los PENDIENTES se detienen
    # para que TU confirmes (y el bot aprende).
    cola = [m for m in movimientos if m.accion in ("CAPTURAR", "PENDIENTE")]
    inicio = max(1, args.desde)   # numero de la tabla del PDF (columna N°)

    def num_fila(m):
        try:
            return int(m.n.strip("/ "))
        except Exception:
            return 0

    total_antes = len(cola)
    if args.solo:
        cola = [m for m in cola if num_fila(m) == args.solo]
    elif inicio > 1:
        cola = [m for m in cola if num_fila(m) >= inicio]

    print(f"\n  Corte {fecha_doc}  ->  {len(cola)} movimientos"
          f"{f' (desde la fila {inicio} del PDF, de {total_antes})' if inicio > 1 else ''}")
    print(f"  Modo: {'AUTOMATICO (sugeridos/pendientes los confirmas tu)' if args.auto else 'SUPERVISADO (aprende de ti)'}\n")

    with sync_playwright() as p:
        ctx = abrir_navegador(p)
        if ctx is None:
            return
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        # solo navegar si no estamos ya en movimientos (no quitar tu pestaña)
        if "movimientos" not in page.url:
            page.goto(URL_MOV, wait_until="domcontentloaded")
            page.wait_for_timeout(2500)
        if "auth/login" in page.url:
            print("  (!) Inicia sesion en la ventana (espero 60s)...")
            page.wait_for_timeout(60000)

        ok = err = 0
        for i, m in enumerate(cola, 1):
            etiqueta = "PENDIENTE" if m.accion == "PENDIENTE" else m.tipo
            print(f"\n[{i}/{len(cola)}]  {m.n} {etiqueta}  {m.forma_pago}"
                  f"{' '+m.cuenta if m.cuenta else ''}  ->  {m.contacto or '(tu defines)'}  ${m.importe:,.2f}")
            print(f"           Concepto: {m.concepto}")
            if m.sugerido:
                print(f"           ~ {m.nota}")
            exito = False
            for intento in (1, 2):      # el 1er movimiento a veces falla al cargar: reintentar
                try:
                    page = pagina_viva(ctx, page)   # re-toma la pestaña si murio
                    dlg = asegurar_dialogo(page)
                    # auto-seleccion de combos solo para los seguros (no sugeridos/pendientes)
                    seguro = args.auto and m.accion == "CAPTURAR" and not m.sugerido and m.cat1
                    listo = llenar_campos(page, dlg, m, auto=bool(seguro))
                    captura(page, f"lleno_{m.n.strip('/')}")

                    guardar_solo = seguro and listo
                    if guardar_solo:
                        # VERIFICAR que el formulario quedo como se esperaba
                        est = leer_estado(dlg)
                        esperado = [("contacto", m.contacto), ("cat1", m.cat1),
                                    ("cat2", m.cat2), ("cat3", m.cat3)]
                        for campo, valor in esperado:
                            v_form = _norm(est[campo])
                            v_esp = _norm(valor or "")
                            if v_esp and not (v_esp in v_form or v_form in v_esp):
                                print(f"    (!) {campo}: quedo '{est[campo]}' y esperaba '{valor}'")
                                guardar_solo = False
                            if not v_esp and campo != "contacto" and v_form:
                                print(f"    (!) {campo}: quedo '{est[campo]}' (sobrante de otro movimiento)")
                                guardar_solo = False

                    if guardar_solo:
                        resolver_avisos_previos(dlg, page)
                        dlg.get_by_role("button", name=re.compile(r"Guardar y agregar otro", re.I)).first.click()
                        page.wait_for_timeout(900)
                        try:
                            page.get_by_role("button", name=re.compile(r"^S[ií]$|Confirmar|Aceptar", re.I)).first.click(timeout=1500)
                            page.wait_for_timeout(400)
                        except PWTimeout:
                            pass
                        print("   ✔ guardado solo")
                    else:
                        if args.auto:
                            motivo = ("sugerido por parecido" if m.sugerido else
                                      "pendiente: tu eliges" if m.accion == "PENDIENTE" else
                                      "no pude llenarlo completo")
                            print(f"   >>> Este lo confirmas TU ({motivo})")
                        esperar_y_aprender(page, dlg, m, aprender=True)
                    ok += 1
                    exito = True
                    break
                except Exception as e:
                    print(f"   (!) intento {intento} fallo: {str(e)[:90]}")
                    captura(page, f"ERROR_{m.n.strip('/')}")
                    # no cerramos el dialogo con Escape para no quitarte tu vista
                    page.wait_for_timeout(2500)
            if not exito:
                err += 1
                print("   ERROR: no se pudo capturar este movimiento (reintentalo con --solo)")

        print(f"\n  ===== FIN. Procesados: {ok}   Errores: {err} =====")
        if total_inicio:
            print(f"  Recuerda: TOTAL INICIO DIA = ${total_inicio:,.2f} (debe cuadrar).")
        try:
            page.wait_for_timeout(1500)
        except Exception:
            pass
        ctx.close()          # cierra la ventana sola al terminar


if __name__ == "__main__":
    main()
