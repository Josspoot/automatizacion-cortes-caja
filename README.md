# Automatización de captura contable (corte de caja / reportes de sucursal)

Conjunto de bots en Python + [Playwright](https://playwright.dev/python/) que
automatizan la captura **manual y repetitiva** de movimientos contables desde
archivos PDF/Excel hacia un ERP web y/o un libro Excel con macros (VBA), sin
usar la API del sistema (porque normalmente no existe una para este tipo de
plataformas internas tipo "formulario web").

Nació para digitalizar los cortes de caja diarios y los reportes de ingresos/
gastos de una pequeña empresa (una agencia de motocicletas), pero la lógica es
genérica: **cualquier negocio que reciba datos en PDF/Excel y tenga que
teclearlos a mano en un sistema web o en un Excel con formato fijo** puede
adaptar esto.

## ¿Cuándo usar esto?

Úsalo si tu caso se parece a:

- Recibes periódicamente un PDF o Excel con movimientos (ventas, gastos,
  depósitos) que **alguien tiene que capturar a mano** en un sistema web (ERP,
  CRM, plataforma de contabilidad) que no tiene API pública.
- El sistema web tiene un formulario con combos/selects relativamente
  consistentes (mismo proveedor/categoría se repite mes a mes).
- Además (opcional) necesitas el **mismo dato también en un Excel** con
  macros/plantilla fija (doble registro, por ejemplo para auditoría o respaldo
  local).
- Quieres evitar volver a capturar lo mismo dos veces si el proceso se
  interrumpe (control de duplicados).

**No** es para ti si necesitas:
- Integración en tiempo real / vía API (si el sistema tiene API, úsala en vez
  de esto — Playwright es el último recurso cuando no hay otra forma).
- Un sistema genérico "no-code": esto requiere ajustar manualmente los
  selectores de Playwright (`bot.py`) a **tu** plataforma, porque cada
  formulario web es distinto.

## Cómo funciona (arquitectura)

```
PDF / Excel de origen
        │
        ▼
  leer_corte.py / adaptador de reporte   (parseo del archivo de origen)
        │
        ▼
  conocimiento.json  ──────────────►  "cerebro": mapea texto del concepto
  conocimiento_excel.json             (proveedor/cliente) a la selección
                                       exacta de combos (proveedor + hasta
                                       3 niveles de categoría)
        │
        ▼
  bot.py (web, Playwright)   /   bot_excel.py (Excel, AppleScript)
        │
        ▼
  registro.json / registro_reporte.json   (dedup: qué ya se capturó y dónde,
                                            para no repetirlo si se re-corre)
```

- **`bot.py`** — captura cortes de caja PDF en la plataforma web. Dos modos:
  *supervisado* (se detiene en cada movimiento para que confirmes) y
  *automático* (`--auto`, llena y guarda solo lo que el "cerebro" ya conoce al
  100%; lo ambiguo o nuevo se detiene para que tú decidas, y luego **aprende**
  guardando tu elección en `conocimiento.json`).
- **`bot_excel.py`** — lo mismo pero escribiendo directo en un libro `.xlsm`
  con macros, usando Excel real vía AppleScript (macOS) para no romper
  fórmulas/tablas/segmentadores.
- **`bot_reporte.py`** — variante para reportes de sucursal en Excel (detalle
  por ticket) que captura en **las dos plataformas a la vez** y lleva su
  propio registro de qué-ya-se-capturó-dónde (`registro_reporte.json`), porque
  el candado del Excel no sirve para saber si la web ya tiene el dato.
- **`leer_corte.py`** — parsea los PDF de corte de caja con `pdfplumber`.
- **`config.py`** — toda la configuración específica de tu negocio (URL,
  rutas) sale de variables de entorno (`.env`), el código no tiene nada
  hardcodeado.

### El "cerebro" (`conocimiento.json`)

Es un diccionario `texto del concepto → (proveedor/cliente, categoría 1, 2, 3)`.
La primera vez que aparece un proveedor/concepto nuevo, el bot no lo captura
solo: se detiene, tú lo completas en el formulario, y el bot **aprende** esa
elección para la próxima vez. Con el tiempo, la gran mayoría de movimientos se
capturan 100% automático y solo lo realmente nuevo requiere intervención.

Revisa `conocimiento.example.json`, `conocimiento_excel.example.json` y
`categorias.example.json` para ver la forma exacta que debe tener tu propia
copia (sin datos reales).

### Control de duplicados

Cada movimiento capturado se identifica con una clave determinística
(`fecha|orden|banco|importe`) y se guarda en `registro.json` /
`registro_reporte.json`. Si vuelves a correr el mismo PDF/Excel, lo ya
capturado se salta automáticamente — así puedes reintentar sin miedo a
duplicar.

## Instalación

```bash
git clone <este-repo>
cd automatizacion-cortes-caja
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

Copia `.env.example` a `.env` y llena al menos `CORTES_URL_BASE` con la URL de
tu plataforma. Copia también las plantillas de conocimiento:

```bash
cp .env.example .env
cp conocimiento.example.json conocimiento.json
cp conocimiento_excel.example.json conocimiento_excel.json
cp categorias.example.json categorias.json
```

Edítalas con tus propios proveedores/clientes/categorías — empiezan casi
vacías a propósito, se van llenando solas conforme usas el bot.

## Primer uso: sesión iniciada

El bot usa un **perfil de navegador persistente** (`.perfil_navegador/`, fuera
del repo) para no pedirte login cada vez. La primera vez:

```bash
python abrir.py
```

Esto abre una ventana de Chromium; inicia sesión manualmente ahí, déjala
abierta. De ahí en adelante el bot reutiliza esa sesión guardada.

## Uso

**Siempre corre primero en modo ENSAYO (sin flags de escritura)** para ver qué
va a capturar antes de tocar nada real:

```bash
# corte de caja individual
python bot.py "cortes_pdf/CC 12-06-2026.pdf"          # supervisado
python bot.py "cortes_pdf/CC 12-06-2026.pdf" --auto    # automático

# registro doble (web + excel) de un corte
python capturar.py 12                                   # busca el PDF del día 12

# reportes de sucursal (xlsx con detalle por ticket)
python bot_reporte.py                   # ENSAYO
python bot_reporte.py --excel           # escribe en el Excel
python bot_reporte.py --web             # captura en la web
python bot_reporte.py --web --solo 7    # solo el movimiento 7 (recaptura puntual)
```

## Adaptar esto a tu propio sistema

1. **Selectores de Playwright** (`bot.py`): el formulario de cada ERP web es
   distinto. Vas a necesitar abrir el inspector del navegador y reescribir los
   `page.locator(...)` para que apunten a los campos de tu plataforma.
2. **Estructura del Excel** (`bot_excel.py`): ajusta los nombres de hoja,
   columnas y el AppleScript si tu libro no tiene la misma plantilla.
3. **Formato del PDF/Excel de origen** (`leer_corte.py` / adaptadores de
   `bot_reporte.py`): el parseo asume un formato de corte de caja específico;
   para un formato distinto, escribe tu propio adaptador que produzca objetos
   `Movimiento` (ver `leer_corte.py`) y reutiliza `bot.py`/`bot_excel.py` tal
   cual.

## Lección aprendida: verifica, no confíes ciegamente en el log

Este bot tuvo, en producción, casos ocasionales de **falso positivo de
guardado**: el log decía "guardado" pero el movimiento no había quedado
realmente capturado en la web (race condition del formulario/SPA). La
recomendación, si adaptas esto a tu caso: después de cada lote, haz una
verificación independiente (buscar cada movimiento capturado directamente en
la plataforma, no solo confiar en la salida del bot) antes de dar el lote por
cerrado.

## Nota sobre licencia

Este repo se publica como referencia/portafolio, no bajo una licencia abierta
explícita.
