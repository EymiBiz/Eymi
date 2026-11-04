"""
BOT DE SALIDAS · SOFTCOM ALMOVI20 · v3.8

Qué hace:
  1. Lee TODAS las salidas pendientes de la hoja PARTE. Una salida = un PEDIDO en el mostrador:
     filas seguidas del mismo trabajador, área y fecha (Glosa 1 = nombre del trabajador).
  2. Para cada una:
       - si hace falta, cambia la fecha de SOFTCOM (Selección Almacén y Fecha) y abre ALMOVI20;
       - llena la cabecera con ENTER: PR, RQ, CC, solicitante, glosa (trabajador);
       - en "Detalles de Salida" escribe cada producto: código ⏎ ⏎ ⏎ cantidad ⏎;
       - graba solo (o pregunta, en modo 'preguntar'): Alt+F → Aceptar → ¿Imprimir? No
         → anota el N° PS-… en la columna P del Excel y lo guarda.

Uso (CMD, en Descargas):
  py -3.14 bot_salidas.py            todas las pendientes, AUTOMÁTICO (graba sin preguntar)
  py -3.14 bot_salidas.py preguntar  igual, pero pregunta SI antes de grabar cada una
  py -3.14 bot_salidas.py cabecera   solo llena la cabecera de la primera y para
  py -3.14 bot_salidas.py menu       guarda el menú de SOFTCOM en el Escritorio (softcom_menu.txt)
  py -3.14 bot_salidas.py revisar    solo muestra qué salidas registraría (no toca SOFTCOM)
No toques mouse ni teclado mientras escribe.

Cambios v3.8 (sobre v3.7):
  - Si escribes mal el modo (ej. "preguntr") se detiene: antes grababa en AUTOMÁTICO.
  - Apenas SOFTCOM da el N° PS, lo anota en bot_registro.txt CON las filas del Excel.
    Si el bot se cae antes de guardar el Excel, al volver a correr lo detecta y
    ofrece anotarlo, en vez de registrar la salida dos veces.
  - Lee de vuelta cada campo de la cabecera y para si SOFTCOM no lo dejó igual.
  - Si Excel está ocupado (editando una celda) reintenta en vez de caerse.
  - Lee todas las filas del PARTE (ya no se corta en la fila 804).
  - Avisa si la cantidad no es válida o el stock estimado queda negativo.
  - Mensajes claros si falta un botón en un cuadro de SOFTCOM; Ctrl+C para limpio.
"""
import os, re, sys, time, ctypes, warnings, datetime
from ctypes import wintypes
warnings.filterwarnings("ignore")

# ------------------------------------------------------------------ CONFIGURACIÓN
RUTA_EXCEL = r"C:\Users\Logistica4\Downloads\PARTE_SALIDA_ALMACEN_v7.xlsx"
PREGUNTAR_ANTES_DE_GRABAR = False   # False = graba solo, sin pedir SI  ·  True = pregunta en cada salida
PAUSA = 0.35                 # segundos tras cada tecla/campo (súbelo a 0.6 si SOFTCOM va lento)

# Posiciones medidas en tu PC (respecto a la esquina de cada ventana)
A = {  # Registro de Salida - ALMOVI20
    "mov": (139, 55), "fecha": (452, 76), "tipo_conv": (139, 121),
    "tipo_ref1": (139, 206), "nro_ref1": (399, 206), "cc": (139, 268),
    "sol": (139, 310), "glosa1": (139, 331), "glosa2": (139, 351), "grabar": (497, 466),
}
D = {  # Detalles de Salida
    "items": (582, 83), "nro_doc": (151, 71), "col_codigo": 10, "col_cantidad": 506,
}
ENTERS_CODIGO_A_CANTIDAD = 3   # en tu grabación: código ⏎ ⏎ ⏎ y recién la cantidad
# ---------------------------------------------------------------------------------

try:
    import psutil, win32gui, win32process, win32con
    from pywinauto.keyboard import send_keys
    from pywinauto.controls.hwndwrapper import HwndWrapper
except ImportError as e:
    print("Falta una librería:", e); input("Enter para cerrar..."); sys.exit(1)


class Alto(Exception):
    """El bot se detiene a propósito, con un mensaje claro."""


# ============================================================ SOFTCOM: ventanas y foco
def ventana_principal():
    """Busca la ventana principal VISIBLE de SOFTCOM (no depende del nombre del .exe
    ni de si hay otro SOFTCOM escondido abierto). Devuelve (hwnd, pid)."""
    cands = []
    def cb(h, _):
        try:
            if win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) == "ThunderRT6MDIForm":
                t = win32gui.GetWindowText(h).upper()
                pid = win32process.GetWindowThreadProcessId(h)[1]
                tiene_almovi = [0]
                def hijas(c, _):
                    if "ALMOVI20" in win32gui.GetWindowText(c):
                        tiene_almovi[0] = 1
                    return True
                win32gui.EnumChildWindows(h, hijas, None)
                puntos = tiene_almovi[0] * 10 + ("ALMACEN" in t) * 2 + ("SOFTCOM" in t or "SISTEMA COMERCIAL" in t)
                cands.append((puntos, h, pid))
        except Exception:
            pass
        return True
    win32gui.EnumWindows(cb, None)
    if not cands:
        return 0, None
    cands.sort(reverse=True)
    return cands[0][1], cands[0][2]


PID = None


def formularios():
    """{hwnd: título} de las ventanas visibles de SOFTCOM (principal, hijas MDI y mensajes)."""
    res = {}
    def top(h, _):
        try:
            if win32process.GetWindowThreadProcessId(h)[1] == PID and win32gui.IsWindowVisible(h):
                res[h] = win32gui.GetWindowText(h)
                def hijas(c, _):
                    if win32gui.GetClassName(c) == "ThunderRT6FormDC" and win32gui.IsWindowVisible(c):
                        res[c] = win32gui.GetWindowText(c)
                    return True
                win32gui.EnumChildWindows(h, hijas, None)
        except Exception:
            pass
        return True
    win32gui.EnumWindows(top, None)
    return res


def form(titulo):
    for h, t in formularios().items():
        if titulo in t and win32gui.GetClassName(h) != "#32770":
            return h
    return 0


def texto(h):
    try:
        n = win32gui.SendMessage(h, win32con.WM_GETTEXTLENGTH, 0, 0)
        buf = ctypes.create_unicode_buffer(n + 1)
        ctypes.windll.user32.SendMessageW(h, win32con.WM_GETTEXT, n + 1, buf)
        return buf.value
    except Exception:
        return ""


def rel(h, f):
    fl, ft, _, _ = win32gui.GetWindowRect(f)
    l, t, _, _ = win32gui.GetWindowRect(h)
    return l - fl, t - ft


def control(f, pos, clase=None, tol=8):
    encontrados = []
    def cb(h, _):
        try:
            if clase and win32gui.GetClassName(h) != clase:
                return True
            dx, dy = rel(h, f)
            if abs(dx - pos[0]) <= tol and abs(dy - pos[1]) <= tol and win32gui.IsWindowVisible(h):
                encontrados.append(h)
        except Exception:
            pass
        return True
    win32gui.EnumChildWindows(f, cb, None)
    return encontrados[0] if encontrados else 0


class GUITHREADINFO(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.DWORD), ("flags", wintypes.DWORD), ("hwndActive", wintypes.HWND),
                ("hwndFocus", wintypes.HWND), ("hwndCapture", wintypes.HWND), ("hwndMenuOwner", wintypes.HWND),
                ("hwndMoveSize", wintypes.HWND), ("hwndCaret", wintypes.HWND), ("rcCaret", wintypes.RECT)]


TID = None


def foco():
    gi = GUITHREADINFO(); gi.cbSize = ctypes.sizeof(GUITHREADINFO)
    if ctypes.windll.user32.GetGUIThreadInfo(TID, ctypes.byref(gi)):
        return gi.hwndFocus or 0
    return 0


def foco_en(f, pos, tol=8):
    h = foco()
    if not h or not f:
        return False
    try:
        dx, dy = rel(h, f)
        return abs(dx - pos[0]) <= tol and abs(dy - pos[1]) <= tol
    except Exception:
        return False


def casilla(det, x, tol=12):
    """Si el cursor está en la casilla de edición de la columna x del detalle, la devuelve."""
    h = foco()
    try:
        if h and win32gui.GetClassName(h) == "ThunderRT6TextBox" and abs(rel(h, det)[0] - x) <= tol:
            return h
    except Exception:
        pass
    return 0


def donde_esta_foco():
    h = foco()
    if not h:
        return "(ninguno)"
    f = h
    while f and win32gui.GetClassName(f) not in ("ThunderRT6FormDC", "#32770"):
        f = win32gui.GetParent(f)
    if not f:
        return win32gui.GetClassName(h)
    dx, dy = rel(h, f)
    return f"{win32gui.GetClassName(h)} pos=({dx},{dy}) en '{win32gui.GetWindowText(f)}'"


def mensaje():
    """Cuadro de mensaje de SOFTCOM abierto → (hwnd, texto) o (0, '')."""
    for h, t in formularios().items():
        if win32gui.GetClassName(h) == "#32770":
            partes = []
            def cb(c, _):
                if win32gui.GetClassName(c) == "Static" and texto(c).strip():
                    partes.append(texto(c).strip())
                return True
            win32gui.EnumChildWindows(h, cb, None)
            return h, " ".join(partes)
    return 0, ""


def boton(dialogo, etiqueta):
    hallado = []
    def cb(c, _):
        if win32gui.GetClassName(c) == "Button" and texto(c).replace("&", "").strip().lower() == etiqueta.lower():
            hallado.append(c)
        return True
    win32gui.EnumChildWindows(dialogo, cb, None)
    return hallado[0] if hallado else 0


def pulsar(dialogo, *etiquetas):
    """Hace clic en el primer botón que exista de las etiquetas dadas; si no hay, se detiene."""
    for e in etiquetas:
        b = boton(dialogo, e)
        if b:
            HwndWrapper(b).click()
            time.sleep(0.5)
            return
    raise Alto(f"En el cuadro «{mensaje()[1]}» no encontré el botón {' / '.join(etiquetas)}. Púlsalo tú.")


def esperar(cond, seg=3.0):
    fin = time.time() + seg
    while time.time() < fin:
        r = cond()
        if r:
            return r
        time.sleep(0.08)
    return cond()


def esc(txt):
    return "".join("{" + ch + "}" if ch in "+^%~(){}[]" else ch for ch in str(txt))


def softcom_al_frente():
    try:
        return win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[1] == PID
    except Exception:
        return False


def asegurar_softcom():
    """Si moviste el foco a otra ventana (clic en Excel, Chrome…), lo devuelve a SOFTCOM
    al mismo campo donde estaba. Si el cursor cambió de campo, se detiene."""
    if softcom_al_frente():
        return
    esperado = foco()
    print("\n   ⏸ Hiciste clic fuera de SOFTCOM. Lo traigo de vuelta…", end="")
    for _ in range(20):
        try:
            HwndWrapper(PRINCIPAL).set_focus()
        except Exception:
            pass
        time.sleep(0.5)
        if softcom_al_frente():
            break
    if not softcom_al_frente():
        raise Alto("No pude volver a SOFTCOM. No hagas clic en otras ventanas mientras escribe.")
    if esperado and foco() != esperado:
        raise Alto(f"El cursor de SOFTCOM cambió de campo ({donde_esta_foco()}). Revisa la pantalla.")
    print(" listo, sigo.")


def teclas(s):
    asegurar_softcom()
    send_keys(s, pause=0.03, with_spaces=True)
    time.sleep(PAUSA)


def sin_mensaje(donde):
    h, t = mensaje()
    if h:
        raise Alto(f"SOFTCOM mostró un mensaje {donde}:\n      «{t}»\n   Ciérralo y revisa ese dato.")


# ============================================================ EXCEL
def com(fn, intentos=10):
    """Ejecuta una llamada a Excel; si Excel está ocupado (editando una celda, un cuadro
    abierto) reintenta unos segundos en vez de caerse."""
    for k in range(intentos):
        try:
            return fn()
        except Exception as e:
            ocupado = any(x in str(e).lower() for x in ("rechaz", "reject", "-2147418111", "-2146777998", "busy", "ocupad"))
            if not ocupado or k == intentos - 1:
                raise
            if k == 0:
                print("\n   ⏸ Excel está ocupado (¿editando una celda?). Presiona Enter o Esc en Excel…", end="")
            time.sleep(1.5)


def abrir_libro():
    import xlwings as xw
    nombre = os.path.basename(RUTA_EXCEL).lower()
    for app in xw.apps:
        for b in app.books:
            if b.name.lower() == nombre:
                return b
    if not os.path.exists(RUTA_EXCEL):
        raise Alto(f"No encuentro el Excel en {RUTA_EXCEL}. Cambia RUTA_EXCEL al inicio del bot.")
    print("(Abriendo el Excel…)")
    return xw.Book(RUTA_EXCEL)


def leer_salidas(libro, hechas=(), avisar=True):
    """Arma TODAS las salidas pendientes desde la hoja PARTE.
    Fecha / trabajador / área se toman de las columnas B, F y G (vacío = el de arriba).
    Una salida = filas SEGUIDAS con la misma fecha, trabajador y área (igual que la columna O)."""
    try:
        com(lambda: libro.app.calculate())        # solo lo que cambió (CalculateFull era lento)
    except Exception:
        pass
    cfg = {r[0]: r[1] for r in com(lambda: libro.sheets["CONFIG"].range("A2:C9").value) if r[0]}
    ccosto = {}
    for r in com(lambda: libro.sheets["CCOSTO"].range("A2:D60").value):
        if r[0] not in (None, ""):
            cod = str(int(r[0])) if isinstance(r[0], float) and r[0].is_integer() else str(r[0]).strip()
            ccosto[cod] = (r[1], r[3])
    hoja = libro.sheets["PARTE"]
    ultima = max(804, com(lambda: hoja.used_range.last_cell.row))
    parte = com(lambda: hoja.range(f"A5:S{ultima}").value)

    def cc_cod(v):
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v or "").split(" - ")[0].strip()

    # Una salida = un PEDIDO en el mostrador: filas SEGUIDAS del mismo trabajador, área y fecha.
    fecha_ant = trab_ant = cc_ant = None
    grupos, malas, clave_prev = [], [], None
    for i, f in enumerate(parte, start=5):
        fecha_ant = f[1] if f[1] not in (None, "") else fecha_ant
        trab_ant = f[5] if str(f[5] or "").strip() else trab_ant
        cc_ant = f[6] if f[6] not in (None, "") else cc_ant
        if not (f[2] or f[3]):
            continue                                   # fila vacía
        if f[15] not in (None, "") or i in hechas:
            clave_prev = None; continue                # ya REGISTRADA
        if str(f[13] or "").strip().upper() != "OK":
            malas.append(f"   fila {i}: {str(f[8] or f[2] or f[3])[:40]:<40} · {f[13] or '(sin calcular)'}")
            clave_prev = None; continue
        if not isinstance(fecha_ant, (datetime.date, datetime.datetime)):
            malas.append(f"   fila {i}: sin FECHA válida"); clave_prev = None; continue
        fe = fecha_ant.date() if isinstance(fecha_ant, datetime.datetime) else fecha_ant
        clave = (fe, str(trab_ant or "").strip().upper(), cc_cod(cc_ant))
        if clave != clave_prev:
            grupos.append((clave, []))
        grupos[-1][1].append((i, f, clave[1]))
        clave_prev = clave
    if malas and avisar:
        print("Filas que NO voy a registrar (corrígelas en el Excel):")
        print("\n".join(malas)); print()

    salidas = []
    for (fe, _trab, cc), filas in grupos:
        area, sol_area = ccosto.get(cc, ("", ""))
        sol = filas[0][1][17] if filas[0][1][17] not in (None, "") else sol_area
        sol = f"{int(sol):03d}" if isinstance(sol, (int, float)) else str(sol or "").strip().zfill(3)
        if not sol.strip("0"):
            raise Alto(f"El área {cc} no tiene SOLICITANTE en la hoja CCOSTO.")
        nombres, lineas = [], []
        for i, f, t in filas:
            if t and t not in nombres:
                nombres.append(t)
            codigo = str(f[7] or f[3] or "").strip().upper()
            if not codigo:
                raise Alto(f"La fila {i} no tiene CÓDIGO SOFTCOM (columna H). En Excel presiona F9, guarda y vuelve a correr.")
            try:
                cant = float(f[4])
            except (TypeError, ValueError):
                raise Alto(f"La fila {i} tiene una CANTIDAD que no es número: '{f[4]}'.")
            if cant <= 0:
                raise Alto(f"La fila {i} tiene CANTIDAD {f[4]} (debe ser mayor que 0).")
            stock = f[10] if isinstance(f[10], (int, float)) else None
            lineas.append(dict(fila=i, codigo=codigo, desc=str(f[8] or ""), stock=stock,
                               cant=str(int(cant)) if cant.is_integer() else repr(round(cant, 4))))
        salidas.append(dict(fecha=datetime.datetime.combine(fe, datetime.time()), cc=cc, area=area, sol=sol,
                            glosa=", ".join(nombres)[:60], lineas=lineas,
                            mov=str(cfg.get("Código de movimiento", "PR")),
                            ref=str(cfg.get("Tipo doc. referencia 1", "RQ"))))
    salidas.sort(key=lambda x: x["fecha"])   # por fecha (así cambia la fecha de SOFTCOM lo menos posible)
    return salidas


def anotar_en_excel(libro, filas, nro):
    hoja = libro.sheets["PARTE"]
    for fila in filas:
        com(lambda: setattr(hoja.range(f"P{fila}"), "value", nro))
    com(lambda: libro.save())


# ============================================================ HISTORIAL (bot_registro.txt)
REGISTRO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bot_registro.txt")


def anotar_en_registro(s, nro):
    """Se llama APENAS SOFTCOM da el número: si después falla Excel, aquí queda la prueba."""
    linea = (f"{datetime.datetime.now():%d/%m/%Y %H:%M}\t{nro}\t{s['fecha']:%d/%m/%Y}\t{s['cc']}\t"
             f"{s['glosa']}\t" + "; ".join(f"{l['codigo']} x{l['cant']}" for l in s['lineas'])
             + "\tfilas=" + ",".join(str(l["fila"]) for l in s["lineas"]))
    try:
        with open(REGISTRO, "a", encoding="utf-8") as fh:
            fh.write(linea + "\n")
            fh.flush(); os.fsync(fh.fileno())
    except Exception as e:
        print(f"  ⚠ No pude escribir en {REGISTRO} ({e}). Anota tú: {nro}")


def recuperar_pendientes(libro):
    """Busca en el historial salidas que SOFTCOM ya grabó pero que no quedaron en la columna P
    (el bot se cayó entre grabar y guardar el Excel). Así no se registran dos veces."""
    if not os.path.exists(REGISTRO):
        return
    hoja = libro.sheets["PARTE"]
    faltan = []
    try:
        lineas = open(REGISTRO, encoding="utf-8").read().splitlines()
    except Exception:
        return
    for linea in lineas[-200:]:
        partes = linea.split("\t")
        if len(partes) < 7 or not partes[-1].startswith("filas="):
            continue                                       # líneas viejas (v3.7) no traen filas
        nro, prods = partes[1], partes[5].split("; ")
        filas = [int(x) for x in partes[-1][6:].split(",") if x.strip().isdigit()]
        vacias = []
        for fila, prod in zip(filas, prods):
            p = com(lambda: hoja.range(f"P{fila}").value)
            h = com(lambda: hoja.range(f"H{fila}").value)
            if p in (None, "") and str(h or "").strip().upper() == prod.split(" x")[0]:
                vacias.append(fila)
        if vacias:
            faltan.append((nro, vacias, linea))
    if not faltan:
        return
    print("⚠ SOFTCOM ya grabó estas salidas pero NO quedaron anotadas en la columna P:")
    for nro, filas, _ in faltan:
        print(f"   {nro} → filas {', '.join(map(str, filas))}")
    resp = input("¿Las anoto yo en el Excel? SI = anotar · otra cosa = parar y revisar a mano: ").strip().upper()
    if resp != "SI":
        raise Alto("Revisa esas filas: escribe el N° PS en la columna P antes de volver a correr el bot.")
    for nro, filas, _ in faltan:
        anotar_en_excel(libro, filas, nro)
    print("   ✔ Anotadas y Excel guardado.\n")


# ============================================================ MENÚ Y FECHA DE SOFTCOM
MENU_FECHA = ""     # si el bot no encuentra solo el menú, aquí va la ruta exacta, ej. "Almacén > Selección Almacén y Fecha"
MENU_SALIDA = ""    # ej. "Almacén > Movimientos > Registro de Salida"
PRINCIPAL = 0


def _norm(t):
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFD", t.lower()) if unicodedata.category(c) != "Mn")


_u = ctypes.windll.user32
_u.GetMenu.restype = wintypes.HMENU; _u.GetMenu.argtypes = [wintypes.HWND]
_u.GetSubMenu.restype = wintypes.HMENU; _u.GetSubMenu.argtypes = [wintypes.HMENU, ctypes.c_int]
_u.GetMenuItemCount.argtypes = [wintypes.HMENU]
_u.GetMenuItemID.restype = wintypes.UINT; _u.GetMenuItemID.argtypes = [wintypes.HMENU, ctypes.c_int]
_u.GetMenuStringW.argtypes = [wintypes.HMENU, wintypes.UINT, wintypes.LPWSTR, ctypes.c_int, wintypes.UINT]


def items_menu(hmenu=None, ruta=()):
    u = _u
    if hmenu is None:
        hmenu = u.GetMenu(PRINCIPAL)
    if not hmenu:
        return
    for i in range(max(u.GetMenuItemCount(hmenu), 0)):
        buf = ctypes.create_unicode_buffer(256)
        u.GetMenuStringW(hmenu, i, buf, 256, 0x400)
        txt = buf.value.replace("&", "").split("\t")[0].strip()
        sub = u.GetSubMenu(hmenu, i)
        if sub:
            yield from items_menu(sub, ruta + (txt,))
        elif txt:
            yield ruta + (txt,), u.GetMenuItemID(hmenu, i)


def buscar_menu(fija, debe, no_debe=()):
    items = list(items_menu())
    if fija:
        for ruta, mid in items:
            if _norm(" > ".join(ruta)) == _norm(fija):
                return ruta, mid
    cand = [(r, m) for r, m in items if all(d in _norm(" ".join(r)) for d in debe)
            and not any(x in _norm(r[-1]) for x in no_debe)]
    if len(cand) > 1:
        cand = [(r, m) for r, m in cand if _norm(r[0]).startswith("almac")] or cand
    if len(cand) > 1:
        # preferir la opción que más se parece a "Registro de Salida" / "Selección Almacén y Fecha"
        def puntos(rm):
            ult = _norm(rm[0][-1])
            p = 0
            if "salida" in debe:
                p += 5 * ult.startswith("registro de salida") + 3 * ("registro" in ult) + 2 * (ult in ("salida", "salidas"))
            if "fecha" in debe:
                p += 5 * ("seleccion" in ult) + 3 * ("almacen" in ult) + 2 * ("cambio" in ult)
            return p
        cand.sort(key=puntos, reverse=True)
        if puntos(cand[0]) > puntos(cand[1]):
            cand = cand[:1]
    if len(cand) == 1:
        print(f"   · menú: {' > '.join(cand[0][0])}")
        return cand[0]
    lista = "\n".join("      " + " > ".join(r) for r, _ in (cand or items))
    raise Alto(f"No sé cuál opción del menú usar ({' + '.join(debe)}). Opciones:\n{lista}\n"
               f"   Mándale a Claude una captura de esto.")


def abrir_por_menu(ruta_mid, titulo):
    ruta, mid = ruta_mid
    win32gui.PostMessage(PRINCIPAL, win32con.WM_COMMAND, mid, 0)
    h = esperar(lambda: form(titulo), 6)
    if not h:
        raise Alto(f"Usé el menú '{' > '.join(ruta)}' pero no se abrió '{titulo}'.")
    time.sleep(0.5)
    return h


def cerrar_almovi20():
    alm = form("ALMOVI20")
    if alm:
        win32gui.PostMessage(alm, win32con.WM_CLOSE, 0, 0)
        if not esperar(lambda: not form("ALMOVI20"), 4):
            raise Alto("No pude cerrar ALMOVI20.")
        h, t = mensaje()
        if h:
            raise Alto(f"Al cerrar ALMOVI20 SOFTCOM preguntó: «{t}». Respóndelo tú y vuelve a correr.")


def cambiar_fecha(fecha):
    print(f"   · Cambiando la fecha de SOFTCOM a {fecha:%d/%m/%Y}…")
    cerrar_almovi20()
    sel = abrir_por_menu(buscar_menu(MENU_FECHA, ["fecha"], ["reporte", "consulta", "listado"]), "ALSELE")
    mask = control(sel, (162, 65), "MSMaskWndClass")
    if not mask:
        raise Alto("No encontré el campo de fecha en 'Selección Almacén y Fecha'.")
    objetivo = fecha.strftime("%d/%m/%y")
    HwndWrapper(sel).set_focus(); time.sleep(0.2)
    HwndWrapper(mask).set_focus(); time.sleep(0.2)
    win32gui.SendMessage(mask, win32con.WM_SETTEXT, 0, objetivo); time.sleep(0.2)
    if texto(mask).strip() != objetivo:
        teclas("{HOME}"); teclas(fecha.strftime("%d%m%y"))
    if texto(mask).strip() != objetivo:
        raise Alto(f"No pude escribir la fecha (quedó '{texto(mask)}'). Cámbiala tú y vuelve a correr.")
    teclas("{ENTER}")
    if form("ALSELE") and not mensaje()[0]:
        b = control(sel, (85, 118), "ThunderRT6CommandButton")
        if not b:
            raise Alto("No encontré el botón Aceptar de 'Selección Almacén y Fecha'.")
        HwndWrapper(b).click(); time.sleep(0.5)
    dlg = esperar(lambda: (lambda m: m if m[0] else None)(mensaje()), 4)
    if dlg:
        if "guardar" in dlg[1].lower():
            pulsar(dlg[0], "Sí", "Si")
        else:
            raise Alto(f"SOFTCOM dijo: «{dlg[1]}» al cambiar la fecha.")
    if not esperar(lambda: not form("ALSELE"), 4):
        raise Alto("La ventana 'Selección Almacén y Fecha' no se cerró.")


RECUERDO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bot_menu_salida.txt")
NO_SALIDA = ["reporte", "consulta", "listado", "anula", "kardex", "guia", "imprim", "elimin", "borra",
             "proceso", "cierre", "regener", "recalcul", "actualiz", "transfer", "genera", "masiv", "import"]


def abrir_almovi20():
    """Abre Registro de Salida (ALMOVI20) desde el menú.
    1) usa la opción que ya aprendió (bot_menu_salida.txt);
    2) si no, prueba las opciones del menú que dicen 'salida' hasta que se abra ALMOVI20 y la recuerda;
    3) si nada funciona, te pide abrirla a mano y sigue."""
    items = list(items_menu())
    fija = MENU_SALIDA
    try:
        if not fija and os.path.exists(RECUERDO):
            fija = open(RECUERDO, encoding="utf-8").read().strip()
    except Exception:
        pass
    cand = []
    if fija:
        cand += [(r, m) for r, m in items if _norm(" > ".join(r)) == _norm(fija)]
    otras = [(r, m) for r, m in items if "salida" in _norm(" ".join(r))
             and not any(x in _norm(r[-1]) for x in NO_SALIDA) and (r, m) not in cand]
    def puntos(rm):
        ult = _norm(rm[0][-1])
        return (5 * ult.startswith("registro de salida") + 3 * ("registro" in ult)
                + 2 * (ult in ("salida", "salidas")) + 1 * _norm(rm[0][0]).startswith("almac"))
    otras.sort(key=puntos, reverse=True)
    for ruta, mid in (cand + otras)[:5]:
        antes = set(formularios())
        win32gui.PostMessage(PRINCIPAL, win32con.WM_COMMAND, mid, 0)
        alm = esperar(lambda: form("ALMOVI20"), 4)
        if alm:
            print(f"   · menú: {' > '.join(ruta)}")
            try:
                open(RECUERDO, "w", encoding="utf-8").write(" > ".join(ruta))
            except Exception:
                pass
            time.sleep(0.5)
            return alm
        # se abrió otra ventana: la cierro y pruebo la siguiente opción
        for h in set(formularios()) - antes:
            try:
                if win32gui.GetClassName(h) == "#32770":
                    b = boton(h, "No") or boton(h, "Cancelar") or boton(h, "Salir") or boton(h, "Aceptar")
                    if b:
                        HwndWrapper(b).click()
                else:
                    win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
            except Exception:
                pass
        time.sleep(0.6)
    print("\n   ⚠ No pude abrir 'Registro de Salida' solo.")
    if otras or cand:
        print("   Probé: " + " | ".join(" > ".join(r) for r, _ in (cand + otras)[:5]))
    input("   → Ábrela TÚ en SOFTCOM (Registro de Salida) y presiona Enter aquí para seguir...")
    alm = esperar(lambda: form("ALMOVI20"), 5)
    if not alm:
        raise Alto("Sigo sin ver ALMOVI20 abierta.")
    return alm


def preparar_almovi20(fecha):
    """Deja ALMOVI20 abierta, vacía y con la fecha de la salida."""
    objetivo = fecha.strftime("%d/%m/%Y")
    alm = form("ALMOVI20")
    if alm:
        if texto(control(alm, A["mov"], "ThunderRT6TextBox")).strip():
            raise Alto("ALMOVI20 no está vacía (Código de Movimiento tiene algo). Termínala o ciérrala y vuelve a correr.")
        if texto(control(alm, A["fecha"], "MSMaskWndClass")).strip() == objetivo:
            return alm
        cambiar_fecha(fecha)
    else:
        # sin ALMOVI20: miro la fecha abriéndola; si no coincide, cambio la fecha
        alm = abrir_almovi20()
        if texto(control(alm, A["fecha"], "MSMaskWndClass")).strip() == objetivo:
            return alm
        cambiar_fecha(fecha)
    alm = abrir_almovi20()
    if texto(control(alm, A["fecha"], "MSMaskWndClass")).strip() != objetivo:
        raise Alto(f"Cambié la fecha pero ALMOVI20 sigue con {texto(control(alm, A['fecha'], 'MSMaskWndClass'))}.")
    return alm


# ============================================================ PASOS EN SOFTCOM
def avanzar_hasta(f, pos, nombre, maximo=7):
    """Presiona ENTER (como tú) hasta que el cursor llegue al campo pedido."""
    for _ in range(maximo):
        if foco_en(f, pos):
            return
        antes = foco()
        teclas("{ENTER}")
        sin_mensaje(f"camino a {nombre}")
        esperar(lambda: foco() != antes, 1.5)
    if not foco_en(f, pos):
        raise Alto(f"No llegué al campo {nombre}. El cursor está en: {donde_esta_foco()}")


def _igual(leido, valor):
    a, b = _norm(leido).strip().upper(), _norm(str(valor)).strip().upper()
    return a == b or (len(a) >= 20 and b.startswith(a))   # tolera que SOFTCOM corte una glosa larga


def escribir_y_enter(f, valor, pos_siguiente, nombre):
    campo = foco()
    teclas(esc(valor))
    leido = texto(campo).strip() if campo else str(valor)
    if not _igual(leido, valor):
        raise Alto(f"En {nombre} escribí '{valor}' pero SOFTCOM muestra '{leido}'.\n"
                   f"   (¿Bloq Mayús, teclado en otro idioma, o alguien tocó el teclado?) No grabé nada.")
    teclas("{ENTER}")
    sin_mensaje(f"después de {nombre} = {valor}")
    if not esperar(lambda: foco_en(f, pos_siguiente), 3):
        raise Alto(f"SOFTCOM no aceptó {nombre} = '{valor}'. El cursor quedó en: {donde_esta_foco()}")
    print(f"   ✓ {nombre:<14} {valor}")


def cabecera(s):
    alm = form("ALMOVI20")
    if not alm:
        raise Alto("No veo abierta 'Registro de Salida - ALMOVI20'. Ábrela vacía y vuelve a correr.")
    mov = control(alm, A["mov"], "ThunderRT6TextBox")
    if texto(mov).strip():
        raise Alto("ALMOVI20 no está vacía (Código de Movimiento ya tiene algo). Ciérrala, ábrela de nuevo y vuelve a correr.")
    fecha_sc = texto(control(alm, A["fecha"], "MSMaskWndClass")).strip()
    if fecha_sc != s["fecha"].strftime("%d/%m/%Y"):
        raise Alto(f"ALMOVI20 está con fecha {fecha_sc} y la salida es del {s['fecha']:%d/%m/%Y}.")
    HwndWrapper(alm).set_focus(); time.sleep(0.3)
    HwndWrapper(mov).set_focus(); time.sleep(0.3)
    if not foco_en(alm, A["mov"]):
        raise Alto(f"No pude poner el cursor en Código de Movimiento ({donde_esta_foco()}).")

    escribir_y_enter(alm, s["mov"], A["tipo_ref1"], "Movimiento")
    conv = texto(control(alm, A["tipo_conv"], "ThunderRT6ComboBox")).strip()
    if not conv:
        raise Alto("SOFTCOM no validó el movimiento (Tipo Conversión quedó vacío).")
    escribir_y_enter(alm, s["ref"], A["nro_ref1"], "Tipo Ref. 1")
    avanzar_hasta(alm, A["cc"], "Centro de Costo")
    escribir_y_enter(alm, s["cc"], A["sol"], "Centro Costo")
    escribir_y_enter(alm, s["sol"], A["glosa1"], "Solicitante")
    escribir_y_enter(alm, s["glosa"], A["glosa2"], "Glosa 1")
    avanzar_hasta(alm, A["grabar"], "botón Grabar")
    return alm


def detalle(s):
    teclas("{ENTER}")  # Enter sobre el botón Grabar → abre Detalles de Salida
    det = esperar(lambda: form("Detalles de Salida"), 5)
    sin_mensaje("al pasar al detalle")
    if not det:
        raise Alto("No se abrió 'Detalles de Salida'.")
    items = control(det, D["items"], "ThunderRT6TextBox")
    time.sleep(0.4)
    for k, ln in enumerate(s["lineas"], 1):
        antes = texto(items)
        if k > 1:
            teclas("{HOME}")
        teclas("{ENTER}")
        ed = esperar(lambda: casilla(det, D["col_codigo"]), 2)
        if not ed:
            raise Alto(f"No se abrió la casilla de código (línea {k}). Cursor en: {donde_esta_foco()}")
        if texto(ed).strip():  # cayó sobre una fila ya llena → bajar a la siguiente
            teclas("{ESC}{DOWN}{ENTER}")
            ed = esperar(lambda: casilla(det, D["col_codigo"]), 2)
            if not ed or texto(ed).strip():
                raise Alto(f"No encontré una fila vacía para la línea {k}. Cursor en: {donde_esta_foco()}")
        teclas(esc(ln["codigo"]))
        if texto(ed).strip().upper() != ln["codigo"]:
            raise Alto(f"El código no quedó bien escrito: '{texto(ed).strip()}' (debía ser {ln['codigo']}).")
        for _ in range(ENTERS_CODIGO_A_CANTIDAD):
            teclas("{ENTER}")
            sin_mensaje(f"con el código {ln['codigo']}")
        teclas(esc(ln["cant"]))
        if not casilla(det, D["col_cantidad"], tol=15):
            raise Alto(f"La cantidad no cayó en la columna Cantidad (línea {k}). Cursor en: {donde_esta_foco()}\n"
                       f"   Revisa la línea en pantalla y no grabes.")
        teclas("{ENTER}")
        sin_mensaje(f"con la cantidad de {ln['codigo']}")
        if not esperar(lambda: texto(items) != antes, 3):
            raise Alto(f"SOFTCOM no registró la línea {k} ({ln['codigo']}).")
        print(f"   ✓ línea {k}: {ln['codigo']}  x {ln['cant']}   {ln['desc'][:40]}")
    nro = texto(control(det, D["nro_doc"], "ThunderRT6TextBox")).strip()
    return det, nro


def grabar(det, nro_pantalla=""):
    HwndWrapper(det).set_focus(); time.sleep(0.4)
    teclas("%f")  # Alt+F
    dlg = esperar(lambda: (lambda m: m if "Documento Generado" in m[1] else None)(mensaje()), 15)
    if not dlg:
        raise Alto(f"No apareció 'Documento Generado' (mensaje en pantalla: {mensaje()[1] or '(ninguno)'}).\n"
                   f"   OJO: puede que SOFTCOM SÍ la haya grabado. Antes de volver a correr, revisa en SOFTCOM\n"
                   f"   si existe el documento PS-{nro_pantalla or '…'} y, si existe, escríbelo en la columna P.")
    m = re.search(r"PS-\d+", dlg[1])
    nro = m.group(0) if m else dlg[1].replace("Documento Generado Número", "").strip()
    if not m:
        print(f"  ⚠ No reconocí el número en «{dlg[1]}». Lo anoto tal cual: revisa la columna P.")
    elif nro_pantalla and nro_pantalla.lstrip("0") and nro[3:].lstrip("0") != nro_pantalla.lstrip("0"):
        print(f"  ⚠ La pantalla decía N° {nro_pantalla} y SOFTCOM generó {nro}. Me quedo con {nro}.")
    pulsar(dlg[0], "Aceptar")
    imp = esperar(lambda: (lambda m: m if "Imprimir" in m[1] else None)(mensaje()), 5)
    if imp:
        pulsar(imp[0], "No")
    return nro


# ============================================================ PRINCIPAL
MODOS = ("completo", "preguntar", "cabecera", "menu", "revisar")


def main():
    global PID, TID, PRINCIPAL
    modo = (sys.argv[1].lower() if len(sys.argv) > 1 else "completo")
    if modo not in MODOS:
        raise Alto(f"No conozco el modo '{modo}'. Usa uno de: {', '.join(MODOS)} (o nada = completo).")
    print(f"=== BOT DE SALIDAS v3.8 · modo {modo.upper()} ===\n")
    PRINCIPAL, PID = ventana_principal()
    if not PID:
        raise Alto("No veo SOFTCOM abierto (ventana principal del almacén).")
    try:
        print(f"(SOFTCOM: {psutil.Process(PID).name()} · '{win32gui.GetWindowText(PRINCIPAL)[:60]}')")
    except Exception:
        pass
    TID = win32process.GetWindowThreadProcessId(PRINCIPAL)[0]

    if modo == "menu":   # diagnóstico: lista el menú de SOFTCOM
        ruta = os.path.join(os.path.expanduser("~"), "Desktop", "softcom_menu.txt")
        with open(ruta, "w", encoding="utf-8") as fh:
            for r, mid in items_menu():
                fh.write(f"{mid:>6}  {' > '.join(r)}\n")
        print("Menú guardado en", ruta); return

    libro = abrir_libro()
    recuperar_pendientes(libro)
    salidas = leer_salidas(libro)
    if not salidas:
        raise Alto("No hay salidas pendientes (todas las filas en OK ya tienen N° DOC).")
    print(f"Salidas pendientes: {len(salidas)}")
    for k, s in enumerate(salidas, 1):
        print(f"   {k}. {s['fecha']:%d/%m} · {s['cc']:<5} {str(s['area'])[:28]:<28} · {len(s['lineas'])} productos · {s['glosa'][:35]}")
        for ln in s["lineas"]:
            if ln["stock"] is not None and ln["stock"] < 0:
                print(f"        ⚠ fila {ln['fila']}: {ln['codigo']} queda con stock estimado {ln['stock']:g} (SOFTCOM podría rechazarla)")
    if modo == "revisar":
        print("\nModo REVISAR: no toqué SOFTCOM."); return
    print("\nNo toques mouse ni teclado mientras escribe en SOFTCOM.\n")

    todas_inicio = not PREGUNTAR_ANTES_DE_GRABAR and modo != "preguntar"
    if todas_inicio:
        print("Modo AUTOMÁTICO: graba cada salida sin preguntar. Para detenerlo: cierra esta ventana negra.\n")
    hechas, todas, n_ok = set(), todas_inicio, 0
    resumen = []
    while True:
        pend = leer_salidas(libro, hechas, avisar=False)
        if not pend:
            break
        s = pend[0]
        print(f"\n▶ Salida {s['fecha']:%d/%m/%Y} · CC {s['cc']} {s['area']} · solicitante {s['sol']}")
        print(f"  Glosa: {s['glosa']}")
        preparar_almovi20(s["fecha"])
        cabecera(s)
        if modo == "cabecera":
            print("\nCabecera lista (cursor en el botón Grabar). NO entré al detalle. Cierra ALMOVI20 con S, sin grabar.")
            return
        det, nro_pantalla = detalle(s)
        if not todas:
            print("\nREVISA LA PANTALLA DE SOFTCOM.")
            resp = input("¿Grabo? SI = esta · TODAS = esta y las demás sin preguntar · otra cosa = parar: ").strip().upper()
            if resp == "TODAS":
                todas = True
            elif resp != "SI":
                print("\nNo grabé. La salida quedó en pantalla: termínala o anúlala tú a mano.")
                break
        nro = grabar(det, nro_pantalla)
        anotar_en_registro(s, nro)          # lo PRIMERO: si algo falla después, queda la prueba
        n_ok += 1
        hechas.update(ln["fila"] for ln in s["lineas"])
        print(f"  ✔ SOFTCOM generó {nro}")
        resumen.append(f"   {nro}  {s['fecha']:%d/%m}  {s['cc']:<5} {s['glosa'][:30]:<30} {len(s['lineas'])} prod.")
        try:
            anotar_en_excel(libro, [ln["fila"] for ln in s["lineas"]], nro)
            print(f"  ✔ Anotado en la columna P ({len(s['lineas'])} filas) y Excel guardado.")
        except Exception as e:
            print(f"  ⚠ No pude anotarlo en el Excel ({e}). Escribe {nro} a mano en la columna P de esas filas.")
            break
    print(f"\nListo: {n_ok} salida(s) registrada(s).")
    if resumen:
        print("\n".join(resumen))
        print(f"\n(Historial en {REGISTRO})")


if __name__ == "__main__":
    try:
        main()
    except Alto as a:
        print(f"\n■ ME DETUVE: {a}")
    except KeyboardInterrupt:
        print("\n■ Detenido con Ctrl+C. Revisa SOFTCOM: la salida en pantalla NO se grabó si no viste 'SOFTCOM generó'.")
    except Exception:
        import traceback
        print("\n*** OCURRIÓ UN ERROR. Mándale a Claude una captura de esto: ***\n")
        traceback.print_exc()
    input("\nPresiona Enter para cerrar...")
