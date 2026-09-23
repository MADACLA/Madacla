# -*- coding: utf-8 -*-
"""La caja de Madacla: cobrar, ticket y cierre. Puente hasta Verifactu.

    python caja.py

Abre un servidor SOLO en este ordenador (127.0.0.1) y la pantalla se ve en
el navegador a pantalla completa (arrancar-caja.cmd lo abre solo). No
necesita internet ni la torre de casa: todo vive en este ordenador.

LO QUE HACE
  - Cobrar: botones por familia (Ropa, Bolsos...) + precio tecleado, varias
    prendas por ticket, efectivo o tarjeta. Con efectivo se abre el cajon.
  - Ticket numerado (2026-000001...) con los datos legales de la factura
    simplificada: titular, NIF, direccion, fecha y hora, lineas, total e
    «IVA incluido». Ticket regalo (sin precios), reimprimir, anular.
  - Cierre de caja del dia y Excel (CSV) del mes para la gestora.
  - Copia de seguridad diaria en copias/ y, si esta puesta, en Google Drive.

LO QUE NO SE PUEDE HACER, A PROPOSITO
  - Borrar ni cambiar una venta. La base de datos lo impide con triggers:
    si Maricarmen se equivoca, «Anular» crea un ticket nuevo que anula al
    anterior (lineas en negativo). Asi cuadra todo con Hacienda.
  - Cada ticket lleva un hash encadenado con el anterior: si alguien tocara
    el fichero a mano, la cadena se rompe y se nota.

CADUCA: a la madre de Claudia, autonoma, Verifactu le obliga desde el
1 de julio de 2027. Este programa NO esta certificado: ese dia hay que
pasar a uno que lo este (Agora u otro). Hasta entonces, es legal.
"""
from __future__ import annotations

import csv
import ctypes
import hashlib
import json
import os
import shutil
import subprocess
import sqlite3
import sys
import threading
import time
from datetime import datetime, date
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, parse_qs
from urllib.request import urlopen

AQUI = Path(__file__).resolve().parent
CONFIG = AQUI / "config.json"
CONFIG_EJEMPLO = AQUI / "config.ejemplo.json"
DATOS = AQUI / "datos"
DB = DATOS / "ventas.sqlite"
COPIAS = AQUI / "copias"
LOGO_TICKET = AQUI / "logo-ticket.bin"
REGISTRO = AQUI / "registro-caja.txt"

ANCHO = 42          # columnas del ticket (Bixolon SRP-350plus, letra normal)
COPIAS_QUE_GUARDAR = 60


# ---------------------------------------------------------------------------
# CONFIGURACION
# ---------------------------------------------------------------------------

def leer_config() -> dict:
    fuente = CONFIG if CONFIG.exists() else CONFIG_EJEMPLO
    with open(fuente, encoding="utf-8") as f:
        cfg = json.load(f)
    cfg.setdefault("familias", ["Ropa", "Bolsos", "Zapatos", "Bisutería"])
    # «Otros» siempre (Claudia, 2026-09-16): para lo que no encaje en ninguna
    # (legumbres, un detalle...). Va en el programa y no en el config.json
    # del Lenovo, que no se toca en las actualizaciones.
    if "Otros" not in cfg["familias"]:
        cfg["familias"].append("Otros")
    cfg.setdefault("pie", [])
    cfg.setdefault("iva", 21)
    cfg.setdefault("pin", "0000")
    cfg.setdefault("impresora", "")
    cfg.setdefault("cajon", True)
    cfg.setdefault("carpeta_drive", "")
    cfg.setdefault("puerto", 8791)
    return cfg


CFG = leer_config()


def apuntar(texto: str) -> None:
    linea = f"{datetime.now():%Y-%m-%d %H:%M:%S}  {texto}"
    print(linea)
    try:
        with open(REGISTRO, "a", encoding="utf-8") as f:
            f.write(linea + "\n")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# BASE DE DATOS (ventas que no se borran)
# ---------------------------------------------------------------------------

ESQUEMA = """
CREATE TABLE IF NOT EXISTS tickets (
    id            INTEGER PRIMARY KEY,
    numero        TEXT NOT NULL UNIQUE,
    fecha         TEXT NOT NULL,
    hora          TEXT NOT NULL,
    tipo          TEXT NOT NULL CHECK (tipo IN ('venta', 'anulacion')),
    anula_a       TEXT,
    pago          TEXT NOT NULL CHECK (pago IN ('efectivo', 'tarjeta')),
    total         REAL NOT NULL,
    hash_anterior TEXT NOT NULL,
    hash          TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS lineas (
    id        INTEGER PRIMARY KEY,
    ticket_id INTEGER NOT NULL REFERENCES tickets(id),
    orden     INTEGER NOT NULL,
    familia   TEXT NOT NULL,
    importe   REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS cierres (
    id          INTEGER PRIMARY KEY,
    fecha       TEXT NOT NULL UNIQUE,
    hora        TEXT NOT NULL,
    efectivo    REAL NOT NULL,
    tarjeta     REAL NOT NULL,
    total       REAL NOT NULL,
    tickets     INTEGER NOT NULL,
    anulaciones INTEGER NOT NULL
);
-- LA LIBRETA DE «SE LO LLEVA A PROBAR» (2026-09-16). Aqui SI se puede
-- corregir y borrar: no es una venta, es la libreta de fiados de toda la
-- vida. Lo que se queda la clienta se cobra con un ticket normal, y esa
-- venta si va a las tablas de arriba, intocable.
CREATE TABLE IF NOT EXISTS prestamos (
    id        INTEGER PRIMARY KEY,
    fecha     TEXT NOT NULL,
    hora      TEXT NOT NULL,
    nombre    TEXT NOT NULL,
    nota      TEXT NOT NULL DEFAULT '',
    estado    TEXT NOT NULL DEFAULT 'abierto' CHECK (estado IN ('abierto', 'cerrado')),
    cerrado   TEXT
);
CREATE TABLE IF NOT EXISTS prestamo_lineas (
    id          INTEGER PRIMARY KEY,
    prestamo_id INTEGER NOT NULL REFERENCES prestamos(id) ON DELETE CASCADE,
    orden       INTEGER NOT NULL,
    familia     TEXT NOT NULL,
    importe     REAL NOT NULL,
    estado      TEXT NOT NULL DEFAULT 'fuera' CHECK (estado IN ('fuera', 'devuelto', 'vendido')),
    ticket      TEXT
);
-- Nada de esto se toca nunca: ni borrar ni editar. Un error se arregla con
-- un ticket de anulacion, no reescribiendo la historia.
CREATE TRIGGER IF NOT EXISTS tickets_no_borrar BEFORE DELETE ON tickets
    BEGIN SELECT RAISE(ABORT, 'Las ventas no se borran'); END;
CREATE TRIGGER IF NOT EXISTS tickets_no_editar BEFORE UPDATE ON tickets
    BEGIN SELECT RAISE(ABORT, 'Las ventas no se cambian'); END;
CREATE TRIGGER IF NOT EXISTS lineas_no_borrar BEFORE DELETE ON lineas
    BEGIN SELECT RAISE(ABORT, 'Las lineas no se borran'); END;
CREATE TRIGGER IF NOT EXISTS lineas_no_editar BEFORE UPDATE ON lineas
    BEGIN SELECT RAISE(ABORT, 'Las lineas no se cambian'); END;
CREATE TRIGGER IF NOT EXISTS cierres_no_borrar BEFORE DELETE ON cierres
    BEGIN SELECT RAISE(ABORT, 'Los cierres no se borran'); END;
CREATE TRIGGER IF NOT EXISTS cierres_no_editar BEFORE UPDATE ON cierres
    BEGIN SELECT RAISE(ABORT, 'Los cierres no se cambian'); END;
"""

_cerrojo_db = threading.Lock()


def conectar() -> sqlite3.Connection:
    DATOS.mkdir(exist_ok=True)
    con = sqlite3.connect(DB, timeout=10)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA journal_mode = WAL")
    con.execute("PRAGMA foreign_keys = ON")
    return con


def preparar_db() -> None:
    with conectar() as con:
        con.executescript(ESQUEMA)


def redondear(x: float) -> float:
    return round(float(x) + 1e-9, 2)


def hash_de(hash_anterior: str, numero: str, fecha: str, hora: str, tipo: str,
            pago: str, total: float, lineas: list[dict]) -> str:
    cuerpo = json.dumps([hash_anterior, numero, fecha, hora, tipo, pago,
                         f"{total:.2f}",
                         [[l["familia"], f"{l['importe']:.2f}"] for l in lineas]],
                        ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(cuerpo.encode("utf-8")).hexdigest()


def crear_ticket(lineas: list[dict], pago: str, tipo: str = "venta",
                 anula_a: str | None = None) -> dict:
    """Guarda un ticket (venta o anulacion). Devuelve el ticket completo."""
    if pago not in ("efectivo", "tarjeta"):
        raise ValueError("Forma de pago no valida")
    limpias = []
    for l in lineas:
        familia = str(l.get("familia", "")).strip()
        importe = redondear(l.get("importe", 0))
        if not familia or importe == 0:
            raise ValueError("Cada linea necesita familia e importe")
        limpias.append({"familia": familia, "importe": importe})
    if not limpias:
        raise ValueError("El ticket esta vacio")
    total = redondear(sum(l["importe"] for l in limpias))

    ahora = datetime.now()
    fecha, hora = ahora.strftime("%Y-%m-%d"), ahora.strftime("%H:%M:%S")
    with _cerrojo_db, conectar() as con:
        con.execute("BEGIN IMMEDIATE")
        anyo = ahora.strftime("%Y")
        cuantos = con.execute(
            "SELECT COUNT(*) FROM tickets WHERE numero LIKE ?", (anyo + "-%",)
        ).fetchone()[0]
        numero = f"{anyo}-{cuantos + 1:06d}"
        ultimo = con.execute(
            "SELECT hash FROM tickets ORDER BY id DESC LIMIT 1").fetchone()
        hash_anterior = ultimo["hash"] if ultimo else "0" * 64
        h = hash_de(hash_anterior, numero, fecha, hora, tipo, pago, total, limpias)
        cur = con.execute(
            "INSERT INTO tickets (numero, fecha, hora, tipo, anula_a, pago, total, "
            "hash_anterior, hash) VALUES (?,?,?,?,?,?,?,?,?)",
            (numero, fecha, hora, tipo, anula_a, pago, total, hash_anterior, h))
        tid = cur.lastrowid
        for i, l in enumerate(limpias, 1):
            con.execute(
                "INSERT INTO lineas (ticket_id, orden, familia, importe) VALUES (?,?,?,?)",
                (tid, i, l["familia"], l["importe"]))
    return leer_ticket(numero)


def leer_ticket(numero: str) -> dict | None:
    with conectar() as con:
        t = con.execute("SELECT * FROM tickets WHERE numero = ?", (numero,)).fetchone()
        if not t:
            return None
        lineas = con.execute(
            "SELECT familia, importe FROM lineas WHERE ticket_id = ? ORDER BY orden",
            (t["id"],)).fetchall()
        anulado_por = con.execute(
            "SELECT numero FROM tickets WHERE anula_a = ?", (numero,)).fetchone()
    return {
        "numero": t["numero"], "fecha": t["fecha"], "hora": t["hora"],
        "tipo": t["tipo"], "anula_a": t["anula_a"], "pago": t["pago"],
        "total": t["total"],
        "anulado_por": anulado_por["numero"] if anulado_por else None,
        "lineas": [{"familia": l["familia"], "importe": l["importe"]} for l in lineas],
    }


def anular_ticket(numero: str) -> dict:
    original = leer_ticket(numero)
    if not original:
        raise ValueError("Ese ticket no existe")
    if original["tipo"] != "venta":
        raise ValueError("Eso ya es una anulacion")
    if original["anulado_por"]:
        raise ValueError(f"Ya estaba anulado por el {original['anulado_por']}")
    negativas = [{"familia": l["familia"], "importe": -l["importe"]}
                 for l in original["lineas"]]
    return crear_ticket(negativas, original["pago"], "anulacion", numero)


def resumen_dia(fecha: str) -> dict:
    with conectar() as con:
        filas = con.execute(
            "SELECT * FROM tickets WHERE fecha = ? ORDER BY id", (fecha,)).fetchall()
        cierre = con.execute("SELECT * FROM cierres WHERE fecha = ?", (fecha,)).fetchone()
    efectivo = redondear(sum(t["total"] for t in filas if t["pago"] == "efectivo"))
    tarjeta = redondear(sum(t["total"] for t in filas if t["pago"] == "tarjeta"))
    return {
        "fecha": fecha,
        "efectivo": efectivo, "tarjeta": tarjeta,
        "total": redondear(efectivo + tarjeta),
        "ventas": sum(1 for t in filas if t["tipo"] == "venta"),
        "anulaciones": sum(1 for t in filas if t["tipo"] == "anulacion"),
        "tickets": [leer_ticket(t["numero"]) for t in filas],
        "cerrado": bool(cierre),
        "cierre_hora": cierre["hora"] if cierre else None,
    }


def cerrar_caja(fecha: str) -> dict:
    r = resumen_dia(fecha)
    if r["cerrado"]:
        raise ValueError(f"La caja del {fecha} ya estaba cerrada a las {r['cierre_hora']}")
    with _cerrojo_db, conectar() as con:
        con.execute(
            "INSERT INTO cierres (fecha, hora, efectivo, tarjeta, total, tickets, anulaciones) "
            "VALUES (?,?,?,?,?,?,?)",
            (fecha, datetime.now().strftime("%H:%M:%S"), r["efectivo"], r["tarjeta"],
             r["total"], r["ventas"], r["anulaciones"]))
    return resumen_dia(fecha)


def exportar_mes(mes: str) -> Path:
    """CSV (se abre en Excel) con todos los tickets del mes: para la gestora."""
    COPIAS.mkdir(exist_ok=True)
    destino = COPIAS / f"ventas-{mes}.csv"
    with conectar() as con:
        filas = con.execute(
            "SELECT t.numero, t.fecha, t.hora, t.tipo, t.anula_a, t.pago, t.total, "
            "l.orden, l.familia, l.importe FROM tickets t JOIN lineas l ON l.ticket_id = t.id "
            "WHERE t.fecha LIKE ? ORDER BY t.id, l.orden", (mes + "-%",)).fetchall()
        cierres = con.execute(
            "SELECT * FROM cierres WHERE fecha LIKE ? ORDER BY fecha", (mes + "-%",)).fetchall()
    # utf-8-sig y punto y coma: asi Excel en español lo abre bien a la primera
    with open(destino, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Ticket", "Fecha", "Hora", "Tipo", "Anula a", "Pago",
                    "Total ticket", "Linea", "Familia", "Importe"])
        for r in filas:
            w.writerow([r["numero"], r["fecha"], r["hora"], r["tipo"], r["anula_a"] or "",
                        r["pago"], f"{r['total']:.2f}".replace(".", ","), r["orden"],
                        r["familia"], f"{r['importe']:.2f}".replace(".", ",")])
        w.writerow([])
        w.writerow(["CIERRES DEL MES"])
        w.writerow(["Fecha", "Hora", "Efectivo", "Tarjeta", "Total", "Ventas", "Anulaciones"])
        for c in cierres:
            w.writerow([c["fecha"], c["hora"],
                        f"{c['efectivo']:.2f}".replace(".", ","),
                        f"{c['tarjeta']:.2f}".replace(".", ","),
                        f"{c['total']:.2f}".replace(".", ","),
                        c["tickets"], c["anulaciones"]])
    llevar_a_drive(destino)
    return destino


def comprobar_cadena() -> dict:
    """Recorre todos los tickets y comprueba que la cadena de hashes cuadra."""
    with conectar() as con:
        filas = con.execute("SELECT numero FROM tickets ORDER BY id").fetchall()
    anterior = "0" * 64
    for f in filas:
        t = leer_ticket(f["numero"])
        with conectar() as con:
            guardado = con.execute(
                "SELECT hash, hash_anterior FROM tickets WHERE numero = ?",
                (t["numero"],)).fetchone()
        esperado = hash_de(anterior, t["numero"], t["fecha"], t["hora"], t["tipo"],
                           t["pago"], t["total"], t["lineas"])
        if guardado["hash_anterior"] != anterior or guardado["hash"] != esperado:
            return {"ok": False, "roto_en": t["numero"], "tickets": len(filas)}
        anterior = guardado["hash"]
    return {"ok": True, "tickets": len(filas)}


# ---------------------------------------------------------------------------
# «SE LO LLEVA A PROBAR» (la libreta de fiados)
#
# Claudia (2026-09-16): «viene una chica que mi madre conoce, se lleva cuatro
# vestidos a probarselos en casa; mi madre apunta en una libreta los
# vestidos, el nombre y el precio. Manana devuelve tres y paga el que se
# queda, y lo tacha». Eso, tal cual, pero en la caja: se apunta con un
# toque, se marca lo devuelto, y lo que se queda pasa al ticket de siempre.
# ---------------------------------------------------------------------------

def crear_prestamo(nombre: str, nota: str, lineas: list[dict]) -> dict:
    nombre = " ".join(str(nombre or "").split())
    if not nombre:
        raise ValueError("Falta el nombre de la clienta")
    limpias = []
    for l in lineas:
        familia = str(l.get("familia", "")).strip()
        importe = redondear(l.get("importe", 0))
        if not familia or importe <= 0:
            raise ValueError("Cada prenda necesita familia y precio")
        limpias.append({"familia": familia, "importe": importe})
    if not limpias:
        raise ValueError("No hay ninguna prenda apuntada")
    ahora = datetime.now()
    with _cerrojo_db, conectar() as con:
        cur = con.execute(
            "INSERT INTO prestamos (fecha, hora, nombre, nota) VALUES (?,?,?,?)",
            (ahora.strftime("%Y-%m-%d"), ahora.strftime("%H:%M:%S"), nombre,
             str(nota or "").strip()[:200]))
        pid = cur.lastrowid
        for i, l in enumerate(limpias, 1):
            con.execute(
                "INSERT INTO prestamo_lineas (prestamo_id, orden, familia, importe) VALUES (?,?,?,?)",
                (pid, i, l["familia"], l["importe"]))
    return leer_prestamo(pid)


def leer_prestamo(pid: int) -> dict | None:
    with conectar() as con:
        p = con.execute("SELECT * FROM prestamos WHERE id = ?", (pid,)).fetchone()
        if not p:
            return None
        lineas = con.execute(
            "SELECT id, familia, importe, estado, ticket FROM prestamo_lineas "
            "WHERE prestamo_id = ? ORDER BY orden", (pid,)).fetchall()
    fuera = [l for l in lineas if l["estado"] == "fuera"]
    dias = (date.today() - datetime.strptime(p["fecha"], "%Y-%m-%d").date()).days
    return {
        "id": p["id"], "fecha": p["fecha"], "hora": p["hora"], "nombre": p["nombre"],
        "nota": p["nota"], "estado": p["estado"], "cerrado": p["cerrado"], "dias": dias,
        "pendiente": redondear(sum(l["importe"] for l in fuera)),
        "lineas": [dict(l) for l in lineas],
    }


def listar_prestamos(abiertos: bool = True) -> list[dict]:
    with conectar() as con:
        if abiertos:
            filas = con.execute(
                "SELECT id FROM prestamos WHERE estado = 'abierto' ORDER BY fecha, hora").fetchall()
        else:
            filas = con.execute(
                "SELECT id FROM prestamos WHERE estado = 'cerrado' ORDER BY cerrado DESC LIMIT 60").fetchall()
    return [leer_prestamo(f["id"]) for f in filas]


def nombres_prestamos() -> list[str]:
    """Las clientas que ya han pasado por la libreta, las ultimas primero:
    para apuntar a la habitual con un toque en vez de teclear."""
    with conectar() as con:
        filas = con.execute(
            "SELECT nombre, MAX(id) AS u FROM prestamos GROUP BY lower(nombre) "
            "ORDER BY u DESC LIMIT 40").fetchall()
    return [f["nombre"] for f in filas]


def _revisar_cierre(pid: int) -> None:
    """Si ya no queda nada fuera, la hoja de la libreta se cierra sola."""
    with conectar() as con:
        fuera = con.execute(
            "SELECT COUNT(*) FROM prestamo_lineas WHERE prestamo_id = ? AND estado = 'fuera'",
            (pid,)).fetchone()[0]
        if fuera == 0:
            con.execute("UPDATE prestamos SET estado = 'cerrado', cerrado = ? WHERE id = ?",
                        (datetime.now().strftime("%Y-%m-%d %H:%M"), pid))
        else:
            con.execute("UPDATE prestamos SET estado = 'abierto', cerrado = NULL WHERE id = ?", (pid,))


def marcar_linea(linea_id: int, estado: str) -> dict:
    """«Devuelve» (fuera -> devuelto) o deshacer (devuelto -> fuera)."""
    if estado not in ("fuera", "devuelto"):
        raise ValueError("Estado no valido")
    with _cerrojo_db, conectar() as con:
        l = con.execute("SELECT prestamo_id, estado FROM prestamo_lineas WHERE id = ?",
                        (linea_id,)).fetchone()
        if not l:
            raise ValueError("Esa prenda no esta en la libreta")
        if l["estado"] == "vendido":
            raise ValueError("Esa prenda ya se cobro")
        con.execute("UPDATE prestamo_lineas SET estado = ? WHERE id = ?", (estado, linea_id))
    _revisar_cierre(l["prestamo_id"])
    return leer_prestamo(l["prestamo_id"])


def vender_lineas(linea_ids: list, ticket: str) -> list[int]:
    """Lo que se ha cobrado con un ticket de verdad se tacha en la libreta."""
    tocados = set()
    with _cerrojo_db, conectar() as con:
        for lid in linea_ids:
            l = con.execute("SELECT prestamo_id FROM prestamo_lineas WHERE id = ? AND estado = 'fuera'",
                            (int(lid),)).fetchone()
            if not l:
                continue
            con.execute("UPDATE prestamo_lineas SET estado = 'vendido', ticket = ? WHERE id = ?",
                        (ticket, int(lid)))
            tocados.add(l["prestamo_id"])
    for pid in tocados:
        _revisar_cierre(pid)
    return sorted(tocados)


def borrar_prestamo(pid: int) -> None:
    with _cerrojo_db, conectar() as con:
        con.execute("DELETE FROM prestamo_lineas WHERE prestamo_id = ?", (pid,))
        con.execute("DELETE FROM prestamos WHERE id = ?", (pid,))


def lineas_resguardo(p: dict) -> list[tuple[str, dict]]:
    """El papelito para la clienta: que se lleva, a que precio y hasta cuando.
    NO es una factura y lo dice."""
    L: list[tuple[str, dict]] = []

    def centrada(texto: str, **estilo):
        for trozo in partir(texto, ANCHO):
            L.append((trozo, {"c": True, **estilo}))

    L.append(("LOGO", {"logo": True}))
    centrada(CFG["nombre"], n=True)
    centrada(CFG["direccion"])
    L.append(("", {}))
    L.append(("PRUEBA EN CASA", {"c": True, "g": True}))
    centrada("Esto no es una factura")
    L.append((fecha_bonita(p["fecha"], p["hora"]), {"c": True}))
    L.append(("", {}))
    L.append((f"Para: {p['nombre']}", {"n": True}))
    if p.get("nota"):
        for trozo in partir(p["nota"], ANCHO):
            L.append((trozo, {}))
    L.append(("-", {"s": True}))
    for l in p["lineas"]:
        if l["estado"] == "devuelto":
            continue
        precio = euros(l["importe"])
        nombre = l["familia"] + (" (pagado)" if l["estado"] == "vendido" else "")
        hueco = ANCHO - len(nombre) - len(precio)
        L.append((f"{nombre}{' ' * max(1, hueco)}{precio}", {}))
    L.append(("-", {"s": True}))
    total = euros(p["pendiente"])
    L.append((f"PENDIENTE{' ' * max(1, ANCHO // 2 - 9 - len(total))}{total}", {"g": True, "n": True}))
    L.append(("", {}))
    # Sin fecha limite (Claudia, 2026-09-16): «suelen traerlo a los pocos
    # dias y no es como que haya un limite».
    centrada("Cuando te decidas, pasate por la tienda")
    L.append(("", {}))
    for pie in CFG["pie"]:
        centrada(pie)
    return L


def imprimir_resguardo(p: dict) -> dict:
    L = lineas_resguardo(p)
    ok, aviso = imprimir_raw(escpos_de(L))
    if not ok:
        apuntar(f"[prueba-casa] resguardo de {p['nombre']}: {aviso}")
    return {"impreso": ok, "aviso": aviso, "texto": texto_de(L)}


# ---------------------------------------------------------------------------
# COPIAS DE SEGURIDAD
# ---------------------------------------------------------------------------

def carpeta_drive() -> Path | None:
    ruta = (CFG.get("carpeta_drive") or "").strip()
    if not ruta:
        return None
    p = Path(ruta)
    return p if p.is_dir() else None


def llevar_a_drive(fichero: Path) -> bool:
    d = carpeta_drive()
    if not d:
        return False
    try:
        shutil.copy2(fichero, d / fichero.name)
        return True
    except Exception as fallo:
        apuntar(f"[copia] no he podido llevar {fichero.name} a Drive: {fallo}")
        return False


def hacer_copia() -> dict:
    """Copia de hoy de la base de datos (se sobrescribe si ya la habia)."""
    COPIAS.mkdir(exist_ok=True)
    if not DB.exists():
        return {"ok": False, "motivo": "todavia no hay ventas"}
    destino = COPIAS / f"ventas-{date.today():%Y-%m-%d}.sqlite"
    with _cerrojo_db:
        origen = sqlite3.connect(DB)
        copia = sqlite3.connect(destino)
        with copia:
            origen.backup(copia)
        copia.close()
        origen.close()
    # No acumular: las ultimas COPIAS_QUE_GUARDAR y ya
    viejas = sorted(COPIAS.glob("ventas-*.sqlite"))[:-COPIAS_QUE_GUARDAR]
    for v in viejas:
        v.unlink(missing_ok=True)
    en_drive = llevar_a_drive(destino)
    if en_drive:
        # Y el mes en curso en CSV, que es lo que mira la gestora
        try:
            exportar_mes(date.today().strftime("%Y-%m"))
        except Exception as fallo:
            apuntar(f"[copia] csv del mes: {fallo}")
    apuntar(f"[copia] {destino.name}" + (" + Drive" if en_drive else " (Drive no configurado)"))
    return {"ok": True, "fichero": destino.name, "drive": en_drive,
            "cuando": datetime.now().strftime("%H:%M")}


ULTIMA_COPIA: dict = {}


def vigilar_copias() -> None:
    def ronda():
        global ULTIMA_COPIA
        while True:
            try:
                ULTIMA_COPIA = hacer_copia()
            except Exception as fallo:
                apuntar(f"[copia] fallo: {fallo}")
            time.sleep(6 * 3600)
    threading.Thread(target=ronda, daemon=True).start()


# ---------------------------------------------------------------------------
# EL TICKET (texto para la pantalla y bytes ESC/POS para la impresora)
# ---------------------------------------------------------------------------

def euros(x: float) -> str:
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + " €"


def fecha_bonita(fecha: str, hora: str) -> str:
    d = datetime.strptime(fecha, "%Y-%m-%d")
    return f"{d:%d/%m/%Y}  {hora[:5]}"


def lineas_ticket(t: dict, regalo: bool = False) -> list[tuple[str, dict]]:
    """Lo que va en el ticket, linea a linea, con su estilo.

    estilo: c=centrado, g=grande (doble alto y ancho), n=negrita, s=separador
    """
    L: list[tuple[str, dict]] = []

    def centrada(texto: str, **estilo):
        # una linea larga (la direccion) se parte en varias de ANCHO
        for trozo in partir(texto, ANCHO):
            L.append((trozo, {"c": True, **estilo}))

    L.append(("LOGO", {"logo": True}))
    centrada(CFG["lema"])
    centrada(CFG["nombre"], n=True)
    centrada(CFG["titular"])
    centrada(f"NIF {CFG['nif']}")
    centrada(CFG["direccion"])
    L.append(("", {}))
    if regalo:
        L.append(("TICKET REGALO", {"c": True, "g": True}))
        L.append((f"Ref. {t['numero']}", {"c": True}))
    else:
        titulo = "ANULACION" if t["tipo"] == "anulacion" else "FACTURA SIMPLIFICADA"
        L.append((titulo, {"c": True, "n": True}))
        L.append((f"Nº {t['numero']}", {"c": True, "n": True}))
        if t["tipo"] == "anulacion":
            L.append((f"Anula al ticket {t['anula_a']}", {"c": True}))
    L.append((fecha_bonita(t["fecha"], t["hora"]), {"c": True}))
    L.append(("-", {"s": True}))
    for l in t["lineas"]:
        nombre = l["familia"]
        if regalo:
            L.append((nombre, {}))
        else:
            precio = euros(l["importe"])
            hueco = ANCHO - len(nombre) - len(precio)
            L.append((f"{nombre}{' ' * max(1, hueco)}{precio}", {}))
    L.append(("-", {"s": True}))
    if not regalo:
        # En letra doble caben ANCHO // 2 columnas: TOTAL a la izquierda,
        # importe pegado a la derecha, todo en UNA linea (si se pasa, la
        # impresora parte el importe en dos y queda «36,» / «00 €»).
        total = euros(t["total"])
        L.append((f"TOTAL{' ' * max(1, ANCHO // 2 - 5 - len(total))}{total}", {"g": True, "n": True}))
        L.append((f"IVA {CFG['iva']}% incluido", {"c": True}))
        L.append((f"Pago: {'Efectivo' if t['pago'] == 'efectivo' else 'Tarjeta'}", {"c": True}))
        L.append(("", {}))
        L.append(("Cambios y devoluciones con este ticket", {"c": True}))
    L.append(("", {}))
    for p in CFG["pie"]:
        centrada(p)
    return L


def partir(texto: str, ancho: int) -> list[str]:
    """Parte un texto en lineas de como mucho `ancho` letras, por palabras."""
    lineas, actual = [], ""
    for palabra in texto.split():
        if actual and len(actual) + 1 + len(palabra) > ancho:
            lineas.append(actual)
            actual = palabra
        else:
            actual = f"{actual} {palabra}".strip()
    if actual or not lineas:
        lineas.append(actual)
    return lineas


def ticket_como_texto(t: dict, regalo: bool = False) -> str:
    return texto_de(lineas_ticket(t, regalo))


def texto_de(L: list[tuple[str, dict]]) -> str:
    out = []
    for texto, e in L:
        if e.get("logo"):
            out.append("MADACLA".center(ANCHO))
            continue
        if e.get("s"):
            out.append("-" * ANCHO)
            continue
        if e.get("g"):
            out.append(texto.strip().center(ANCHO))
            continue
        out.append(texto.center(ANCHO) if e.get("c") else texto)
    return "\n".join(out)


def ticket_como_escpos(t: dict, regalo: bool = False) -> bytes:
    # Codigo de barras con el numero del ticket (CODE128), para que
    # Maricarmen lo pase por el lector y salga la venta en la caja.
    return escpos_de(lineas_ticket(t, regalo), codigo_barras(t["numero"]))


def escpos_de(L: list[tuple[str, dict]], barras: bytes = b"") -> bytes:
    ESC, GS = b"\x1b", b"\x1d"
    b = bytearray()
    b += ESC + b"@"                 # reset
    b += ESC + b"t\x13"             # pagina de codigos 19 = CP858 (lleva el €)
    for texto, e in L:
        if e.get("logo"):
            if LOGO_TICKET.exists():
                b += ESC + b"a\x01" + LOGO_TICKET.read_bytes() + b"\n"
            else:
                b += ESC + b"a\x01" + GS + b"!\x11" + b"MADACLA\n" + GS + b"!\x00"
            continue
        if e.get("s"):
            b += ESC + b"a\x00" + ("-" * ANCHO).encode("cp858") + b"\n"
            continue
        b += ESC + (b"a\x01" if e.get("c") or e.get("g") else b"a\x00")
        if e.get("n"):
            b += ESC + b"E\x01"
        if e.get("g"):
            b += GS + b"!\x11"
            texto = texto.strip()
        b += texto.encode("cp858", errors="replace") + b"\n"
        if e.get("g"):
            b += GS + b"!\x00"
        if e.get("n"):
            b += ESC + b"E\x00"
    if barras:
        b += b"\n" + barras
    b += b"\n\n\n\n"
    b += GS + b"V\x42\x00"          # corte parcial
    return bytes(b)


def codigo_barras(numero: str) -> bytes:
    """CODE128 centrado con «T» + los digitos del numero (T2026000006).

    La T delante es la senal para la pantalla de que lo que llega es un
    lector y no alguien tecleando un precio.
    """
    ESC, GS = b"\x1b", b"\x1d"
    datos = b"{B" + ("T" + numero.replace("-", "")).encode("ascii")
    return (ESC + b"a\x01"
            + GS + b"h\x46"                        # alto 70 puntos
            + GS + b"w\x02"                        # barras de ancho 2
            + GS + b"H\x00"                        # sin el numero debajo (ya va en letras)
            + GS + b"k\x49" + bytes([len(datos)]) + datos
            + b"\n")


def numero_normal(texto: str) -> str:
    """Lo que teclee o lea el lector -> «2026-000006».

    Vale «T2026000006», «2026000006», «2026-000006» o solo «6» (se
    entiende que es de este anno).
    """
    d = "".join(c for c in str(texto) if c.isdigit())
    if not d:
        raise ValueError("Escribe el numero del ticket")
    if len(d) >= 10:
        return f"{d[:4]}-{int(d[4:]):06d}"
    return f"{date.today().year}-{int(d):06d}"


PULSO_CAJON = b"\x1b\x70\x00\x19\xfa"   # ESC p 0 25 250: abre el cajon


def imprimir_raw(datos: bytes) -> tuple[bool, str]:
    """Manda bytes tal cual a la impresora de Windows que diga la config."""
    nombre = (CFG.get("impresora") or "").strip()
    if not nombre:
        return False, "sin impresora configurada"
    if sys.platform != "win32":
        return False, "solo imprime en Windows"
    try:
        winspool = ctypes.WinDLL("winspool.drv")

        class DOC_INFO_1(ctypes.Structure):
            _fields_ = [("pDocName", ctypes.c_wchar_p),
                        ("pOutputFile", ctypes.c_wchar_p),
                        ("pDatatype", ctypes.c_wchar_p)]

        h = ctypes.c_void_p()
        if not winspool.OpenPrinterW(nombre, ctypes.byref(h), None):
            return False, f"no encuentro la impresora «{nombre}»"
        try:
            doc = DOC_INFO_1("Ticket Madacla", None, "RAW")
            if not winspool.StartDocPrinterW(h, 1, ctypes.byref(doc)):
                return False, "la impresora no acepta el trabajo"
            winspool.StartPagePrinter(h)
            escritos = ctypes.c_ulong(0)
            buf = ctypes.create_string_buffer(datos, len(datos))
            winspool.WritePrinter(h, buf, len(datos), ctypes.byref(escritos))
            winspool.EndPagePrinter(h)
            winspool.EndDocPrinter(h)
        finally:
            winspool.ClosePrinter(h)
        return True, "impreso"
    except Exception as fallo:
        return False, f"error al imprimir: {fallo}"


def imprimir_ticket(t: dict, regalo: bool = False, abrir_cajon: bool = False) -> dict:
    datos = ticket_como_escpos(t, regalo)
    if abrir_cajon and CFG.get("cajon", True):
        datos = PULSO_CAJON + datos
    ok, aviso = imprimir_raw(datos)
    if not ok:
        apuntar(f"[ticket] {t['numero']}{' regalo' if regalo else ''}: {aviso}")
    return {"impreso": ok, "aviso": aviso, "texto": ticket_como_texto(t, regalo)}


def abrir_cajon_solo() -> dict:
    ok, aviso = imprimir_raw(PULSO_CAJON)
    return {"abierto": ok, "aviso": aviso}


def listar_impresoras() -> list[str]:
    """Las impresoras que Windows conoce, por nombre (para elegir la de tickets)."""
    if sys.platform != "win32":
        return []
    try:
        winspool = ctypes.WinDLL("winspool.drv")

        class PRINTER_INFO_4(ctypes.Structure):
            _fields_ = [("pPrinterName", ctypes.c_wchar_p),
                        ("pServerName", ctypes.c_wchar_p),
                        ("Attributes", ctypes.c_ulong)]

        LOCAL_Y_CONECTADAS = 2 | 4
        hacen_falta, cuantas = ctypes.c_ulong(0), ctypes.c_ulong(0)
        winspool.EnumPrintersW(LOCAL_Y_CONECTADAS, None, 4, None, 0,
                               ctypes.byref(hacen_falta), ctypes.byref(cuantas))
        if not hacen_falta.value:
            return []
        buf = ctypes.create_string_buffer(hacen_falta.value)
        if not winspool.EnumPrintersW(LOCAL_Y_CONECTADAS, None, 4, buf, hacen_falta.value,
                                      ctypes.byref(hacen_falta), ctypes.byref(cuantas)):
            return []
        filas = ctypes.cast(buf, ctypes.POINTER(PRINTER_INFO_4))
        return [filas[i].pPrinterName for i in range(cuantas.value) if filas[i].pPrinterName]
    except Exception as fallo:
        apuntar(f"[impresoras] no he podido listarlas: {fallo}")
        return []


def guardar_ajustes(cambios: dict) -> None:
    """Cambia la config en caliente y la deja escrita en config.json."""
    permitidos = {"impresora", "cajon", "carpeta_drive"}
    for k, v in cambios.items():
        if k in permitidos:
            CFG[k] = v
    with open(CONFIG, "w", encoding="utf-8") as f:
        json.dump(CFG, f, ensure_ascii=False, indent=2)
    apuntar(f"[ajustes] {', '.join(f'{k}={v!r}' for k, v in cambios.items() if k in permitidos)}")


def imprimir_prueba() -> dict:
    """Un ticket corto para comprobar impresora y cajon."""
    ESC, GS = b"\x1b", b"\x1d"
    b = bytearray(PULSO_CAJON if CFG.get("cajon", True) else b"")
    b += ESC + b"@" + ESC + b"t\x13" + ESC + b"a\x01"
    if LOGO_TICKET.exists():
        b += LOGO_TICKET.read_bytes() + b"\n"
    b += ESC + b"E\x01" + GS + b"!\x11" + b"PRUEBA\n" + GS + b"!\x00" + ESC + b"E\x00"
    b += f"{datetime.now():%d/%m/%Y  %H:%M}\n\n".encode("cp858")
    b += "Si lees esto, la impresora funciona.\n".encode("cp858")
    b += "El cajon deberia haberse abierto.\n".encode("cp858")
    b += b"\n\n\n\n" + GS + b"V\x42\x00"
    ok, aviso = imprimir_raw(bytes(b))
    apuntar(f"[prueba] {aviso}")
    return {"impreso": ok, "aviso": aviso}


# ---------------------------------------------------------------------------
# EL SERVIDOR (solo escucha en este ordenador)
# ---------------------------------------------------------------------------

TIPOS = {".html": "text/html; charset=utf-8", ".woff2": "font/woff2",
         ".png": "image/png", ".css": "text/css; charset=utf-8",
         ".js": "text/javascript; charset=utf-8", ".svg": "image/svg+xml"}


def pin_ok(pin) -> bool:
    return str(pin or "") == str(CFG.get("pin", ""))


class Manejador(BaseHTTPRequestHandler):
    def log_message(self, *_):      # sin ruido en la consola
        pass

    def json(self, codigo: int, datos) -> None:
        cuerpo = json.dumps(datos, ensure_ascii=False).encode("utf-8")
        self.send_response(codigo)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(cuerpo)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(cuerpo)

    def fichero(self, ruta: Path) -> None:
        if not ruta.is_file() or AQUI not in ruta.resolve().parents:
            self.send_error(404)
            return
        datos = ruta.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", TIPOS.get(ruta.suffix, "application/octet-stream"))
        self.send_header("Content-Length", str(len(datos)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(datos)

    def leer_json(self) -> dict:
        n = int(self.headers.get("Content-Length") or 0)
        if not n:
            return {}
        try:
            return json.loads(self.rfile.read(n).decode("utf-8"))
        except Exception:
            return {}

    def do_GET(self):
        u = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        if u.path in ("/", "/index.html"):
            return self.fichero(AQUI / "pantalla.html")
        if u.path.startswith("/fuentes/") or u.path == "/logo.png":
            return self.fichero(AQUI / u.path.lstrip("/"))
        if u.path == "/api/config":
            return self.json(200, {
                "nombre": CFG["nombre"], "lema": CFG["lema"],
                "familias": CFG["familias"],
                "impresora": bool((CFG.get("impresora") or "").strip()),
                "drive": bool(carpeta_drive()),
                "version": version_local(),
            })
        if u.path == "/api/estado":
            return self.json(200, {
                "impresora": (CFG.get("impresora") or "").strip() or None,
                "drive": str(carpeta_drive()) if carpeta_drive() else None,
                "ultima_copia": ULTIMA_COPIA,
                "cadena": comprobar_cadena(),
            })
        if u.path == "/api/dia":
            if not pin_ok(q.get("pin")):
                return self.json(403, {"error": "PIN incorrecto"})
            fecha = q.get("fecha") or date.today().isoformat()
            return self.json(200, resumen_dia(fecha))
        if u.path == "/api/impresoras":
            return self.json(200, {"impresoras": listar_impresoras(),
                                   "elegida": (CFG.get("impresora") or "").strip()})
        if u.path == "/api/ticket":
            try:
                numero = numero_normal(q.get("numero", ""))
            except ValueError as fallo:
                return self.json(400, {"error": str(fallo)})
            t = leer_ticket(numero)
            if not t:
                return self.json(404, {"error": f"No hay ningún ticket {numero}"})
            return self.json(200, t)
        if u.path == "/api/prestamos":
            return self.json(200, {"abiertos": listar_prestamos(True),
                                   "cerrados": listar_prestamos(False) if q.get("todos") else []})
        if u.path == "/api/prestamos/nombres":
            return self.json(200, {"nombres": nombres_prestamos()})
        if u.path == "/api/ultimo":
            with conectar() as con:
                t = con.execute("SELECT numero FROM tickets WHERE tipo='venta' "
                                "ORDER BY id DESC LIMIT 1").fetchone()
            return self.json(200, leer_ticket(t["numero"]) if t else None)
        self.send_error(404)

    def do_POST(self):
        if self.path.startswith("/api/latido"):
            # la pantalla avisa cada 30 s de si esta en reposo (ticket vacio)
            d = self.leer_json()
            ESTADO["ultimo_latido"] = time.time()
            ESTADO["reposo"] = bool(d.get("reposo"))
            return self.json(200, {"version": version_local()})
        ESTADO["ultimo_post"] = time.time()
        u = urlparse(self.path)
        d = self.leer_json()
        try:
            if u.path == "/api/venta":
                t = crear_ticket(d.get("lineas", []), d.get("pago", ""))
                imp = imprimir_ticket(t, abrir_cajon=(t["pago"] == "efectivo"))
                apuntar(f"[venta] {t['numero']} {euros(t['total'])} {t['pago']}")
                return self.json(200, {"ticket": t, **imp})
            if u.path == "/api/regalo":
                t = leer_ticket(d.get("numero", ""))
                if not t:
                    return self.json(404, {"error": "Ese ticket no existe"})
                return self.json(200, {"ticket": t, **imprimir_ticket(t, regalo=True)})
            if u.path == "/api/reimprimir":
                t = leer_ticket(d.get("numero", ""))
                if not t:
                    return self.json(404, {"error": "Ese ticket no existe"})
                return self.json(200, {"ticket": t, **imprimir_ticket(t)})
            if u.path == "/api/pin":
                return self.json(200, {"ok": pin_ok(d.get("pin"))})
            if u.path == "/api/anular":
                if not pin_ok(d.get("pin")):
                    return self.json(403, {"error": "PIN incorrecto"})
                t = anular_ticket(d.get("numero", ""))
                apuntar(f"[anulacion] {t['numero']} anula {t['anula_a']}")
                return self.json(200, {"ticket": t, **imprimir_ticket(t)})
            if u.path == "/api/cierre":
                if not pin_ok(d.get("pin")):
                    return self.json(403, {"error": "PIN incorrecto"})
                fecha = d.get("fecha") or date.today().isoformat()
                r = cerrar_caja(fecha)
                copia = hacer_copia()
                apuntar(f"[cierre] {fecha}: {euros(r['total'])} "
                        f"(efectivo {euros(r['efectivo'])}, tarjeta {euros(r['tarjeta'])})")
                return self.json(200, {"resumen": r, "copia": copia, **imprimir_cierre(r)})
            if u.path == "/api/cierre/imprimir":
                if not pin_ok(d.get("pin")):
                    return self.json(403, {"error": "PIN incorrecto"})
                r = resumen_dia(d.get("fecha") or date.today().isoformat())
                return self.json(200, imprimir_cierre(r))
            if u.path == "/api/exportar":
                if not pin_ok(d.get("pin")):
                    return self.json(403, {"error": "PIN incorrecto"})
                mes = d.get("mes") or date.today().strftime("%Y-%m")
                destino = exportar_mes(mes)
                return self.json(200, {"fichero": str(destino), "drive": bool(carpeta_drive())})
            if u.path == "/api/cajon":
                if not pin_ok(d.get("pin")):
                    return self.json(403, {"error": "PIN incorrecto"})
                return self.json(200, abrir_cajon_solo())
            if u.path == "/api/ajustes":
                if not pin_ok(d.get("pin")):
                    return self.json(403, {"error": "PIN incorrecto"})
                guardar_ajustes({k: v for k, v in d.items() if k != "pin"})
                return self.json(200, {"ok": True, "impresora": (CFG.get("impresora") or "").strip()})
            if u.path == "/api/prueba":
                if not pin_ok(d.get("pin")):
                    return self.json(403, {"error": "PIN incorrecto"})
                return self.json(200, imprimir_prueba())
            if u.path == "/api/copia":
                return self.json(200, hacer_copia())
            # --- la libreta de «se lo lleva a probar» ---
            if u.path == "/api/prestamo/nuevo":
                p = crear_prestamo(d.get("nombre", ""), d.get("nota", ""), d.get("lineas", []))
                apuntar(f"[prueba-casa] {p['nombre']}: {len(p['lineas'])} prenda(s), {euros(p['pendiente'])}")
                return self.json(200, {"prestamo": p, **imprimir_resguardo(p)})
            if u.path == "/api/prestamo/linea":
                p = marcar_linea(int(d.get("linea_id", 0)), d.get("estado", ""))
                return self.json(200, {"prestamo": p})
            if u.path == "/api/prestamo/vendido":
                tocados = vender_lineas(d.get("linea_ids", []), str(d.get("ticket", "")))
                apuntar(f"[prueba-casa] cobradas {len(d.get('linea_ids', []))} prenda(s) en el {d.get('ticket')}")
                return self.json(200, {"prestamos": tocados})
            if u.path == "/api/prestamo/imprimir":
                p = leer_prestamo(int(d.get("id", 0)))
                if not p:
                    return self.json(404, {"error": "Esa hoja ya no esta"})
                return self.json(200, imprimir_resguardo(p))
            if u.path == "/api/prestamo/borrar":
                if not pin_ok(d.get("pin")):
                    return self.json(403, {"error": "PIN incorrecto"})
                borrar_prestamo(int(d.get("id", 0)))
                apuntar(f"[prueba-casa] borrada la hoja {d.get('id')}")
                return self.json(200, {"ok": True})
        except ValueError as fallo:
            return self.json(400, {"error": str(fallo)})
        except Exception as fallo:
            apuntar(f"[error] {u.path}: {fallo}")
            return self.json(500, {"error": f"Algo ha fallado: {fallo}"})
        self.send_error(404)


def lineas_cierre(r: dict) -> list[tuple[str, dict]]:
    """El resumen del dia en papel, como el que sacaba la caja vieja: lo
    que se ha cobrado en efectivo y con tarjeta, por familias, y del ticket
    al ticket. Maricarmen lo guarda cada noche."""
    d = datetime.strptime(r["fecha"], "%Y-%m-%d")
    L: list[tuple[str, dict]] = []

    def fila(izq: str, der: str, **estilo):
        L.append((f"{izq}{' ' * max(1, ANCHO - len(izq) - len(der))}{der}", estilo))

    L.append(("LOGO", {"logo": True}))
    for trozo in partir(CFG["nombre"], ANCHO):
        L.append((trozo, {"c": True, "n": True}))
    L.append(("", {}))
    L.append(("CIERRE DE CAJA", {"c": True, "g": True, "n": True}))
    L.append((f"{d:%d/%m/%Y}" + (f"  ·  {r['cierre_hora'][:5]}" if r.get("cierre_hora") else ""), {"c": True}))
    L.append(("-", {"s": True}))

    if r["tickets"]:
        fila("Tickets", f"{r['tickets'][0]['numero']} al {r['tickets'][-1]['numero'][5:]}")
    fila("Ventas", str(r["ventas"]))
    if r["anulaciones"]:
        fila("Anulaciones", str(r["anulaciones"]))
    L.append(("-", {"s": True}))

    # por familias (las anulaciones restan, porque van en negativo)
    familias: dict[str, float] = {}
    for t in r["tickets"]:
        for l in t["lineas"]:
            familias[l["familia"]] = familias.get(l["familia"], 0) + l["importe"]
    if familias:
        L.append(("POR FAMILIAS", {"n": True}))
        for nombre, imp in familias.items():
            if round(imp, 2):
                fila(nombre, euros(redondear(imp)))
        L.append(("-", {"s": True}))

    fila("Efectivo", euros(r["efectivo"]), n=True)
    fila("Tarjeta", euros(r["tarjeta"]), n=True)
    L.append(("-", {"s": True}))
    total = euros(r["total"])
    L.append((f"TOTAL{' ' * max(1, ANCHO // 2 - 5 - len(total))}{total}", {"g": True, "n": True}))
    return L


def texto_cierre(r: dict) -> str:
    return texto_de(lineas_cierre(r))


def imprimir_cierre(r: dict) -> dict:
    L = lineas_cierre(r)
    ok, aviso = imprimir_raw(escpos_de(L))
    if not ok:
        apuntar(f"[cierre] resumen del {r['fecha']}: {aviso}")
    return {"impreso": ok, "aviso": aviso, "texto": texto_de(L)}


# ---------------------------------------------------------------------------
# SE ACTUALIZA SOLA (2026-09-19)
# ---------------------------------------------------------------------------
# Maricarmen no puede bajar zips. Cada 20 minutos la caja mira en
# madacla.es/caja-act/version.json si hay version nueva; si la hay, baja los
# ficheros, comprueba su huella (sha256) y que el programa nuevo arranca, se
# guarda una copia del actual en anteriores/ultima/ y los pone. Luego:
#   - la pantalla se recarga sola cuando el ticket en curso esta vacio;
#   - el programa se reinicia solo cuando lleva 3 minutos sin cobrar y la
#     pantalla esta en reposo (lo reinicia el bucle de arrancar-caja.cmd,
#     que tambien vuelve a la copia anterior si el nuevo se cae al arrancar).
# config.json (NIF, PIN, impresora) NUNCA se toca.

URL_ACTUALIZACION = os.environ.get("CAJA_URL_ACT", "https://madacla.es/caja-act/")   # la variable, solo para probar
FICHEROS_ACTUALIZABLES = ("caja.py", "pantalla.html", "logo.png", "logo-ticket.bin")
VERSION_FICH = AQUI / "version.txt"
ANTERIORES = AQUI / "anteriores"
CADA_CUANTO = int(os.environ.get("CAJA_CADA", str(20 * 60)))


def version_local() -> int:
    try:
        return int(VERSION_FICH.read_text(encoding="utf-8").strip())
    except Exception:
        return 0


VERSION = version_local()          # la que esta corriendo AHORA
ESTADO = {"ultimo_post": 0.0, "ultimo_latido": 0.0, "reposo": False,
          "pendiente": False, "reiniciar": False}
SERVIDOR = None


def _bajar(nombre: str) -> bytes:
    with urlopen(URL_ACTUALIZACION + nombre + f"?v={int(time.time())}", timeout=30) as r:
        return r.read()


def buscar_actualizacion() -> str:
    """Baja y pone la version nueva si la hay. Devuelve que ha pasado."""
    try:
        manifiesto = json.loads(_bajar("version.json").decode("utf-8"))
        nueva = int(manifiesto["version"])
    except Exception as fallo:
        return f"no he podido mirar si hay version nueva: {fallo}"
    if nueva <= version_local():
        return "al dia"
    try:   # una version que ya se cayo al arrancar no se vuelve a poner
        if nueva <= int((ANTERIORES / "mala.txt").read_text(encoding="utf-8").strip()):
            return "al dia"
    except Exception:
        pass
    huellas = manifiesto.get("ficheros", {})
    if not huellas or any(n not in FICHEROS_ACTUALIZABLES for n in huellas):
        return f"version {nueva}: lista de ficheros rara, no la pongo"

    temporal = AQUI / "actualizacion"
    shutil.rmtree(temporal, ignore_errors=True)
    temporal.mkdir()
    try:
        for nombre, huella in huellas.items():
            datos = _bajar(nombre)
            if hashlib.sha256(datos).hexdigest() != huella:
                return f"version {nueva}: {nombre} ha llegado mal, no la pongo"
            (temporal / nombre).write_bytes(datos)
        # el programa nuevo tiene que poder arrancar antes de ponerlo
        for conf in (CONFIG, CONFIG_EJEMPLO):
            if conf.exists():
                shutil.copy2(conf, temporal / conf.name)
        if "caja.py" in huellas:
            prueba = subprocess.run(
                [sys.executable, str(temporal / "caja.py"), "--comprobar"],
                capture_output=True, text=True, timeout=60, cwd=str(AQUI))
            if prueba.returncode != 0:
                return (f"version {nueva}: el programa nuevo no pasa la prueba, "
                        f"no lo pongo: {(prueba.stderr or prueba.stdout)[-300:]}")
        # copia de lo de ahora, por si hay que volver atras
        ultima = ANTERIORES / "ultima"
        shutil.rmtree(ultima, ignore_errors=True)
        ultima.mkdir(parents=True)
        for nombre in huellas:
            if (AQUI / nombre).exists():
                shutil.copy2(AQUI / nombre, ultima / nombre)
        (ultima / "version.txt").write_text(str(version_local()), encoding="utf-8")
        for nombre in huellas:
            os.replace(temporal / nombre, AQUI / nombre)
        VERSION_FICH.write_text(str(nueva), encoding="utf-8")
    finally:
        shutil.rmtree(temporal, ignore_errors=True)
    ESTADO["pendiente"] = "caja.py" in huellas
    return f"puesta la version {nueva}"


def en_reposo() -> bool:
    """Se puede reiniciar sin molestar: 3 min sin cobrar y la pantalla vacia
    (o la pantalla cerrada, sin latir desde hace mas de 2 minutos)."""
    ahora = time.time()
    if ahora - ESTADO["ultimo_post"] < 180:
        return False
    pantalla_viva = ahora - ESTADO["ultimo_latido"] < 120
    return ESTADO["reposo"] or not pantalla_viva


def vigilar_actualizaciones() -> None:
    def ronda():
        time.sleep(int(os.environ.get("CAJA_ESPERA", "60")))
        buena_desde = time.time()
        while True:
            # si esta version lleva 10 minutos bien, la copia anterior se archiva
            ultima = ANTERIORES / "ultima"
            if ultima.exists() and time.time() - buena_desde > 600:
                try:
                    os.replace(ultima, ANTERIORES / f"{datetime.now():%Y%m%d-%H%M}")
                except Exception:
                    pass
            if not ESTADO["pendiente"]:
                que = buscar_actualizacion()
                if que != "al dia":
                    apuntar(f"[actualizacion] {que}")
            if ESTADO["pendiente"] and en_reposo():
                if os.environ.get("CAJA_BUCLE") == "1" and SERVIDOR is not None:
                    apuntar("[actualizacion] me reinicio con el programa nuevo")
                    ESTADO["reiniciar"] = True
                    SERVIDOR.shutdown()
                    return
            time.sleep(int(os.environ.get("CAJA_ESPERA", "60")) if ESTADO["pendiente"] else CADA_CUANTO)
    threading.Thread(target=ronda, daemon=True).start()


def quitar_marca_de_internet() -> None:
    """Lo que viene en un zip bajado de internet lleva una marca (el flujo
    «Zone.Identifier») y Windows pregunta «¿Ejecutar?» al encender. Se quita
    de los ficheros de la caja para que arranque sola sin preguntar."""
    if sys.platform != "win32":
        return
    for f in AQUI.iterdir():
        if f.is_file():
            try:
                os.remove(f"{f}:Zone.Identifier")
            except OSError:
                pass


def comprobar_arranque() -> None:
    """caja.py --comprobar: lo que usa el actualizador para fiarse de una
    version nueva antes de ponerla. Monta un ticket de mentira."""
    t = {"numero": "2026-000001", "fecha": "2026-01-01", "hora": "10:00:00",
         "tipo": "venta", "anula_a": None, "pago": "efectivo", "total": 1.0,
         "anulado_por": None, "lineas": [{"familia": "Ropa", "importe": 1.0}]}
    ticket_como_escpos(t)
    ticket_como_escpos(t, regalo=True)
    ticket_como_texto(t)
    print("ok")


def main() -> None:
    global SERVIDOR
    if "--comprobar" in sys.argv:
        comprobar_arranque()
        return
    quitar_marca_de_internet()
    preparar_db()
    vigilar_copias()
    vigilar_actualizaciones()
    puerto = int(CFG.get("puerto", 8791))
    servidor = ThreadingHTTPServer(("127.0.0.1", puerto), Manejador)
    SERVIDOR = servidor
    apuntar(f"Version {VERSION}")
    apuntar(f"Caja Madacla en http://127.0.0.1:{puerto}  "
            f"(impresora: {CFG.get('impresora') or 'ninguna'}; "
            f"Drive: {carpeta_drive() or 'no'})")
    if not CONFIG.exists():
        apuntar("AVISO: no hay config.json, uso config.ejemplo.json (datos de ejemplo)")
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        pass
    servidor.server_close()
    if ESTADO["reiniciar"]:
        sys.exit(3)          # el bucle de arrancar-caja.cmd me vuelve a abrir


if __name__ == "__main__":
    main()
