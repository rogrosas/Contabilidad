# -*- coding: utf-8 -*-
"""Base de datos SQLite: esquema, migraciones y consultas."""
import contextlib
import os
import sqlite3

from .util import CATEGORIAS_GASTO, IVA_POR_DEFECTO, redondear

# Esquema original (versión 0). Las migraciones lo llevan a la versión actual.
ESQUEMA_V0 = [
    "CREATE TABLE IF NOT EXISTS config(clave TEXT PRIMARY KEY, valor TEXT)",
    """CREATE TABLE IF NOT EXISTS productos(
        id INTEGER PRIMARY KEY,
        nombre TEXT NOT NULL UNIQUE COLLATE NOCASE,
        unidad TEXT NOT NULL DEFAULT 'unid',
        categoria TEXT NOT NULL DEFAULT '',
        precio_compra REAL NOT NULL DEFAULT 0,
        precio_venta REAL NOT NULL DEFAULT 0,
        stock_inicial REAL NOT NULL DEFAULT 0,
        activo INTEGER NOT NULL DEFAULT 1)""",
    """CREATE TABLE IF NOT EXISTS movimientos(
        id INTEGER PRIMARY KEY,
        tipo TEXT NOT NULL CHECK (tipo IN ('venta','compra','merma')),
        fecha TEXT NOT NULL,
        producto_id INTEGER NOT NULL REFERENCES productos(id),
        cantidad REAL NOT NULL,
        precio_unit REAL NOT NULL DEFAULT 0,
        costo_unit REAL NOT NULL DEFAULT 0,
        nota TEXT NOT NULL DEFAULT '')""",
]


def _migracion_1(con):
    """Tickets (varias líneas por venta/compra), medios de pago, ajustes de stock,
    stock mínimo, gastos y cierres de caja."""
    con.execute("""CREATE TABLE tickets(
        id INTEGER PRIMARY KEY,
        tipo TEXT NOT NULL CHECK (tipo IN ('venta','compra')),
        fecha TEXT NOT NULL,
        medio_pago TEXT NOT NULL DEFAULT 'Efectivo',
        nota TEXT NOT NULL DEFAULT '',
        pago_con REAL,
        creado TEXT NOT NULL DEFAULT (datetime('now','localtime')))""")
    con.execute("CREATE INDEX idx_tickets_tipo_fecha ON tickets(tipo, fecha)")
    con.execute("""CREATE TABLE movimientos_nueva(
        id INTEGER PRIMARY KEY,
        tipo TEXT NOT NULL CHECK (tipo IN ('venta','compra','merma','ajuste')),
        fecha TEXT NOT NULL,
        producto_id INTEGER NOT NULL REFERENCES productos(id),
        cantidad REAL NOT NULL,
        precio_unit REAL NOT NULL DEFAULT 0,
        costo_unit REAL NOT NULL DEFAULT 0,
        nota TEXT NOT NULL DEFAULT '',
        ticket_id INTEGER REFERENCES tickets(id) ON DELETE CASCADE)""")
    con.execute("""INSERT INTO movimientos_nueva(id, tipo, fecha, producto_id, cantidad, precio_unit, costo_unit, nota)
                   SELECT id, tipo, fecha, producto_id, cantidad, precio_unit, costo_unit, nota FROM movimientos""")
    con.execute("DROP TABLE movimientos")
    con.execute("ALTER TABLE movimientos_nueva RENAME TO movimientos")
    con.execute("CREATE INDEX idx_mov_tipo_fecha ON movimientos(tipo, fecha)")
    con.execute("CREATE INDEX idx_mov_fecha ON movimientos(fecha)")
    con.execute("CREATE INDEX idx_mov_producto ON movimientos(producto_id)")
    con.execute("CREATE INDEX idx_mov_ticket ON movimientos(ticket_id)")
    # Cada venta/compra anterior pasa a ser un ticket de una sola línea, pagado en efectivo.
    for mid, tipo, fecha, nota in con.execute(
            "SELECT id, tipo, fecha, nota FROM movimientos WHERE tipo IN ('venta','compra') ORDER BY id").fetchall():
        tid = con.execute("INSERT INTO tickets(tipo, fecha, nota) VALUES(?,?,?)", (tipo, fecha, nota)).lastrowid
        con.execute("UPDATE movimientos SET ticket_id=? WHERE id=?", (tid, mid))
    con.execute("ALTER TABLE productos ADD COLUMN stock_minimo REAL NOT NULL DEFAULT 0")
    con.execute("""CREATE TABLE gastos(
        id INTEGER PRIMARY KEY,
        fecha TEXT NOT NULL,
        categoria TEXT NOT NULL DEFAULT 'Otros',
        descripcion TEXT NOT NULL DEFAULT '',
        monto REAL NOT NULL,
        medio_pago TEXT NOT NULL DEFAULT 'Efectivo')""")
    con.execute("CREATE INDEX idx_gastos_fecha ON gastos(fecha)")
    con.execute("""CREATE TABLE cierres(
        fecha TEXT PRIMARY KEY,
        inicial REAL NOT NULL DEFAULT 0,
        ventas_ef REAL NOT NULL DEFAULT 0,
        compras_ef REAL NOT NULL DEFAULT 0,
        gastos_ef REAL NOT NULL DEFAULT 0,
        retiros REAL NOT NULL DEFAULT 0,
        esperado REAL NOT NULL DEFAULT 0,
        contado REAL NOT NULL DEFAULT 0,
        diferencia REAL NOT NULL DEFAULT 0,
        nota TEXT NOT NULL DEFAULT '',
        guardado TEXT NOT NULL DEFAULT (datetime('now','localtime')))""")


def _migracion_2(con):
    """IVA: tasa por línea (incluida en el precio), productos exentos, compras con/sin factura y
    gastos con IVA crédito. Los datos anteriores quedan como afectos a la tasa general."""
    tasa = IVA_POR_DEFECTO
    r = con.execute("SELECT valor FROM config WHERE clave='iva_tasa'").fetchone()
    if r:
        tasa = float(r[0])
    con.execute("ALTER TABLE productos ADD COLUMN afecto_iva INTEGER NOT NULL DEFAULT 1")
    con.execute("ALTER TABLE movimientos ADD COLUMN iva_tasa REAL NOT NULL DEFAULT 0")
    con.execute("UPDATE movimientos SET iva_tasa=?", (tasa,))
    con.execute("ALTER TABLE tickets ADD COLUMN con_factura INTEGER NOT NULL DEFAULT 1")
    con.execute("ALTER TABLE gastos ADD COLUMN iva_tasa REAL NOT NULL DEFAULT 0")


MIGRACIONES = [_migracion_1, _migracion_2]
VERSION = len(MIGRACIONES)

STOCK_SQL = """SELECT p.*, p.stock_inicial + COALESCE((
    SELECT SUM(CASE WHEN m.tipo IN ('compra','ajuste') THEN m.cantidad ELSE -m.cantidad END)
    FROM movimientos m WHERE m.producto_id = p.id), 0) AS stock
FROM productos p"""

# Los precios se guardan con IVA incluido; m.iva_tasa es la tasa (%) de la línea (0 = exento).
# Ventas y compras se informan con IVA; costo y mermas, netos (el IVA de la compra se recupera como crédito).
AGREGADOS_SQL = """
    COALESCE(SUM(CASE WHEN m.tipo='venta'  THEN m.cantidad END), 0) AS cant_v,
    COALESCE(SUM(CASE WHEN m.tipo='venta'  THEN m.cantidad*m.precio_unit END), 0) AS ventas,
    COALESCE(SUM(CASE WHEN m.tipo='venta'  THEN m.cantidad*m.precio_unit/(1+m.iva_tasa/100.0) END), 0) AS ventas_netas,
    COALESCE(SUM(CASE WHEN m.tipo='venta'  THEN m.cantidad*m.costo_unit/(1+m.iva_tasa/100.0) END), 0) AS costo,
    COALESCE(SUM(CASE WHEN m.tipo='compra' THEN m.cantidad END), 0) AS cant_c,
    COALESCE(SUM(CASE WHEN m.tipo='compra' THEN m.cantidad*m.precio_unit END), 0) AS compras,
    COALESCE(SUM(CASE WHEN m.tipo='compra' AND tk.con_factura=1
                      THEN m.cantidad*m.precio_unit*m.iva_tasa/(100.0+m.iva_tasa) END), 0) AS iva_credito_c,
    COALESCE(SUM(CASE WHEN m.tipo='merma'  THEN m.cantidad END), 0) AS cant_m,
    COALESCE(SUM(CASE WHEN m.tipo='merma'  THEN m.cantidad*m.costo_unit/(1+m.iva_tasa/100.0) END), 0) AS mermas
"""
JOIN_TICKET = "LEFT JOIN tickets tk ON tk.id = m.ticket_id"


class ArchivoInvalido(Exception):
    pass


def _iso(d):
    return d.isoformat()


class DB:
    def __init__(self, ruta, antes_de_migrar=None):
        os.makedirs(os.path.dirname(ruta), exist_ok=True)
        self.ruta = ruta
        self.con = sqlite3.connect(ruta, isolation_level=None)  # transacciones manuales con tx()
        self.con.row_factory = sqlite3.Row
        self._preparar(antes_de_migrar)

    def _preparar(self, antes_de_migrar=None):
        self.con.execute("PRAGMA foreign_keys = OFF")
        for sql in ESQUEMA_V0:
            self.con.execute(sql)
        version = self.con.execute("PRAGMA user_version").fetchone()[0]
        if version < VERSION:
            con_datos = self.con.execute("SELECT COUNT(*) FROM productos").fetchone()[0] > 0
            if con_datos and antes_de_migrar:
                antes_de_migrar(self)
            for i in range(version, VERSION):
                with self.tx():
                    MIGRACIONES[i](self.con)
                    self.con.execute(f"PRAGMA user_version = {i + 1}")
        self.con.execute("PRAGMA foreign_keys = ON")

    @contextlib.contextmanager
    def tx(self):
        """Transacción: todo o nada."""
        if self.con.in_transaction:
            yield
            return
        self.con.execute("BEGIN IMMEDIATE")
        try:
            yield
        except BaseException:
            self.con.execute("ROLLBACK")
            raise
        self.con.execute("COMMIT")

    def q(self, sql, a=()):
        return self.con.execute(sql, a).fetchall()

    def q1(self, sql, a=()):
        return self.con.execute(sql, a).fetchone()

    def x(self, sql, a=()):
        return self.con.execute(sql, a).lastrowid

    def cerrar(self):
        self.con.close()

    # ------------------------------------------------------------ config
    def config(self, clave, defecto=""):
        r = self.q1("SELECT valor FROM config WHERE clave=?", (clave,))
        return r[0] if r and r[0] is not None else defecto

    def set_config(self, clave, valor):
        self.x("INSERT OR REPLACE INTO config(clave, valor) VALUES(?,?)", (clave, valor))

    def tasa_iva(self):
        """Tasa general de IVA en % (configurable)."""
        try:
            return float(self.config("iva_tasa", IVA_POR_DEFECTO))
        except ValueError:
            return IVA_POR_DEFECTO

    def tasa_producto(self, pid):
        r = self.q1("SELECT afecto_iva FROM productos WHERE id=?", (pid,))
        return self.tasa_iva() if r and r[0] else 0.0

    # ------------------------------------------------------------ productos
    def productos(self, buscar="", inactivos=False, solo_bajo=False):
        b = f"%{buscar.strip()}%"
        sql = f"SELECT * FROM ({STOCK_SQL}) WHERE (nombre LIKE ? OR categoria LIKE ?)"
        if not inactivos:
            sql += " AND activo = 1"
        if solo_bajo:
            sql += " AND activo = 1 AND stock_minimo > 0 AND stock <= stock_minimo"
        return self.q(sql + " ORDER BY nombre COLLATE NOCASE", (b, b))

    def producto(self, pid):
        return self.q1(STOCK_SQL + " WHERE p.id=?", (pid,))

    def producto_por_nombre(self, nombre):
        return self.q1(STOCK_SQL + " WHERE p.nombre = ?", (nombre.strip(),))

    def nombres_activos(self):
        return [r[0] for r in self.q("SELECT nombre FROM productos WHERE activo=1 ORDER BY nombre COLLATE NOCASE")]

    def categorias_producto(self):
        return [r[0] for r in self.q("SELECT DISTINCT categoria FROM productos WHERE categoria<>'' "
                                     "ORDER BY categoria COLLATE NOCASE")]

    def guardar_producto(self, pid, nombre, unidad, categoria, pc, pv, si, minimo=0.0, afecto=True):
        with self.tx():
            if pid:
                self.x("""UPDATE productos SET nombre=?, unidad=?, categoria=?, precio_compra=?,
                          precio_venta=?, stock_inicial=?, stock_minimo=?, afecto_iva=? WHERE id=?""",
                       (nombre, unidad, categoria, pc, pv, si, minimo, 1 if afecto else 0, pid))
            else:
                pid = self.x("""INSERT INTO productos(nombre, unidad, categoria, precio_compra, precio_venta,
                                stock_inicial, stock_minimo, afecto_iva) VALUES(?,?,?,?,?,?,?,?)""",
                             (nombre, unidad, categoria, pc, pv, si, minimo, 1 if afecto else 0))
            if pc > 0:
                self._completar_costos_cero(pid, pc)
        return pid

    def _completar_costos_cero(self, pid, costo):
        """Ventas y mermas registradas cuando el producto aún no tenía costo: se les asigna el costo
        actual, para que la ganancia no quede inflada."""
        self.x("UPDATE movimientos SET costo_unit=? WHERE producto_id=? AND tipo='venta' AND costo_unit=0",
               (costo, pid))
        self.x("""UPDATE movimientos SET costo_unit=?, precio_unit=? WHERE producto_id=? AND tipo IN ('merma','ajuste')
                  AND costo_unit=0 AND precio_unit=0""", (costo, costo, pid))

    def _actualizar_costo(self, pid):
        """El costo del producto es el precio de su compra más reciente (por fecha)."""
        r = self.q1("""SELECT precio_unit FROM movimientos WHERE producto_id=? AND tipo='compra' AND precio_unit>0
                       ORDER BY fecha DESC, id DESC LIMIT 1""", (pid,))
        if r:
            self.x("UPDATE productos SET precio_compra=? WHERE id=?", (r[0], pid))
            self._completar_costos_cero(pid, r[0])

    def usos_producto(self, pid):
        return self.q1("SELECT COUNT(*) FROM movimientos WHERE producto_id=?", (pid,))[0]

    def eliminar_producto(self, pid):
        self.x("DELETE FROM productos WHERE id=?", (pid,))

    def set_activo(self, pid, activo):
        self.x("UPDATE productos SET activo=? WHERE id=?", (1 if activo else 0, pid))

    def cantidad_stock_bajo(self):
        return self.q1(f"""SELECT COUNT(*) FROM ({STOCK_SQL})
                           WHERE activo=1 AND stock_minimo>0 AND stock<=stock_minimo""")[0]

    def kardex(self, pid):
        """Movimientos del producto con el stock acumulado."""
        p = self.producto(pid)
        filas = self.q("""SELECT m.*, t.medio_pago FROM movimientos m LEFT JOIN tickets t ON t.id=m.ticket_id
                          WHERE m.producto_id=? ORDER BY m.fecha, m.id""", (pid,))
        stock = p["stock_inicial"]
        out = []
        for m in filas:
            stock += m["cantidad"] if m["tipo"] in ("compra", "ajuste") else -m["cantidad"]
            out.append((m, stock))
        return p, out

    def ajustar_stock(self, pid, contado, fecha):
        """Registra la diferencia entre el stock contado y el del sistema. Faltante = merma; sobrante = ajuste."""
        p = self.producto(pid)
        dif = round(contado - p["stock"], 6)
        if dif == 0:
            return 0.0
        costo = p["precio_compra"]
        tasa = self.tasa_producto(pid)
        if dif < 0:
            self.x("""INSERT INTO movimientos(tipo, fecha, producto_id, cantidad, precio_unit, costo_unit, nota, iva_tasa)
                      VALUES('merma',?,?,?,?,?,'Ajuste de inventario (faltante)',?)""",
                   (_iso(fecha), pid, -dif, costo, costo, tasa))
        else:
            self.x("""INSERT INTO movimientos(tipo, fecha, producto_id, cantidad, precio_unit, costo_unit, nota, iva_tasa)
                      VALUES('ajuste',?,?,?,?,?,'Ajuste de inventario (sobrante)',?)""",
                   (_iso(fecha), pid, dif, costo, costo, tasa))
        return dif

    # ------------------------------------------------------------ tickets (ventas y compras)
    def ticket(self, tid):
        return self.q1("SELECT * FROM tickets WHERE id=?", (tid,))

    def ticket_items(self, tid):
        return self.q("""SELECT m.*, p.nombre, p.unidad FROM movimientos m JOIN productos p ON p.id=m.producto_id
                         WHERE m.ticket_id=? ORDER BY m.id""", (tid,))

    def tickets(self, tipo, d, h, buscar=""):
        b = f"%{buscar.strip()}%"
        return self.q("""SELECT t.*, SUM(m.cantidad*m.precio_unit) AS total,
                                SUM(m.cantidad*m.precio_unit*m.iva_tasa/(100.0+m.iva_tasa)) AS iva,
                                SUM(m.cantidad*(m.precio_unit-m.costo_unit)/(1+m.iva_tasa/100.0)) AS ganancia,
                                GROUP_CONCAT(printf('%s ×%g', p.nombre, m.cantidad), ', ') AS resumen
                         FROM tickets t JOIN movimientos m ON m.ticket_id=t.id JOIN productos p ON p.id=m.producto_id
                         WHERE t.tipo=? AND t.fecha BETWEEN ? AND ?
                         GROUP BY t.id
                         HAVING resumen LIKE ? OR t.nota LIKE ? OR t.medio_pago LIKE ?
                         ORDER BY t.fecha DESC, t.id DESC""", (tipo, _iso(d), _iso(h), b, b, b))

    def guardar_ticket(self, tid, tipo, fecha, medio, nota, pago_con, items, con_factura=True):
        """items: lista de dict(pid, cantidad, precio, costo, tasa). costo=None usa el costo actual del producto;
        tasa=None usa la tasa de IVA actual del producto. Los precios incluyen IVA."""
        if not items:
            raise ValueError("El ticket no tiene productos.")
        with self.tx():
            afectados = set()
            if tid:
                afectados = {r[0] for r in self.q("SELECT producto_id FROM movimientos WHERE ticket_id=?", (tid,))}
                self.x("UPDATE tickets SET fecha=?, medio_pago=?, nota=?, pago_con=?, con_factura=? WHERE id=?",
                       (_iso(fecha), medio, nota, pago_con, 1 if con_factura else 0, tid))
                self.x("DELETE FROM movimientos WHERE ticket_id=?", (tid,))
            else:
                tid = self.x("INSERT INTO tickets(tipo, fecha, medio_pago, nota, pago_con, con_factura) "
                             "VALUES(?,?,?,?,?,?)", (tipo, _iso(fecha), medio, nota, pago_con, 1 if con_factura else 0))
            for it in items:
                if tipo == "venta":
                    costo = it.get("costo")
                    if costo is None:
                        costo = self.q1("SELECT precio_compra FROM productos WHERE id=?", (it["pid"],))[0]
                else:
                    costo = it["precio"]
                tasa = it.get("tasa")
                if tasa is None:
                    tasa = self.tasa_producto(it["pid"])
                self.x("""INSERT INTO movimientos(tipo, fecha, producto_id, cantidad, precio_unit, costo_unit, nota,
                          ticket_id, iva_tasa) VALUES(?,?,?,?,?,?,?,?,?)""",
                       (tipo, _iso(fecha), it["pid"], it["cantidad"], it["precio"], costo, nota, tid, tasa))
                afectados.add(it["pid"])
            if tipo == "compra":
                for pid in afectados:
                    self._actualizar_costo(pid)
        return tid

    def eliminar_tickets(self, ids):
        with self.tx():
            compras = set()
            for tid in ids:
                t = self.ticket(tid)
                if not t:
                    continue
                if t["tipo"] == "compra":
                    compras |= {r[0] for r in self.q("SELECT producto_id FROM movimientos WHERE ticket_id=?", (tid,))}
                self.x("DELETE FROM movimientos WHERE ticket_id=?", (tid,))
                self.x("DELETE FROM tickets WHERE id=?", (tid,))
            for pid in compras:
                self._actualizar_costo(pid)

    # ------------------------------------------------------------ movimientos sueltos (mermas)
    def mov(self, mid):
        return self.q1("""SELECT m.*, p.nombre, p.unidad FROM movimientos m
                          JOIN productos p ON p.id = m.producto_id WHERE m.id=?""", (mid,))

    def movimientos(self, tipo, d, h, buscar=""):
        """Líneas de un tipo en el período (para ventas y compras incluye datos del ticket)."""
        b = f"%{buscar.strip()}%"
        return self.q("""SELECT m.*, p.nombre, p.unidad, t.medio_pago, t.con_factura FROM movimientos m
                         JOIN productos p ON p.id = m.producto_id
                         LEFT JOIN tickets t ON t.id = m.ticket_id
                         WHERE m.tipo=? AND m.fecha BETWEEN ? AND ? AND (p.nombre LIKE ? OR m.nota LIKE ?)
                         ORDER BY m.fecha DESC, m.id DESC""", (tipo, _iso(d), _iso(h), b, b))

    def guardar_mov(self, mid, tipo, fecha, pid, cant, precio, costo, nota):
        tasa = self.tasa_producto(pid)
        if mid:
            self.x("""UPDATE movimientos SET fecha=?, producto_id=?, cantidad=?, precio_unit=?,
                      costo_unit=?, nota=?, iva_tasa=? WHERE id=?""",
                   (_iso(fecha), pid, cant, precio, costo, nota, tasa, mid))
            return mid
        return self.x("""INSERT INTO movimientos(tipo, fecha, producto_id, cantidad, precio_unit, costo_unit, nota, iva_tasa)
                         VALUES(?,?,?,?,?,?,?,?)""", (tipo, _iso(fecha), pid, cant, precio, costo, nota, tasa))

    def eliminar_movs(self, ids):
        with self.tx():
            for i in ids:
                self.x("DELETE FROM movimientos WHERE id=? AND ticket_id IS NULL", (i,))

    # ------------------------------------------------------------ gastos
    def gasto(self, gid):
        return self.q1("SELECT * FROM gastos WHERE id=?", (gid,))

    def gastos(self, d, h, buscar=""):
        b = f"%{buscar.strip()}%"
        return self.q("""SELECT * FROM gastos WHERE fecha BETWEEN ? AND ?
                         AND (categoria LIKE ? OR descripcion LIKE ? OR medio_pago LIKE ?)
                         ORDER BY fecha DESC, id DESC""", (_iso(d), _iso(h), b, b, b))

    def guardar_gasto(self, gid, fecha, categoria, descripcion, monto, medio, iva_tasa=0.0):
        """monto con IVA incluido; iva_tasa > 0 si tiene factura (genera IVA crédito)."""
        if gid:
            self.x("UPDATE gastos SET fecha=?, categoria=?, descripcion=?, monto=?, medio_pago=?, iva_tasa=? WHERE id=?",
                   (_iso(fecha), categoria, descripcion, monto, medio, iva_tasa, gid))
            return gid
        return self.x("INSERT INTO gastos(fecha, categoria, descripcion, monto, medio_pago, iva_tasa) VALUES(?,?,?,?,?,?)",
                      (_iso(fecha), categoria, descripcion, monto, medio, iva_tasa))

    def ultimo_iva_gasto(self, categoria):
        """Si el último gasto de esa categoría tuvo factura con IVA (None si no hay)."""
        r = self.q1("SELECT iva_tasa FROM gastos WHERE categoria=? ORDER BY id DESC LIMIT 1", (categoria,))
        return None if r is None else r[0] > 0

    def gastos_totales(self, d, h):
        r = self.q1("""SELECT COALESCE(SUM(monto),0), COALESCE(SUM(monto/(1+iva_tasa/100.0)),0)
                       FROM gastos WHERE fecha BETWEEN ? AND ?""", (_iso(d), _iso(h)))
        return dict(bruto=r[0], neto=r[1], iva=r[0] - r[1])

    def eliminar_gastos(self, ids):
        with self.tx():
            for i in ids:
                self.x("DELETE FROM gastos WHERE id=?", (i,))

    def categorias_gasto(self):
        usadas = [r[0] for r in self.q("SELECT DISTINCT categoria FROM gastos ORDER BY categoria COLLATE NOCASE")]
        return CATEGORIAS_GASTO + [c for c in usadas if c not in CATEGORIAS_GASTO]

    def total_gastos(self, d, h):
        return self.q1("SELECT COALESCE(SUM(monto),0) FROM gastos WHERE fecha BETWEEN ? AND ?", (_iso(d), _iso(h)))[0]

    def gastos_por_categoria(self, d, h):
        return self.q("""SELECT categoria, SUM(monto) AS total, COUNT(*) AS n FROM gastos
                         WHERE fecha BETWEEN ? AND ? GROUP BY categoria ORDER BY total DESC""", (_iso(d), _iso(h)))

    # ------------------------------------------------------------ caja
    def resumen_dia(self, fecha):
        """{medio: {'venta': x, 'compra': y, 'gasto': z}} del día."""
        f = _iso(fecha)
        out = {}
        for medio, tipo, total in self.q(
                """SELECT t.medio_pago, t.tipo, SUM(m.cantidad*m.precio_unit) FROM tickets t
                   JOIN movimientos m ON m.ticket_id=t.id WHERE t.fecha=? GROUP BY t.medio_pago, t.tipo
                   UNION ALL
                   SELECT medio_pago, 'gasto', SUM(monto) FROM gastos WHERE fecha=? GROUP BY medio_pago""", (f, f)):
            out.setdefault(medio, {"venta": 0.0, "compra": 0.0, "gasto": 0.0})[tipo] += total or 0
        return out

    def cierre(self, fecha):
        return self.q1("SELECT * FROM cierres WHERE fecha=?", (_iso(fecha),))

    def cierre_anterior(self, fecha):
        return self.q1("SELECT * FROM cierres WHERE fecha<? ORDER BY fecha DESC LIMIT 1", (_iso(fecha),))

    def guardar_cierre(self, fecha, inicial, ventas_ef, compras_ef, gastos_ef, retiros, esperado, contado, nota):
        self.x("""INSERT OR REPLACE INTO cierres(fecha, inicial, ventas_ef, compras_ef, gastos_ef, retiros,
                  esperado, contado, diferencia, nota, guardado)
                  VALUES(?,?,?,?,?,?,?,?,?,?,datetime('now','localtime'))""",
               (_iso(fecha), inicial, ventas_ef, compras_ef, gastos_ef, retiros, esperado, contado,
                round(contado - esperado, 2), nota))

    def eliminar_cierre(self, fecha):
        self.x("DELETE FROM cierres WHERE fecha=?", (_iso(fecha),))

    def cierres(self, d, h):
        return self.q("SELECT * FROM cierres WHERE fecha BETWEEN ? AND ? ORDER BY fecha DESC", (_iso(d), _iso(h)))

    # ------------------------------------------------------------ reportes
    def por_producto(self, d, h):
        filas = self.q(f"""SELECT p.id, p.nombre, p.unidad, p.categoria, p.afecto_iva, {AGREGADOS_SQL}
                           FROM movimientos m JOIN productos p ON p.id = m.producto_id {JOIN_TICKET}
                           WHERE m.fecha BETWEEN ? AND ? AND m.tipo <> 'ajuste'
                           GROUP BY p.id ORDER BY ventas DESC, p.nombre COLLATE NOCASE""", (_iso(d), _iso(h)))
        return [dict(r) for r in filas]

    def por_categoria(self, d, h):
        filas = self.q(f"""SELECT CASE WHEN p.categoria='' THEN '(sin categoría)' ELSE p.categoria END AS categoria,
                           {AGREGADOS_SQL}
                           FROM movimientos m JOIN productos p ON p.id = m.producto_id {JOIN_TICKET}
                           WHERE m.fecha BETWEEN ? AND ? AND m.tipo <> 'ajuste'
                           GROUP BY 1 ORDER BY ventas DESC""", (_iso(d), _iso(h)))
        return [dict(r) for r in filas]

    def por_mes(self, d, h):
        filas = {r["mes"]: dict(r) for r in self.q(
            f"""SELECT substr(m.fecha, 1, 7) AS mes, {AGREGADOS_SQL}
                FROM movimientos m {JOIN_TICKET} WHERE m.fecha BETWEEN ? AND ? AND m.tipo <> 'ajuste'
                GROUP BY mes""", (_iso(d), _iso(h)))}
        vacio = dict(cant_v=0, ventas=0, ventas_netas=0, costo=0, cant_c=0, compras=0, iva_credito_c=0,
                     cant_m=0, mermas=0)
        for r in filas.values():
            r["gastos"] = dict(bruto=0.0, neto=0.0, iva=0.0)
            r["tickets"] = 0
        for mes, bruto, neto in self.q("""SELECT substr(fecha,1,7), SUM(monto), SUM(monto/(1+iva_tasa/100.0))
                                          FROM gastos WHERE fecha BETWEEN ? AND ? GROUP BY 1""", (_iso(d), _iso(h))):
            fila = filas.setdefault(mes, dict(vacio, mes=mes, tickets=0))
            fila["gastos"] = dict(bruto=bruto, neto=neto, iva=bruto - neto)
        for mes, n in self.q("""SELECT substr(fecha,1,7), COUNT(*) FROM tickets WHERE tipo='venta'
                                AND fecha BETWEEN ? AND ? GROUP BY 1""", (_iso(d), _iso(h))):
            if mes in filas:
                filas[mes]["tickets"] = n
        return [filas[k] for k in sorted(filas)]

    def iva_mensual(self):
        """IVA de cada mes (estilo Formulario 29) desde el primer registro, arrastrando el remanente de
        crédito fiscal al mes siguiente. Montos en pesos enteros. No aplica reajuste UTM."""
        meses = {}
        for r in self.q(f"""SELECT substr(m.fecha,1,7) AS mes,
                COALESCE(SUM(CASE WHEN m.tipo='venta' AND m.iva_tasa>0
                                  THEN m.cantidad*m.precio_unit/(1+m.iva_tasa/100.0) END),0) AS afectas,
                COALESCE(SUM(CASE WHEN m.tipo='venta' AND m.iva_tasa=0 THEN m.cantidad*m.precio_unit END),0) AS exentas,
                COALESCE(SUM(CASE WHEN m.tipo='venta'
                                  THEN m.cantidad*m.precio_unit*m.iva_tasa/(100.0+m.iva_tasa) END),0) AS debito,
                COALESCE(SUM(CASE WHEN m.tipo='compra' AND tk.con_factura=1
                                  THEN m.cantidad*m.precio_unit*m.iva_tasa/(100.0+m.iva_tasa) END),0) AS credito_c
                FROM movimientos m {JOIN_TICKET} WHERE m.tipo IN ('venta','compra') GROUP BY mes"""):
            meses[r["mes"]] = dict(mes=r["mes"], afectas=r["afectas"], exentas=r["exentas"], debito=r["debito"],
                                   credito_c=r["credito_c"], credito_g=0.0)
        for mes, iva in self.q("SELECT substr(fecha,1,7), SUM(monto*iva_tasa/(100.0+iva_tasa)) FROM gastos GROUP BY 1"):
            meses.setdefault(mes, dict(mes=mes, afectas=0.0, exentas=0.0, debito=0.0, credito_c=0.0))["credito_g"] = iva
        remanente = 0
        out = []
        for mes in sorted(meses):
            r = meses[mes]
            for k in ("afectas", "exentas", "debito", "credito_c", "credito_g"):
                r[k] = redondear(r[k])
            r["credito"] = r["credito_c"] + r["credito_g"]
            r["remanente_anterior"] = remanente
            saldo = r["debito"] - r["credito"] - remanente
            r["pagar"] = max(saldo, 0)
            remanente = max(-saldo, 0)
            r["remanente"] = remanente
            out.append(r)
        return out

    def ventas_por_medio(self, d, h):
        return self.q("""SELECT t.medio_pago, SUM(m.cantidad*m.precio_unit) AS total, COUNT(DISTINCT t.id) AS n
                         FROM tickets t JOIN movimientos m ON m.ticket_id=t.id
                         WHERE t.tipo='venta' AND t.fecha BETWEEN ? AND ?
                         GROUP BY t.medio_pago ORDER BY total DESC""", (_iso(d), _iso(h)))

    def cantidad_tickets(self, d, h, tipo="venta"):
        return self.q1("SELECT COUNT(*) FROM tickets WHERE tipo=? AND fecha BETWEEN ? AND ?",
                       (tipo, _iso(d), _iso(h)))[0]

    def ventas_sin_costo(self, d, h):
        return self.q1("""SELECT COUNT(*), COUNT(DISTINCT producto_id) FROM movimientos
                          WHERE tipo='venta' AND costo_unit=0 AND fecha BETWEEN ? AND ?""", (_iso(d), _iso(h)))

    def conteos(self, d, h):
        return {r[0]: r[1] for r in self.q("""SELECT tipo, COUNT(*) FROM movimientos
                                               WHERE fecha BETWEEN ? AND ? GROUP BY tipo""", (_iso(d), _iso(h)))}

    # ------------------------------------------------------------ respaldos
    def respaldo(self, destino):
        dst = sqlite3.connect(destino)
        try:
            self.con.backup(dst)
        finally:
            dst.close()

    @staticmethod
    def validar_archivo(ruta):
        """Comprueba que el archivo sea una base de este programa. Devuelve un resumen."""
        try:
            con = sqlite3.connect(f"file:{ruta}?mode=ro", uri=True)
            try:
                tablas = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
                if not {"productos", "movimientos"} <= tablas:
                    raise ArchivoInvalido("El archivo no es una base de datos de este programa.")
                if con.execute("PRAGMA quick_check").fetchone()[0] != "ok":
                    raise ArchivoInvalido("El archivo está dañado.")
                return dict(productos=con.execute("SELECT COUNT(*) FROM productos").fetchone()[0],
                            movimientos=con.execute("SELECT COUNT(*) FROM movimientos").fetchone()[0],
                            ultima=con.execute("SELECT MAX(fecha) FROM movimientos").fetchone()[0])
            finally:
                con.close()
        except sqlite3.DatabaseError as ex:
            raise ArchivoInvalido(f"No se pudo leer el archivo: {ex}") from ex

    def restaurar(self, ruta):
        """Reemplaza el contenido actual por el del archivo indicado (ya validado)."""
        src = sqlite3.connect(f"file:{ruta}?mode=ro", uri=True)
        try:
            src.backup(self.con)
        finally:
            src.close()
        self._preparar()
