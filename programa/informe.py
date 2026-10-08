# -*- coding: utf-8 -*-
"""Generación del reporte contable en Excel (montos en pesos, IVA incluido en los precios)."""
import datetime as dt

from . import excel_simple as xl
from .util import MESES, fmt_fecha, hoy, iso_a_fecha, neto, totales

H, N, M, D, MT, Q, T, TI, P, PT, QT, S = (xl.HEADER, xl.NORMAL, xl.MONEY, xl.DATE, xl.MONEY_TOTAL, xl.QTY,
                                          xl.TOTAL, xl.TITLE, xl.PCT, xl.PCT_TOTAL, xl.QTY_TOTAL, xl.SUBTITLE)
NOTA_IVA = ("Cálculo referencial para el Formulario 29 (precios con IVA incluido). No incluye el reajuste UTM "
            "del remanente, PPM ni otros impuestos: verifique con su contador antes de declarar.")


def _cabecera(hoja, columnas, anchos):
    hoja.fila(*[(c, H) for c in columnas])
    hoja.congelar = 1
    hoja.anchos = dict(enumerate(anchos))


def _nombre_mes(mes):
    y, mo = map(int, mes.split("-"))
    return f"{MESES[mo - 1]} {y}"


def generar_excel(db, d, h, descripcion, ruta):
    negocio = db.config("negocio", "Mi Negocio")
    prods = db.por_producto(d, h)
    t = totales(prods, db.gastos_totales(d, h), db.cantidad_tickets(d, h))
    libro = xl.Libro()

    _resumen(libro, db, d, h, descripcion, negocio, t)
    meses = db.por_mes(d, h)
    if len(meses) > 1:
        _por_mes(libro, meses)
    _iva_mensual(libro, db, d, h)
    _por_producto(libro, db, prods, t)
    _por_categoria(libro, db.por_categoria(d, h))
    _ventas(libro, db.movimientos("venta", d, h))
    _compras(libro, db.movimientos("compra", d, h))
    _mermas(libro, db.movimientos("merma", d, h))
    _gastos(libro, db.gastos(d, h))
    cierres = db.cierres(d, h)
    if cierres:
        _caja(libro, cierres)
    _inventario(libro, db)
    libro.guardar(ruta)


def _resumen(libro, db, d, h, descripcion, negocio, t):
    s = libro.hoja("Resumen")
    s.anchos = {0: 38, 1: 18, 2: 66}
    s.fila((f"{negocio} — Reporte contable", TI))
    s.fila((f"Período: {descripcion}", S))
    s.fila((f"Generado el {fmt_fecha(hoy())} a las {dt.datetime.now():%H:%M}  ·  IVA {db.tasa_iva():g}% incluido "
            "en los precios", S))

    def seccion(titulo):
        s.fila()
        s.fila((titulo, H), ("Monto", H), ("Qué significa", H))

    def linea(etiqueta, valor, explicacion, destacado=False, estilo=None):
        s.fila((etiqueta, T if destacado else N), (valor, estilo or (MT if destacado else M)), (explicacion, S))

    seccion("RESULTADO DEL NEGOCIO (sin IVA)")
    linea("Ventas (con IVA)", t["ventas"], "Total cobrado a los clientes")
    linea("− IVA débito", t["iva_debito"], "IVA incluido en las ventas: no es del negocio, se paga al SII")
    linea("Ventas netas", t["ventas_netas"], "Ventas sin IVA", destacado=True)
    linea("− Costo de lo vendido (neto)", t["costo"], "Lo que costaron, sin IVA, los productos vendidos")
    linea("Ganancia bruta", t["ganancia"], "Ventas netas − costo de lo vendido", destacado=True)
    linea("Margen bruto", t["margen"], "Ganancia bruta / ventas netas", estilo=P)
    linea("− Mermas (al costo neto)", t["mermas"], "Mercadería perdida: vencida, rota, faltantes de inventario")
    linea("− Gastos del negocio (netos)", t["gastos"], "Gastos sin el IVA recuperable de las facturas")
    linea("Resultado neto", t["resultado"], "Lo que realmente ganó el negocio en el período", destacado=True)

    seccion("IVA DEL PERÍODO")
    linea("IVA débito (ventas)", t["iva_debito"], "IVA cobrado en las ventas de productos afectos")
    linea("IVA crédito (compras con factura)", t["iva_credito_c"], "IVA pagado en compras de mercadería con factura")
    linea("IVA crédito (gastos con factura)", t["iva_credito_g"], "IVA pagado en gastos con factura (luz, internet…)")
    linea("IVA crédito total", t["iva_credito"], "")
    linea("IVA del período (débito − crédito)", t["iva_pagar"],
          "Positivo: IVA a pagar. Negativo: crédito a favor (remanente). Detalle mes a mes en «IVA mensual (F29)»",
          destacado=True)

    seccion("MOVIMIENTO DE DINERO (con IVA)")
    linea("Ventas (con IVA)", t["ventas"], "")
    linea("Compras de mercadería (con IVA)", t["compras"], "Total pagado a proveedores")
    linea("Gastos (con IVA)", t["gastos_bruto"], "Total pagado en gastos")
    linea("Flujo de caja", t["flujo"], "Ventas − compras − gastos: dinero que entró menos el que salió",
          destacado=True)
    s.fila("Cantidad de ventas (tickets)", t["tickets"])
    linea("Venta promedio por ticket", t["ticket_prom"], "")

    sin_costo = db.ventas_sin_costo(d, h)
    if sin_costo[0]:
        s.fila()
        s.fila((f"⚠ Hay {sin_costo[0]} línea(s) de venta de {sin_costo[1]} producto(s) sin costo cargado: "
                "la ganancia de esos productos aparece más alta de lo real.", S))

    medios = db.ventas_por_medio(d, h)
    if medios:
        s.fila()
        s.fila(("Ventas por medio de pago", H), ("Monto", H), ("Cantidad de ventas", H))
        for r in medios:
            s.fila(r["medio_pago"], (r["total"], M), r["n"])

    cats = db.gastos_por_categoria(d, h)
    if cats:
        s.fila()
        s.fila(("Gastos por categoría (con IVA)", H), ("Monto", H), ("Cantidad", H))
        for r in cats:
            s.fila(r["categoria"], (r["total"], M), r["n"])
    s.fila()
    s.fila((NOTA_IVA, S))


def _por_mes(libro, meses):
    s = libro.hoja("Por mes")
    _cabecera(s, ["Mes", "Ventas (con IVA)", "Ventas netas", "IVA débito", "Costo de lo vendido", "Ganancia bruta",
                  "Margen %", "Mermas", "Gastos netos", "Resultado neto", "IVA crédito", "IVA del mes",
                  "Compras (con IVA)", "Gastos (con IVA)", "Flujo de caja", "Nº ventas"],
              [16, 15, 14, 13, 15, 14, 9, 12, 13, 14, 13, 13, 15, 14, 14, 9])
    r = 1
    for m in meses:
        r += 1
        tm = totales([m], m["gastos"])
        s.fila(_nombre_mes(m["mes"]),
               (m["ventas"], M), (m["ventas_netas"], M), (xl.F(f"B{r}-C{r}", tm["iva_debito"]), M),
               (m["costo"], M), (xl.F(f"C{r}-E{r}", tm["ganancia"]), M),
               (xl.F(f"IF(C{r}=0,0,F{r}/C{r})", tm["margen"]), P),
               (m["mermas"], M), (tm["gastos"], M), (xl.F(f"F{r}-H{r}-I{r}", tm["resultado"]), M),
               (tm["iva_credito"], M), (xl.F(f"D{r}-K{r}", tm["iva_pagar"]), M),
               (m["compras"], M), (tm["gastos_bruto"], M), (xl.F(f"B{r}-M{r}-N{r}", tm["flujo"]), M),
               m["tickets"])
    u = r
    tt = totales(meses, {k: sum(m["gastos"][k] for m in meses) for k in ("bruto", "neto", "iva")})
    suma = lambda c, v: (xl.F(f"SUM({c}2:{c}{u})", v), MT)  # noqa: E731
    s.fila(("TOTAL", T), suma("B", tt["ventas"]), suma("C", tt["ventas_netas"]), suma("D", tt["iva_debito"]),
           suma("E", tt["costo"]), suma("F", tt["ganancia"]),
           (xl.F(f"IF(C{u + 1}=0,0,F{u + 1}/C{u + 1})", tt["margen"]), PT),
           suma("H", tt["mermas"]), suma("I", tt["gastos"]), suma("J", tt["resultado"]),
           suma("K", tt["iva_credito"]), suma("L", tt["iva_pagar"]), suma("M", tt["compras"]),
           suma("N", tt["gastos_bruto"]), suma("O", tt["flujo"]),
           (xl.F(f"SUM(P2:P{u})", sum(m["tickets"] for m in meses)), QT))


def _iva_mensual(libro, db, d, h):
    desde, hasta = d.isoformat()[:7], h.isoformat()[:7]
    filas = [r for r in db.iva_mensual() if desde <= r["mes"] <= hasta]
    s = libro.hoja("IVA mensual (F29)")
    _cabecera(s, ["Mes", "Ventas afectas (neto)", "Ventas exentas", "IVA débito", "IVA crédito compras",
                  "IVA crédito gastos", "IVA crédito total", "Remanente mes anterior", "IVA a pagar",
                  "Remanente para el mes siguiente"], [16, 17, 14, 14, 16, 15, 15, 17, 14, 18])
    if not filas:
        s.fila(("Sin movimientos con IVA en este período", S))
        return
    r = 1
    for f in filas:
        r += 1
        rem_ant = f["remanente_anterior"] if r == 2 else xl.F(f"J{r - 1}", f["remanente_anterior"])
        s.fila(_nombre_mes(f["mes"]), (f["afectas"], M), (f["exentas"], M), (f["debito"], M),
               (f["credito_c"], M), (f["credito_g"], M), (xl.F(f"E{r}+F{r}", f["credito"]), M),
               (rem_ant, M), (xl.F(f"MAX(D{r}-G{r}-H{r},0)", f["pagar"]), MT),
               (xl.F(f"MAX(G{r}+H{r}-D{r},0)", f["remanente"]), M))
    u = r
    s.fila(("TOTAL", T), (xl.F(f"SUM(B2:B{u})", sum(f["afectas"] for f in filas)), MT),
           (xl.F(f"SUM(C2:C{u})", sum(f["exentas"] for f in filas)), MT),
           (xl.F(f"SUM(D2:D{u})", sum(f["debito"] for f in filas)), MT),
           (xl.F(f"SUM(E2:E{u})", sum(f["credito_c"] for f in filas)), MT),
           (xl.F(f"SUM(F2:F{u})", sum(f["credito_g"] for f in filas)), MT),
           (xl.F(f"SUM(G2:G{u})", sum(f["credito"] for f in filas)), MT), ("", T),
           (xl.F(f"SUM(I2:I{u})", sum(f["pagar"] for f in filas)), MT), ("", T))
    s.fila()
    s.fila((NOTA_IVA, S))
    s.fila(("El remanente del primer mes mostrado viene de los meses anteriores registrados en el programa.", S))


def _por_producto(libro, db, prods, t):
    s = libro.hoja("Por producto")
    _cabecera(s, ["Producto", "Categoría", "Unidad", "IVA", "Cant. vendida", "Ventas (con IVA)", "Ventas netas",
                  "IVA débito", "Costo neto", "Ganancia", "Margen %", "Cant. comprada", "Compras (con IVA)",
                  "Cant. merma", "Mermas (neto)", "Stock actual"],
              [30, 15, 8, 8, 11, 15, 14, 12, 13, 13, 9, 12, 15, 10, 13, 11])
    stock = {p["id"]: p["stock"] for p in db.productos(inactivos=True)}
    if not prods:
        s.fila(("Sin movimientos en este período", S))
        return
    r = 1
    for p in prods:
        r += 1
        gan = p["ventas_netas"] - p["costo"]
        s.fila(p["nombre"], p["categoria"], p["unidad"], "Afecto" if p["afecto_iva"] else "Exento",
               (round(p["cant_v"], 3), Q), (p["ventas"], M), (p["ventas_netas"], M),
               (xl.F(f"F{r}-G{r}", p["ventas"] - p["ventas_netas"]), M), (p["costo"], M),
               (xl.F(f"G{r}-I{r}", gan), M),
               (xl.F(f"IF(G{r}=0,0,J{r}/G{r})", gan / p["ventas_netas"] if p["ventas_netas"] else 0), P),
               (round(p["cant_c"], 3), Q), (p["compras"], M), (round(p["cant_m"], 3), Q), (p["mermas"], M),
               (round(stock.get(p["id"], 0), 3), Q))
    u = r
    s.filtro = f"A1:P{u}"
    suma = lambda c, v: (xl.F(f"SUM({c}2:{c}{u})", v), MT)  # noqa: E731
    s.fila(("TOTAL", T), ("", T), ("", T), ("", T), ("", T), suma("F", t["ventas"]), suma("G", t["ventas_netas"]),
           suma("H", t["iva_debito"]), suma("I", t["costo"]), suma("J", t["ganancia"]),
           (xl.F(f"IF(G{u + 1}=0,0,J{u + 1}/G{u + 1})", t["margen"]), PT), ("", T), suma("M", t["compras"]),
           ("", T), suma("O", t["mermas"]), ("", T))


def _por_categoria(libro, cats):
    if not cats:
        return
    s = libro.hoja("Por categoría")
    _cabecera(s, ["Categoría", "Ventas (con IVA)", "Ventas netas", "Costo neto", "Ganancia", "Margen %",
                  "Compras (con IVA)", "Mermas (neto)"], [24, 16, 15, 15, 15, 10, 16, 14])
    r = 1
    for c in cats:
        r += 1
        gan = c["ventas_netas"] - c["costo"]
        s.fila(c["categoria"], (c["ventas"], M), (c["ventas_netas"], M), (c["costo"], M),
               (xl.F(f"C{r}-D{r}", gan), M),
               (xl.F(f"IF(C{r}=0,0,E{r}/C{r})", gan / c["ventas_netas"] if c["ventas_netas"] else 0), P),
               (c["compras"], M), (c["mermas"], M))
    u = r
    tc = totales(cats)
    suma = lambda col, v: (xl.F(f"SUM({col}2:{col}{u})", v), MT)  # noqa: E731
    s.fila(("TOTAL", T), suma("B", tc["ventas"]), suma("C", tc["ventas_netas"]), suma("D", tc["costo"]),
           suma("E", tc["ganancia"]), (xl.F(f"IF(C{u + 1}=0,0,E{u + 1}/C{u + 1})", tc["margen"]), PT),
           suma("G", tc["compras"]), suma("H", tc["mermas"]))


def _sin_registros(s, filas, texto="Sin registros en este período"):
    if not filas:
        s.fila((texto, S))
        return True
    return False


def _ventas(libro, filas):
    s = libro.hoja("Ventas")
    _cabecera(s, ["Fecha", "Producto", "Unidad", "Cantidad", "Precio unit.", "Total", "IVA %", "Neto", "IVA",
                  "Costo unit.", "Ganancia (neta)", "Ticket Nº", "Medio de pago", "Cliente / nota"],
              [11, 30, 8, 10, 12, 13, 7, 13, 12, 12, 14, 9, 17, 24])
    if _sin_registros(s, filas):
        return
    r = 1
    tot = tot_neto = gan_tot = 0.0
    for m in reversed(filas):  # orden cronológico
        r += 1
        total = m["cantidad"] * m["precio_unit"]
        n = neto(total, m["iva_tasa"])
        gan = n - neto(m["cantidad"] * m["costo_unit"], m["iva_tasa"])
        tot, tot_neto, gan_tot = tot + total, tot_neto + n, gan_tot + gan
        s.fila((iso_a_fecha(m["fecha"]), D), m["nombre"], m["unidad"], (m["cantidad"], Q), (m["precio_unit"], M),
               (xl.F(f"D{r}*E{r}", total), M), (m["iva_tasa"], Q), (xl.F(f"F{r}/(1+G{r}/100)", n), M),
               (xl.F(f"F{r}-H{r}", total - n), M), (m["costo_unit"], M),
               (xl.F(f"H{r}-D{r}*J{r}/(1+G{r}/100)", gan), M), m["ticket_id"], m["medio_pago"] or "", m["nota"])
    u = r
    s.filtro = f"A1:N{u}"
    s.fila(("TOTAL", T), (f"{len(filas)} líneas", T), ("", T), ("", T), ("", T),
           (xl.F(f"SUM(F2:F{u})", tot), MT), ("", T), (xl.F(f"SUM(H2:H{u})", tot_neto), MT),
           (xl.F(f"SUM(I2:I{u})", tot - tot_neto), MT), ("", T), (xl.F(f"SUM(K2:K{u})", gan_tot), MT),
           ("", T), ("", T), ("", T))


def _compras(libro, filas):
    s = libro.hoja("Compras")
    _cabecera(s, ["Fecha", "Producto", "Unidad", "Cantidad", "Costo unit.", "Total", "Factura", "IVA %", "Neto",
                  "IVA crédito", "Ticket Nº", "Medio de pago", "Proveedor / nota"],
              [11, 30, 8, 10, 12, 13, 8, 7, 13, 12, 9, 17, 24])
    if _sin_registros(s, filas):
        return
    r = 1
    tot = tot_neto = cred = 0.0
    for m in reversed(filas):
        r += 1
        total = m["cantidad"] * m["precio_unit"]
        n = neto(total, m["iva_tasa"])
        factura = bool(m["con_factura"])
        tot, tot_neto, cred = tot + total, tot_neto + n, cred + (total - n if factura else 0)
        s.fila((iso_a_fecha(m["fecha"]), D), m["nombre"], m["unidad"], (m["cantidad"], Q), (m["precio_unit"], M),
               (xl.F(f"D{r}*E{r}", total), M), "Sí" if factura else "No", (m["iva_tasa"], Q),
               (xl.F(f"F{r}/(1+H{r}/100)", n), M), (xl.F(f'IF(G{r}="Sí",F{r}-I{r},0)', total - n if factura else 0), M),
               m["ticket_id"], m["medio_pago"] or "", m["nota"])
    u = r
    s.filtro = f"A1:M{u}"
    s.fila(("TOTAL", T), (f"{len(filas)} líneas", T), ("", T), ("", T), ("", T),
           (xl.F(f"SUM(F2:F{u})", tot), MT), ("", T), ("", T), (xl.F(f"SUM(I2:I{u})", tot_neto), MT),
           (xl.F(f"SUM(J2:J{u})", cred), MT), ("", T), ("", T), ("", T))


def _mermas(libro, filas):
    s = libro.hoja("Mermas")
    _cabecera(s, ["Fecha", "Producto", "Unidad", "Cantidad", "Costo unit.", "Total (con IVA)", "IVA %",
                  "Pérdida neta", "Motivo"], [11, 30, 8, 10, 12, 14, 7, 13, 30])
    if _sin_registros(s, filas):
        return
    r = 1
    tot = tot_neto = 0.0
    for m in reversed(filas):
        r += 1
        total = m["cantidad"] * m["costo_unit"]
        n = neto(total, m["iva_tasa"])
        tot, tot_neto = tot + total, tot_neto + n
        s.fila((iso_a_fecha(m["fecha"]), D), m["nombre"], m["unidad"], (m["cantidad"], Q), (m["costo_unit"], M),
               (xl.F(f"D{r}*E{r}", total), M), (m["iva_tasa"], Q), (xl.F(f"F{r}/(1+G{r}/100)", n), M), m["nota"])
    u = r
    s.filtro = f"A1:I{u}"
    s.fila(("TOTAL", T), (f"{len(filas)} registros", T), ("", T), ("", T), ("", T),
           (xl.F(f"SUM(F2:F{u})", tot), MT), ("", T), (xl.F(f"SUM(H2:H{u})", tot_neto), MT), ("", T))


def _gastos(libro, filas):
    s = libro.hoja("Gastos")
    _cabecera(s, ["Fecha", "Categoría", "Descripción", "Medio de pago", "Monto (con IVA)", "Factura", "IVA %",
                  "Neto", "IVA crédito"], [11, 20, 36, 17, 15, 8, 7, 14, 13])
    if _sin_registros(s, filas, "Sin gastos en este período"):
        return
    r = 1
    tot = tot_neto = 0.0
    for g in reversed(filas):
        r += 1
        n = neto(g["monto"], g["iva_tasa"])
        tot, tot_neto = tot + g["monto"], tot_neto + n
        s.fila((iso_a_fecha(g["fecha"]), D), g["categoria"], g["descripcion"], g["medio_pago"], (g["monto"], M),
               "Sí" if g["iva_tasa"] > 0 else "No", (g["iva_tasa"], Q), (xl.F(f"E{r}/(1+G{r}/100)", n), M),
               (xl.F(f"E{r}-H{r}", g["monto"] - n), M))
    u = r
    s.filtro = f"A1:I{u}"
    s.fila(("TOTAL", T), (f"{len(filas)} gastos", T), ("", T), ("", T), (xl.F(f"SUM(E2:E{u})", tot), MT),
           ("", T), ("", T), (xl.F(f"SUM(H2:H{u})", tot_neto), MT), (xl.F(f"SUM(I2:I{u})", tot - tot_neto), MT))


def _caja(libro, cierres):
    s = libro.hoja("Cierres de caja")
    _cabecera(s, ["Fecha", "Efectivo inicial", "Ventas efectivo", "Compras efectivo", "Gastos efectivo",
                  "Retiros", "Esperado", "Contado", "Diferencia", "Nota"],
              [11, 15, 15, 15, 15, 13, 15, 15, 13, 30])
    r = 1
    for c in reversed(cierres):
        r += 1
        s.fila((iso_a_fecha(c["fecha"]), D), (c["inicial"], M), (c["ventas_ef"], M), (c["compras_ef"], M),
               (c["gastos_ef"], M), (c["retiros"], M), (c["esperado"], M), (c["contado"], M),
               (xl.F(f"H{r}-G{r}", c["diferencia"]), M), c["nota"])
    u = r
    s.fila(("TOTAL", T), ("", T), (xl.F(f"SUM(C2:C{u})", sum(c["ventas_ef"] for c in cierres)), MT),
           (xl.F(f"SUM(D2:D{u})", sum(c["compras_ef"] for c in cierres)), MT),
           (xl.F(f"SUM(E2:E{u})", sum(c["gastos_ef"] for c in cierres)), MT),
           (xl.F(f"SUM(F2:F{u})", sum(c["retiros"] for c in cierres)), MT), ("", T), ("", T),
           (xl.F(f"SUM(I2:I{u})", sum(c["diferencia"] for c in cierres)), MT), ("", T))


def _inventario(libro, db):
    s = libro.hoja("Inventario actual")
    _cabecera(s, ["Producto", "Categoría", "Unidad", "IVA", "Stock", "Stock mínimo", "Costo unit.",
                  "Valor al costo", "Precio venta", "Estado"], [30, 15, 8, 8, 10, 12, 12, 15, 13, 14])
    prods = db.productos()
    if _sin_registros(s, prods, "Sin productos"):
        return
    r = 1
    for p in prods:
        r += 1
        if p["stock"] <= 0:
            estado = "Sin stock"
        elif p["stock_minimo"] > 0 and p["stock"] <= p["stock_minimo"]:
            estado = "Reponer"
        else:
            estado = "OK"
        s.fila(p["nombre"], p["categoria"], p["unidad"], "Afecto" if p["afecto_iva"] else "Exento",
               (round(p["stock"], 3), Q), (p["stock_minimo"], Q), (p["precio_compra"], M),
               (xl.F(f"MAX(E{r},0)*G{r}", max(p["stock"], 0) * p["precio_compra"]), M), (p["precio_venta"], M), estado)
    u = r
    s.filtro = f"A1:J{u}"
    s.fila(("TOTAL", T), ("", T), ("", T), ("", T), ("", T), ("", T), ("", T),
           (xl.F(f"SUM(H2:H{u})", sum(max(p["stock"], 0) * p["precio_compra"] for p in prods)), MT), ("", T), ("", T))
