# -*- coding: utf-8 -*-
"""Constantes y funciones de formato/lectura compartidas."""
import datetime as dt
import math
import re

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
DIAS_CORTOS = ["Lu", "Ma", "Mi", "Ju", "Vi", "Sá", "Do"]
UNIDADES = ["unid", "kg", "g", "lt", "ml", "paquete", "caja", "docena", "bolsa", "metro"]
MEDIOS_PAGO = ["Efectivo", "Tarjeta de débito", "Tarjeta de crédito", "Transferencia", "Otro"]
CATEGORIAS_GASTO = ["Alquiler", "Luz", "Agua", "Gas", "Internet / Teléfono", "Sueldos", "Impuestos",
                    "Transporte", "Mantenimiento", "Limpieza", "Otros"]
EFECTIVO = "Efectivo"
IVA_POR_DEFECTO = 19.0
GASTOS_SIN_IVA = {"Alquiler", "Sueldos", "Impuestos"}  # por defecto sin factura con IVA


def hoy():
    return dt.date.today()


def fmt_fecha(d):
    return d.strftime("%d/%m/%Y")


def iso_a_fecha(s):
    return dt.date.fromisoformat(s)


def fmt_iso(s):
    """'2026-10-08' -> '08/10/2026'."""
    return fmt_fecha(iso_a_fecha(s))


def parse_fecha(s):
    s = (s or "").strip()
    for f in ("%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y", "%d.%m.%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def parse_num(s):
    """Formato chileno: coma decimal y punto de miles. Acepta '1.500', '12,5', '1.234,5', '$ 100'.
    Un punto suelto con 1 ó 2 decimales ('1.5') se toma como decimal. Devuelve None si no es un número."""
    s = str(s or "").strip().replace(" ", "").replace("$", "")
    if not s:
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    try:
        n = float(s)
    except ValueError:
        return None
    return n if n == n and abs(n) != float("inf") else None


def redondear(x):
    """Redondeo a pesos enteros (mitad hacia arriba)."""
    return math.floor(x + 0.5)


def _formato_cl(texto):
    """'1,234.5' -> '1.234,5'."""
    return texto.replace(",", "_").replace(".", ",").replace("_", ".")


def money(x):
    v = redondear(x)
    return ("-$ " if v < 0 else "$ ") + f"{abs(v):,}".replace(",", ".")


def qty(x):
    return _formato_cl(f"{x:,.3f}".rstrip("0").rstrip("."))


def num_txt(x):
    """Número para un campo editable (sin separador de miles, coma decimal)."""
    return f"{x:.3f}".rstrip("0").rstrip(".").replace(".", ",")


def pct(x):
    return f"{x * 100:.1f} %".replace(".", ",")


def neto(bruto, tasa):
    """Monto sin IVA a partir de un monto con IVA incluido (tasa en %)."""
    return bruto / (1 + tasa / 100.0)


GASTOS_CERO = dict(bruto=0.0, neto=0.0, iva=0.0)


def totales(prods, gastos=None, tickets=0):
    """Indicadores del período a partir de las filas por producto (o por mes).
    Ventas y compras se informan con IVA; costos, mermas, gastos y ganancia, sin IVA."""
    gastos = gastos or GASTOS_CERO
    t = dict(ventas=0.0, ventas_netas=0.0, costo=0.0, compras=0.0, iva_credito_c=0.0, mermas=0.0)
    for r in prods:
        for k in t:
            t[k] += r.get(k) or 0
    t["iva_debito"] = t["ventas"] - t["ventas_netas"]
    t["gastos"] = gastos["neto"]
    t["gastos_bruto"] = gastos["bruto"]
    t["iva_credito_g"] = gastos["iva"]
    t["iva_credito"] = t["iva_credito_c"] + gastos["iva"]
    t["iva_pagar"] = t["iva_debito"] - t["iva_credito"]
    t["ganancia"] = t["ventas_netas"] - t["costo"]
    t["resultado"] = t["ganancia"] - t["mermas"] - t["gastos"]
    t["flujo"] = t["ventas"] - t["compras"] - t["gastos_bruto"]
    t["margen"] = t["ganancia"] / t["ventas_netas"] if t["ventas_netas"] else 0.0
    t["tickets"] = tickets
    t["ticket_prom"] = t["ventas"] / tickets if tickets else 0.0
    return t
