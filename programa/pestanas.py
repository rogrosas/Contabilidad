# -*- coding: utf-8 -*-
"""Pestañas de la ventana principal."""
import datetime as dt
import os
import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from .informe import generar_excel
from .util import (EFECTIVO, GASTOS_SIN_IVA, MEDIOS_PAGO, UNIDADES, fmt_fecha, fmt_iso, hoy, money, neto,
                   num_txt, parse_num, pct, qty, redondear, totales)
from .widgets import CampoFecha, PeriodoSelector, ProductoCombo, crear_tabla, limpiar_tabla

MAX_FILAS = 3000  # filas mostradas en las listas (los totales usan todas)


def _aviso(widget, msg, foco=None):
    messagebox.showwarning("Revise los datos", msg, parent=widget)
    if foco is not None:
        foco.focus_set()
    return "break"


def _bind_enter(widget, func):
    widget.bind("<Return>", func)
    widget.bind("<KP_Enter>", func)


# =====================================================================================
# Ventas y compras (ticket con varios productos)
# =====================================================================================
class TicketTab(ttk.Frame):
    def __init__(self, master, app, tipo):
        super().__init__(master, padding=(8, 6))
        self.app, self.db, self.tipo = app, app.db, tipo
        self.es_venta = tipo == "venta"
        self.nombre = "venta" if self.es_venta else "compra"
        self.items = []
        self.tid = None
        self.orig_cant = {}
        self._pid = None

        self.v_prod, self.v_cant, self.v_precio = tk.StringVar(), tk.StringVar(), tk.StringVar()
        self.v_info, self.v_aviso = tk.StringVar(), tk.StringVar()
        self.v_medio = tk.StringVar(value=EFECTIVO)
        self.v_nota, self.v_paga = tk.StringVar(), tk.StringVar()
        self.v_total, self.v_vuelto = tk.StringVar(value=money(0)), tk.StringVar()
        self.v_desglose = tk.StringVar()
        self.v_factura = tk.BooleanVar(value=True)
        self.v_buscar = tk.StringVar()

        arriba = ttk.Frame(self)
        arriba.pack(fill="x")

        # ---- carga de productos y carrito
        self.marco = ttk.LabelFrame(arriba, padding=(8, 4))
        self.marco.pack(side="left", fill="both", expand=True)
        fila = ttk.Frame(self.marco)
        fila.pack(fill="x")
        precio_txt = "Precio unit." if self.es_venta else "Costo unit."
        for c, txt in enumerate(["Producto (escriba y Enter)", "Cantidad", precio_txt]):
            ttk.Label(fila, text=txt).grid(row=0, column=c, sticky="w", padx=3)
        self.cb_prod = ProductoCombo(fila, textvariable=self.v_prod, width=38)
        self.e_cant = ttk.Entry(fila, textvariable=self.v_cant, width=9)
        self.e_precio = ttk.Entry(fila, textvariable=self.v_precio, width=12)
        b_agregar = ttk.Button(fila, text="+ Agregar", command=self.agregar)
        for c, w in enumerate([self.cb_prod, self.e_cant, self.e_precio, b_agregar]):
            w.grid(row=1, column=c, sticky="ew", padx=3)
        fila.columnconfigure(0, weight=1)
        info = ttk.Frame(self.marco)
        info.pack(fill="x", pady=(2, 2))
        ttk.Label(info, textvariable=self.v_info, foreground="#555555").pack(side="left", padx=3)
        ttk.Label(info, textvariable=self.v_aviso, foreground="#b00020").pack(side="left", padx=6)

        cols = [("producto", "Producto", 260, "w"), ("cantidad", "Cantidad", 80, "e"), ("unidad", "Unid.", 60, "center"),
                ("precio", precio_txt, 100, "e"), ("subtotal", "Subtotal", 110, "e")]
        marco_c, self.carrito = crear_tabla(self.marco, cols, seleccion="browse", alto=4)
        marco_c.pack(fill="both", expand=True)
        self.carrito.bind("<Delete>", lambda e: self.quitar_linea())
        self.carrito.bind("<Double-1>", lambda e: self.editar_linea())
        bl = ttk.Frame(self.marco)
        bl.pack(fill="x", pady=(4, 0))
        ttk.Button(bl, text="Quitar línea (Supr)", command=self.quitar_linea).pack(side="left")
        ttk.Button(bl, text="Corregir línea", command=self.editar_linea).pack(side="left", padx=4)

        # ---- panel de cobro
        panel = ttk.LabelFrame(arriba, text="Cobro" if self.es_venta else "Pago al proveedor", padding=8)
        panel.pack(side="left", fill="y", padx=(8, 0))
        ttk.Label(panel, text="Fecha").grid(row=0, column=0, sticky="w")
        self.c_fecha = CampoFecha(panel)
        self.c_fecha.grid(row=0, column=1, sticky="ew", pady=2)
        ttk.Label(panel, text="Medio de pago").grid(row=1, column=0, sticky="w", padx=(0, 6))
        cm = ttk.Combobox(panel, textvariable=self.v_medio, values=MEDIOS_PAGO, state="readonly", width=18)
        cm.grid(row=1, column=1, sticky="ew", pady=2)
        cm.bind("<<ComboboxSelected>>", lambda e: self._medio_cambiado())
        ttk.Label(panel, text="Cliente / nota" if self.es_venta else "Proveedor / nota").grid(row=2, column=0, sticky="w")
        self.e_nota = ttk.Entry(panel, textvariable=self.v_nota, width=20)
        self.e_nota.grid(row=2, column=1, sticky="ew", pady=2)
        if not self.es_venta:
            ttk.Checkbutton(panel, text="Con factura (genera IVA crédito)", variable=self.v_factura,
                            command=self._dibujar_carrito).grid(row=3, column=0, columnspan=2, sticky="w", pady=2)
        ttk.Label(panel, text="TOTAL", style="Sub.TLabel").grid(row=4, column=0, sticky="w", pady=(4, 0))
        ttk.Label(panel, textvariable=self.v_total, style="Grande.TLabel").grid(row=4, column=1, sticky="e", pady=(4, 0))
        ttk.Label(panel, textvariable=self.v_desglose, foreground="#555555").grid(row=5, column=0, columnspan=2,
                                                                                 sticky="e")
        self.l_paga = ttk.Label(panel, text="Paga con")
        self.e_paga = ttk.Entry(panel, textvariable=self.v_paga, width=12, justify="right")
        self.l_vuelto_t = ttk.Label(panel, text="Vuelto")
        self.l_vuelto = ttk.Label(panel, textvariable=self.v_vuelto, style="Vuelto.TLabel")
        if self.es_venta:
            self.l_paga.grid(row=6, column=0, sticky="w")
            self.e_paga.grid(row=6, column=1, sticky="e", pady=2)
            self.l_vuelto_t.grid(row=7, column=0, sticky="w")
            self.l_vuelto.grid(row=7, column=1, sticky="e")
        self.b_guardar = ttk.Button(panel, style="Cobrar.TButton", command=self.guardar)
        self.b_guardar.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(6, 2))
        ttk.Button(panel, text="Cancelar ticket", command=self.cancelar).grid(row=9, column=0, columnspan=2, sticky="ew")
        panel.columnconfigure(1, weight=1)

        # ---- eventos
        self.cb_prod.bind("<<ComboboxSelected>>", lambda e: self._producto_elegido())
        self.cb_prod.bind("<FocusOut>", lambda e: self._producto_elegido())
        _bind_enter(self.cb_prod, self._enter_producto)
        _bind_enter(self.e_cant, self._enter_cantidad)
        _bind_enter(self.e_precio, lambda e: self.agregar())
        _bind_enter(self.e_paga, lambda e: self.guardar())
        _bind_enter(self.e_nota, lambda e: self.guardar())
        self.v_paga.trace_add("write", lambda *a: self._calcular_vuelto())

        # ---- historial
        filtros = ttk.Frame(self)
        filtros.pack(fill="x", pady=(8, 4))
        ttk.Label(filtros, text=f"{'Ventas' if self.es_venta else 'Compras'} registradas —",
                  style="Sub.TLabel").pack(side="left", padx=(0, 8))
        self.periodo = PeriodoSelector(filtros, al_cambiar=self.cargar_historial)
        self.periodo.pack(side="left")
        ttk.Label(filtros, text="Buscar:").pack(side="left", padx=(14, 3))
        eb = ttk.Entry(filtros, textvariable=self.v_buscar, width=18)
        eb.pack(side="left")
        _bind_enter(eb, lambda e: self.cargar_historial())
        ttk.Button(filtros, text="Buscar", command=self.cargar_historial).pack(side="left", padx=3)

        cols = [("num", "Nº", 55, "center"), ("fecha", "Fecha", 88, "center"), ("resumen", "Productos", 300, "w"),
                ("medio", "Medio de pago", 120, "center"), ("total", "Total", 100, "e"),
                ("iva", "IVA" if self.es_venta else "IVA crédito", 90, "e")]
        if self.es_venta:
            cols.append(("ganancia", "Ganancia neta", 105, "e"))
        cols.append(("nota", "Cliente / nota" if self.es_venta else "Proveedor / nota", 150, "w"))
        marco_h, self.historial = crear_tabla(self, cols)
        marco_h.pack(fill="both", expand=True)
        self.historial.bind("<Double-1>", lambda e: self.editar_ticket())
        self.historial.bind("<Delete>", lambda e: self.eliminar_tickets())
        pie = ttk.Frame(self)
        pie.pack(fill="x", pady=(4, 0), side="bottom", before=marco_h)
        ttk.Button(pie, text="Abrir / corregir", command=self.editar_ticket).pack(side="left")
        ttk.Button(pie, text="Eliminar", command=self.eliminar_tickets).pack(side="left", padx=4)
        self.l_totales = ttk.Label(pie, style="Total.TLabel")
        self.l_totales.pack(side="right")

        self._modo_nuevo()

    # ------------------------------------------------------------ estado general
    def refrescar(self):
        self.cb_prod.set_items(self.db.nombres_activos())
        self.cargar_historial()

    def hay_cambios(self):
        return bool(self.items) or bool(self.v_prod.get().strip())

    def _modo_nuevo(self):
        self.tid = None
        self.orig_cant = {}
        self.marco.config(text=f"Nueva {self.nombre}")
        self.b_guardar.config(text=("Cobrar y guardar  (F12)" if self.es_venta else "Guardar compra  (F12)"))

    def _limpiar_linea(self):
        self._pid = None
        for v in (self.v_prod, self.v_cant, self.v_precio, self.v_info, self.v_aviso):
            v.set("")
        self.cb_prod.restaurar()

    def limpiar(self, foco=True):
        self.items = []
        self._limpiar_linea()
        self.v_nota.set("")
        self.v_paga.set("")
        self.v_medio.set(EFECTIVO)
        self.v_factura.set(True)
        self._modo_nuevo()
        self._medio_cambiado()
        self._dibujar_carrito()
        if foco:
            self.cb_prod.focus_set()

    def cancelar(self):
        if self.items and not messagebox.askyesno("Cancelar", "¿Descartar el ticket actual?", parent=self):
            return
        self.limpiar()

    # ------------------------------------------------------------ carga de líneas
    def _producto_elegido(self):
        nombre = self.v_prod.get().strip()
        p = self.db.producto_por_nombre(nombre) if nombre else None
        if not p:
            self._pid = None
            self.v_info.set("Producto nuevo: se creará al agregarlo" if nombre else "")
            return
        if p["id"] != self._pid:
            self._pid = p["id"]
            self.v_prod.set(p["nombre"])
            precio = p["precio_venta"] if self.es_venta else p["precio_compra"]
            self.v_precio.set(num_txt(precio) if precio else "")
        minimo = f" (mín. {qty(p['stock_minimo'])})" if p["stock_minimo"] else ""
        self.v_info.set(f"Stock: {qty(p['stock'])} {p['unidad']}{minimo}   ·   Costo: {money(p['precio_compra'])}"
                        f"   ·   Venta: {money(p['precio_venta'])}")

    def _enter_producto(self, e=None):
        texto = self.v_prod.get().strip()
        if not texto:  # Enter con el campo vacío: pasar al cobro
            if self.items:
                destino = self.e_paga if (self.es_venta and self.v_medio.get() == EFECTIVO) else self.b_guardar
                destino.focus_set()
            return "break"
        exacto = self.db.producto_por_nombre(texto)
        opciones = self.cb_prod.coincidencias()
        if not exacto and len(opciones) == 1:
            self.v_prod.set(opciones[0])
        elif not exacto and len(opciones) > 1:
            self.cb_prod["values"] = opciones
            self.cb_prod.event_generate("<Down>")  # despliega la lista para elegir
            return "break"
        self._producto_elegido()
        if not self.v_cant.get().strip():
            self.v_cant.set("1")
        self.e_cant.focus_set()
        self.e_cant.select_range(0, "end")
        return "break"

    def _enter_cantidad(self, e=None):
        if not self.v_precio.get().strip():
            self.e_precio.focus_set()
            return "break"
        self.agregar()
        return "break"

    def agregar(self):
        """Pasa la línea escrita al ticket. Devuelve True si se agregó."""
        nombre = self.v_prod.get().strip()
        if not nombre:
            _aviso(self, "Elija o escriba un producto.", self.cb_prod)
            return False
        cant = parse_num(self.v_cant.get())
        if not cant or cant <= 0:
            _aviso(self, "Ingrese una cantidad mayor que cero.", self.e_cant)
            return False
        precio = parse_num(self.v_precio.get())
        if precio is None or precio < 0:
            _aviso(self, "Ingrese el precio unitario.", self.e_precio)
            return False
        p = self.db.producto_por_nombre(nombre)
        if not p:
            if not messagebox.askyesno("Producto nuevo", f"El producto «{nombre}» no existe.\n¿Desea crearlo?",
                                       parent=self):
                return False
            pv, pc = (precio, 0.0) if self.es_venta else (0.0, precio)
            p = self.db.producto(self.db.guardar_producto(None, nombre, "unid", "", pc, pv, 0))
            self.cb_prod.set_items(self.db.nombres_activos())
        for it in self.items:
            if it["pid"] == p["id"] and abs(it["precio"] - precio) < 1e-9:
                it["cantidad"] += cant
                break
        else:
            self.items.append(dict(pid=p["id"], nombre=p["nombre"], unidad=p["unidad"], cantidad=cant,
                                   precio=precio, costo=None, tasa=self.db.tasa_producto(p["id"])))
        self._limpiar_linea()
        self._dibujar_carrito()
        self.cb_prod.focus_set()
        return True

    def quitar_linea(self):
        sel = self.carrito.selection()
        if sel:
            del self.items[int(sel[0])]
            self._dibujar_carrito()

    def editar_linea(self):
        sel = self.carrito.selection()
        if not sel:
            return
        it = self.items.pop(int(sel[0]))
        self._pid = it["pid"]
        self.v_prod.set(it["nombre"])
        self.v_cant.set(num_txt(it["cantidad"]))
        self.v_precio.set(num_txt(it["precio"]))
        self._producto_elegido()
        self._dibujar_carrito()
        self.e_cant.focus_set()
        self.e_cant.select_range(0, "end")

    def _dibujar_carrito(self):
        limpiar_tabla(self.carrito)
        total = 0.0
        faltan = []
        pedidos = {}
        for it in self.items:
            pedidos[it["pid"]] = pedidos.get(it["pid"], 0) + it["cantidad"]
        for i, it in enumerate(self.items):
            sub = it["cantidad"] * it["precio"]
            total += sub
            tags = []
            if self.es_venta:
                p = self.db.producto(it["pid"])
                disponible = p["stock"] + self.orig_cant.get(it["pid"], 0)
                if pedidos[it["pid"]] > disponible + 1e-9:
                    tags.append("alerta")
                    if p["nombre"] not in faltan:
                        faltan.append(p["nombre"])
            self.carrito.insert("", "end", iid=str(i), tags=tags, values=(
                it["nombre"], qty(it["cantidad"]), it["unidad"], money(it["precio"]), money(sub)))
        self.v_total.set(money(total))
        self.v_aviso.set(f"⚠ Stock insuficiente: {', '.join(faltan)}" if faltan else "")
        n = sum(neto(it["cantidad"] * it["precio"], it["tasa"]) for it in self.items)
        if not self.items:
            self.v_desglose.set("")
        elif self.es_venta:
            self.v_desglose.set(f"Neto {money(n)}  ·  IVA {money(total - n)}")
        elif self.v_factura.get():
            self.v_desglose.set(f"Neto {money(n)}  ·  IVA crédito {money(total - n)}")
        else:
            self.v_desglose.set("Sin factura: no genera IVA crédito")
        self._calcular_vuelto()

    def total(self):
        """Total del ticket en pesos enteros."""
        return redondear(sum(it["cantidad"] * it["precio"] for it in self.items))

    def _medio_cambiado(self):
        if not self.es_venta:
            return
        if self.v_medio.get() == EFECTIVO:
            for w in (self.l_paga, self.e_paga, self.l_vuelto_t, self.l_vuelto):
                w.grid()
        else:
            self.v_paga.set("")
            for w in (self.l_paga, self.e_paga, self.l_vuelto_t, self.l_vuelto):
                w.grid_remove()

    def _calcular_vuelto(self):
        paga = parse_num(self.v_paga.get())
        if paga is None:
            self.v_vuelto.set("")
            return
        dif = paga - self.total()
        self.v_vuelto.set(money(dif) if dif >= -0.005 else f"Faltan {money(-dif)}")

    # ------------------------------------------------------------ guardar
    def guardar(self, e=None):
        if self.v_prod.get().strip() and self.v_cant.get().strip():
            if not self.agregar():  # línea escrita pero sin agregar: se agrega primero
                return "break"
        if not self.items:
            return _aviso(self, "Agregue al menos un producto al ticket.", self.cb_prod)
        fecha = self.c_fecha.fecha()
        if not fecha:
            return _aviso(self, "Fecha inválida. Use el formato dd/mm/aaaa.", self.c_fecha.entry)
        total = self.total()
        pago_con = None
        if self.es_venta and self.v_medio.get() == EFECTIVO and self.v_paga.get().strip():
            pago_con = parse_num(self.v_paga.get())
            if pago_con is None:
                return _aviso(self, "El monto de «Paga con» no es válido.", self.e_paga)
            if pago_con < total - 0.005:
                return _aviso(self, f"El cliente paga {money(pago_con)} pero el total es {money(total)}.", self.e_paga)
        editando = self.tid is not None
        tid = self.db.guardar_ticket(self.tid, self.tipo, fecha, self.v_medio.get(), self.v_nota.get().strip(),
                                     pago_con, self.items, con_factura=self.es_venta or self.v_factura.get())
        msg = f"{self.nombre.capitalize()} Nº {tid} {'actualizada' if editando else 'guardada'}: {money(total)}"
        if pago_con is not None:
            msg += f"   ·   Vuelto: {money(pago_con - total)}"
        try:
            d, h = self.periodo.rango()
            if not d <= fecha <= h:
                msg += f"   (fecha {fmt_fecha(fecha)}: fuera del período que se muestra)"
        except ValueError:
            pass
        self.app.estado(msg)
        self.limpiar()
        self.cargar_historial()
        self.app.actualizar_alertas()
        return "break"

    # ------------------------------------------------------------ historial
    def cargar_historial(self):
        try:
            d, h = self.periodo.rango()
        except ValueError as ex:
            messagebox.showerror("Período", str(ex), parent=self)
            return
        filas = self.db.tickets(self.tipo, d, h, self.v_buscar.get())
        limpiar_tabla(self.historial)
        total = gan = iva_total = 0.0
        for i, t in enumerate(filas):
            iva = t["iva"] if (self.es_venta or t["con_factura"]) else 0.0
            total += t["total"]
            gan += t["ganancia"]
            iva_total += iva
            if i >= MAX_FILAS:
                continue
            vals = [t["id"], fmt_iso(t["fecha"]), t["resumen"], t["medio_pago"], money(t["total"]),
                    money(iva) if (self.es_venta or t["con_factura"]) else "sin factura"]
            if self.es_venta:
                vals.append(money(t["ganancia"]))
            vals.append(t["nota"])
            self.historial.insert("", "end", iid=str(t["id"]), values=vals, tags=("par",) if i % 2 else ())
        txt = f"{len(filas)} {'ventas' if self.es_venta else 'compras'}   ·   Total: {money(total)}"
        txt += f"   ·   {'IVA' if self.es_venta else 'IVA crédito'}: {money(iva_total)}"
        if self.es_venta:
            txt += f"   ·   Ganancia neta: {money(gan)}"
        if len(filas) > MAX_FILAS:
            txt = f"(mostrando {MAX_FILAS})   " + txt
        self.l_totales.config(text=txt)

    def editar_ticket(self):
        sel = self.historial.selection()
        if not sel:
            return
        tid = int(sel[0])
        if self.items and self.tid != tid and not messagebox.askyesno(
                "Ticket en curso", "Hay un ticket sin guardar. ¿Descartarlo y abrir el seleccionado?", parent=self):
            return
        t = self.db.ticket(tid)
        if not t:
            return
        self.limpiar()
        self.tid = tid
        for m in self.db.ticket_items(tid):
            self.items.append(dict(pid=m["producto_id"], nombre=m["nombre"], unidad=m["unidad"],
                                   cantidad=m["cantidad"], precio=m["precio_unit"],
                                   costo=m["costo_unit"] if self.es_venta else None, tasa=m["iva_tasa"]))
            self.orig_cant[m["producto_id"]] = self.orig_cant.get(m["producto_id"], 0) + m["cantidad"]
        self.c_fecha.set(dt.date.fromisoformat(t["fecha"]))
        self.v_medio.set(t["medio_pago"])
        self._medio_cambiado()
        self.v_nota.set(t["nota"])
        self.v_factura.set(bool(t["con_factura"]))
        self.v_paga.set(num_txt(t["pago_con"]) if t["pago_con"] else "")
        self.marco.config(text=f"Corrigiendo {self.nombre} Nº {tid}")
        self.b_guardar.config(text="Guardar cambios  (F12)")
        self._dibujar_carrito()

    def eliminar_tickets(self):
        sel = [int(i) for i in self.historial.selection()]
        if not sel:
            return
        txt = f"la {self.nombre} Nº {sel[0]}" if len(sel) == 1 else f"estas {len(sel)} {self.nombre}s"
        if not messagebox.askyesno("Eliminar", f"¿Eliminar {txt}? No se puede deshacer.", parent=self):
            return
        self.db.eliminar_tickets(sel)
        if self.tid in sel:
            self.limpiar()
        self.cargar_historial()
        self.app.actualizar_alertas()
        self.app.estado(f"{len(sel)} {self.nombre}(s) eliminada(s).")


# =====================================================================================
# Mermas (una línea por registro)
# =====================================================================================
class MermasTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=(8, 6))
        self.app, self.db = app, app.db
        self.edit_id = None
        self.edit_orig = None
        self._pid = None
        self.v_prod, self.v_cant, self.v_costo, self.v_nota = (tk.StringVar() for _ in range(4))
        self.v_total, self.v_info, self.v_aviso, self.v_buscar = (tk.StringVar() for _ in range(4))

        self.form = ttk.LabelFrame(self, padding=8)
        self.form.pack(fill="x")
        f = self.form
        for c, txt in enumerate(["Fecha", "Producto", "Cantidad", "Costo unit.", "Pérdida", "Motivo (vencido, roto…)"]):
            ttk.Label(f, text=txt).grid(row=0, column=c, sticky="w", padx=3)
        self.c_fecha = CampoFecha(f)
        self.cb_prod = ProductoCombo(f, textvariable=self.v_prod, width=34)
        self.e_cant = ttk.Entry(f, textvariable=self.v_cant, width=10)
        self.e_costo = ttk.Entry(f, textvariable=self.v_costo, width=12)
        l_total = ttk.Label(f, textvariable=self.v_total, width=14, style="Total.TLabel")
        self.e_nota = ttk.Entry(f, textvariable=self.v_nota, width=28)
        for c, w in enumerate([self.c_fecha, self.cb_prod, self.e_cant, self.e_costo, l_total, self.e_nota]):
            w.grid(row=1, column=c, sticky="ew", padx=3)
        f.columnconfigure(5, weight=1)
        barra = ttk.Frame(f)
        barra.grid(row=2, column=0, columnspan=6, sticky="ew", pady=(8, 0))
        self.b_guardar = ttk.Button(barra, style="Accent.TButton", command=self.guardar)
        self.b_guardar.pack(side="left")
        self.b_cancelar = ttk.Button(barra, text="Cancelar corrección", command=self.limpiar)
        ttk.Label(barra, textvariable=self.v_info, foreground="#555555").pack(side="left", padx=10)
        ttk.Label(barra, textvariable=self.v_aviso, foreground="#b00020").pack(side="left")

        self.cb_prod.bind("<<ComboboxSelected>>", lambda e: self._producto_elegido())
        self.cb_prod.bind("<FocusOut>", lambda e: self._producto_elegido())
        _bind_enter(self.cb_prod, self._enter_producto)
        self.c_fecha.bind_enter(self.guardar)
        for w in (self.e_cant, self.e_costo, self.e_nota):
            _bind_enter(w, self.guardar)
        self.v_cant.trace_add("write", lambda *a: self._recalcular())
        self.v_costo.trace_add("write", lambda *a: self._recalcular())

        filtros = ttk.Frame(self)
        filtros.pack(fill="x", pady=(10, 4))
        self.periodo = PeriodoSelector(filtros, al_cambiar=self.cargar_lista)
        self.periodo.pack(side="left")
        ttk.Label(filtros, text="Buscar:").pack(side="left", padx=(14, 3))
        eb = ttk.Entry(filtros, textvariable=self.v_buscar, width=18)
        eb.pack(side="left")
        _bind_enter(eb, lambda e: self.cargar_lista())
        ttk.Button(filtros, text="Buscar", command=self.cargar_lista).pack(side="left", padx=3)

        cols = [("fecha", "Fecha", 90, "center"), ("producto", "Producto", 240, "w"), ("cantidad", "Cantidad", 80, "e"),
                ("unidad", "Unid.", 60, "center"), ("costo", "Costo unit.", 100, "e"), ("total", "Pérdida", 110, "e"),
                ("nota", "Motivo", 240, "w")]
        marco, self.tabla = crear_tabla(self, cols)
        marco.pack(fill="both", expand=True)
        self.tabla.bind("<Double-1>", lambda e: self.editar())
        self.tabla.bind("<Delete>", lambda e: self.eliminar())
        pie = ttk.Frame(self)
        pie.pack(fill="x", pady=(4, 0), side="bottom", before=marco)
        ttk.Button(pie, text="Corregir", command=self.editar).pack(side="left")
        ttk.Button(pie, text="Eliminar", command=self.eliminar).pack(side="left", padx=4)
        self.l_totales = ttk.Label(pie, style="Total.TLabel")
        self.l_totales.pack(side="right")
        self.limpiar(foco=False)

    def refrescar(self):
        self.cb_prod.set_items(self.db.nombres_activos())
        self.cargar_lista()

    def hay_cambios(self):
        return bool(self.v_prod.get().strip() and self.v_cant.get().strip())

    def cargar_lista(self):
        try:
            d, h = self.periodo.rango()
        except ValueError as ex:
            messagebox.showerror("Período", str(ex), parent=self)
            return
        filas = self.db.movimientos("merma", d, h, self.v_buscar.get())
        limpiar_tabla(self.tabla)
        total = 0.0
        for i, m in enumerate(filas):
            t = m["cantidad"] * m["costo_unit"]
            total += t
            if i < MAX_FILAS:
                self.tabla.insert("", "end", iid=str(m["id"]), tags=("par",) if i % 2 else (), values=(
                    fmt_iso(m["fecha"]), m["nombre"], qty(m["cantidad"]), m["unidad"], money(m["costo_unit"]),
                    money(t), m["nota"]))
        self.l_totales.config(text=f"{len(filas)} registros   ·   Pérdida total: {money(total)}")

    def _producto_elegido(self):
        nombre = self.v_prod.get().strip()
        p = self.db.producto_por_nombre(nombre) if nombre else None
        if not p:
            self._pid = None
            self.v_info.set("Producto nuevo: se creará al guardar" if nombre else "")
            self._recalcular()
            return
        if p["id"] != self._pid:
            self._pid = p["id"]
            self.v_prod.set(p["nombre"])
            self.v_costo.set(num_txt(p["precio_compra"]) if p["precio_compra"] else "")
        self.v_info.set(f"Stock: {qty(p['stock'])} {p['unidad']}   ·   Costo: {money(p['precio_compra'])}")
        self._recalcular()

    def _enter_producto(self, e=None):
        texto = self.v_prod.get().strip()
        exacto = self.db.producto_por_nombre(texto) if texto else None
        opciones = self.cb_prod.coincidencias()
        if texto and not exacto and len(opciones) == 1:
            self.v_prod.set(opciones[0])
        elif texto and not exacto and len(opciones) > 1:
            self.cb_prod["values"] = opciones
            self.cb_prod.event_generate("<Down>")
            return "break"
        self._producto_elegido()
        self.e_cant.focus_set()
        return "break"

    def _recalcular(self):
        cant = parse_num(self.v_cant.get()) or 0
        costo = parse_num(self.v_costo.get()) or 0
        self.v_total.set(money(cant * costo))
        aviso = ""
        if self._pid and cant > 0:
            disponible = self.db.producto(self._pid)["stock"]
            if self.edit_orig and self.edit_orig["producto_id"] == self._pid:
                disponible += self.edit_orig["cantidad"]
            if cant > disponible + 1e-9:
                aviso = f"⚠ Stock insuficiente (hay {qty(disponible)})"
        self.v_aviso.set(aviso)

    def guardar(self, e=None):
        fecha = self.c_fecha.fecha()
        if not fecha:
            return _aviso(self, "Fecha inválida. Use el formato dd/mm/aaaa.", self.c_fecha.entry)
        nombre = self.v_prod.get().strip()
        if not nombre:
            return _aviso(self, "Elija o escriba un producto.", self.cb_prod)
        cant = parse_num(self.v_cant.get())
        if not cant or cant <= 0:
            return _aviso(self, "Ingrese una cantidad mayor que cero.", self.e_cant)
        costo = parse_num(self.v_costo.get()) if self.v_costo.get().strip() else 0.0
        if costo is None or costo < 0:
            return _aviso(self, "Revise el costo unitario.", self.e_costo)
        p = self.db.producto_por_nombre(nombre)
        if not p:
            if not messagebox.askyesno("Producto nuevo", f"El producto «{nombre}» no existe.\n¿Desea crearlo?",
                                       parent=self):
                return "break"
            p = self.db.producto(self.db.guardar_producto(None, nombre, "unid", "", costo, 0, 0))
            self.cb_prod.set_items(self.db.nombres_activos())
        self.db.guardar_mov(self.edit_id, "merma", fecha, p["id"], cant, costo, costo, self.v_nota.get().strip())
        self.app.estado(f"Merma guardada: {p['nombre']} × {qty(cant)} = {money(cant * costo)}")
        self.limpiar()
        self.cargar_lista()
        self.app.actualizar_alertas()
        return "break"

    def limpiar(self, foco=True):
        self.edit_id = self.edit_orig = self._pid = None
        for v in (self.v_prod, self.v_cant, self.v_costo, self.v_nota, self.v_info, self.v_aviso):
            v.set("")
        self.cb_prod.restaurar()
        self.form.config(text="Registrar merma (mercadería perdida)")
        self.b_guardar.config(text="Guardar  (Enter)")
        self.b_cancelar.pack_forget()
        if foco:
            self.cb_prod.focus_set()

    def editar(self):
        sel = self.tabla.selection()
        if not sel:
            return
        m = self.db.mov(int(sel[0]))
        if not m:
            return
        self.limpiar()
        self.edit_id, self.edit_orig, self._pid = m["id"], m, m["producto_id"]
        self.c_fecha.set(dt.date.fromisoformat(m["fecha"]))
        self.v_prod.set(m["nombre"])
        self.v_cant.set(num_txt(m["cantidad"]))
        self.v_costo.set(num_txt(m["costo_unit"]))
        self.v_nota.set(m["nota"])
        self._producto_elegido()
        self.form.config(text=f"Corrigiendo merma del {fmt_iso(m['fecha'])}")
        self.b_guardar.config(text="Guardar cambios")
        self.b_cancelar.pack(side="left", padx=4, after=self.b_guardar)
        self.e_cant.focus_set()

    def eliminar(self):
        sel = self.tabla.selection()
        if not sel:
            return
        if messagebox.askyesno("Eliminar", f"¿Eliminar {len(sel)} registro(s)? No se puede deshacer.", parent=self):
            self.db.eliminar_movs([int(i) for i in sel])
            if self.edit_id and str(self.edit_id) in sel:
                self.limpiar()
            self.cargar_lista()
            self.app.actualizar_alertas()


# =====================================================================================
# Gastos del negocio
# =====================================================================================
class GastosTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=(8, 6))
        self.app, self.db = app, app.db
        self.edit_id = None
        self.v_cat, self.v_desc, self.v_monto, self.v_buscar = (tk.StringVar() for _ in range(4))
        self.v_medio = tk.StringVar(value=EFECTIVO)
        self.v_factura = tk.BooleanVar(value=True)
        self.v_desglose = tk.StringVar()

        self.form = ttk.LabelFrame(self, padding=8)
        self.form.pack(fill="x")
        f = self.form
        for c, txt in enumerate(["Fecha", "Categoría", "Descripción", "Monto (con IVA)", "Medio de pago"]):
            ttk.Label(f, text=txt).grid(row=0, column=c, sticky="w", padx=3)
        self.c_fecha = CampoFecha(f)
        self.cb_cat = ttk.Combobox(f, textvariable=self.v_cat, width=20)
        self.e_desc = ttk.Entry(f, textvariable=self.v_desc, width=36)
        self.e_monto = ttk.Entry(f, textvariable=self.v_monto, width=13)
        cm = ttk.Combobox(f, textvariable=self.v_medio, values=MEDIOS_PAGO, state="readonly", width=18)
        for c, w in enumerate([self.c_fecha, self.cb_cat, self.e_desc, self.e_monto, cm]):
            w.grid(row=1, column=c, sticky="ew", padx=3)
        f.columnconfigure(2, weight=1)
        barra = ttk.Frame(f)
        barra.grid(row=2, column=0, columnspan=5, sticky="ew", pady=(8, 0))
        self.b_guardar = ttk.Button(barra, style="Accent.TButton", command=self.guardar)
        self.b_guardar.pack(side="left")
        self.b_cancelar = ttk.Button(barra, text="Cancelar corrección", command=self.limpiar)
        self.chk_factura = ttk.Checkbutton(barra, text="Con factura (genera IVA crédito)", variable=self.v_factura,
                                           command=self._desglose)
        self.chk_factura.pack(side="left", padx=(12, 4))
        ttk.Label(barra, textvariable=self.v_desglose, foreground="#555555").pack(side="left", padx=4)
        ttk.Label(barra, foreground="#777777", text="(alquiler, luz, sueldos, impuestos… no la mercadería)"
                  ).pack(side="right")
        self.cb_cat.bind("<<ComboboxSelected>>", lambda e: self._categoria_elegida())
        self.cb_cat.bind("<FocusOut>", lambda e: self._categoria_elegida())
        self.v_monto.trace_add("write", lambda *a: self._desglose())
        self.c_fecha.bind_enter(self.guardar)
        for w in (self.cb_cat, self.e_desc, self.e_monto):
            _bind_enter(w, self.guardar)

        filtros = ttk.Frame(self)
        filtros.pack(fill="x", pady=(10, 4))
        self.periodo = PeriodoSelector(filtros, al_cambiar=self.cargar_lista)
        self.periodo.pack(side="left")
        ttk.Label(filtros, text="Buscar:").pack(side="left", padx=(14, 3))
        eb = ttk.Entry(filtros, textvariable=self.v_buscar, width=18)
        eb.pack(side="left")
        _bind_enter(eb, lambda e: self.cargar_lista())
        ttk.Button(filtros, text="Buscar", command=self.cargar_lista).pack(side="left", padx=3)

        cuerpo = ttk.Frame(self)
        cuerpo.pack(fill="both", expand=True)
        cols = [("fecha", "Fecha", 88, "center"), ("cat", "Categoría", 140, "w"), ("desc", "Descripción", 240, "w"),
                ("medio", "Medio de pago", 120, "center"), ("monto", "Monto", 100, "e"), ("neto", "Neto", 95, "e"),
                ("iva", "IVA crédito", 90, "e")]
        marco, self.tabla = crear_tabla(cuerpo, cols)
        marco.pack(side="left", fill="both", expand=True)
        self.tabla.bind("<Double-1>", lambda e: self.editar())
        self.tabla.bind("<Delete>", lambda e: self.eliminar())
        marco_c, self.t_cat = crear_tabla(cuerpo, [("cat", "Por categoría", 150, "w"), ("monto", "Total", 110, "e")],
                                          seleccion="none")
        marco_c.pack(side="left", fill="y", padx=(8, 0))

        pie = ttk.Frame(self)
        pie.pack(fill="x", pady=(4, 0), side="bottom", before=cuerpo)
        ttk.Button(pie, text="Corregir", command=self.editar).pack(side="left")
        ttk.Button(pie, text="Eliminar", command=self.eliminar).pack(side="left", padx=4)
        self.l_totales = ttk.Label(pie, style="Total.TLabel")
        self.l_totales.pack(side="right")
        self.limpiar(foco=False)

    def refrescar(self):
        self.cb_cat["values"] = self.db.categorias_gasto()
        self.cargar_lista()

    def hay_cambios(self):
        return bool(self.v_monto.get().strip())

    def cargar_lista(self):
        try:
            d, h = self.periodo.rango()
        except ValueError as ex:
            messagebox.showerror("Período", str(ex), parent=self)
            return
        filas = self.db.gastos(d, h, self.v_buscar.get())
        limpiar_tabla(self.tabla)
        for i, g in enumerate(filas[:MAX_FILAS]):
            n = neto(g["monto"], g["iva_tasa"])
            self.tabla.insert("", "end", iid=str(g["id"]), tags=("par",) if i % 2 else (), values=(
                fmt_iso(g["fecha"]), g["categoria"], g["descripcion"], g["medio_pago"], money(g["monto"]),
                money(n), money(g["monto"] - n) if g["iva_tasa"] else "—"))
        limpiar_tabla(self.t_cat)
        for c in self.db.gastos_por_categoria(d, h):
            self.t_cat.insert("", "end", values=(c["categoria"], money(c["total"])))
        bruto = sum(g["monto"] for g in filas)
        iva = sum(g["monto"] - neto(g["monto"], g["iva_tasa"]) for g in filas)
        self.l_totales.config(text=f"{len(filas)} gastos   ·   Total: {money(bruto)}   ·   IVA crédito: {money(iva)}")

    def _categoria_elegida(self):
        """Propone «con factura» según el último gasto de la categoría (o según el tipo de gasto)."""
        if self.edit_id:
            return
        cat = self.v_cat.get().strip()
        if not cat:
            return
        previo = self.db.ultimo_iva_gasto(cat)
        self.v_factura.set(previo if previo is not None else cat not in GASTOS_SIN_IVA)
        self._desglose()

    def _desglose(self):
        monto = parse_num(self.v_monto.get())
        if not monto or monto <= 0:
            self.v_desglose.set("")
        elif self.v_factura.get():
            n = neto(monto, self.db.tasa_iva())
            self.v_desglose.set(f"Neto {money(n)}  ·  IVA {money(monto - n)}")
        else:
            self.v_desglose.set("Sin IVA crédito")

    def guardar(self, e=None):
        fecha = self.c_fecha.fecha()
        if not fecha:
            return _aviso(self, "Fecha inválida. Use el formato dd/mm/aaaa.", self.c_fecha.entry)
        cat = self.v_cat.get().strip()
        if not cat:
            return _aviso(self, "Elija o escriba una categoría (ej. Luz, Alquiler).", self.cb_cat)
        monto = parse_num(self.v_monto.get())
        if not monto or monto <= 0:
            return _aviso(self, "Ingrese un monto mayor que cero.", self.e_monto)
        self.db.guardar_gasto(self.edit_id, fecha, cat, self.v_desc.get().strip(), monto, self.v_medio.get(),
                              self.db.tasa_iva() if self.v_factura.get() else 0.0)
        self.app.estado(f"Gasto guardado: {cat} {money(monto)}")
        self.limpiar()
        self.refrescar()
        return "break"

    def limpiar(self, foco=True):
        self.edit_id = None
        for v in (self.v_cat, self.v_desc, self.v_monto, self.v_desglose):
            v.set("")
        self.v_medio.set(EFECTIVO)
        self.v_factura.set(True)
        self.form.config(text="Registrar gasto")
        self.b_guardar.config(text="Guardar  (Enter)")
        self.b_cancelar.pack_forget()
        if foco:
            self.cb_cat.focus_set()

    def editar(self):
        sel = self.tabla.selection()
        if not sel:
            return
        g = self.db.gasto(int(sel[0]))
        if not g:
            return
        self.limpiar()
        self.edit_id = g["id"]
        self.c_fecha.set(dt.date.fromisoformat(g["fecha"]))
        self.v_cat.set(g["categoria"])
        self.v_desc.set(g["descripcion"])
        self.v_monto.set(num_txt(g["monto"]))
        self.v_medio.set(g["medio_pago"])
        self.v_factura.set(g["iva_tasa"] > 0)
        self._desglose()
        self.form.config(text=f"Corrigiendo gasto del {fmt_iso(g['fecha'])}")
        self.b_guardar.config(text="Guardar cambios")
        self.b_cancelar.pack(side="left", padx=4, after=self.b_guardar)
        self.e_monto.focus_set()

    def eliminar(self):
        sel = self.tabla.selection()
        if not sel:
            return
        if messagebox.askyesno("Eliminar", f"¿Eliminar {len(sel)} gasto(s)? No se puede deshacer.", parent=self):
            self.db.eliminar_gastos([int(i) for i in sel])
            if self.edit_id and str(self.edit_id) in sel:
                self.limpiar()
            self.cargar_lista()


# =====================================================================================
# Caja diaria
# =====================================================================================
class CajaTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=(8, 6))
        self.app, self.db = app, app.db
        self.ef = dict(venta=0.0, compra=0.0, gasto=0.0)
        self.v_inicial, self.v_retiros, self.v_contado, self.v_nota = (tk.StringVar() for _ in range(4))
        self.v_ventas_ef, self.v_compras_ef, self.v_gastos_ef = (tk.StringVar() for _ in range(3))
        self.v_esperado, self.v_dif, self.v_estado = (tk.StringVar() for _ in range(3))

        top = ttk.Frame(self)
        top.pack(fill="x")
        ttk.Label(top, text="Día:", style="Sub.TLabel").pack(side="left")
        self.c_fecha = CampoFecha(top, al_cambiar=self.cargar)
        self.c_fecha.pack(side="left", padx=4)
        self.c_fecha.bind_enter(lambda e: self.cargar())
        ttk.Button(top, text="◀", width=3, command=lambda: self.c_fecha.mover(-1)).pack(side="left")
        ttk.Button(top, text="Hoy", width=6, command=lambda: self.c_fecha.set(hoy(), avisar=True)).pack(side="left", padx=2)
        ttk.Button(top, text="▶", width=3, command=lambda: self.c_fecha.mover(1)).pack(side="left")

        medio = ttk.Frame(self)
        medio.pack(fill="x", pady=8)
        izq = ttk.LabelFrame(medio, text="Movimientos del día por medio de pago", padding=8)
        izq.pack(side="left", fill="both", expand=True)
        cols = [("medio", "Medio de pago", 160, "w"), ("ventas", "Ventas", 120, "e"),
                ("compras", "Compras", 120, "e"), ("gastos", "Gastos", 120, "e")]
        marco, self.t_medios = crear_tabla(izq, cols, seleccion="none", alto=7)
        marco.pack(fill="both", expand=True)
        self.l_resumen = ttk.Label(izq, style="Total.TLabel")
        self.l_resumen.pack(anchor="w", pady=(6, 0))

        der = ttk.LabelFrame(medio, text="Cierre de caja (efectivo)", padding=10)
        der.pack(side="left", fill="y", padx=(8, 0))
        filas = [
            ("Efectivo al abrir la caja", "entry", self.v_inicial),
            ("+ Ventas en efectivo", "label", self.v_ventas_ef),
            ("− Compras pagadas en efectivo", "label", self.v_compras_ef),
            ("− Gastos pagados en efectivo", "label", self.v_gastos_ef),
            ("− Retiros (dinero que se sacó)", "entry", self.v_retiros),
            ("= Debería haber en la caja", "total", self.v_esperado),
            ("Efectivo contado", "entry", self.v_contado),
            ("Diferencia", "dif", self.v_dif),
            ("Nota", "nota", self.v_nota),
        ]
        self.entradas = []
        for r, (txt, clase, var) in enumerate(filas):
            ttk.Label(der, text=txt, style="Sub.TLabel" if clase in ("total", "dif") else "TLabel").grid(
                row=r, column=0, sticky="w", pady=3, padx=(0, 10))
            if clase in ("entry", "nota"):
                w = ttk.Entry(der, textvariable=var, width=14 if clase == "entry" else 24,
                              justify="right" if clase == "entry" else "left")
                _bind_enter(w, lambda e: self.guardar())
                self.entradas.append(w)
            elif clase == "dif":
                w = tk.Label(der, textvariable=var, font=self.app.f_media, bg=self.app.fondo)
                self.l_dif = w
            else:
                w = ttk.Label(der, textvariable=var, style="Total.TLabel" if clase == "total" else "TLabel")
            w.grid(row=r, column=1, sticky="e", pady=3)
            if r == 5:
                ttk.Separator(der).grid(row=r, column=0, columnspan=2, sticky="new")
        for v in (self.v_inicial, self.v_retiros, self.v_contado):
            v.trace_add("write", lambda *a: self._recalcular())
        bot = ttk.Frame(der)
        bot.grid(row=len(filas), column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.b_guardar = ttk.Button(bot, text="Guardar cierre", style="Accent.TButton", command=self.guardar)
        self.b_guardar.pack(side="left")
        ttk.Button(bot, text="Borrar cierre", command=self.borrar).pack(side="left", padx=4)
        ttk.Label(der, textvariable=self.v_estado, foreground="#555555", wraplength=330).grid(
            row=len(filas) + 1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        ttk.Label(self, text="Cierres de los últimos 90 días (doble clic para abrir)", style="Sub.TLabel").pack(anchor="w")
        cols = [("fecha", "Fecha", 90, "center"), ("inicial", "Al abrir", 110, "e"), ("ventas", "Ventas efectivo", 120, "e"),
                ("esperado", "Debería haber", 120, "e"), ("contado", "Contado", 110, "e"),
                ("dif", "Diferencia", 110, "e"), ("nota", "Nota", 200, "w")]
        marco, self.t_cierres = crear_tabla(self, cols, seleccion="browse")
        marco.pack(fill="both", expand=True)
        self.t_cierres.bind("<Double-1>", self._abrir_cierre)

    def refrescar(self):
        self.cargar()

    def hay_cambios(self):
        return False

    def cargar(self):
        fecha = self.c_fecha.fecha()
        if not fecha:
            messagebox.showwarning("Fecha", "Fecha inválida. Use el formato dd/mm/aaaa.", parent=self)
            return
        res = self.db.resumen_dia(fecha)
        limpiar_tabla(self.t_medios)
        tv = tc = tg = 0.0
        for i, medio in enumerate(sorted(res, key=lambda m: (m != EFECTIVO, m))):
            r = res[medio]
            tv, tc, tg = tv + r["venta"], tc + r["compra"], tg + r["gasto"]
            self.t_medios.insert("", "end", tags=("par",) if i % 2 else (), values=(
                medio, money(r["venta"]), money(r["compra"]), money(r["gasto"])))
        self.t_medios.insert("", "end", tags=("total",), values=("TOTAL", money(tv), money(tc), money(tg)))
        n = self.db.cantidad_tickets(fecha, fecha)
        self.l_resumen.config(text=f"{n} ventas en el día   ·   Vendido: {money(tv)}")

        ef = res.get(EFECTIVO, {})
        self.ef = dict(venta=ef.get("venta", 0.0), compra=ef.get("compra", 0.0), gasto=ef.get("gasto", 0.0))
        self.v_ventas_ef.set(money(self.ef["venta"]))
        self.v_compras_ef.set(money(self.ef["compra"]))
        self.v_gastos_ef.set(money(self.ef["gasto"]))

        c = self.db.cierre(fecha)
        if c:
            self.v_inicial.set(num_txt(c["inicial"]))
            self.v_retiros.set(num_txt(c["retiros"]) if c["retiros"] else "")
            self.v_contado.set(num_txt(c["contado"]))
            self.v_nota.set(c["nota"])
            estado = f"Cierre guardado el {fmt_iso(c['guardado'][:10])} a las {c['guardado'][11:16]}."
            if (abs(c["ventas_ef"] - self.ef["venta"]) > 0.005 or abs(c["compras_ef"] - self.ef["compra"]) > 0.005
                    or abs(c["gastos_ef"] - self.ef["gasto"]) > 0.005):
                estado += ("\n⚠ Se registraron o corrigieron movimientos en efectivo después de cerrar. "
                           "Revise y vuelva a guardar el cierre.")
            self.v_estado.set(estado)
            self.b_guardar.config(text="Actualizar cierre")
        else:
            ant = self.db.cierre_anterior(fecha)
            self.v_inicial.set(num_txt(ant["contado"]) if ant else "")
            self.v_retiros.set("")
            self.v_contado.set("")
            self.v_nota.set("")
            self.v_estado.set(f"Sin cierre para este día. El efectivo inicial se tomó del cierre del "
                              f"{fmt_iso(ant['fecha'])}." if ant else "Sin cierre para este día.")
            self.b_guardar.config(text="Guardar cierre")
        self._recalcular()
        self._cargar_cierres()

    def _cargar_cierres(self):
        limpiar_tabla(self.t_cierres)
        for i, c in enumerate(self.db.cierres(hoy() - dt.timedelta(days=90), dt.date(9999, 12, 31))):
            tags = ["par"] if i % 2 else []
            if abs(c["diferencia"]) > 0.005:
                tags.append("alerta")
            self.t_cierres.insert("", "end", iid=c["fecha"], tags=tags, values=(
                fmt_iso(c["fecha"]), money(c["inicial"]), money(c["ventas_ef"]), money(c["esperado"]),
                money(c["contado"]), money(c["diferencia"]), c["nota"]))

    def _abrir_cierre(self, e=None):
        sel = self.t_cierres.selection()
        if sel:
            self.c_fecha.set(dt.date.fromisoformat(sel[0]), avisar=True)

    def _esperado(self):
        inicial = parse_num(self.v_inicial.get()) or 0
        retiros = parse_num(self.v_retiros.get()) or 0
        return round(inicial + self.ef["venta"] - self.ef["compra"] - self.ef["gasto"] - retiros, 2)

    def _recalcular(self):
        esperado = self._esperado()
        self.v_esperado.set(money(esperado))
        contado = parse_num(self.v_contado.get())
        if contado is None:
            self.v_dif.set("—")
            self.l_dif.config(fg="#555555")
            return
        dif = round(contado - esperado, 2)
        if abs(dif) < 0.5:
            self.v_dif.set("Cuadra ✔")
            self.l_dif.config(fg="#1e6b34")
        elif dif < 0:
            self.v_dif.set(f"Faltan {money(-dif)}")
            self.l_dif.config(fg="#b00020")
        else:
            self.v_dif.set(f"Sobran {money(dif)}")
            self.l_dif.config(fg="#1f4e9c")

    def guardar(self):
        fecha = self.c_fecha.fecha()
        if not fecha:
            return _aviso(self, "Fecha inválida.", self.c_fecha.entry)
        valores = {}
        for nombre, var, obligatorio in (("efectivo inicial", self.v_inicial, False),
                                         ("retiros", self.v_retiros, False),
                                         ("efectivo contado", self.v_contado, True)):
            txt = var.get().strip()
            if not txt and not obligatorio:
                valores[nombre] = 0.0
                continue
            n = parse_num(txt)
            if n is None or n < 0:
                return _aviso(self, f"Ingrese el {nombre} (un número mayor o igual a cero).")
            valores[nombre] = n
        self.db.guardar_cierre(fecha, valores["efectivo inicial"], self.ef["venta"], self.ef["compra"],
                               self.ef["gasto"], valores["retiros"], self._esperado(), valores["efectivo contado"],
                               self.v_nota.get().strip())
        self.app.estado(f"Cierre de caja del {fmt_fecha(fecha)} guardado.")
        self.cargar()
        return "break"

    def borrar(self):
        fecha = self.c_fecha.fecha()
        if fecha and self.db.cierre(fecha) and messagebox.askyesno(
                "Borrar cierre", f"¿Borrar el cierre de caja del {fmt_fecha(fecha)}?", parent=self):
            self.db.eliminar_cierre(fecha)
            self.cargar()


# =====================================================================================
# Productos
# =====================================================================================
class ProductosTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=(8, 6))
        self.app, self.db = app, app.db
        self.pid = None
        self.v = {k: tk.StringVar() for k in ("nombre", "unidad", "categoria", "iva", "pc", "pv", "si", "min")}
        self.v_buscar = tk.StringVar()
        self.v_inactivos = tk.BooleanVar(value=False)
        self.v_bajo = tk.BooleanVar(value=False)

        self.form = ttk.LabelFrame(self, padding=8)
        self.form.pack(fill="x")
        f = self.form
        campos = [("Nombre", "nombre", 28), ("Unidad", "unidad", 7), ("Categoría", "categoria", 13),
                  ("IVA", "iva", 11), ("Costo con IVA", "pc", 11), ("Precio venta", "pv", 11),
                  ("Stock inicial", "si", 8), ("Stock mínimo", "min", 8)]
        self.entradas = {}
        for c, (txt, k, w) in enumerate(campos):
            ttk.Label(f, text=txt).grid(row=0, column=c, sticky="w", padx=3)
            if k == "unidad":
                e = ttk.Combobox(f, textvariable=self.v[k], values=UNIDADES, width=w)
            elif k == "categoria":
                e = ttk.Combobox(f, textvariable=self.v[k], width=w)
            elif k == "iva":
                e = ttk.Combobox(f, textvariable=self.v[k], values=[], state="readonly", width=w)
            else:
                e = ttk.Entry(f, textvariable=self.v[k], width=w)
            e.grid(row=1, column=c, sticky="ew", padx=3)
            _bind_enter(e, self.guardar)
            self.entradas[k] = e
        f.columnconfigure(0, weight=1)
        barra = ttk.Frame(f)
        barra.grid(row=2, column=0, columnspan=8, sticky="ew", pady=(8, 0))
        self.b_guardar = ttk.Button(barra, style="Accent.TButton", command=self.guardar)
        self.b_guardar.pack(side="left")
        ttk.Button(barra, text="Nuevo / Limpiar", command=self.limpiar).pack(side="left", padx=4)
        self.l_margen = ttk.Label(barra, foreground="#555555")
        self.l_margen.pack(side="left", padx=10)
        for k in ("pc", "pv"):
            self.v[k].trace_add("write", lambda *a: self._margen())

        filtros = ttk.Frame(self)
        filtros.pack(fill="x", pady=(10, 4))
        ttk.Label(filtros, text="Buscar:").pack(side="left")
        eb = ttk.Entry(filtros, textvariable=self.v_buscar, width=24)
        eb.pack(side="left", padx=3)
        eb.bind("<KeyRelease>", lambda e: self.cargar_lista())
        ttk.Checkbutton(filtros, text="Solo los que hay que reponer", variable=self.v_bajo,
                        command=self.cargar_lista).pack(side="left", padx=10)
        ttk.Checkbutton(filtros, text="Mostrar desactivados", variable=self.v_inactivos,
                        command=self.cargar_lista).pack(side="left")

        cols = [("nombre", "Producto", 220, "w"), ("categoria", "Categoría", 105, "w"), ("unidad", "Unid.", 55, "center"),
                ("iva", "IVA", 65, "center"), ("pc", "Costo", 85, "e"), ("pv", "Precio venta", 95, "e"),
                ("margen", "Margen", 65, "e"),
                ("stock", "Stock", 75, "e"), ("min", "Mínimo", 65, "e"), ("valor", "Valor stock", 105, "e"),
                ("estado", "Estado", 95, "center")]
        marco, self.tabla = crear_tabla(self, cols, seleccion="browse")
        marco.pack(fill="both", expand=True)
        self.tabla.bind("<<TreeviewSelect>>", self._seleccionar)
        self.tabla.bind("<Double-1>", lambda e: self.historial())
        self.tabla.bind("<Delete>", lambda e: self.eliminar())

        pie = ttk.Frame(self)
        pie.pack(fill="x", pady=(4, 0), side="bottom", before=marco)
        ttk.Button(pie, text="Ajustar stock (conteo)…", command=self.ajustar).pack(side="left")
        ttk.Button(pie, text="Ver historial", command=self.historial).pack(side="left", padx=4)
        self.b_activo = ttk.Button(pie, text="Desactivar", command=self.alternar_activo)
        self.b_activo.pack(side="left")
        ttk.Button(pie, text="Eliminar", command=self.eliminar).pack(side="left", padx=4)
        self.l_tot = ttk.Label(pie, style="Total.TLabel")
        self.l_tot.pack(side="right")
        self.limpiar(foco=False)

    def _opciones_iva(self):
        return [f"Afecto {self.db.tasa_iva():g}%", "Exento"]

    def refrescar(self):
        self.entradas["categoria"]["values"] = self.db.categorias_producto()
        self.entradas["iva"]["values"] = self._opciones_iva()
        if not self.pid:
            self.v["iva"].set(self._opciones_iva()[0])
        self.cargar_lista()

    def hay_cambios(self):
        return False

    def mostrar_stock_bajo(self):
        self.v_bajo.set(True)
        self.v_inactivos.set(False)
        self.v_buscar.set("")
        self.cargar_lista()

    def cargar_lista(self):
        filas = self.db.productos(self.v_buscar.get(), self.v_inactivos.get(), self.v_bajo.get())
        limpiar_tabla(self.tabla)
        valor_total = 0.0
        for i, p in enumerate(filas):
            valor = max(p["stock"], 0) * p["precio_compra"]
            tags = ["par"] if i % 2 else []
            if not p["activo"]:
                estado = "Desactivado"
                tags.append("inactivo")
            else:
                valor_total += valor
                if p["stock"] <= 0:
                    estado = "Sin stock"
                    tags.append("alerta")
                elif p["stock_minimo"] > 0 and p["stock"] <= p["stock_minimo"]:
                    estado = "Reponer"
                    tags.append("bajo")
                else:
                    estado = "OK"
            margen = (p["precio_venta"] - p["precio_compra"]) / p["precio_venta"] if p["precio_venta"] else None
            self.tabla.insert("", "end", iid=str(p["id"]), tags=tags, values=(
                p["nombre"], p["categoria"], p["unidad"], "Afecto" if p["afecto_iva"] else "Exento",
                money(p["precio_compra"]), money(p["precio_venta"]),
                pct(margen) if margen is not None else "—", qty(p["stock"]),
                qty(p["stock_minimo"]) if p["stock_minimo"] else "", money(valor), estado))
        self.l_tot.config(text=f"{len(filas)} productos   ·   Valor del inventario (al costo): {money(valor_total)}")
        if self.pid and self.tabla.exists(str(self.pid)):
            self.tabla.selection_set(str(self.pid))

    def _margen(self):
        pc, pv = parse_num(self.v["pc"].get()) or 0, parse_num(self.v["pv"].get()) or 0
        self.l_margen.config(text=f"Margen: {pct((pv - pc) / pv)}  (gana {money(pv - pc)} por unidad)" if pv else "")

    def _seleccionar(self, e=None):
        sel = self.tabla.selection()
        if not sel:
            return
        p = self.db.producto(int(sel[0]))
        if not p:
            return
        self.pid = p["id"]
        for k, campo in (("nombre", "nombre"), ("unidad", "unidad"), ("categoria", "categoria")):
            self.v[k].set(p[campo])
        self.v["iva"].set(self._opciones_iva()[0 if p["afecto_iva"] else 1])
        self.v["pc"].set(num_txt(p["precio_compra"]))
        self.v["pv"].set(num_txt(p["precio_venta"]))
        self.v["si"].set(num_txt(p["stock_inicial"]))
        self.v["min"].set(num_txt(p["stock_minimo"]) if p["stock_minimo"] else "")
        self.form.config(text=f"Corrigiendo: {p['nombre']}")
        self.b_guardar.config(text="Guardar cambios")
        self.b_activo.config(text="Desactivar" if p["activo"] else "Reactivar")

    def limpiar(self, foco=True):
        self.pid = None
        for k, v in self.v.items():
            v.set("unid" if k == "unidad" else "")
        self.v["iva"].set(self._opciones_iva()[0])
        self.tabla.selection_remove(self.tabla.selection())
        self.form.config(text="Nuevo producto")
        self.b_guardar.config(text="Guardar  (Enter)")
        if foco:
            self.entradas["nombre"].focus_set()

    def guardar(self, e=None):
        nombre = self.v["nombre"].get().strip()
        if not nombre:
            return _aviso(self, "Escriba el nombre del producto.", self.entradas["nombre"])
        nums = {}
        for k, etiqueta in (("pc", "costo"), ("pv", "precio de venta"), ("si", "stock inicial"),
                            ("min", "stock mínimo")):
            txt = self.v[k].get().strip()
            n = parse_num(txt) if txt else 0.0
            if n is None or (k != "si" and n < 0):
                return _aviso(self, f"Revise el {etiqueta}.", self.entradas[k])
            nums[k] = n
        try:
            pid = self.db.guardar_producto(self.pid, nombre, self.v["unidad"].get().strip() or "unid",
                                           self.v["categoria"].get().strip(), nums["pc"], nums["pv"], nums["si"],
                                           nums["min"], afecto=self.v["iva"].get() != "Exento")
        except sqlite3.IntegrityError:
            return _aviso(self, f"Ya existe un producto llamado «{nombre}».", self.entradas["nombre"])
        self.app.estado(f"Producto guardado: {nombre}")
        self.pid = pid
        self.refrescar()
        self.limpiar()
        self.app.actualizar_alertas()
        return "break"

    def _elegido(self):
        if not self.pid:
            messagebox.showinfo("Productos", "Primero seleccione un producto de la lista.", parent=self)
            return None
        return self.db.producto(self.pid)

    def ajustar(self):
        p = self._elegido()
        if not p:
            return
        txt = simpledialog.askstring(
            "Ajustar stock",
            f"{p['nombre']}\n\nSegún el sistema hay: {qty(p['stock'])} {p['unidad']}\n\n"
            "¿Cuánto hay realmente? (cuente la mercadería)", parent=self)
        if txt is None:
            return
        contado = parse_num(txt)
        if contado is None or contado < 0:
            messagebox.showwarning("Ajustar stock", "Ingrese un número válido.", parent=self)
            return
        dif = self.db.ajustar_stock(self.pid, contado, hoy())
        if dif == 0:
            msg = "El stock coincide. No hubo que ajustar nada."
        elif dif < 0:
            msg = (f"Faltaban {qty(-dif)} {p['unidad']}. Se registró como merma "
                   f"({money(-dif * p['precio_compra'])}).")
        else:
            msg = f"Sobraban {qty(dif)} {p['unidad']}. Se sumaron al stock."
        messagebox.showinfo("Ajustar stock", msg, parent=self)
        self.cargar_lista()
        self.app.actualizar_alertas()

    def historial(self):
        p = self._elegido()
        if p:
            VentanaKardex(self, self.db, p["id"])

    def eliminar(self):
        p = self._elegido()
        if not p:
            return
        usos = self.db.usos_producto(self.pid)
        if usos:
            if messagebox.askyesno(
                    "No se puede eliminar",
                    f"«{p['nombre']}» tiene {usos} movimiento(s) registrados, por eso no se puede borrar.\n\n"
                    "¿Desea desactivarlo? Deja de aparecer en las listas pero se conserva el historial.", parent=self):
                self.db.set_activo(self.pid, False)
                self.cargar_lista()
            return
        if messagebox.askyesno("Eliminar", f"¿Eliminar el producto «{p['nombre']}»?", parent=self):
            self.db.eliminar_producto(self.pid)
            self.limpiar()
            self.cargar_lista()

    def alternar_activo(self):
        p = self._elegido()
        if not p:
            return
        self.db.set_activo(self.pid, not p["activo"])
        self.b_activo.config(text="Reactivar" if p["activo"] else "Desactivar")
        self.cargar_lista()
        self.app.actualizar_alertas()


class VentanaKardex(tk.Toplevel):
    """Historial de movimientos de un producto con el stock acumulado."""
    TIPOS = {"venta": "Venta", "compra": "Compra", "merma": "Merma", "ajuste": "Ajuste (+)"}

    def __init__(self, master, db, pid):
        super().__init__(master)
        p, filas = db.kardex(pid)
        self.title(f"Historial — {p['nombre']}")
        self.geometry("860x480")
        self.transient(master.winfo_toplevel())
        ttk.Label(self, padding=8, style="Sub.TLabel",
                  text=f"{p['nombre']}   ·   Stock inicial: {qty(p['stock_inicial'])}   ·   "
                       f"Stock actual: {qty(p['stock'])} {p['unidad']}").pack(anchor="w")
        cols = [("fecha", "Fecha", 90, "center"), ("tipo", "Movimiento", 100, "w"), ("ticket", "Ticket", 60, "center"),
                ("entra", "Entra", 80, "e"), ("sale", "Sale", 80, "e"), ("stock", "Stock", 80, "e"),
                ("precio", "Precio unit.", 100, "e"), ("detalle", "Detalle", 220, "w")]
        marco, tabla = crear_tabla(self, cols, seleccion="browse")
        marco.pack(fill="both", expand=True, padx=8)
        ultimo = None
        for i, (m, stock) in enumerate(filas):
            entra = m["tipo"] in ("compra", "ajuste")
            detalle = " · ".join(x for x in (m["medio_pago"] or "", m["nota"]) if x)
            ultimo = tabla.insert("", "end", tags=("par",) if i % 2 else (), values=(
                fmt_iso(m["fecha"]), self.TIPOS[m["tipo"]], m["ticket_id"] or "",
                qty(m["cantidad"]) if entra else "", "" if entra else qty(m["cantidad"]), qty(stock),
                money(m["precio_unit"]), detalle))
        if ultimo:
            tabla.see(ultimo)
        ttk.Button(self, text="Cerrar", command=self.destroy).pack(pady=8)
        self.bind("<Escape>", lambda e: self.destroy())


# =====================================================================================
# Reportes y Excel
# =====================================================================================
class ReportesTab(ttk.Frame):
    COLUMNAS = 6
    TARJETAS = [
        ("ventas", "Ventas (con IVA)", "#1e6b34"), ("ventas_netas", "Ventas netas (sin IVA)", "#1e6b34"),
        ("costo", "Costo de lo vendido (neto)", "#555555"), ("ganancia", "Ganancia bruta", "#1f4e9c"),
        ("margen", "Margen bruto", "#1f4e9c"), ("tickets", "Cantidad de ventas", "#555555"),
        ("mermas", "Mermas (neto)", "#b00020"), ("gastos", "Gastos (netos)", "#b00020"),
        ("resultado", "Resultado neto", "#1f4e9c"), ("iva_debito", "IVA débito (ventas)", "#8a5a00"),
        ("iva_credito", "IVA crédito (compras y gastos)", "#8a5a00"), ("iva_pagar", "IVA a pagar", "#8a5a00"),
    ]

    def __init__(self, master, app):
        super().__init__(master, padding=(8, 6))
        self.app, self.db = app, app.db

        top = ttk.Frame(self)
        top.pack(fill="x")
        self.periodo = PeriodoSelector(top, al_cambiar=self.refrescar)
        self.periodo.pack(side="left")
        ttk.Button(top, text="Exportar a Excel", style="Accent.TButton", command=self.exportar).pack(side="right")

        tarjetas = ttk.Frame(self)
        tarjetas.pack(fill="x", pady=(8, 2))
        self.valores = {}
        for i, (k, titulo, color) in enumerate(self.TARJETAS):
            card = tk.Frame(tarjetas, bg="white", highlightbackground="#d0d7e2", highlightthickness=1, padx=10, pady=4)
            card.grid(row=i // self.COLUMNAS, column=i % self.COLUMNAS, sticky="nsew", padx=3, pady=3)
            titulo_lbl = tk.Label(card, text=titulo, bg="white", fg="#555555", font=self.app.f_chica)
            titulo_lbl.pack(anchor="w")
            lbl = tk.Label(card, text="—", bg="white", fg=color, font=self.app.f_tarjeta)
            lbl.pack(anchor="w")
            self.valores[k] = lbl
            if k == "iva_pagar":
                self.titulo_iva = titulo_lbl
            tarjetas.columnconfigure(i % self.COLUMNAS, weight=1, uniform="tarjetas")

        self.l_dinero = ttk.Label(self, foreground="#333333")
        self.l_dinero.pack(anchor="w", pady=(2, 0))
        self.l_medios = ttk.Label(self, foreground="#333333")
        self.l_medios.pack(anchor="w")
        self.l_alerta = ttk.Label(self, foreground="#b00020")
        self.l_alerta.pack(anchor="w")

        ttk.Label(self, text="Detalle por producto en el período", style="Sub.TLabel").pack(anchor="w", pady=(4, 2))
        cols = [("producto", "Producto", 210, "w"), ("cant_v", "Vendido", 70, "e"),
                ("ventas", "Ventas c/IVA", 100, "e"), ("iva", "IVA", 85, "e"), ("ganancia", "Ganancia neta", 100, "e"),
                ("margen", "Margen", 65, "e"), ("cant_c", "Comprado", 75, "e"), ("compras", "Compras", 100, "e"),
                ("cant_m", "Merma", 60, "e"), ("mermas", "Mermas", 85, "e"), ("stock", "Stock actual", 85, "e")]
        marco, self.tabla = crear_tabla(self, cols, seleccion="browse")
        marco.pack(fill="both", expand=True)
        ttk.Label(self, foreground="#666666", text=(
            "Precios con IVA incluido. Ganancia y resultado se calculan sin IVA:  resultado neto = ventas netas − "
            "costo − mermas − gastos.   IVA a pagar = débito − crédito.")).pack(anchor="w", pady=(4, 0), side="bottom", before=marco)

    def hay_cambios(self):
        return False

    def refrescar(self):
        try:
            d, h = self.periodo.rango()
        except ValueError as ex:
            messagebox.showerror("Período", str(ex), parent=self)
            return
        prods = self.db.por_producto(d, h)
        t = totales(prods, self.db.gastos_totales(d, h), self.db.cantidad_tickets(d, h))
        for k, lbl in self.valores.items():
            if k == "margen":
                lbl.config(text=pct(t[k]))
            elif k == "tickets":
                lbl.config(text=f"{t['tickets']}  (prom. {money(t['ticket_prom'])})" if t["tickets"] else "0")
            else:
                lbl.config(text=money(t[k]))
        # IVA a pagar: si el período es un mes, se usa el cálculo mensual con remanente (como el F29).
        self.titulo_iva.config(text="IVA a pagar")
        if self.periodo.modo.get() == "Mes":
            mes = d.isoformat()[:7]
            f29 = next((r for r in self.db.iva_mensual() if r["mes"] == mes), None)
            if f29:
                if f29["pagar"] > 0 or not f29["remanente"]:
                    self.valores["iva_pagar"].config(text=money(f29["pagar"]))
                    if f29["remanente_anterior"]:
                        self.titulo_iva.config(text=f"IVA a pagar (descontado remanente {money(f29['remanente_anterior'])})")
                else:
                    self.titulo_iva.config(text="IVA a pagar: $ 0 — crédito a favor")
                    self.valores["iva_pagar"].config(text=money(f29["remanente"]))
        elif t["iva_pagar"] < 0:
            self.titulo_iva.config(text="IVA: crédito a favor")
            self.valores["iva_pagar"].config(text=money(-t["iva_pagar"]))
        self.l_dinero.config(text=(
            f"Dinero (con IVA):   Compras {money(t['compras'])}   ·   Gastos {money(t['gastos_bruto'])}   ·   "
            f"Flujo de caja (ventas − compras − gastos) {money(t['flujo'])}"))
        medios = self.db.ventas_por_medio(d, h)
        self.l_medios.config(text="Ventas por medio de pago:   " + "   ·   ".join(
            f"{r['medio_pago']}: {money(r['total'])}" for r in medios) if medios else "")
        sin_costo = self.db.ventas_sin_costo(d, h)
        self.l_alerta.config(text=(
            f"⚠ {sin_costo[1]} producto(s) vendido(s) no tienen costo cargado: la ganancia aparece más alta de lo "
            f"real. Cargue el costo en Productos o registre su compra." if sin_costo[0] else ""))
        stock = {p["id"]: p["stock"] for p in self.db.productos(inactivos=True)}
        limpiar_tabla(self.tabla)
        for i, p in enumerate(prods):
            gan = p["ventas_netas"] - p["costo"]
            self.tabla.insert("", "end", tags=("par",) if i % 2 else (), values=(
                p["nombre"], qty(p["cant_v"]), money(p["ventas"]), money(p["ventas"] - p["ventas_netas"]), money(gan),
                pct(gan / p["ventas_netas"]) if p["ventas_netas"] else "—", qty(p["cant_c"]), money(p["compras"]),
                qty(p["cant_m"]), money(p["mermas"]), qty(stock.get(p["id"], 0))))

    def exportar(self):
        try:
            d, h = self.periodo.rango()
            desc = self.periodo.descripcion()
            sufijo = self.periodo.sufijo_archivo()
        except ValueError as ex:
            messagebox.showerror("Período", str(ex), parent=self)
            return
        os.makedirs(self.app.REPORT_DIR, exist_ok=True)
        ruta = filedialog.asksaveasfilename(
            parent=self, title="Guardar reporte Excel", initialdir=self.app.REPORT_DIR,
            initialfile=f"Contabilidad_{sufijo}.xlsx", defaultextension=".xlsx",
            filetypes=[("Libro de Excel", "*.xlsx")])
        if not ruta:
            return
        self.app.config(cursor="watch")
        self.app.update_idletasks()
        try:
            generar_excel(self.db, d, h, desc, ruta)
        except PermissionError:
            messagebox.showerror("No se pudo guardar",
                                 "El archivo está abierto en Excel u otro programa.\nCiérrelo e intente de nuevo.",
                                 parent=self)
            return
        finally:
            self.app.config(cursor="")
        self.app.estado(f"Excel generado: {ruta}")
        if messagebox.askyesno("Listo", f"Reporte guardado en:\n{ruta}\n\n¿Abrirlo ahora?", parent=self):
            try:
                os.startfile(ruta)
            except OSError:
                messagebox.showinfo("Abrir", "No hay un programa instalado para abrir archivos .xlsx.", parent=self)
