# -*- coding: utf-8 -*-
"""Componentes de interfaz reutilizables."""
import calendar
import datetime as dt
import tkinter as tk
from tkinter import ttk

from .util import DIAS_CORTOS, MESES, fmt_fecha, hoy, parse_fecha


def crear_tabla(master, columnas, seleccion="extended", alto=None):
    """columnas: lista de (id, título, ancho, alineación 'w'|'e'|'center'). Devuelve (marco, tabla)."""
    marco = ttk.Frame(master)
    kw = {"height": alto} if alto else {}
    tabla = ttk.Treeview(marco, columns=[c[0] for c in columnas], show="headings", selectmode=seleccion, **kw)
    for cid, titulo, ancho, alin in columnas:
        tabla.heading(cid, text=titulo, anchor=alin)
        tabla.column(cid, width=ancho, minwidth=40, anchor=alin, stretch=(alin == "w"))
    sb = ttk.Scrollbar(marco, orient="vertical", command=tabla.yview)
    tabla.configure(yscrollcommand=sb.set)
    tabla.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    tabla.tag_configure("par", background="#f2f6fb")
    tabla.tag_configure("alerta", foreground="#b00020")
    tabla.tag_configure("bajo", foreground="#b35c00")
    tabla.tag_configure("inactivo", foreground="#999999")
    tabla.tag_configure("total", font=("Segoe UI", 10, "bold"))
    return marco, tabla


def limpiar_tabla(tabla):
    tabla.delete(*tabla.get_children())


class ProductoCombo(ttk.Combobox):
    """Combobox que filtra la lista a medida que se escribe."""

    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.todos = []
        self.bind("<KeyRelease>", self._filtrar)

    def set_items(self, nombres):
        self.todos = nombres
        self["values"] = nombres

    def coincidencias(self):
        t = self.get().strip().lower()
        return [n for n in self.todos if t in n.lower()] if t else list(self.todos)

    def restaurar(self):
        self["values"] = self.todos

    def _filtrar(self, e):
        if e.keysym in ("Up", "Down", "Return", "KP_Enter", "Escape", "Tab", "ISO_Left_Tab", "F12"):
            return
        self["values"] = self.coincidencias()


class Calendario(tk.Toplevel):
    """Ventanita para elegir una fecha con el mouse."""

    def __init__(self, campo):
        super().__init__(campo)
        self.campo = campo
        self.withdraw()
        self.title("Elegir fecha")
        self.resizable(False, False)
        self.transient(campo.winfo_toplevel())
        self.configure(bg="white", padx=6, pady=6)
        self.sel = campo.fecha() or hoy()
        self.y, self.m = self.sel.year, self.sel.month

        cab = tk.Frame(self, bg="white")
        cab.pack(fill="x")
        tk.Button(cab, text="◀", width=3, relief="flat", bg="#e8eef7", command=lambda: self._mover(-1)).pack(side="left")
        self.titulo = tk.Label(cab, bg="white", font=("Segoe UI", 11, "bold"), width=16)
        self.titulo.pack(side="left", expand=True)
        tk.Button(cab, text="▶", width=3, relief="flat", bg="#e8eef7", command=lambda: self._mover(1)).pack(side="right")

        self.grilla = tk.Frame(self, bg="white")
        self.grilla.pack(pady=4)
        tk.Button(self, text="Hoy", relief="flat", bg="#e8eef7",
                  command=lambda: self._elegir(hoy())).pack(fill="x")

        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Prior>", lambda e: self._mover(-1))
        self.bind("<Next>", lambda e: self._mover(1))
        self._dibujar()

        self.update_idletasks()
        x = campo.winfo_rootx()
        y = campo.winfo_rooty() + campo.winfo_height() + 2
        if y + self.winfo_reqheight() > self.winfo_screenheight() - 40:
            y = campo.winfo_rooty() - self.winfo_reqheight() - 30
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.deiconify()
        self.grab_set()
        self.focus_set()

    def _mover(self, delta):
        m = self.m - 1 + delta
        self.y += m // 12
        self.m = m % 12 + 1
        self._dibujar()

    def _dibujar(self):
        for w in self.grilla.winfo_children():
            w.destroy()
        self.titulo.config(text=f"{MESES[self.m - 1]} {self.y}")
        for c, nombre in enumerate(DIAS_CORTOS):
            tk.Label(self.grilla, text=nombre, bg="white", fg="#666666",
                     font=("Segoe UI", 9, "bold")).grid(row=0, column=c, padx=1)
        t = hoy()
        for r, semana in enumerate(calendar.Calendar(0).monthdatescalendar(self.y, self.m), start=1):
            for c, d in enumerate(semana):
                bg, fg = "white", ("#222222" if d.month == self.m else "#aaaaaa")
                if d == t:
                    bg = "#fff2c2"
                if d == self.sel:
                    bg, fg = "#305496", "white"
                tk.Button(self.grilla, text=str(d.day), width=3, relief="flat", bg=bg, fg=fg,
                          activebackground="#c9d7ee", font=("Segoe UI", 10),
                          command=lambda d=d: self._elegir(d)).grid(row=r, column=c, padx=1, pady=1)

    def _elegir(self, d):
        self.grab_release()
        self.destroy()
        self.campo.set(d, avisar=True)
        self.campo.entry.focus_set()


class CampoFecha(ttk.Frame):
    """Campo de fecha con botón de calendario. Flechas ↑/↓ suman o restan un día."""

    def __init__(self, master, al_cambiar=None, ancho=11, fecha=None):
        super().__init__(master)
        self.al_cambiar = al_cambiar
        self.var = tk.StringVar(value=fmt_fecha(fecha or hoy()))
        self.entry = ttk.Entry(self, textvariable=self.var, width=ancho)
        self.entry.pack(side="left", fill="x", expand=True)
        self.boton = ttk.Button(self, text="▾", width=2, command=self.abrir, takefocus=False)
        self.boton.pack(side="left", padx=(2, 0))
        self.entry.bind("<Up>", lambda e: self.mover(1))
        self.entry.bind("<Down>", lambda e: self.mover(-1))
        self.entry.bind("<F4>", lambda e: self.abrir())
        self.entry.bind("<Alt-Down>", lambda e: self.abrir())

    def fecha(self):
        return parse_fecha(self.var.get())

    def set(self, d, avisar=False):
        self.var.set(fmt_fecha(d))
        if avisar and self.al_cambiar:
            self.al_cambiar()

    def mover(self, dias):
        self.set((self.fecha() or hoy()) + dt.timedelta(days=dias), avisar=True)
        return "break"

    def abrir(self):
        Calendario(self)
        return "break"

    def bind_enter(self, func):
        self.entry.bind("<Return>", func)
        self.entry.bind("<KP_Enter>", func)


class PeriodoSelector(ttk.Frame):
    """Elige un mes, un rango de fechas o todo."""

    def __init__(self, master, al_cambiar=None):
        super().__init__(master)
        self.al_cambiar = al_cambiar
        t = hoy()
        self.modo = tk.StringVar(value="Mes")
        self.mes = tk.StringVar(value=MESES[t.month - 1])
        self.anio = tk.StringVar(value=str(t.year))

        ttk.Label(self, text="Período:").pack(side="left")
        cb = ttk.Combobox(self, textvariable=self.modo, values=["Mes", "Día", "Rango de fechas", "Todo"],
                          state="readonly", width=14)
        cb.pack(side="left", padx=4)
        cb.bind("<<ComboboxSelected>>", lambda e: self._cambiar_modo())

        self.f_mes = ttk.Frame(self)
        ttk.Button(self.f_mes, text="◀", width=3, command=lambda: self._mover(-1)).pack(side="left")
        cm = ttk.Combobox(self.f_mes, textvariable=self.mes, values=MESES, state="readonly", width=11)
        cm.pack(side="left", padx=2)
        cm.bind("<<ComboboxSelected>>", lambda e: self._avisar())
        sp = ttk.Spinbox(self.f_mes, textvariable=self.anio, from_=2000, to=2100, width=6, command=self._avisar)
        sp.pack(side="left", padx=2)
        sp.bind("<Return>", lambda e: self._avisar())
        ttk.Button(self.f_mes, text="▶", width=3, command=lambda: self._mover(1)).pack(side="left")
        ttk.Button(self.f_mes, text="Este mes", command=self._este_mes).pack(side="left", padx=(6, 0))

        self.f_dia = ttk.Frame(self)
        self.dia = CampoFecha(self.f_dia, al_cambiar=self._avisar)
        self.dia.pack(side="left")
        self.dia.bind_enter(lambda e: self._avisar())
        ttk.Button(self.f_dia, text="Hoy", command=lambda: self.dia.set(hoy(), avisar=True)).pack(side="left", padx=4)

        self.f_rango = ttk.Frame(self)
        ttk.Label(self.f_rango, text="Desde").pack(side="left")
        self.desde = CampoFecha(self.f_rango, al_cambiar=self._avisar, fecha=t.replace(day=1))
        self.desde.pack(side="left", padx=3)
        ttk.Label(self.f_rango, text="Hasta").pack(side="left")
        self.hasta = CampoFecha(self.f_rango, al_cambiar=self._avisar)
        self.hasta.pack(side="left", padx=3)
        for c in (self.desde, self.hasta):
            c.bind_enter(lambda e: self._avisar())

        self._cambiar_modo(inicial=True)

    def _cambiar_modo(self, inicial=False):
        for f in (self.f_mes, self.f_dia, self.f_rango):
            f.pack_forget()
        m = self.modo.get()
        if m == "Mes":
            self.f_mes.pack(side="left")
        elif m == "Día":
            self.f_dia.pack(side="left")
        elif m.startswith("Rango"):
            self.f_rango.pack(side="left")
        if not inicial:
            self._avisar()

    def _mover(self, delta):
        try:
            y = int(self.anio.get())
        except ValueError:
            y = hoy().year
        m = MESES.index(self.mes.get()) + delta
        self.mes.set(MESES[m % 12])
        self.anio.set(str(y + m // 12))
        self._avisar()

    def _este_mes(self):
        t = hoy()
        self.mes.set(MESES[t.month - 1])
        self.anio.set(str(t.year))
        self._avisar()

    def _avisar(self):
        if self.al_cambiar:
            self.al_cambiar()

    def rango(self):
        """(desde, hasta) como date. Lanza ValueError si los datos no son válidos."""
        m = self.modo.get()
        if m == "Mes":
            try:
                y = int(self.anio.get())
                if not 1900 <= y <= 9999:
                    raise ValueError
            except ValueError:
                raise ValueError("El año no es válido.") from None
            mi = MESES.index(self.mes.get()) + 1
            return dt.date(y, mi, 1), dt.date(y, mi, calendar.monthrange(y, mi)[1])
        if m == "Todo":
            return dt.date(1900, 1, 1), dt.date(9999, 12, 31)
        if m == "Día":
            d = self.dia.fecha()
            if not d:
                raise ValueError("Fecha inválida. Use el formato dd/mm/aaaa (ej. 15/03/2026).")
            return d, d
        d, h = self.desde.fecha(), self.hasta.fecha()
        if not d or not h:
            raise ValueError("Fecha inválida. Use el formato dd/mm/aaaa (ej. 15/03/2026).")
        return (d, h) if d <= h else (h, d)

    def descripcion(self):
        d, h = self.rango()
        m = self.modo.get()
        if m == "Mes":
            return f"{self.mes.get()} {self.anio.get()}"
        if m == "Todo":
            return "Todos los registros"
        if m == "Día":
            return fmt_fecha(d)
        return f"{fmt_fecha(d)} al {fmt_fecha(h)}"

    def sufijo_archivo(self):
        d, h = self.rango()
        m = self.modo.get()
        if m == "Mes":
            return d.strftime("%Y-%m")
        if m == "Todo":
            return "completo"
        if m == "Día":
            return d.isoformat()
        return f"{d.isoformat()}_a_{h.isoformat()}"
