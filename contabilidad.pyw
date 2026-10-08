# -*- coding: utf-8 -*-
"""Contabilidad del negocio: ventas, compras, mermas, gastos y caja.

Tkinter + SQLite, sin dependencias externas. Pensado para equipos modestos.
Los datos se guardan en ./datos/negocio.db y los reportes Excel en ./reportes/.
"""
import datetime as dt
import os
import re
import sqlite3
import sys
import traceback
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, simpledialog, ttk

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

from programa.modelo import DB, ArchivoInvalido  # noqa: E402
from programa.pestanas import (CajaTab, GastosTab, MermasTab, ProductosTab, ReportesTab,  # noqa: E402
                               TicketTab)
from programa.util import fmt_iso, hoy, parse_num  # noqa: E402

VERSION = "2.0"
DATA_DIR = os.path.join(BASE, "datos")
BACKUP_DIR = os.path.join(DATA_DIR, "respaldos")
REPORT_DIR = os.path.join(BASE, "reportes")
DB_PATH = os.path.join(DATA_DIR, "negocio.db")
LOG_PATH = os.path.join(DATA_DIR, "errores.log")
RESPALDOS_A_GUARDAR = 30
DIAS_AVISO_RESPALDO_EXTERNO = 7
_RESPALDO_DIARIO = re.compile(r"^negocio_\d{4}-\d{2}-\d{2}_\d{6}\.db$")


def registrar_error(texto):
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"\n===== {dt.datetime.now():%Y-%m-%d %H:%M:%S} (v{VERSION}) =====\n{texto}\n")
    except OSError:
        pass


def _marca_tiempo():
    return f"{dt.datetime.now():%Y-%m-%d_%H%M%S}"


class App(tk.Tk):
    REPORT_DIR = REPORT_DIR

    def __init__(self, db_path=None):
        super().__init__()
        self.withdraw()
        self.db = DB(db_path or DB_PATH, antes_de_migrar=self._respaldo_previo_a_migrar)
        self._estilos()
        self._titulo()
        self.geometry("1280x720")
        self.minsize(1000, 600)

        self._menu()
        # La barra de estado se empaqueta antes que las pestañas para que nunca quede sin espacio.
        barra = ttk.Frame(self)
        barra.pack(fill="x", side="bottom")
        self.l_alerta = tk.Label(barra, fg="#b35c00", bg=self.fondo, cursor="hand2", font=("Segoe UI", 10, "bold"),
                                 padx=10)
        self.l_alerta.pack(side="right")
        self.l_alerta.bind("<Button-1>", lambda e: self.ir_a_stock_bajo())
        self.v_estado = tk.StringVar(value="Listo.")
        ttk.Label(barra, textvariable=self.v_estado, anchor="w", padding=(8, 3)).pack(side="left", fill="x", expand=True)

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=6, pady=(6, 0))
        self.tabs = [
            ("Ventas", TicketTab(self.nb, self, "venta")),
            ("Compras", TicketTab(self.nb, self, "compra")),
            ("Mermas", MermasTab(self.nb, self)),
            ("Gastos", GastosTab(self.nb, self)),
            ("Caja", CajaTab(self.nb, self)),
            ("Productos", ProductosTab(self.nb, self)),
            ("Reportes y Excel", ReportesTab(self.nb, self)),
        ]
        for i, (txt, tab) in enumerate(self.tabs, start=1):
            self.nb.add(tab, text=f"  {txt}  ")
            self.bind(f"<Control-Key-{i}>", lambda e, t=tab: self.nb.select(t))
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self.tab_actual().refrescar())

        self.bind("<F5>", lambda e: self.tab_actual().refrescar())
        self.bind("<F12>", self._f12)
        self.protocol("WM_DELETE_WINDOW", self.salir)

        try:
            self.state("zoomed")
        except tk.TclError:
            pass
        self.deiconify()
        self.after(50, self._al_iniciar)

    # ------------------------------------------------------------ apariencia
    def _estilos(self):
        st = ttk.Style(self)
        if "vista" in st.theme_names():
            st.theme_use("vista")
        for nombre in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                tkfont.nametofont(nombre).configure(family="Segoe UI", size=10)
            except tk.TclError:
                pass
        self.fondo = st.lookup("TFrame", "background") or "SystemButtonFace"
        self.f_grande = tkfont.Font(family="Segoe UI", size=15, weight="bold")
        self.f_media = tkfont.Font(family="Segoe UI", size=13, weight="bold")
        self.f_chica = tkfont.Font(family="Segoe UI", size=9)
        self.f_tarjeta = tkfont.Font(family="Segoe UI", size=14, weight="bold")
        st.configure("Treeview", rowheight=23, font=("Segoe UI", 10))
        st.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        st.configure("Total.TLabel", font=("Segoe UI", 10, "bold"))
        st.configure("Sub.TLabel", font=("Segoe UI", 10, "bold"), foreground="#333333")
        st.configure("Grande.TLabel", font=("Segoe UI", 18, "bold"), foreground="#1e6b34")
        st.configure("Vuelto.TLabel", font=("Segoe UI", 13, "bold"), foreground="#1f4e9c")
        st.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
        st.configure("Cobrar.TButton", font=("Segoe UI", 11, "bold"), padding=(10, 8))
        st.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=(10, 5))

    def _titulo(self):
        self.title(f"Contabilidad — {self.db.config('negocio', 'Mi Negocio')}")

    def _menu(self):
        barra = tk.Menu(self)
        m = tk.Menu(barra, tearoff=0)
        m.add_command(label="Nombre del negocio…", command=self.cambiar_nombre)
        m.add_command(label="Tasa de IVA…", command=self.cambiar_iva)
        m.add_separator()
        m.add_command(label="Crear respaldo ahora", command=self.respaldo_manual)
        m.add_command(label="Guardar respaldo en pendrive u otra carpeta…", command=self.respaldo_externo)
        m.add_command(label="Restaurar un respaldo…", command=self.restaurar)
        m.add_separator()
        m.add_command(label="Abrir carpeta de respaldos", command=lambda: self._abrir(BACKUP_DIR))
        m.add_command(label="Abrir carpeta de reportes", command=lambda: self._abrir(REPORT_DIR))
        m.add_separator()
        m.add_command(label="Salir", command=self.salir)
        barra.add_cascade(label="Archivo", menu=m)
        a = tk.Menu(barra, tearoff=0)
        a.add_command(label="Cómo se usa", command=self.ayuda)
        a.add_command(label="Atajos de teclado", command=self.atajos)
        a.add_separator()
        a.add_command(label="Ver registro de errores", command=self._ver_errores)
        a.add_command(label="Acerca de", command=lambda: messagebox.showinfo(
            "Acerca de", f"Contabilidad del negocio\nVersión {VERSION}\n\nDatos: {DB_PATH}", parent=self))
        barra.add_cascade(label="Ayuda", menu=a)
        self.config(menu=barra)

    # ------------------------------------------------------------ general
    def tab_actual(self):
        return self.nametowidget(self.nb.select())

    def estado(self, txt):
        self.v_estado.set(txt)

    def _f12(self, e=None):
        tab = self.tab_actual()
        if isinstance(tab, TicketTab):
            tab.guardar()
        return "break"

    def _al_iniciar(self):
        self.tab_actual().refrescar()
        self.actualizar_alertas()
        self.respaldo_automatico()
        self.tabs[0][1].cb_prod.focus_set()

    def actualizar_alertas(self):
        n = self.db.cantidad_stock_bajo()
        self.l_alerta.config(text=f"⚠ {n} producto{'s' if n != 1 else ''} para reponer" if n else "")

    def ir_a_stock_bajo(self):
        tab = dict(self.tabs)["Productos"]
        self.nb.select(tab)
        tab.mostrar_stock_bajo()

    def report_callback_exception(self, tipo, valor, tb):
        """Cualquier error inesperado: se guarda en el registro y se avisa (sin consola no se vería)."""
        registrar_error("".join(traceback.format_exception(tipo, valor, tb)))
        if isinstance(valor, sqlite3.OperationalError) and "locked" in str(valor):
            msg = ("La base de datos está ocupada. ¿Está el programa abierto dos veces?\n"
                   "Cierre la otra ventana e intente de nuevo.")
        else:
            msg = (f"Ocurrió un error inesperado:\n\n{valor}\n\n"
                   f"Los datos guardados están a salvo. El detalle quedó en:\n{LOG_PATH}")
        messagebox.showerror("Error", msg, parent=self)

    def _abrir(self, carpeta):
        os.makedirs(carpeta, exist_ok=True)
        os.startfile(carpeta)

    def _ver_errores(self):
        if os.path.exists(LOG_PATH):
            os.startfile(LOG_PATH)
        else:
            messagebox.showinfo("Errores", "No se registró ningún error. 👍", parent=self)

    def cambiar_nombre(self):
        nuevo = simpledialog.askstring("Nombre del negocio", "Nombre que aparecerá en los reportes:",
                                       initialvalue=self.db.config("negocio", "Mi Negocio"), parent=self)
        if nuevo and nuevo.strip():
            self.db.set_config("negocio", nuevo.strip())
            self._titulo()

    def cambiar_iva(self):
        actual = self.db.tasa_iva()
        txt = simpledialog.askstring(
            "Tasa de IVA", f"Tasa de IVA general en % (actual: {actual:g}%).\n\n"
            "Se aplica a las ventas, compras y gastos que registre desde ahora.\n"
            "Lo ya registrado conserva la tasa con que se guardó.\n"
            "Los productos marcados como «Exento» no llevan IVA.", initialvalue=f"{actual:g}", parent=self)
        if txt is None:
            return
        tasa = parse_num(txt)
        if tasa is None or not 0 <= tasa < 100:
            messagebox.showwarning("Tasa de IVA", "Ingrese un porcentaje entre 0 y 99.", parent=self)
            return
        self.db.set_config("iva_tasa", f"{tasa:g}")
        self.tab_actual().refrescar()
        self.estado(f"Tasa de IVA: {tasa:g}%")

    # ------------------------------------------------------------ respaldos
    @staticmethod
    def _respaldo_previo_a_migrar(db):
        os.makedirs(BACKUP_DIR, exist_ok=True)
        db.respaldo(os.path.join(BACKUP_DIR, f"antes_de_actualizar_{_marca_tiempo()}.db"))

    def _hacer_respaldo(self):
        os.makedirs(BACKUP_DIR, exist_ok=True)
        ruta = os.path.join(BACKUP_DIR, f"negocio_{_marca_tiempo()}.db")
        self.db.respaldo(ruta)
        diarios = sorted(f for f in os.listdir(BACKUP_DIR) if _RESPALDO_DIARIO.match(f))
        for f in diarios[:-RESPALDOS_A_GUARDAR]:
            try:
                os.remove(os.path.join(BACKUP_DIR, f))
            except OSError:
                pass
        return ruta

    def respaldo_automatico(self):
        """Un respaldo por día al abrir el programa, y recordatorio de copia externa."""
        try:
            prefijo = f"negocio_{hoy():%Y-%m-%d}_"
            if not (os.path.isdir(BACKUP_DIR) and any(f.startswith(prefijo) for f in os.listdir(BACKUP_DIR))):
                self._hacer_respaldo()
                self.estado("Respaldo diario creado automáticamente.")
        except (OSError, sqlite3.Error) as ex:
            registrar_error(f"Respaldo automático: {ex}")
            self.estado(f"No se pudo crear el respaldo automático: {ex}")
            return
        ultimo = self.db.config("ultimo_respaldo_externo", "")
        hay_datos = self.db.q1("SELECT COUNT(*) FROM movimientos")[0] > 0
        if hay_datos and (not ultimo or (hoy() - dt.date.fromisoformat(ultimo)).days >= DIAS_AVISO_RESPALDO_EXTERNO):
            cuando = f"el {fmt_iso(ultimo)}" if ultimo else "nunca"
            self.estado(f"Recordatorio: la última copia en pendrive fue {cuando}. "
                        "Use Archivo → Guardar respaldo en pendrive.")

    def respaldo_manual(self):
        try:
            ruta = self._hacer_respaldo()
            messagebox.showinfo("Respaldo", f"Respaldo creado:\n{ruta}", parent=self)
        except (OSError, sqlite3.Error) as ex:
            messagebox.showerror("Respaldo", f"No se pudo crear el respaldo:\n{ex}", parent=self)

    def respaldo_externo(self):
        carpeta = filedialog.askdirectory(
            parent=self, title="Elija el pendrive o la carpeta donde guardar la copia",
            initialdir=self.db.config("carpeta_respaldo_externo", "") or None, mustexist=True)
        if not carpeta:
            return
        ruta = os.path.join(carpeta, f"respaldo_contabilidad_{_marca_tiempo()}.db")
        try:
            self.db.respaldo(ruta)
            DB.validar_archivo(ruta)
        except (OSError, sqlite3.Error, ArchivoInvalido) as ex:
            messagebox.showerror("Respaldo", f"No se pudo guardar la copia en esa carpeta:\n{ex}", parent=self)
            return
        self.db.set_config("carpeta_respaldo_externo", carpeta)
        self.db.set_config("ultimo_respaldo_externo", hoy().isoformat())
        self.estado(f"Copia guardada en {ruta}")
        messagebox.showinfo("Respaldo", f"Copia guardada y verificada:\n{ruta}", parent=self)

    def restaurar(self):
        ruta = filedialog.askopenfilename(
            parent=self, title="Elija el respaldo a restaurar", initialdir=BACKUP_DIR,
            filetypes=[("Respaldo de la contabilidad", "*.db"), ("Todos los archivos", "*.*")])
        if not ruta:
            return
        if os.path.normcase(os.path.abspath(ruta)) == os.path.normcase(os.path.abspath(self.db.ruta)):
            messagebox.showwarning("Restaurar", "Ese es el archivo de datos en uso. Elija un respaldo.", parent=self)
            return
        try:
            info = DB.validar_archivo(ruta)
        except ArchivoInvalido as ex:
            messagebox.showerror("Restaurar", str(ex), parent=self)
            return
        ultima = fmt_iso(info["ultima"]) if info["ultima"] else "—"
        if not messagebox.askyesno(
                "Restaurar respaldo",
                f"Se reemplazarán TODOS los datos actuales por los del respaldo:\n\n"
                f"   {info['productos']} productos\n   {info['movimientos']} movimientos\n"
                f"   Último movimiento: {ultima}\n\n"
                "Antes se guardará una copia de los datos actuales por si acaso.\n¿Continuar?",
                icon="warning", parent=self):
            return
        if any(tab.hay_cambios() for _, tab in self.tabs):
            if not messagebox.askyesno("Restaurar", "Hay datos escritos sin guardar que se perderán. ¿Continuar?",
                                       parent=self):
                return
        try:
            os.makedirs(BACKUP_DIR, exist_ok=True)
            seguro = os.path.join(BACKUP_DIR, f"antes_de_restaurar_{_marca_tiempo()}.db")
            self.db.respaldo(seguro)
            self.db.restaurar(ruta)
        except (OSError, sqlite3.Error) as ex:
            registrar_error(traceback.format_exc())
            messagebox.showerror("Restaurar", f"No se pudo restaurar:\n{ex}", parent=self)
            return
        for _, tab in self.tabs:
            if hasattr(tab, "limpiar"):
                tab.limpiar(foco=False)
        self._titulo()
        self.tab_actual().refrescar()
        self.actualizar_alertas()
        messagebox.showinfo("Restaurar", f"Datos restaurados.\nLa copia de los datos anteriores quedó en:\n{seguro}",
                            parent=self)

    # ------------------------------------------------------------ ayuda y salida
    def ayuda(self):
        messagebox.showinfo("Cómo se usa", (
            "1. PRODUCTOS: cargue sus productos con costo y precio de venta.\n\n"
            "2. VENTAS: escriba el producto y Enter, la cantidad y Enter. Repita con cada producto. "
            "Al terminar, Enter con el campo vacío, escriba con cuánto paga el cliente y Enter "
            "(o F12) para guardar.\n\n"
            "3. COMPRAS: igual que ventas. El costo de cada producto se actualiza solo.\n\n"
            "4. MERMAS: mercadería que se perdió (vencida, rota…).\n\n"
            "5. GASTOS: alquiler, luz, sueldos, impuestos…\n\n"
            "6. CAJA: al final del día cuente el efectivo y compárelo con lo que debería haber.\n\n"
            "7. REPORTES: elija el mes o el rango de fechas y exporte a Excel.\n\n"
            "IVA: todos los precios se cargan CON IVA incluido. El programa separa el neto y el IVA, "
            "calcula el IVA débito (ventas), el crédito (compras y gastos con factura) y el IVA a pagar "
            "de cada mes. Marque como «Exento» los productos que no llevan IVA."), parent=self)

    def atajos(self):
        messagebox.showinfo("Atajos de teclado", (
            "Ctrl+1 … Ctrl+7   Cambiar de pestaña\n"
            "F12                    Cobrar / guardar el ticket\n"
            "Enter                 Pasar al campo siguiente / agregar / guardar\n"
            "Supr                  Quitar línea o eliminar registro seleccionado\n"
            "Doble clic           Abrir o corregir un registro\n"
            "F5                     Actualizar\n\n"
            "En los campos de fecha:\n"
            "  ↑ / ↓  suma o resta un día     F4  abre el calendario\n\n"
            "Fechas: dd/mm/aaaa (también 5/3/26). Números: coma o punto decimal."), parent=self)

    def salir(self):
        pendientes = [txt for txt, tab in self.tabs if tab.hay_cambios()]
        if pendientes and not messagebox.askyesno(
                "Salir", f"Hay datos sin guardar en: {', '.join(pendientes)}.\n¿Salir igual?", parent=self):
            return
        try:
            self.db.cerrar()
        finally:
            self.destroy()


def main():
    try:  # texto nítido en pantallas con escalado
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    try:
        app = App()
    except Exception as ex:
        registrar_error(traceback.format_exc())
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("No se pudo iniciar",
                             f"El programa no pudo abrirse:\n\n{ex}\n\nDetalle guardado en:\n{LOG_PATH}")
        root.destroy()
        return
    app.mainloop()


if __name__ == "__main__":
    main()
