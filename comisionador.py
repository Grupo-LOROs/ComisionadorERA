import tkinter as tk
from tkinter import ttk, messagebox
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from datetime import date

META_MENSUAL_DEFAULT = Decimal("1666666.67")
LIMITE_2 = Decimal("3000000.00")

# ----- Ajuste de proporción de espacio -----
# Aproximadamente 45% para inputs y 55% para espacio restante
INPUT_W = 9
SPACER_W = 11

MONTHS = [
    ("Enero", 1), ("Febrero", 2), ("Marzo", 3), ("Abril", 4),
    ("Mayo", 5), ("Junio", 6), ("Julio", 7), ("Agosto", 8),
    ("Septiembre", 9), ("Octubre", 10), ("Noviembre", 11), ("Diciembre", 12),
]
MONTH_NAMES = [m[0] for m in MONTHS]
MONTH_NAME_TO_NUM = {name: num for name, num in MONTHS}
MONTH_NUM_TO_NAME = {num: name for name, num in MONTHS}


# ---------- Utilidades ----------
def parse_money(s: str) -> Decimal:
    s = (s or "").strip().replace("$", "").replace(",", "")
    if s == "":
        raise ValueError("Campo vacío")
    try:
        return Decimal(s)
    except InvalidOperation:
        raise ValueError("Número inválido")


def fmt_money(x: Decimal) -> str:
    x = x.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return f"${x:,.2f}"


def fmt_pct(x: Decimal) -> str:
    x = (x * Decimal("100")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return f"{x}%"


def meses_desde_ingreso(inicio: date, periodo: date) -> int:
    delta = (periodo.year - inicio.year) * 12 + (periodo.month - inicio.month) + 1
    return max(delta, 0)


def smart_year_values(start=2018, end=2035) -> list[str]:
    years = list(range(start, end + 1))
    now_y = date.today().year
    if now_y not in years:
        years.append(now_y)
        years.sort()
    return [str(y) for y in years]


def configure_input_ratio(frame: ttk.Frame):
    """
    3 columnas:
      0: labels (fijo)
      1: inputs (porcentaje)
      2: spacer (absorbe el resto)
    """
    frame.grid_columnconfigure(0, weight=0)
    frame.grid_columnconfigure(1, weight=INPUT_W)
    frame.grid_columnconfigure(2, weight=SPACER_W)


# ---------- Reglas ----------
def tasa_asesor(cobrado_mes: Decimal) -> Decimal:
    if cobrado_mes <= META_MENSUAL_DEFAULT:
        return Decimal("0.055")
    elif cobrado_mes <= LIMITE_2:
        return Decimal("0.065")
    else:
        return Decimal("0.08")


def comision_asesor(utilidad_bruta_mes: Decimal, cobrado_mes: Decimal):
    tasa = tasa_asesor(cobrado_mes)
    com = (utilidad_bruta_mes * tasa).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return tasa, com


def tasa_coordinador(meses: int, cumplimiento: Decimal) -> Decimal:
    if meses <= 5:
        return Decimal("0.30")
    if cumplimiento < Decimal("0.80"):
        return Decimal("0.20")
    elif cumplimiento < Decimal("1.00"):
        return Decimal("0.30")
    else:
        return Decimal("0.40")


# ---------- App ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Calculadora de Comisiones")
        self.minsize(820, 480)

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.container = ttk.Frame(self, padding=12)
        self.container.grid(row=0, column=0, sticky="nsew")
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for F in (StartFrame, AsesorFrame, CoordinadorFrame):
            frame = F(self.container, self)
            self.frames[F.__name__] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show("StartFrame")
        self.after(0, self.center_window)

    def show(self, name: str):
        self.frames[name].tkraise()

    def center_window(self):
        self.update_idletasks()
        w = self.winfo_width() or 880
        h = self.winfo_height() or 520
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")


class StartFrame(ttk.Frame):
    def __init__(self, parent, app: App):
        super().__init__(parent)
        self.app = app

        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        card = ttk.Frame(self)
        card.grid(row=1, column=0, sticky="n", pady=8)

        ttk.Label(card, text="Selecciona tipo de usuario", font=("Segoe UI", 16, "bold")).grid(
            row=0, column=0, columnspan=2, pady=(0, 16)
        )

        ttk.Button(card, text="Asesor", width=20, command=lambda: app.show("AsesorFrame")).grid(
            row=1, column=0, padx=10, pady=6
        )
        ttk.Button(card, text="Coordinador", width=20, command=lambda: app.show("CoordinadorFrame")).grid(
            row=1, column=1, padx=10, pady=6
        )

        info = (
            "Reglas:\n"
            "- Comisión asesor = Utilidad bruta mensual * % según cobrado mensual.\n"
            "- Comisión coordinador = % sobre la comisión del asesor.\n"
        )
        ttk.Label(card, text=info, justify="left").grid(row=2, column=0, columnspan=2, pady=(16, 0))


class AsesorFrame(ttk.Frame):
    def __init__(self, parent, app: App):
        super().__init__(parent)
        self.app = app

        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        header = ttk.Frame(self)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)

        ttk.Button(header, text="← Regresar", command=lambda: app.show("StartFrame")).grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Asesor", font=("Segoe UI", 14, "bold")).grid(row=0, column=1, sticky="w", padx=12)

        form = ttk.LabelFrame(self, text="Datos del mes", padding=12)
        form.grid(row=1, column=0, sticky="ew", pady=12)
        configure_input_ratio(form)

        self.cobrado_var = tk.StringVar()
        self.utilidad_var = tk.StringVar()

        ttk.Label(form, text="Cobrado total del mes:").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.cobrado_var).grid(row=0, column=1, sticky="ew", pady=6)

        ttk.Label(form, text="Utilidad bruta del mes:").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.utilidad_var).grid(row=1, column=1, sticky="ew", pady=6)

        # Botón solo en la columna de input (no ocupa todo el ancho del frame)
        ttk.Button(form, text="Calcular", command=self.calcular).grid(row=2, column=1, sticky="ew", pady=10)

        out = ttk.LabelFrame(self, text="Resultado", padding=12)
        out.grid(row=2, column=0, sticky="nsew")
        out.grid_columnconfigure(0, weight=1)
        out.grid_rowconfigure(0, weight=1)

        self.out_lbl = ttk.Label(out, text="—", justify="left")
        self.out_lbl.grid(row=0, column=0, sticky="nw")

    def calcular(self):
        try:
            cobrado = parse_money(self.cobrado_var.get())
            utilidad = parse_money(self.utilidad_var.get())
            tasa, com = comision_asesor(utilidad, cobrado)

            txt = (
                f"Cobrado: {fmt_money(cobrado)}\n"
                f"Utilidad bruta: {fmt_money(utilidad)}\n"
                f"% Asesor: {fmt_pct(tasa)}\n"
                f"Comisión Asesor: {fmt_money(com)}"
            )
            self.out_lbl.config(text=txt)
        except Exception as e:
            messagebox.showerror("Error", str(e))


class CoordinadorFrame(ttk.Frame):
    def __init__(self, parent, app: App):
        super().__init__(parent)
        self.app = app
        self.total_coord = Decimal("0.00")

        self.grid_rowconfigure(3, weight=1)
        self.grid_columnconfigure(0, weight=1)

        header = ttk.Frame(self)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(1, weight=1)

        ttk.Button(header, text="← Regresar", command=self._back).grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Coordinador", font=("Segoe UI", 14, "bold")).grid(row=0, column=1, sticky="w", padx=12)

        now = date.today()
        years = smart_year_values(2018, 2035)

        # Parámetros
        top = ttk.LabelFrame(self, text="Parámetros", padding=12)
        top.grid(row=1, column=0, sticky="ew", pady=(12, 8))
        configure_input_ratio(top)

        self.periodo_mes = tk.StringVar(value=MONTH_NUM_TO_NAME[now.month])
        self.periodo_anio = tk.StringVar(value=str(now.year))
        self.meta_var = tk.StringVar(value=str(META_MENSUAL_DEFAULT))

        ttk.Label(top, text="Periodo (mes/año):").grid(row=0, column=0, sticky="w", pady=6)

        period_box = ttk.Frame(top)
        period_box.grid(row=0, column=1, sticky="w", pady=6)
        ttk.Combobox(period_box, textvariable=self.periodo_mes, values=MONTH_NAMES, state="readonly", width=14)\
            .grid(row=0, column=0, sticky="w")
        ttk.Combobox(period_box, textvariable=self.periodo_anio, values=years, state="readonly", width=10)\
            .grid(row=0, column=1, sticky="w", padx=(6, 0))

        ttk.Label(top, text="Meta mensual:").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(top, textvariable=self.meta_var).grid(row=1, column=1, sticky="ew", pady=6)

        # Agregar asesor
        add = ttk.LabelFrame(self, text="Agregar asesor (uno por uno)", padding=12)
        add.grid(row=2, column=0, sticky="ew", pady=(8, 8))
        configure_input_ratio(add)

        self.ing_mes = tk.StringVar(value=MONTH_NUM_TO_NAME[now.month])
        self.ing_anio = tk.StringVar(value=str(now.year))
        self.a_cobrado = tk.StringVar()
        self.a_utilidad = tk.StringVar()

        ttk.Label(add, text="Ingreso (mes/año):").grid(row=0, column=0, sticky="w", pady=6)

        ing_box = ttk.Frame(add)
        ing_box.grid(row=0, column=1, sticky="w", pady=6)
        ttk.Combobox(ing_box, textvariable=self.ing_mes, values=MONTH_NAMES, state="readonly", width=14)\
            .grid(row=0, column=0, sticky="w")
        ttk.Combobox(ing_box, textvariable=self.ing_anio, values=years, state="readonly", width=10)\
            .grid(row=0, column=1, sticky="w", padx=(6, 0))

        ttk.Label(add, text="Cobrado del mes:").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(add, textvariable=self.a_cobrado).grid(row=1, column=1, sticky="ew", pady=6)

        ttk.Label(add, text="Utilidad bruta del mes:").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(add, textvariable=self.a_utilidad).grid(row=2, column=1, sticky="ew", pady=6)

        # Botón solo en columna de input (no toma todo el ancho del frame)
        ttk.Button(add, text="Agregar y calcular", command=self.agregar_asesor)\
            .grid(row=3, column=1, pady=8, sticky="ew")

        # Tabla
        table_box = ttk.LabelFrame(self, text="Detalle (asesores capturados)", padding=8)
        table_box.grid(row=3, column=0, sticky="nsew")
        table_box.grid_rowconfigure(0, weight=1)
        table_box.grid_columnconfigure(0, weight=1)

        cols = ("meses", "cobrado", "utilidad", "pct_asesor", "com_asesor", "cumpl", "pct_coord", "com_coord")
        self.tree = ttk.Treeview(table_box, columns=cols, show="headings")

        headers = {
            "meses": "Meses",
            "cobrado": "Cobrado",
            "utilidad": "Utilidad",
            "pct_asesor": "% Asesor",
            "com_asesor": "Com. Asesor",
            "cumpl": "Cumpl.",
            "pct_coord": "% Coord.",
            "com_coord": "Com. Coord."
        }
        widths = {"meses": 60, "cobrado": 120, "utilidad": 120, "pct_asesor": 80,
                  "com_asesor": 120, "cumpl": 80, "pct_coord": 80, "com_coord": 120}

        for c in cols:
            self.tree.heading(c, text=headers[c])
            self.tree.column(c, width=widths[c], anchor="center")

        vsb = ttk.Scrollbar(table_box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        bottom = ttk.Frame(self)
        bottom.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        bottom.grid_columnconfigure(0, weight=1)

        self.total_lbl = ttk.Label(bottom, text="Total comisión coordinador: $0.00", font=("Segoe UI", 11, "bold"))
        self.total_lbl.grid(row=0, column=0, sticky="w")

        ttk.Button(bottom, text="Limpiar", command=self.limpiar).grid(row=0, column=1, sticky="e")

    def _back(self):
        self.limpiar()
        self.app.show("StartFrame")

    def periodo_date(self) -> date:
        mes_num = MONTH_NAME_TO_NUM[self.periodo_mes.get()]
        return date(int(self.periodo_anio.get()), mes_num, 1)

    def agregar_asesor(self):
        try:
            meta = parse_money(self.meta_var.get())
            if meta <= 0:
                raise ValueError("La meta mensual debe ser > 0")

            periodo = self.periodo_date()

            ing_mes_num = MONTH_NAME_TO_NUM[self.ing_mes.get()]
            ingreso = date(int(self.ing_anio.get()), ing_mes_num, 1)

            cobrado = parse_money(self.a_cobrado.get())
            utilidad = parse_money(self.a_utilidad.get())

            meses = meses_desde_ingreso(ingreso, periodo)
            pctA, comA = comision_asesor(utilidad, cobrado)
            cumplimiento = (cobrado / meta)

            pctC = tasa_coordinador(meses, cumplimiento)
            comC = (comA * pctC).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

            self.total_coord += comC
            self.total_lbl.config(text=f"Total comisión coordinador: {fmt_money(self.total_coord)}")

            self.tree.insert("", "end", values=(
                meses,
                fmt_money(cobrado),
                fmt_money(utilidad),
                fmt_pct(pctA),
                fmt_money(comA),
                f"{(cumplimiento*Decimal('100')).quantize(Decimal('0.01'))}%",
                fmt_pct(pctC),
                fmt_money(comC),
            ))

            self.a_cobrado.set("")
            self.a_utilidad.set("")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def limpiar(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.total_coord = Decimal("0.00")
        self.total_lbl.config(text="Total comisión coordinador: $0.00")
        self.a_cobrado.set("")
        self.a_utilidad.set("")
        self.meta_var.set(str(META_MENSUAL_DEFAULT))


if __name__ == "__main__":
    try:
        app = App()
        app.mainloop()
    except Exception as e:
        messagebox.showerror("Error fatal", str(e))
