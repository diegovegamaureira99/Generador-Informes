#!/usr/bin/env python3
"""
Bitácora Apoyo Pedagógico
Aplicación de escritorio para generar informes de clases en PDF.

Requisitos:  Python 3.8+  y  reportlab   ->   pip install reportlab
(tkinter y sqlite3 vienen incluidos con Python)
"""
import os
import re
import sqlite3
import subprocess
import sys
import tkinter as tk
import unicodedata
from datetime import date, datetime
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# --------------------------------------------------------------------------
# Configuración
# --------------------------------------------------------------------------
CARPETA_APP = os.path.dirname(os.path.abspath(sys.argv[0]))

CARPETA_DATOS = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
    "BitacoraApoyoPedagogico"
)

os.makedirs(CARPETA_DATOS, exist_ok=True)

DB_PATH = os.path.join(CARPETA_DATOS, "bitacora.db")
CARPETA_INFORMES_DEFECTO = os.path.join(CARPETA_APP, "Informes")

# Cursos del desplegable = cada nivel combinado con cada letra.
# Edite NIVELES y LETRAS según su establecimiento.
NIVELES = (
    ["Pre-Kínder", "Kínder"]
    + [f"{n}° Básico" for n in range(1, 9)]
    + [f"{n}° Medio" for n in range(1, 5)]
)
CURSOS = [f"{nivel}" for nivel in NIVELES]

# Paleta
BG = "#F3F4F6"
CARD = "#FFFFFF"
BORDER = "#D1D5DB"
TEXT = "#111827"
MUTED = "#6B7280"
PRIMARY = "#2563EB"
PRIMARY_H = "#1D4ED8"
DANGER = "#DC2626"
DANGER_H = "#B91C1C"
FUENTE = "Helvetica"  # se reemplaza al iniciar por la mejor fuente disponible

CAMPOS_NOMBRE = ("nombre", "segundo_nombre", "apellido_paterno", "apellido_materno")


# --------------------------------------------------------------------------
# Utilidades de texto y RUT
# --------------------------------------------------------------------------
def normalizar_nombre(texto):
    """Formato de almacenamiento: minúsculas y sin espacios sobrantes."""
    return " ".join((texto or "").split()).lower()


def titulo(texto):
    """Formato de visualización: Primera Letra De Cada Palabra En Mayúscula."""
    return " ".join((texto or "").split()).title()


def sin_tildes(texto):
    return "".join(
        c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn"
    )


def nombre_valido(texto):
    return bool(texto) and all(c.isalpha() or c in " '-" for c in texto)


def nombre_completo(a):
    return " ".join(titulo(a[c]) for c in CAMPOS_NOMBRE if a[c])


def limpiar_rut(texto):
    """Deja solo dígitos y, a lo más, una K final (dígito verificador). Máx. 9."""
    t = re.sub(r"[^0-9kK]", "", texto).upper()[:9]
    if len(t) > 1:
        t = t[:-1].replace("K", "") + t[-1]
    return t


def formatear_rut(texto):
    """12345678K -> 12.345.678-K (sirve también para RUT incompletos)."""
    t = limpiar_rut(texto)
    if len(t) <= 1:
        return t
    cuerpo, dv = t[:-1], t[-1]
    grupos = []
    while cuerpo:
        grupos.insert(0, cuerpo[-3:])
        cuerpo = cuerpo[:-3]
    return ".".join(grupos) + "-" + dv


def dv_rut(cuerpo):
    suma, mult = 0, 2
    for c in reversed(cuerpo):
        suma += int(c) * mult
        mult = 2 if mult == 7 else mult + 1
    r = 11 - suma % 11
    return "0" if r == 11 else "K" if r == 10 else str(r)


def normalizar_rut(texto):
    """Devuelve el RUT con formato 12.345.678-9, o None si no es válido."""
    t = limpiar_rut(texto)
    if len(t) < 8 or not t[:-1].isdigit():
        return None
    if t[-1] != dv_rut(t[:-1]):
        return None
    return formatear_rut(t)


# --------------------------------------------------------------------------
# Base de datos
# --------------------------------------------------------------------------
class BaseDatos:
    def __init__(self, ruta=DB_PATH):
        self.con = sqlite3.connect(ruta)
        self.con.row_factory = sqlite3.Row
        self.con.execute(
            """CREATE TABLE IF NOT EXISTS alumnos (
                   rut TEXT PRIMARY KEY,
                   nombre TEXT NOT NULL,
                   segundo_nombre TEXT,
                   apellido_paterno TEXT NOT NULL,
                   apellido_materno TEXT,
                   observaciones TEXT
               )"""
        )
        self.con.execute(
            "CREATE TABLE IF NOT EXISTS config (clave TEXT PRIMARY KEY, valor TEXT)"
        )
        self._migrar_quitar_curso()
        self._normalizar_nombres_existentes()
        self.con.commit()

    def _migrar_quitar_curso(self):
        """Bases de versiones antiguas tenían 'curso' en el alumno."""
        cols = [r["name"] for r in self.con.execute("PRAGMA table_info(alumnos)")]
        if "curso" not in cols:
            return
        self.con.executescript(
            """CREATE TABLE alumnos_nueva (
                   rut TEXT PRIMARY KEY, nombre TEXT NOT NULL, segundo_nombre TEXT,
                   apellido_paterno TEXT NOT NULL, apellido_materno TEXT, observaciones TEXT);
               INSERT INTO alumnos_nueva
                   SELECT rut, nombre, segundo_nombre, apellido_paterno, apellido_materno, observaciones
                   FROM alumnos;
               DROP TABLE alumnos;
               ALTER TABLE alumnos_nueva RENAME TO alumnos;"""
        )

    def _normalizar_nombres_existentes(self):
        """Pasa a minúsculas los nombres guardados por versiones anteriores."""
        for r in self.con.execute("SELECT * FROM alumnos").fetchall():
            nuevos = {c: normalizar_nombre(r[c]) for c in CAMPOS_NOMBRE}
            if any(nuevos[c] != (r[c] or "") for c in CAMPOS_NOMBRE):
                self.con.execute(
                    """UPDATE alumnos SET nombre=:nombre, segundo_nombre=:segundo_nombre,
                       apellido_paterno=:apellido_paterno, apellido_materno=:apellido_materno
                       WHERE rut=:rut""",
                    {**nuevos, "rut": r["rut"]},
                )

    def carpeta_informes(self):
        r = self.con.execute("SELECT valor FROM config WHERE clave='carpeta'").fetchone()
        return r["valor"] if r else CARPETA_INFORMES_DEFECTO

    def guardar_carpeta_informes(self, ruta):
        self.con.execute("INSERT OR REPLACE INTO config VALUES ('carpeta', ?)", (ruta,))
        self.con.commit()

    def listar(self):
        filas = self.con.execute("SELECT * FROM alumnos").fetchall()
        return sorted(
            filas,
            key=lambda a: tuple(
                sin_tildes(a[c] or "") for c in ("apellido_paterno", "apellido_materno", "nombre")
            ),
        )

    def obtener(self, rut):
        return self.con.execute("SELECT * FROM alumnos WHERE rut = ?", (rut,)).fetchone()

    def agregar(self, d):
        self.con.execute(
            """INSERT INTO alumnos VALUES
               (:rut, :nombre, :segundo_nombre, :apellido_paterno,
                :apellido_materno, :observaciones)""",
            d,
        )
        self.con.commit()

    def actualizar(self, d):
        self.con.execute(
            """UPDATE alumnos SET nombre=:nombre, segundo_nombre=:segundo_nombre,
               apellido_paterno=:apellido_paterno, apellido_materno=:apellido_materno,
               observaciones=:observaciones WHERE rut=:rut""",
            d,
        )
        self.con.commit()

    def eliminar(self, rut):
        self.con.execute("DELETE FROM alumnos WHERE rut = ?", (rut,))
        self.con.commit()


# --------------------------------------------------------------------------
# Rutas de informes y apertura de archivos
# --------------------------------------------------------------------------
def _seguro(texto):
    """Nombre válido para carpetas/archivos en Windows, macOS y Linux."""
    return re.sub(r'[<>:"/\\|?*]', "_", texto).strip(" .") or "_"


def ruta_informe(base, fecha, curso, alumno):
    """<base>/<año>/<curso>/<alumno>/Bitacora_AAAA-MM-DD[_n].pdf (crea carpetas)."""
    nombre = " ".join(
        titulo(alumno[c])
        for c in ("apellido_paterno", "apellido_materno", "nombre", "segundo_nombre")
        if alumno[c]
    )
    carpeta = os.path.join(
        base, str(fecha.year), _seguro(curso), _seguro(f"{nombre} - {alumno['rut']}")
    )
    os.makedirs(carpeta, exist_ok=True)
    ruta = os.path.join(carpeta, f"Bitacora_{fecha:%Y-%m-%d}.pdf")
    n = 2
    while os.path.exists(ruta):  # varias sesiones el mismo día: no se sobrescribe
        ruta = os.path.join(carpeta, f"Bitacora_{fecha:%Y-%m-%d}_{n}.pdf")
        n += 1
    return ruta


def abrir_archivo(ruta):
    """Abre el archivo con el programa predeterminado del sistema."""
    if sys.platform.startswith("win"):
        os.startfile(ruta)  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.Popen(["open", ruta])
    else:
        subprocess.Popen(["xdg-open", ruta])


# --------------------------------------------------------------------------
# Generación del PDF (basado en la plantilla "Bitácora Apoyo Pedagógico")
# --------------------------------------------------------------------------
def generar_pdf(ruta, fecha, curso, alumno, objetivo, actividades, recursos, observaciones):
    estilos = getSampleStyleSheet()
    titulo_st = ParagraphStyle("t", parent=estilos["Title"], fontSize=18, spaceAfter=6)
    fecha_st = ParagraphStyle("f", parent=estilos["Normal"], alignment=2, fontSize=11)
    etiqueta = ParagraphStyle("e", parent=estilos["Normal"], fontName="Helvetica-Bold", fontSize=11)
    texto = ParagraphStyle("x", parent=estilos["Normal"], fontSize=11, leading=15)

    def p(t):
        return Paragraph(escape(t or "").replace("\n", "<br/>"), texto)

    est = Paragraph(
        escape(nombre_completo(alumno))
        + f"<br/>RUT: {escape(alumno['rut'])} &nbsp;|&nbsp; Curso: {escape(curso)}",
        texto,
    )

    filas = [
        [Paragraph("Estudiante", etiqueta), est],
        [Paragraph("Objetivo de<br/>la sesión", etiqueta), p(objetivo)],
        [Paragraph("Actividades<br/>Desarrolladas", etiqueta), p(actividades)],
        [Paragraph("Recursos utilizados", etiqueta), p(recursos)],
        [Paragraph("Observaciones:", etiqueta), p(observaciones)],
    ]
    tabla = Table(filas, colWidths=[4 * cm, 13.5 * cm])
    tabla.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.8, colors.black),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#EEEEEE")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )

    doc = SimpleDocTemplate(
        ruta, pagesize=letter, leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=2 * cm, title="Bitácora Apoyo Pedagógico",
    )
    doc.build(
        [
            Paragraph(escape(fecha), fecha_st),
            Spacer(1, 10),
            Paragraph("Bitácora Apoyo Pedagógico", titulo_st),
            Spacer(1, 14),
            tabla,
        ]
    )


# --------------------------------------------------------------------------
# Componentes de interfaz reutilizables
# --------------------------------------------------------------------------
def tarjeta(master):
    """Panel blanco con borde fino; devuelve el marco interior donde colocar el contenido."""
    marco = tk.Frame(master, bg=CARD, highlightthickness=1,
                     highlightbackground=BORDER, highlightcolor=BORDER)
    marco.pack(fill="both", expand=True)
    interno = ttk.Frame(marco, style="Card.TFrame", padding=(32, 26))
    interno.pack(fill="both", expand=True)
    return interno


def crear_texto(master, alto):
    """Campo de texto multilínea con el mismo aspecto que los demás campos."""
    t = tk.Text(
        master, height=alto, wrap="word", font=(FUENTE, 10), relief="flat",
        bg=CARD, fg=TEXT, insertbackground=TEXT, padx=8, pady=6, undo=True,
        highlightthickness=1, highlightbackground=BORDER, highlightcolor=PRIMARY,
    )

    def siguiente(e):
        e.widget.tk_focusNext().focus_set()
        return "break"

    def anterior(e):
        e.widget.tk_focusPrev().focus_set()
        return "break"

    t.bind("<Tab>", siguiente)
    for secuencia in ("<Shift-Tab>", "<ISO_Left_Tab>"):
        try:
            t.bind(secuencia, anterior)
        except tk.TclError:
            pass
    return t


class Tile(tk.Frame):
    """Botón grande tipo tarjeta para el menú principal."""

    def __init__(self, master, titulo_, descripcion, comando, acento=PRIMARY):
        super().__init__(master, bg=CARD, highlightthickness=1,
                         highlightbackground=BORDER, cursor="hand2")
        self.acento = acento
        self.comando = comando
        self.lt = tk.Label(self, text=titulo_, bg=CARD, fg=TEXT, anchor="w",
                           font=(FUENTE, 14, "bold"), cursor="hand2")
        self.lt.pack(fill="x", padx=22, pady=(20, 2))
        self.ld = tk.Label(self, text=descripcion, bg=CARD, fg=MUTED, anchor="w",
                           justify="left", wraplength=300, font=(FUENTE, 10), cursor="hand2")
        self.ld.pack(fill="x", padx=22, pady=(0, 20))
        for w in (self, self.lt, self.ld):
            w.bind("<Enter>", self._entrar)
            w.bind("<Leave>", self._salir)
            w.bind("<Button-1>", lambda e: self.comando())

    def _pintar(self, fondo, borde, color_titulo):
        self.configure(bg=fondo, highlightbackground=borde)
        self.lt.configure(bg=fondo, fg=color_titulo)
        self.ld.configure(bg=fondo)

    def _entrar(self, _e=None):
        self._pintar("#F8FAFF" if self.acento == PRIMARY else "#FFF8F8", self.acento, self.acento)

    def _salir(self, _e=None):
        self._pintar(CARD, BORDER, TEXT)


# --------------------------------------------------------------------------
# Aplicación
# --------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bitácora Apoyo Pedagógico")
        self.geometry("920x760")
        self.minsize(800, 680)
        self.configure(bg=BG)
        self.db = BaseDatos()
        self._configurar_estilos()

        cabecera = tk.Frame(self, bg=PRIMARY, height=60)
        cabecera.pack(fill="x")
        cabecera.pack_propagate(False)
        tk.Label(cabecera, text="Bitácora Apoyo Pedagógico", bg=PRIMARY, fg="white",
                 font=(FUENTE, 16, "bold")).pack(side="left", padx=28)

        self.contenedor = ttk.Frame(self, padding=24)
        self.contenedor.pack(fill="both", expand=True)
        self.mostrar(PantallaMenu)

    def _configurar_estilos(self):
        global FUENTE
        disponibles = set(tkfont.families())
        FUENTE = next(
            (f for f in ("Segoe UI", "SF Pro Text", "Helvetica Neue", "Inter",
                         "Ubuntu", "Noto Sans", "DejaVu Sans") if f in disponibles),
            tkfont.nametofont("TkDefaultFont").actual("family"),
        )
        for nombre in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            tkfont.nametofont(nombre).configure(family=FUENTE, size=10)
        self.option_add("*TCombobox*Listbox.font", (FUENTE, 10))
        self.option_add("*TCombobox*Listbox.background", CARD)
        self.option_add("*TCombobox*Listbox.foreground", TEXT)
        self.option_add("*TCombobox*Listbox.selectBackground", PRIMARY)
        self.option_add("*TCombobox*Listbox.selectForeground", "white")

        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".", background=BG, foreground=TEXT, font=(FUENTE, 10))

        # Marcos y etiquetas
        s.configure("TFrame", background=BG)
        s.configure("Card.TFrame", background=CARD)
        s.configure("TLabel", background=BG)
        s.configure("Page.TLabel", background=BG, font=(FUENTE, 20, "bold"))
        s.configure("MutedBG.TLabel", background=BG, foreground=MUTED)
        s.configure("H1.TLabel", background=CARD, font=(FUENTE, 18, "bold"))
        s.configure("Muted.TLabel", background=CARD, foreground=MUTED)
        s.configure("Field.TLabel", background=CARD, foreground="#374151", font=(FUENTE, 9, "bold"))
        s.configure("Hint.TLabel", background=CARD, foreground=MUTED, font=(FUENTE, 9))

        # Botones
        for nombre, bg, hover, pres, fg, borde, borde_h in [
            ("Primary.TButton", PRIMARY, PRIMARY_H, "#1E40AF", "white", PRIMARY, PRIMARY_H),
            ("Danger.TButton", DANGER, DANGER_H, "#991B1B", "white", DANGER, DANGER_H),
            ("Secondary.TButton", CARD, "#F3F4F6", "#E5E7EB", TEXT, BORDER, "#9CA3AF"),
        ]:
            s.configure(nombre, background=bg, foreground=fg, bordercolor=borde,
                        lightcolor=bg, darkcolor=bg, focuscolor=bg, relief="flat",
                        borderwidth=1, padding=(20, 10), font=(FUENTE, 10, "bold"))
            s.map(nombre,
                  background=[("pressed", pres), ("active", hover)],
                  lightcolor=[("pressed", pres), ("active", hover)],
                  darkcolor=[("pressed", pres), ("active", hover)],
                  bordercolor=[("active", borde_h)],
                  foreground=[("disabled", MUTED)])

        # Campos
        campo = dict(fieldbackground=CARD, foreground=TEXT, bordercolor=BORDER,
                     lightcolor=BORDER, darkcolor=BORDER, insertcolor=TEXT, padding=(8, 7))
        s.configure("TEntry", **campo)
        s.map("TEntry",
              bordercolor=[("focus", PRIMARY)], lightcolor=[("focus", PRIMARY)],
              darkcolor=[("focus", PRIMARY)],
              fieldbackground=[("disabled", "#F3F4F6")], foreground=[("disabled", MUTED)])
        s.configure("TCombobox", background=CARD, arrowcolor=MUTED, arrowsize=16, **campo)
        s.map("TCombobox",
              fieldbackground=[("readonly", CARD)],
              selectbackground=[("readonly", CARD)], selectforeground=[("readonly", TEXT)],
              bordercolor=[("focus", PRIMARY)], lightcolor=[("focus", PRIMARY)],
              darkcolor=[("focus", PRIMARY)], background=[("active", "#F3F4F6")])

        # Tabla
        s.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        s.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT,
                    borderwidth=0, rowheight=36, font=(FUENTE, 10))
        s.map("Treeview", background=[("selected", PRIMARY)], foreground=[("selected", "white")])
        s.configure("Treeview.Heading", background="#F3F4F6", foreground=MUTED, relief="flat",
                    padding=(10, 9), font=(FUENTE, 9, "bold"))
        s.map("Treeview.Heading", background=[("active", "#F3F4F6")])
        s.configure("Vertical.TScrollbar", background="#D1D5DB", troughcolor="#F3F4F6",
                    bordercolor="#F3F4F6", lightcolor="#D1D5DB", darkcolor="#D1D5DB",
                    arrowcolor=MUTED, relief="flat")
        s.map("Vertical.TScrollbar", background=[("active", "#9CA3AF")])

    def mostrar(self, clase, **kwargs):
        for w in self.contenedor.winfo_children():
            w.destroy()
        clase(self.contenedor, self, **kwargs).pack(fill="both", expand=True)


# --------------------------------------------------------------------------
# Pantallas
# --------------------------------------------------------------------------
class PantallaMenu(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        ttk.Label(self, text="Inicio", style="Page.TLabel").pack(anchor="w")
        ttk.Label(self, text="¿Qué desea hacer?", style="MutedBG.TLabel").pack(anchor="w", pady=(2, 18))

        grilla = ttk.Frame(self)
        grilla.pack(fill="x")
        for c in (0, 1):
            grilla.columnconfigure(c, weight=1, uniform="tile")
        tiles = [
            ("Generar informe", "Complete la bitácora de una sesión y expórtela a PDF.",
             lambda: app.mostrar(PantallaListaAlumnos, modo="informe"), PRIMARY),
            ("Agregar alumno", "Registre un nuevo estudiante.",
             lambda: app.mostrar(PantallaFormAlumno), PRIMARY),
            ("Editar alumno", "Modifique los datos de un estudiante registrado.",
             lambda: app.mostrar(PantallaListaAlumnos, modo="editar"), PRIMARY),
            ("Eliminar alumno", "Quite a un estudiante del registro.",
             lambda: app.mostrar(PantallaListaAlumnos, modo="eliminar"), DANGER),
        ]
        for i, (t, d, cmd, acento) in enumerate(tiles):
            Tile(grilla, t, d, cmd, acento).grid(row=i // 2, column=i % 2,
                                                  sticky="nsew", padx=8, pady=8)

        pie = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        pie.pack(fill="x", padx=8, pady=(22, 0))
        interno = ttk.Frame(pie, style="Card.TFrame", padding=(20, 14))
        interno.pack(fill="x")
        izq = ttk.Frame(interno, style="Card.TFrame")
        izq.pack(side="left", fill="x", expand=True)
        ttk.Label(izq, text="CARPETA GENERAL DE INFORMES", style="Field.TLabel").pack(anchor="w")
        self.ruta = tk.StringVar(value=app.db.carpeta_informes())
        ttk.Label(izq, textvariable=self.ruta, style="Muted.TLabel", wraplength=560).pack(anchor="w", pady=(3, 0))
        ttk.Button(interno, text="Cambiar carpeta…", style="Secondary.TButton",
                   command=self.cambiar_carpeta).pack(side="right")

    def cambiar_carpeta(self):
        ruta = filedialog.askdirectory(initialdir=self.app.db.carpeta_informes(),
                                       title="Carpeta general de informes")
        if ruta:
            self.app.db.guardar_carpeta_informes(os.path.normpath(ruta))
            self.ruta.set(self.app.db.carpeta_informes())


class PantallaListaAlumnos(ttk.Frame):
    """Listado de alumnos para generar informe, editar o eliminar."""

    MODOS = {
        "informe": ("Generar informe", "Seleccione al estudiante para quien desea generar la bitácora.",
                    "Seleccionar", "Primary.TButton"),
        "editar": ("Editar alumno", "Seleccione al alumno que desea modificar.",
                   "Editar", "Primary.TButton"),
        "eliminar": ("Eliminar alumno", "Seleccione al alumno que desea eliminar del registro.",
                     "Eliminar", "Danger.TButton"),
    }

    def __init__(self, master, app, modo):
        super().__init__(master)
        self.app, self.modo = app, modo
        titulo_, sub, texto_boton, estilo_boton = self.MODOS[modo]
        tarj = tarjeta(self)

        # Pie (se empaqueta primero para que siempre quede visible abajo)
        pie = ttk.Frame(tarj, style="Card.TFrame")
        pie.pack(side="bottom", fill="x", pady=(16, 0))
        ttk.Button(pie, text="← Volver", style="Secondary.TButton",
                   command=lambda: app.mostrar(PantallaMenu)).pack(side="left")
        ttk.Button(pie, text=texto_boton, style=estilo_boton,
                   command=self.seleccionar).pack(side="right")
        if modo == "informe":
            ttk.Button(pie, text="+ Agregar otro estudiante", style="Secondary.TButton",
                       command=lambda: app.mostrar(PantallaFormAlumno, volver_a="informe")
                       ).pack(side="right", padx=10)

        self.lbl_cuenta = ttk.Label(tarj, text="", style="Hint.TLabel")
        self.lbl_cuenta.pack(side="bottom", anchor="w", pady=(8, 0))

        ttk.Label(tarj, text=titulo_, style="H1.TLabel").pack(anchor="w")
        ttk.Label(tarj, text=sub, style="Muted.TLabel").pack(anchor="w", pady=(2, 16))
        ttk.Label(tarj, text="BUSCAR POR NOMBRE O RUT", style="Field.TLabel").pack(anchor="w")
        self.busqueda = tk.StringVar()
        self.busqueda.trace_add("write", self._cargar)
        buscador = ttk.Entry(tarj, textvariable=self.busqueda)
        buscador.pack(fill="x", pady=(4, 12))

        marco = tk.Frame(tarj, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        marco.pack(fill="both", expand=True)
        self.tabla = ttk.Treeview(marco, columns=("rut", "nombre"), show="headings",
                                  selectmode="browse")
        self.tabla.heading("rut", text="RUT", anchor="w")
        self.tabla.heading("nombre", text="NOMBRE", anchor="w")
        self.tabla.column("rut", width=150, stretch=False, anchor="w")
        self.tabla.column("nombre", width=500, anchor="w")
        self.tabla.tag_configure("impar", background="#F9FAFB")
        barra = ttk.Scrollbar(marco, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=barra.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        self.tabla.bind("<Double-1>", lambda e: self.seleccionar())
        self.tabla.bind("<Return>", lambda e: self.seleccionar())

        self.alumnos = app.db.listar()
        self._cargar()
        buscador.focus_set()

    def _cargar(self, *_):
        q = self.busqueda.get().strip().lower()
        q_rut = re.sub(r"[.\-\s]", "", q)
        self.tabla.delete(*self.tabla.get_children())
        n = 0
        for a in self.alumnos:
            nombre = nombre_completo(a)
            rut_plano = re.sub(r"[.\-]", "", a["rut"].lower())
            if q and q not in nombre.lower() and not (q_rut and q_rut in rut_plano):
                continue
            self.tabla.insert("", "end", iid=a["rut"], values=(a["rut"], nombre),
                              tags=("impar",) if n % 2 else ())
            n += 1
        total = len(self.alumnos)
        if total == 0:
            self.lbl_cuenta.configure(text="No hay alumnos registrados todavía.")
        elif q:
            self.lbl_cuenta.configure(text=f"{n} de {total} alumnos")
        else:
            self.lbl_cuenta.configure(text=f"{total} alumno" + ("" if total == 1 else "s"))

    def seleccionar(self):
        sel = self.tabla.selection()
        if not sel:
            messagebox.showinfo("Selección", "Seleccione un estudiante de la lista.")
            return
        rut = sel[0]
        if self.modo == "informe":
            self.app.mostrar(PantallaInforme, rut=rut)
        elif self.modo == "editar":
            self.app.mostrar(PantallaFormAlumno, rut=rut)
        else:
            self._eliminar(rut)

    def _eliminar(self, rut):
        a = self.app.db.obtener(rut)
        if messagebox.askyesno(
            "Eliminar alumno",
            f"¿Desea eliminar a {nombre_completo(a)} (RUT {rut})?\n\n"
            "Esta acción no se puede deshacer. Los informes PDF ya generados "
            "no se borrarán.",
            icon="warning",
        ):
            self.app.db.eliminar(rut)
            self.alumnos = self.app.db.listar()
            self._cargar()


class PantallaFormAlumno(ttk.Frame):
    """Agregar alumno (rut=None) o editar alumno (RUT bloqueado)."""

    def __init__(self, master, app, rut=None, volver_a=None):
        super().__init__(master)
        self.app, self.rut_original, self.volver_a = app, rut, volver_a
        editando = rut is not None
        tarj = tarjeta(self)

        pie = ttk.Frame(tarj, style="Card.TFrame")
        pie.pack(side="bottom", fill="x", pady=(18, 0))
        ttk.Button(pie, text="Cancelar", style="Secondary.TButton",
                   command=self._destino).pack(side="left")
        ttk.Button(pie, text="Guardar cambios", style="Primary.TButton",
                   command=self.guardar).pack(side="right")

        ttk.Label(tarj, text="Editar alumno" if editando else "Agregar alumno",
                  style="H1.TLabel").pack(anchor="w")
        ttk.Label(tarj, text="Los campos marcados con * son obligatorios.",
                  style="Muted.TLabel").pack(anchor="w", pady=(2, 18))

        form = ttk.Frame(tarj, style="Card.TFrame")
        form.pack(fill="both", expand=True)
        for c in (0, 1):
            form.columnconfigure(c, weight=1, uniform="col")
        form.rowconfigure(3, weight=1)

        self.vars = {}
        self.entradas = {}

        def campo(fila, col, clave, etiqueta, ayuda=None):
            sub = ttk.Frame(form, style="Card.TFrame")
            sub.grid(row=fila, column=col, sticky="ew", pady=(0, 16),
                     padx=(0, 12) if col == 0 else (12, 0))
            ttk.Label(sub, text=etiqueta, style="Field.TLabel").pack(anchor="w")
            v = tk.StringVar()
            e = ttk.Entry(sub, textvariable=v)
            e.pack(fill="x", pady=(4, 0))
            if ayuda:
                ttk.Label(sub, text=ayuda, style="Hint.TLabel").pack(anchor="w", pady=(3, 0))
            self.vars[clave], self.entradas[clave] = v, e
            return v, e

        # RUT con autoformato
        ayuda_rut = "El RUT no se puede modificar." if editando else \
            "Escriba solo los dígitos; el formato se aplica solo."
        v_rut, e_rut = campo(0, 0, "rut", "RUT *", ayuda_rut)
        self._formateando = False
        v_rut.trace_add("write", lambda *_: self._autoformato_rut(v_rut, e_rut))
        if editando:
            e_rut.state(["disabled"])

        campo(1, 0, "nombre", "NOMBRE *")
        campo(1, 1, "segundo_nombre", "SEGUNDO NOMBRE (OPCIONAL)")
        campo(2, 0, "apellido_paterno", "APELLIDO PATERNO *")
        campo(2, 1, "apellido_materno", "APELLIDO MATERNO")
        for clave in CAMPOS_NOMBRE:
            self.entradas[clave].bind(
                "<FocusOut>", lambda e, v=self.vars[clave]: v.set(titulo(v.get()))
            )

        sub = ttk.Frame(form, style="Card.TFrame")
        sub.grid(row=3, column=0, columnspan=2, sticky="nsew")
        ttk.Label(sub, text="OBSERVACIONES", style="Field.TLabel").pack(anchor="w")
        self.obs = crear_texto(sub, 5)
        self.obs.pack(fill="both", expand=True, pady=(4, 0))

        if editando:
            a = app.db.obtener(rut)
            v_rut.set(a["rut"])
            for clave in CAMPOS_NOMBRE:
                self.vars[clave].set(titulo(a[clave]))
            self.obs.insert("1.0", a["observaciones"] or "")
            self.entradas["nombre"].focus_set()
        else:
            e_rut.focus_set()

    @staticmethod
    def _autoformato_rut_valor(v, e, bandera):
        pass

   
    def _autoformato_rut(self, var, entrada):
        """Formatea el RUT y deja el cursor al final después de actualizar el campo."""
        if self._formateando:
            return

        actual = var.get()
        nuevo = formatear_rut(actual)

        if nuevo != actual:
            self._formateando = True
            try:
                var.set(nuevo)
            finally:
                self._formateando = False

        # Mover el cursor cuando Tkinter termine de actualizar el campo.
        entrada.after_idle(
            lambda e=entrada: e.icursor(tk.END) if e.winfo_exists() else None
        )


    def _destino(self):
        if self.volver_a == "informe":
            self.app.mostrar(PantallaListaAlumnos, modo="informe")
        elif self.rut_original:
            self.app.mostrar(PantallaListaAlumnos, modo="editar")
        else:
            self.app.mostrar(PantallaMenu)

    def guardar(self):
        datos = {c: normalizar_nombre(self.vars[c].get()) for c in CAMPOS_NOMBRE}
        datos["observaciones"] = self.obs.get("1.0", "end").strip()

        if self.rut_original:
            datos["rut"] = self.rut_original
        else:
            rut = normalizar_rut(self.vars["rut"].get())
            if not rut:
                messagebox.showerror("RUT inválido", "Ingrese un RUT válido (ej: 12.345.678-5).")
                return
            if self.app.db.obtener(rut):
                messagebox.showerror("RUT duplicado", "Ya existe un alumno con ese RUT.")
                return
            datos["rut"] = rut

        for clave, etiqueta, obligatorio in [
            ("nombre", "Nombre", True),
            ("segundo_nombre", "Segundo nombre", False),
            ("apellido_paterno", "Apellido paterno", True),
            ("apellido_materno", "Apellido materno", False),
        ]:
            if not datos[clave]:
                if obligatorio:
                    messagebox.showerror("Campo requerido", f"El campo «{etiqueta}» es obligatorio.")
                    return
            elif not nombre_valido(datos[clave]):
                messagebox.showerror(
                    "Dato inválido",
                    f"El campo «{etiqueta}» solo puede contener letras, espacios, guiones o apóstrofes.",
                )
                return

        if self.rut_original:
            self.app.db.actualizar(datos)
        else:
            self.app.db.agregar(datos)
        messagebox.showinfo("Guardado", "Los cambios se guardaron correctamente.")
        self._destino()


class PantallaInforme(ttk.Frame):
    def __init__(self, master, app, rut):
        super().__init__(master)
        self.app = app
        self.alumno = app.db.obtener(rut)
        tarj = tarjeta(self)

        pie = ttk.Frame(tarj, style="Card.TFrame")
        pie.pack(side="bottom", fill="x", pady=(16, 0))
        ttk.Button(pie, text="← Volver", style="Secondary.TButton",
                   command=lambda: app.mostrar(PantallaListaAlumnos, modo="informe")
                   ).pack(side="left")
        ttk.Button(pie, text="Generar PDF", style="Primary.TButton",
                   command=self.generar).pack(side="right")

        ttk.Label(tarj, text="Bitácora de la sesión", style="H1.TLabel").pack(anchor="w", pady=(0, 12))
        chip_bg = "#EFF6FF"
        chip = tk.Frame(tarj, bg=chip_bg, highlightthickness=1, highlightbackground="#BFDBFE")
        chip.pack(fill="x", pady=(0, 16))
        tk.Label(chip, text=nombre_completo(self.alumno), bg=chip_bg, fg=TEXT,
                 font=(FUENTE, 12, "bold")).pack(side="left", padx=(16, 12), pady=10)
        tk.Label(chip, text=f"RUT {rut}", bg=chip_bg, fg=MUTED,
                 font=(FUENTE, 10)).pack(side="left", pady=10)

        form = ttk.Frame(tarj, style="Card.TFrame")
        form.pack(fill="both", expand=True)
        for c in (0, 1):
            form.columnconfigure(c, weight=1, uniform="col")

        sub = ttk.Frame(form, style="Card.TFrame")
        sub.grid(row=0, column=0, sticky="ew", pady=(0, 12), padx=(0, 12))
        ttk.Label(sub, text="FECHA (DD/MM/AAAA)", style="Field.TLabel").pack(anchor="w")
        self.fecha = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        ttk.Entry(sub, textvariable=self.fecha).pack(fill="x", pady=(4, 0))

        sub = ttk.Frame(form, style="Card.TFrame")
        sub.grid(row=0, column=1, sticky="ew", pady=(0, 12), padx=(12, 0))
        ttk.Label(sub, text="CURSO *", style="Field.TLabel").pack(anchor="w")
        self.curso = tk.StringVar()
        combo = ttk.Combobox(sub, textvariable=self.curso, values=CURSOS,
                             state="readonly", height=15)
        combo.pack(fill="x", pady=(4, 0))

        self.textos = {}
        for fila, (clave, etiqueta, alto, peso) in enumerate(
            [
                ("objetivo", "OBJETIVO DE LA SESIÓN", 2, 2),
                ("actividades", "ACTIVIDADES DESARROLLADAS", 4, 4),
                ("recursos", "RECURSOS UTILIZADOS", 2, 2),
                ("observaciones", "OBSERVACIONES", 3, 3),
            ],
            start=1,
        ):
            sub = ttk.Frame(form, style="Card.TFrame")
            sub.grid(row=fila, column=0, columnspan=2, sticky="nsew", pady=(0, 12))
            ttk.Label(sub, text=etiqueta, style="Field.TLabel").pack(anchor="w")
            t = crear_texto(sub, alto)
            t.pack(fill="both", expand=True, pady=(4, 0))
            form.rowconfigure(fila, weight=peso)
            self.textos[clave] = t
        combo.focus_set()

    def generar(self):
        try:
            fecha = datetime.strptime(self.fecha.get().strip(), "%d/%m/%Y").date()
        except ValueError:
            messagebox.showerror("Fecha inválida", "Ingrese la fecha con el formato DD/MM/AAAA.")
            return
        curso = self.curso.get()
        if not curso:
            messagebox.showerror("Curso", "Seleccione un curso de la lista.")
            return
        valores = {k: t.get("1.0", "end").strip() for k, t in self.textos.items()}
        if not any(valores.values()):
            messagebox.showwarning("Informe vacío", "Complete al menos uno de los campos.")
            return
        try:
            ruta = ruta_informe(self.app.db.carpeta_informes(), fecha, curso, self.alumno)
            generar_pdf(ruta, fecha.strftime("%d/%m/%Y"), curso, self.alumno, **valores)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Error", f"No se pudo generar el PDF:\n{e}")
            return
        self.app.mostrar(PantallaMenu)
        try:
            abrir_archivo(ruta)
        except Exception as e:  # noqa: BLE001
            messagebox.showwarning(
                "Informe generado",
                f"El informe se guardó en:\n{ruta}\n\nNo se pudo abrir automáticamente:\n{e}",
            )


if __name__ == "__main__":
    App().mainloop()