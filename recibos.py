# -*- coding: utf-8 -*-
"""
Recibos de Pago - envia a cada trabajador su recibo por correo.

Como funciona:
  1. Se descarga el RAR con todos los recibos y se extrae en una carpeta.
  2. Esta carpeta queda guardada en el programa.
  3. El programa busca, para cada trabajador, el PDF que lleve su cedula
     en el nombre (ej: Reg_35-2026_T-12345678.pdf).
  4. Se marcan los trabajadores y se le da a "Enviar correos" o "Imprimir".

Solo usa la libreria estandar de Python. No hay que instalar nada.
"""

import csv
import json
import os
import queue
import re
import smtplib
import ssl
import subprocess
import sys
import threading
import time
import tkinter as tk
import zipfile
from email.message import EmailMessage
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

# ---------------------------------------------------------------- archivos

if getattr(sys, "frozen", False):
    BASE = Path(sys.executable).resolve().parent
else:
    BASE = Path(__file__).resolve().parent

ARCHIVO_CONFIG = BASE / "config.json"
ARCHIVO_EMPLEADOS = BASE / "empleados.csv"
ARCHIVO_LOG = BASE / "envios.log"

PROVEEDORES = {
    "Gmail": ("smtp.gmail.com", 587),
    "Outlook / Hotmail": ("smtp-mail.outlook.com", 587),
    "Yahoo": ("smtp.mail.yahoo.com", 587),
    "Otro (escribir a mano)": ("", 587),
}

CONFIG_POR_DEFECTO = {
    "carpeta": "",
    "proveedor": "Gmail",
    "servidor": "smtp.gmail.com",
    "puerto": 587,
    "remitente": "",
    "password": "",
    "recordar_password": False,
    "asunto": "Recibo de pago",
    "cuerpo": (
        "Hola {nombre},\n"
        "\n"
        "Le enviamos adjunto su recibo de pago.\n"
        "\n"
        "Saludos."
    ),
}

# Programas que pueden descomprimir un .rar en Windows.
DESCOMPRESORES = [
    r"C:\Program Files\WinRAR\UnRAR.exe",
    r"C:\Program Files (x86)\WinRAR\UnRAR.exe",
    r"C:\Program Files\WinRAR\WinRAR.exe",
    r"C:\Program Files (x86)\WinRAR\WinRAR.exe",
    r"C:\Program Files\7-Zip\7z.exe",
    r"C:\Program Files (x86)\7-Zip\7z.exe",
]


def cargar_config():
    datos = dict(CONFIG_POR_DEFECTO)
    if ARCHIVO_CONFIG.exists():
        try:
            with open(ARCHIVO_CONFIG, "r", encoding="utf-8") as f:
                datos.update(json.load(f))
        except Exception:
            pass
    if not datos.get("recordar_password"):
        datos["password"] = ""
    return datos


def guardar_config(datos):
    copia = dict(datos)
    if not copia.get("recordar_password"):
        copia["password"] = ""
    try:
        with open(ARCHIVO_CONFIG, "w", encoding="utf-8") as f:
            json.dump(copia, f, indent=2, ensure_ascii=False)
    except Exception as e:
        messagebox.showerror("Recibos", "No se pudo guardar la configuracion:\n%s" % e)


def cargar_empleados():
    lista = []
    if not ARCHIVO_EMPLEADOS.exists():
        return lista
    try:
        with open(ARCHIVO_EMPLEADOS, "r", encoding="utf-8-sig", newline="") as f:
            for fila in csv.DictReader(f):
                nombre = (fila.get("nombre") or "").strip()
                cedula = (fila.get("cedula") or "").strip()
                correo = (fila.get("correo") or "").strip()
                if nombre or cedula or correo:
                    lista.append({"nombre": nombre, "cedula": cedula,
                                  "correo": correo, "marcado": False, "pdfs": []})
    except Exception as e:
        messagebox.showerror("Recibos", "No se pudo leer empleados.csv:\n%s" % e)
    return lista


def guardar_empleados(lista):
    try:
        with open(ARCHIVO_EMPLEADOS, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["nombre", "cedula", "correo"])
            for e in lista:
                w.writerow([e["nombre"], e["cedula"], e["correo"]])
    except Exception as e:
        messagebox.showerror("Recibos", "No se pudo guardar empleados.csv:\n%s" % e)


# ---------------------------------------------------------------- recibos

def solo_digitos(texto):
    return re.sub(r"\D", "", texto or "")


def indexar_recibos(carpeta):
    """Devuelve {numero_de_cedula: [rutas de pdf]} a partir de los nombres."""
    indice = {}
    if not carpeta:
        return indice
    ruta = Path(carpeta)
    if not ruta.is_dir():
        return indice
    for pdf in sorted(ruta.rglob("*.pdf")):
        # Cualquier grupo de 5 digitos o mas del nombre se toma como cedula.
        for numero in re.findall(r"\d{5,}", pdf.stem):
            indice.setdefault(numero, []).append(pdf)
    return indice


def emparejar(lista, indice):
    for e in lista:
        cedula = solo_digitos(e["cedula"])
        e["pdfs"] = list(indice.get(cedula, [])) if cedula else []


def buscar_descompresor():
    for ruta in DESCOMPRESORES:
        if os.path.isfile(ruta):
            return ruta
    return None


def extraer(archivo, destino):
    """Extrae un .rar/.zip/.7z en destino. Devuelve texto de error o None."""
    destino = str(destino)
    if archivo.lower().endswith(".zip"):
        try:
            with zipfile.ZipFile(archivo) as z:
                z.extractall(destino)
            return None
        except Exception as e:
            return str(e)

    programa = buscar_descompresor()
    if not programa:
        return ("No se encontro WinRAR ni 7-Zip en esta computadora.\n\n"
                "Instala uno de los dos, o extrae el RAR a mano (clic derecho "
                "sobre el archivo > Extraer aqui) y luego elige esa carpeta.")

    nombre = os.path.basename(programa).lower()
    if nombre.startswith("7z"):
        cmd = [programa, "x", archivo, "-o" + destino, "-y"]
    else:  # UnRAR.exe / WinRAR.exe
        cmd = [programa, "x", "-o+", archivo, destino + os.sep]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    except Exception as e:
        return str(e)
    if r.returncode != 0:
        return (r.stderr or r.stdout or "codigo %s" % r.returncode).strip()
    return None


# ---------------------------------------------------------------- ventanita

def pedir_datos(padre, titulo, datos=None):
    """Ventana chiquita para agregar o editar un trabajador."""
    datos = datos or {"nombre": "", "cedula": "", "correo": ""}
    win = tk.Toplevel(padre)
    win.title(titulo)
    win.transient(padre)
    win.resizable(False, False)
    win.grab_set()

    campos = {}
    for i, (clave, etiqueta) in enumerate((("nombre", "Nombre y apellido"),
                                           ("cedula", "Cedula"),
                                           ("correo", "Correo"))):
        ttk.Label(win, text=etiqueta + ":").grid(row=i, column=0, sticky="e",
                                                 padx=(12, 6), pady=6)
        ent = ttk.Entry(win, width=38)
        ent.insert(0, datos.get(clave, ""))
        ent.grid(row=i, column=1, padx=(0, 12), pady=6)
        campos[clave] = ent

    resultado = {}

    def aceptar(*_):
        nombre = campos["nombre"].get().strip()
        cedula = campos["cedula"].get().strip()
        if not nombre or not cedula:
            messagebox.showwarning("Recibos", "El nombre y la cedula son obligatorios.",
                                   parent=win)
            return
        resultado.update({"nombre": nombre, "cedula": cedula,
                          "correo": campos["correo"].get().strip()})
        win.destroy()

    barra = ttk.Frame(win)
    barra.grid(row=3, column=0, columnspan=2, pady=(4, 12))
    ttk.Button(barra, text="Guardar", command=aceptar).pack(side="left", padx=6)
    ttk.Button(barra, text="Cancelar", command=win.destroy).pack(side="left", padx=6)

    campos["nombre"].focus_set()
    win.bind("<Return>", aceptar)
    win.bind("<Escape>", lambda *_: win.destroy())
    padre.wait_window(win)
    return resultado or None


# ---------------------------------------------------------------- programa

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Recibos de Pago")
        self.geometry("980x700")
        self.minsize(860, 600)

        self.config_datos = cargar_config()
        self.empleados = cargar_empleados()
        self.mensajes = queue.Queue()
        self.trabajando = False

        self._construir()
        self.refrescar_recibos(avisar=False)
        self.after(200, self._vaciar_cola)
        self.protocol("WM_DELETE_WINDOW", self.cerrar)

    # ---- interfaz

    def _construir(self):
        ttk.Style().configure("Treeview", rowheight=24)
        raiz = ttk.Frame(self, padding=10)
        raiz.pack(fill="both", expand=True)
        raiz.columnconfigure(0, weight=1)
        raiz.rowconfigure(3, weight=1)

        # --- carpeta
        f1 = ttk.LabelFrame(raiz, text="Carpeta con los recibos (PDF)", padding=8)
        f1.grid(row=0, column=0, sticky="ew")
        f1.columnconfigure(0, weight=1)
        self.var_carpeta = tk.StringVar(value=self.config_datos["carpeta"])
        ttk.Entry(f1, textvariable=self.var_carpeta).grid(row=0, column=0, sticky="ew",
                                                          padx=(0, 6))
        ttk.Button(f1, text="Elegir carpeta...",
                   command=self.elegir_carpeta).grid(row=0, column=1, padx=3)
        ttk.Button(f1, text="Importar RAR...",
                   command=self.importar_rar).grid(row=0, column=2, padx=3)
        ttk.Button(f1, text="Volver a buscar",
                   command=lambda: self.refrescar_recibos(True)).grid(row=0, column=3,
                                                                      padx=3)

        # --- correo de quien envia
        f2 = ttk.LabelFrame(raiz, text="Correo de quien envia", padding=8)
        f2.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        f2.columnconfigure(1, weight=1)
        f2.columnconfigure(4, weight=1)

        self.var_proveedor = tk.StringVar(value=self.config_datos["proveedor"])
        self.var_servidor = tk.StringVar(value=self.config_datos["servidor"])
        self.var_puerto = tk.StringVar(value=str(self.config_datos["puerto"]))
        self.var_remitente = tk.StringVar(value=self.config_datos["remitente"])
        self.var_password = tk.StringVar(value=self.config_datos["password"])
        self.var_recordar = tk.BooleanVar(value=self.config_datos["recordar_password"])

        ttk.Label(f2, text="Correo:").grid(row=0, column=0, sticky="e", padx=(0, 6))
        ttk.Entry(f2, textvariable=self.var_remitente).grid(row=0, column=1,
                                                            sticky="ew", padx=(0, 12))
        ttk.Label(f2, text="Contrasena de aplicacion:").grid(row=0, column=3, sticky="e",
                                                             padx=(0, 6))
        ttk.Entry(f2, textvariable=self.var_password, show="*").grid(row=0, column=4,
                                                                     sticky="ew")

        ttk.Label(f2, text="Servicio:").grid(row=1, column=0, sticky="e",
                                             padx=(0, 6), pady=(6, 0))
        combo = ttk.Combobox(f2, textvariable=self.var_proveedor, state="readonly",
                             values=list(PROVEEDORES), width=22)
        combo.grid(row=1, column=1, sticky="w", pady=(6, 0))
        combo.bind("<<ComboboxSelected>>", self.cambiar_proveedor)

        ttk.Label(f2, text="Servidor:").grid(row=1, column=3, sticky="e",
                                             padx=(0, 6), pady=(6, 0))
        caja = ttk.Frame(f2)
        caja.grid(row=1, column=4, sticky="ew", pady=(6, 0))
        ttk.Entry(caja, textvariable=self.var_servidor, width=24).pack(side="left")
        ttk.Label(caja, text=" puerto:").pack(side="left")
        ttk.Entry(caja, textvariable=self.var_puerto, width=6).pack(side="left")

        ttk.Checkbutton(f2, text="Recordar la contrasena en esta computadora",
                        variable=self.var_recordar).grid(row=2, column=1, columnspan=2,
                                                         sticky="w", pady=(6, 0))
        ttk.Button(f2, text="Probar conexion",
                   command=self.probar_conexion).grid(row=2, column=4, sticky="e",
                                                      pady=(6, 0))

        # --- mensaje
        f3 = ttk.LabelFrame(raiz, text="Mensaje (se puede usar {nombre} y {cedula})",
                            padding=8)
        f3.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        f3.columnconfigure(1, weight=1)
        self.var_asunto = tk.StringVar(value=self.config_datos["asunto"])
        ttk.Label(f3, text="Asunto:").grid(row=0, column=0, sticky="e", padx=(0, 6))
        ttk.Entry(f3, textvariable=self.var_asunto).grid(row=0, column=1, sticky="ew")
        self.txt_cuerpo = tk.Text(f3, height=4, wrap="word")
        self.txt_cuerpo.insert("1.0", self.config_datos["cuerpo"])
        self.txt_cuerpo.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        # --- lista
        f4 = ttk.LabelFrame(raiz, text="Trabajadores", padding=8)
        f4.grid(row=3, column=0, sticky="nsew", pady=(8, 0))
        f4.columnconfigure(0, weight=1)
        f4.rowconfigure(1, weight=1)

        barra = ttk.Frame(f4)
        barra.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        ttk.Button(barra, text="Marcar todos",
                   command=lambda: self.marcar(True)).pack(side="left")
        ttk.Button(barra, text="Desmarcar todos",
                   command=lambda: self.marcar(False)).pack(side="left", padx=4)
        ttk.Button(barra, text="Solo los que tienen recibo",
                   command=self.marcar_con_recibo).pack(side="left", padx=4)
        ttk.Button(barra, text="Quitar",
                   command=self.quitar).pack(side="right")
        ttk.Button(barra, text="Editar",
                   command=self.editar).pack(side="right", padx=4)
        ttk.Button(barra, text="Agregar",
                   command=self.agregar).pack(side="right")
        ttk.Button(barra, text="Abrir lista en Excel",
                   command=self.abrir_csv).pack(side="right", padx=12)

        columnas = ("marca", "nombre", "cedula", "correo", "recibo")
        self.tree = ttk.Treeview(f4, columns=columnas, show="headings",
                                 selectmode="extended")
        for clave, titulo, ancho, ancla in (
            ("marca", "", 34, "center"),
            ("nombre", "Nombre y apellido", 300, "w"),
            ("cedula", "Cedula", 110, "w"),
            ("correo", "Correo", 260, "w"),
            ("recibo", "Recibo", 160, "w"),
        ):
            self.tree.heading(clave, text=titulo)
            self.tree.column(clave, width=ancho, anchor=ancla,
                             stretch=(clave in ("nombre", "correo")))
        self.tree.tag_configure("falta", foreground="#b00020")
        self.tree.grid(row=1, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(f4, orient="vertical", command=self.tree.yview)
        scroll.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind("<Button-1>", self.clic_tabla)
        self.tree.bind("<Double-1>", lambda *_: self.editar())
        self.tree.bind("<space>", self.alternar_seleccion)

        # --- acciones
        f5 = ttk.Frame(raiz)
        f5.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        self.btn_enviar = ttk.Button(f5, text="Enviar correos a los marcados",
                                     command=self.enviar)
        self.btn_enviar.pack(side="left")
        self.btn_imprimir = ttk.Button(f5, text="Imprimir los marcados",
                                       command=self.imprimir)
        self.btn_imprimir.pack(side="left", padx=6)
        self.var_estado = tk.StringVar(value="")
        ttk.Label(f5, textvariable=self.var_estado).pack(side="left", padx=12)

        # --- registro
        f6 = ttk.LabelFrame(raiz, text="Registro", padding=6)
        f6.grid(row=5, column=0, sticky="ew", pady=(8, 0))
        f6.columnconfigure(0, weight=1)
        self.txt_log = tk.Text(f6, height=6, wrap="word", state="disabled")
        self.txt_log.grid(row=0, column=0, sticky="ew")
        s2 = ttk.Scrollbar(f6, orient="vertical", command=self.txt_log.yview)
        s2.grid(row=0, column=1, sticky="ns")
        self.txt_log.configure(yscrollcommand=s2.set)

    # ---- utilidades

    def log(self, texto):
        linea = time.strftime("[%H:%M:%S] ") + texto
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", linea + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")
        try:
            with open(ARCHIVO_LOG, "a", encoding="utf-8") as f:
                f.write(time.strftime("%Y-%m-%d %H:%M:%S ") + texto + "\n")
        except Exception:
            pass

    def _vaciar_cola(self):
        while True:
            try:
                tipo, dato = self.mensajes.get_nowait()
            except queue.Empty:
                break
            if tipo == "log":
                self.log(dato)
            elif tipo == "estado":
                self.var_estado.set(dato)
            elif tipo == "fin":
                self.trabajando = False
                self.btn_enviar.configure(state="normal")
                self.btn_imprimir.configure(state="normal")
                self.var_estado.set(dato or "")
        self.after(200, self._vaciar_cola)

    def leer_config(self):
        self.config_datos.update({
            "carpeta": self.var_carpeta.get().strip(),
            "proveedor": self.var_proveedor.get(),
            "servidor": self.var_servidor.get().strip(),
            "puerto": int(self.var_puerto.get()) if self.var_puerto.get().strip().isdigit() else 587,
            "remitente": self.var_remitente.get().strip(),
            "password": self.var_password.get(),
            "recordar_password": bool(self.var_recordar.get()),
            "asunto": self.var_asunto.get(),
            "cuerpo": self.txt_cuerpo.get("1.0", "end-1c"),
        })
        return self.config_datos

    def cerrar(self):
        if self.trabajando and not messagebox.askyesno(
                "Recibos", "Todavia se esta enviando. ¿Cerrar de todas formas?"):
            return
        guardar_config(self.leer_config())
        guardar_empleados(self.empleados)
        self.destroy()

    # ---- carpeta y recibos

    def cambiar_proveedor(self, *_):
        servidor, puerto = PROVEEDORES.get(self.var_proveedor.get(), ("", 587))
        if servidor:
            self.var_servidor.set(servidor)
            self.var_puerto.set(str(puerto))

    def elegir_carpeta(self):
        inicio = self.var_carpeta.get() or str(Path.home())
        carpeta = filedialog.askdirectory(title="Carpeta donde estan los PDF",
                                          initialdir=inicio)
        if carpeta:
            self.var_carpeta.set(carpeta)
            guardar_config(self.leer_config())
            self.refrescar_recibos(True)

    def importar_rar(self):
        archivo = filedialog.askopenfilename(
            title="Elige el RAR que descargaste",
            filetypes=[("Comprimidos", "*.rar *.zip *.7z"), ("Todos", "*.*")])
        if not archivo:
            return
        destino = self.var_carpeta.get().strip()
        if not destino:
            destino = filedialog.askdirectory(
                title="¿En que carpeta se van a guardar los recibos?")
            if not destino:
                return
            self.var_carpeta.set(destino)
        os.makedirs(destino, exist_ok=True)
        self.log("Extrayendo %s ..." % os.path.basename(archivo))
        error = extraer(archivo, destino)
        if error:
            self.log("ERROR al extraer: %s" % error.splitlines()[0])
            messagebox.showerror("Recibos", error)
            return
        guardar_config(self.leer_config())
        self.refrescar_recibos(True)

    def refrescar_recibos(self, avisar=True):
        carpeta = self.var_carpeta.get().strip()
        indice = indexar_recibos(carpeta)
        emparejar(self.empleados, indice)
        self.pintar_tabla()
        total = sum(len(v) for v in indice.values())
        con = sum(1 for e in self.empleados if e["pdfs"])
        if avisar or total:
            self.log("Carpeta: %s | %d PDF encontrados | %d de %d trabajadores "
                     "con recibo" % (carpeta or "(sin elegir)", total, con,
                                     len(self.empleados)))
        if avisar and self.empleados and con == 0 and total:
            messagebox.showwarning(
                "Recibos",
                "Se encontraron PDF pero ninguno coincide con las cedulas de la "
                "lista.\nRevisa que las cedulas esten bien escritas.")

    # ---- tabla

    def pintar_tabla(self):
        seleccion = set(self.tree.selection())
        self.tree.delete(*self.tree.get_children())
        for i, e in enumerate(self.empleados):
            cantidad = len(e["pdfs"])
            if cantidad == 0:
                estado, etiquetas = "sin recibo", ("falta",)
            elif cantidad == 1:
                estado, etiquetas = "listo", ()
            else:
                estado, etiquetas = "%d archivos" % cantidad, ()
            self.tree.insert("", "end", iid=str(i), tags=etiquetas,
                             values=("X" if e["marcado"] else "",
                                     e["nombre"], e["cedula"], e["correo"], estado))
        for iid in seleccion:
            if self.tree.exists(iid):
                self.tree.selection_add(iid)
        marcados = sum(1 for e in self.empleados if e["marcado"])
        self.var_estado.set("%d marcados de %d" % (marcados, len(self.empleados)))

    def clic_tabla(self, evento):
        if self.tree.identify_region(evento.x, evento.y) != "cell":
            return None
        if self.tree.identify_column(evento.x) != "#1":
            return None
        fila = self.tree.identify_row(evento.y)
        if fila:
            i = int(fila)
            self.empleados[i]["marcado"] = not self.empleados[i]["marcado"]
            self.pintar_tabla()
        return "break"

    def alternar_seleccion(self, *_):
        for iid in self.tree.selection():
            i = int(iid)
            self.empleados[i]["marcado"] = not self.empleados[i]["marcado"]
        self.pintar_tabla()
        return "break"

    def marcar(self, valor):
        for e in self.empleados:
            e["marcado"] = valor
        self.pintar_tabla()

    def marcar_con_recibo(self):
        for e in self.empleados:
            e["marcado"] = bool(e["pdfs"]) and bool(e["correo"])
        self.pintar_tabla()

    def agregar(self):
        datos = pedir_datos(self, "Agregar trabajador")
        if datos:
            datos.update({"marcado": False, "pdfs": []})
            self.empleados.append(datos)
            guardar_empleados(self.empleados)
            self.refrescar_recibos(False)

    def editar(self):
        seleccion = self.tree.selection()
        if not seleccion:
            return
        i = int(seleccion[0])
        datos = pedir_datos(self, "Editar trabajador", self.empleados[i])
        if datos:
            self.empleados[i].update(datos)
            guardar_empleados(self.empleados)
            self.refrescar_recibos(False)

    def quitar(self):
        seleccion = self.tree.selection()
        if not seleccion:
            return
        if not messagebox.askyesno("Recibos",
                                   "¿Quitar %d de la lista?" % len(seleccion)):
            return
        for i in sorted((int(x) for x in seleccion), reverse=True):
            del self.empleados[i]
        guardar_empleados(self.empleados)
        self.pintar_tabla()

    def abrir_csv(self):
        guardar_empleados(self.empleados)
        try:
            os.startfile(str(ARCHIVO_EMPLEADOS))
        except Exception as e:
            messagebox.showerror("Recibos", str(e))
            return
        messagebox.showinfo(
            "Recibos",
            "Edita la lista, guardala y cierra Excel.\n"
            "Despues da clic en Aceptar para volver a cargarla.")
        self.empleados = cargar_empleados()
        self.refrescar_recibos(False)

    # ---- envio

    def datos_smtp(self):
        cfg = self.leer_config()
        if not cfg["remitente"] or "@" not in cfg["remitente"]:
            messagebox.showwarning("Recibos", "Escribe el correo de quien envia.")
            return None
        if not cfg["password"]:
            messagebox.showwarning("Recibos", "Escribe la contrasena de aplicacion.")
            return None
        if not cfg["servidor"]:
            messagebox.showwarning("Recibos", "Escribe el servidor de correo.")
            return None
        guardar_config(cfg)
        return cfg

    def conectar(self, cfg):
        if cfg["puerto"] == 465:
            servidor = smtplib.SMTP_SSL(cfg["servidor"], cfg["puerto"], timeout=40)
        else:
            servidor = smtplib.SMTP(cfg["servidor"], cfg["puerto"], timeout=40)
            servidor.starttls(context=ssl.create_default_context())
        servidor.login(cfg["remitente"], cfg["password"])
        return servidor

    def probar_conexion(self):
        cfg = self.datos_smtp()
        if not cfg:
            return
        try:
            servidor = self.conectar(cfg)
            servidor.quit()
        except Exception as e:
            self.log("Conexion fallida: %s" % e)
            messagebox.showerror(
                "Recibos",
                "No se pudo entrar al correo:\n%s\n\n"
                "Si usas Gmail u Outlook necesitas una 'contrasena de aplicacion', "
                "no la contrasena normal. Mira el README." % e)
            return
        self.log("Conexion correcta con %s" % cfg["servidor"])
        messagebox.showinfo("Recibos", "Todo bien, el correo funciona.")

    def enviar(self):
        if self.trabajando:
            return
        cfg = self.datos_smtp()
        if not cfg:
            return
        marcados = [e for e in self.empleados if e["marcado"]]
        if not marcados:
            messagebox.showwarning("Recibos", "No hay nadie marcado.")
            return
        sin_correo = [e["nombre"] for e in marcados if "@" not in e["correo"]]
        sin_pdf = [e["nombre"] for e in marcados if not e["pdfs"]]
        listos = [e for e in marcados if "@" in e["correo"] and e["pdfs"]]
        if not listos:
            messagebox.showwarning("Recibos",
                                   "Ninguno de los marcados tiene correo y recibo.")
            return
        aviso = "Se van a enviar %d correos desde %s." % (len(listos), cfg["remitente"])
        if sin_correo:
            aviso += "\n\nSe saltan por no tener correo: %s" % ", ".join(sin_correo[:8])
        if sin_pdf:
            aviso += "\n\nSe saltan por no tener recibo: %s" % ", ".join(sin_pdf[:8])
        aviso += "\n\n¿Continuar?"
        if not messagebox.askyesno("Recibos", aviso):
            return

        self.trabajando = True
        self.btn_enviar.configure(state="disabled")
        self.btn_imprimir.configure(state="disabled")
        hilo = threading.Thread(target=self._hilo_enviar, args=(cfg, listos), daemon=True)
        hilo.start()

    def _hilo_enviar(self, cfg, listos):
        enviados = fallidos = 0
        try:
            self.mensajes.put(("log", "Entrando a %s ..." % cfg["servidor"]))
            servidor = self.conectar(cfg)
        except Exception as e:
            self.mensajes.put(("log", "ERROR de conexion: %s" % e))
            self.mensajes.put(("fin", "No se pudo conectar"))
            return

        for n, emp in enumerate(listos, 1):
            self.mensajes.put(("estado", "Enviando %d de %d..." % (n, len(listos))))
            try:
                msg = EmailMessage()
                msg["From"] = cfg["remitente"]
                msg["To"] = emp["correo"]
                msg["Subject"] = self._plantilla(cfg["asunto"], emp)
                msg.set_content(self._plantilla(cfg["cuerpo"], emp))
                for pdf in emp["pdfs"]:
                    with open(pdf, "rb") as f:
                        msg.add_attachment(f.read(), maintype="application",
                                           subtype="pdf", filename=pdf.name)
                servidor.send_message(msg)
                enviados += 1
                self.mensajes.put(("log", "OK  %s -> %s" % (emp["nombre"],
                                                            emp["correo"])))
            except Exception as e:
                fallidos += 1
                self.mensajes.put(("log", "FALLO  %s -> %s : %s" % (
                    emp["nombre"], emp["correo"], e)))
        try:
            servidor.quit()
        except Exception:
            pass
        self.mensajes.put(("log", "Terminado: %d enviados, %d con fallo"
                           % (enviados, fallidos)))
        self.mensajes.put(("fin", "%d enviados, %d con fallo" % (enviados, fallidos)))

    @staticmethod
    def _plantilla(texto, emp):
        return (texto.replace("{nombre}", emp["nombre"])
                     .replace("{cedula}", emp["cedula"]))

    # ---- impresion

    def imprimir(self):
        if self.trabajando:
            return
        pdfs = []
        for e in self.empleados:
            if e["marcado"]:
                pdfs.extend(e["pdfs"])
        if not pdfs:
            messagebox.showwarning("Recibos", "No hay recibos marcados para imprimir.")
            return
        if not messagebox.askyesno(
                "Recibos",
                "Se van a mandar %d recibos a la impresora que esta puesta por "
                "defecto en Windows.\n\n¿Continuar?" % len(pdfs)):
            return
        self.trabajando = True
        self.btn_enviar.configure(state="disabled")
        self.btn_imprimir.configure(state="disabled")
        threading.Thread(target=self._hilo_imprimir, args=(pdfs,), daemon=True).start()

    def _hilo_imprimir(self, pdfs):
        ok = malos = 0
        for n, pdf in enumerate(pdfs, 1):
            self.mensajes.put(("estado", "Imprimiendo %d de %d..." % (n, len(pdfs))))
            try:
                os.startfile(str(pdf), "print")
                ok += 1
                self.mensajes.put(("log", "A la impresora: %s" % pdf.name))
            except Exception as e:
                malos += 1
                self.mensajes.put(("log", "No se pudo imprimir %s : %s" % (pdf.name, e)))
            time.sleep(1.5)
        if malos:
            self.mensajes.put((
                "log",
                "Si no imprime nada, abre un PDF, dale clic derecho > Abrir con > "
                "Adobe Reader y ponlo como programa por defecto para PDF."))
        self.mensajes.put(("log", "Impresion: %d enviados, %d con fallo" % (ok, malos)))
        self.mensajes.put(("fin", "%d a la impresora, %d con fallo" % (ok, malos)))


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
