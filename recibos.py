# -*- coding: utf-8 -*-
"""
Recibos de Pago - envia a cada trabajador su recibo por correo.

Como funciona:
  1. Se elige una carpeta una sola vez.
  2. Cada mes se copia ahi el RAR con los recibos. El programa lo abre solo.
  3. Para cada persona de la lista busca el PDF que lleve su cedula en el
     nombre (ej: Reg_35-2026_T-12345678.pdf). Si hay varios, usa el mas nuevo.
  4. Se marcan las personas y se le da a "Enviar" o "Imprimir".

Para enviar se inicia sesion con Gmail (en Configuracion), o se usa una
contrasena de aplicacion para Outlook, Yahoo u otro.

Solo usa la libreria estandar de Python. No hay que instalar nada.
"""

import base64
import csv
import hashlib
import json
import os
import queue
import re
import secrets
import shutil
import smtplib
import ssl
import subprocess
import sys
import threading
import time
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
import zipfile
from email.message import EmailMessage
from http.server import BaseHTTPRequestHandler, HTTPServer
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
ARCHIVO_SESION = BASE / "sesion_gmail.json"

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
                                  "correo": correo, "marcado": False, "pdfs": [],
                                  "estado": ""})
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
        for numero in set(re.findall(r"\d{5,}", pdf.stem)):
            indice.setdefault(numero, []).append(pdf)
    return indice


def _fecha(ruta):
    try:
        return ruta.stat().st_mtime
    except OSError:
        return 0


def emparejar(lista, indice):
    for e in lista:
        cedula = solo_digitos(e["cedula"])
        encontrados = indice.get(cedula, []) if cedula else []
        # Si quedaron recibos de meses anteriores, se usa solo el mas nuevo.
        nuevo = sorted(encontrados, key=lambda p: (_fecha(p), p.name))[-1:]
        if nuevo != e.get("pdfs"):
            e["estado"] = ""  # recibo distinto: todavia no se ha enviado
            if not nuevo:
                e["marcado"] = False
        e["pdfs"] = nuevo


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
        return ("No se encontró WinRAR ni 7-Zip en esta computadora.\n\n"
                "Instala uno de los dos, o abre el RAR a mano (clic derecho "
                "sobre el archivo > Extraer aquí) en esta misma carpeta.")

    os.makedirs(destino, exist_ok=True)
    nombre = os.path.basename(programa).lower()
    if nombre.startswith("7z"):
        cmd = [programa, "x", archivo, "-o" + destino, "-y", "-p-"]
    else:  # UnRAR.exe / WinRAR.exe; -p- para que nunca se quede pidiendo clave
        cmd = [programa, "x", "-o+", "-y", "-p-", archivo, destino + os.sep]
        if nombre == "winrar.exe":
            cmd.insert(1, "-ibck")  # sin ventana de WinRAR
    try:
        # WinRAR y 7-Zip escriben en la codificacion de la consola (oem)
        r = subprocess.run(cmd, capture_output=True, timeout=600,
                           encoding="oem" if os.name == "nt" else None, errors="replace",
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception as e:
        return str(e)
    if r.returncode != 0:
        return (r.stderr or r.stdout or "codigo %s" % r.returncode).strip()
    return None


COMPRIMIDOS = (".rar", ".zip", ".7z")


def extraer_rars_nuevos(carpeta, fallidos):
    """Abre cada RAR de la carpeta en una subcarpeta propia, una sola vez.

    La subcarpeta lleva la fecha del RAR, asi el del mes que viene se abre
    aunque se llame igual. fallidos = {nombre: (tamano, fecha)} para no
    reintentar uno que ya fallo mientras no cambie.
    Devuelve (abiertos, errores, pendientes).
    """
    abiertos, errores, pendientes = [], [], False
    ruta = Path(carpeta)
    for comp in sorted(ruta.iterdir()):
        if not comp.is_file() or comp.suffix.lower() not in COMPRIMIDOS:
            continue
        info = comp.stat()
        fecha = time.strftime("%Y-%m-%d %H.%M", time.localtime(info.st_mtime))
        destino = ruta / ("%s (%s)" % (comp.stem, fecha))
        huella = (info.st_size, info.st_mtime)
        if destino.exists() or fallidos.get(comp.name) == huella:
            continue
        if time.time() - info.st_mtime < 3:  # todavia se esta copiando
            pendientes = True
            continue
        temporal = ruta / (comp.stem + ".abriendo")
        shutil.rmtree(temporal, ignore_errors=True)
        error = extraer(str(comp), temporal)
        if error:
            shutil.rmtree(temporal, ignore_errors=True)
            fallidos[comp.name] = huella
            errores.append((comp.name, error))
        else:
            temporal.rename(destino)
            abiertos.append(comp.name)
    return abiertos, errores, pendientes


def firma_carpeta(carpeta):
    """Lo que hay en la carpeta, para darse cuenta cuando llega algo nuevo."""
    try:
        ruta = Path(carpeta)
        if not carpeta or not ruta.is_dir():
            return None
        return tuple(sorted(
            (str(p), _fecha(p)) for p in ruta.rglob("*")
            if p.suffix.lower() in (".pdf",) + COMPRIMIDOS))
    except OSError:
        return None


# ---------------------------------------------------------------- Gmail

# Iniciar sesion con Google (OAuth para apps de escritorio). Solo se pide
# permiso para ENVIAR correos, nunca para leerlos.
GOOGLE_AUTH = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_TOKEN = "https://oauth2.googleapis.com/token"
GOOGLE_REVOCAR = "https://oauth2.googleapis.com/revoke"
GMAIL_ENVIAR = "https://gmail.googleapis.com/gmail/v1/users/me/messages/send"
PERMISO_ENVIAR = "https://www.googleapis.com/auth/gmail.send"
PERMISOS = "openid email " + PERMISO_ENVIAR

PAGINA_LISTO = """<!doctype html><html><head><meta charset="utf-8">
<title>Recibos de Pago</title></head>
<body style="font-family:Segoe UI,Arial,sans-serif;text-align:center;padding-top:90px">
<h2>Listo.</h2><p>Ya puedes cerrar esta pesta&ntilde;a y volver al programa.</p>
</body></html>"""


def buscar_client_secret():
    """(client_id, client_secret) del archivo que se baja de Google Cloud."""
    carpetas = [BASE]
    if hasattr(sys, "_MEIPASS"):  # dentro del .exe
        carpetas.append(Path(sys._MEIPASS))
    for carpeta in carpetas:
        for ruta in sorted(carpeta.glob("client_secret*.json")):
            try:
                with open(ruta, "r", encoding="utf-8") as f:
                    datos = json.load(f)
                datos = datos.get("installed") or datos
                return datos["client_id"], datos.get("client_secret", "")
            except Exception:
                continue
    return None


def cargar_sesion_gmail():
    try:
        with open(ARCHIVO_SESION, "r", encoding="utf-8") as f:
            datos = json.load(f)
        return datos if datos.get("refresh_token") else None
    except Exception:
        return None


def borrar_sesion_gmail():
    try:
        ARCHIVO_SESION.unlink()
    except FileNotFoundError:
        pass


def _pedir_token(campos):
    datos = urllib.parse.urlencode(campos).encode("ascii")
    with urllib.request.urlopen(GOOGLE_TOKEN, data=datos, timeout=30) as r:
        return json.load(r)


def _correo_del_id_token(id_token):
    try:
        carga = id_token.split(".")[1]
        carga += "=" * (-len(carga) % 4)
        return json.loads(base64.urlsafe_b64decode(carga)).get("email", "")
    except Exception:
        return ""


def gmail_login(espera=300):
    """Abre el navegador para iniciar sesion y guarda la sesion."""
    cliente = buscar_client_secret()
    if not cliente:
        raise RuntimeError("Falta el archivo client_secret.json")
    client_id, client_secret = cliente
    verificador = secrets.token_urlsafe(64)
    reto = base64.urlsafe_b64encode(
        hashlib.sha256(verificador.encode("ascii")).digest()).decode("ascii").rstrip("=")
    estado = secrets.token_urlsafe(16)
    respuesta = {}

    # Google devuelve al navegador a esta direccion local con el codigo.
    class Vuelta(BaseHTTPRequestHandler):
        def do_GET(self):
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            if "code" not in q and "error" not in q:
                self.send_response(404)
                self.end_headers()
                return
            respuesta.update({k: v[0] for k, v in q.items()})
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(PAGINA_LISTO.encode("utf-8"))

        def log_message(self, *args):
            pass

    servidor = HTTPServer(("127.0.0.1", 0), Vuelta)
    servidor.timeout = 1
    redirect = "http://127.0.0.1:%d/" % servidor.server_port
    webbrowser.open(GOOGLE_AUTH + "?" + urllib.parse.urlencode({
        "client_id": client_id,
        "redirect_uri": redirect,
        "response_type": "code",
        "scope": PERMISOS,
        "code_challenge": reto,
        "code_challenge_method": "S256",
        "state": estado,
        "access_type": "offline",
        "prompt": "consent",
    }))
    limite = time.time() + espera
    try:
        while not respuesta and time.time() < limite:
            servidor.handle_request()
    finally:
        servidor.server_close()

    if not respuesta:
        raise RuntimeError("Se acabo el tiempo para iniciar sesion. Intenta de nuevo.")
    if respuesta.get("error") == "access_denied":
        raise RuntimeError("Se cancelo el inicio de sesion.")
    if "error" in respuesta:
        raise RuntimeError("Google respondio: %s" % respuesta["error"])
    if respuesta.get("state") != estado:
        raise RuntimeError("La respuesta del navegador no coincide. Intenta de nuevo.")

    tok = _pedir_token({
        "code": respuesta["code"],
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": redirect,
        "grant_type": "authorization_code",
        "code_verifier": verificador,
    })
    # Google deja desmarcar permisos en la pantalla de inicio de sesion.
    if PERMISO_ENVIAR not in tok.get("scope", "").split():
        raise RuntimeError("No se marco la casilla para enviar correos. Vuelve a "
                           "iniciar sesion y marca esa casilla.")
    if not tok.get("refresh_token"):
        raise RuntimeError("Google no devolvio el permiso para quedarse conectado.")
    sesion = {"correo": _correo_del_id_token(tok.get("id_token", "")),
              "refresh_token": tok["refresh_token"]}
    with open(ARCHIVO_SESION, "w", encoding="utf-8") as f:
        json.dump(sesion, f, indent=2)
    return sesion


def gmail_access_token(sesion):
    """Permiso de envio, dura una hora."""
    client_id, client_secret = buscar_client_secret() or ("", "")
    tok = _pedir_token({
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": sesion["refresh_token"],
        "grant_type": "refresh_token",
    })
    return tok["access_token"]


def gmail_enviar(access_token, msg):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("ascii")
    pedido = urllib.request.Request(
        GMAIL_ENVIAR, data=json.dumps({"raw": raw}).encode("ascii"),
        headers={"Authorization": "Bearer " + access_token,
                 "Content-Type": "application/json"})
    with urllib.request.urlopen(pedido, timeout=60) as r:
        return json.load(r)


def revocar_gmail(refresh_token):
    try:
        datos = urllib.parse.urlencode({"token": refresh_token}).encode("ascii")
        urllib.request.urlopen(GOOGLE_REVOCAR, data=datos, timeout=10).close()
    except Exception:
        pass


def detalle(error):
    """Texto entendible de un error, incluido lo que responde Google."""
    if isinstance(error, urllib.error.HTTPError):
        try:
            cuerpo = json.loads(error.read().decode("utf-8", "replace"))
        except Exception:
            return "HTTP %s" % error.code
        info = cuerpo.get("error")
        if isinstance(info, dict):
            return "%s %s" % (error.code, info.get("message", ""))
        if info == "invalid_grant":
            return "la sesion de Gmail vencio o se cerro. Vuelve a iniciar sesion."
        return "%s %s %s" % (error.code, info or "", cuerpo.get("error_description", ""))
    return str(error)


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

VERDE = "#2e9d4f"
ROJO = "#d64545"
AZUL = "#1a5fb4"
GRIS = "#777777"

# estado de cada persona: (texto en la tabla, color)
ESTADOS = {
    "enviando": ("➜ Enviando...", "enviando"),
    "enviado": ("✔ Enviado", "hecho"),
    "impreso": ("✔ Impreso", "hecho"),
    "fallo": ("✖ No se envió", "fallo"),
    "noimpreso": ("✖ No se imprimió", "fallo"),
}


def _casilla(fondo, borde, palomita=None):
    """Cuadrito de 18x18 para marcar a cada persona."""
    img = tk.PhotoImage(width=18, height=18)
    img.put(borde, to=(1, 1, 17, 17))
    img.put(fondo, to=(3, 3, 15, 15))
    if palomita:
        for x, y in ((4, 8), (5, 9), (6, 10), (7, 11), (8, 10), (9, 9),
                     (10, 8), (11, 7), (12, 6), (13, 5)):
            img.put(palomita, to=(x, y, x + 2, y + 2))
    return img


class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Recibos de Pago")
        # Que quepa arriba de la barra de tareas en pantallas chicas (768 px).
        alto = min(580, self.winfo_screenheight() - 100)
        self.geometry("1040x%d+20+10" % alto)
        self.minsize(880, 420)

        self.config_datos = cargar_config()
        self.empleados = cargar_empleados()
        self.sesion_gmail = cargar_sesion_gmail()
        self.mensajes = queue.Queue()
        self.trabajando = False
        self.refrescando = False
        self.ventana_config = None
        self.btn_gmail = self.btn_salir = None
        self.campos_manuales = []
        self.rars_fallidos = {}
        self.firma = None
        self.hay_pendientes = False

        self.img = {
            "libre": _casilla("white", VERDE),
            "marcado": _casilla(VERDE, VERDE, "white"),
            "falta": _casilla("#fbe3e3", ROJO),
        }
        c = self.config_datos
        self.var_carpeta = tk.StringVar(value=c["carpeta"])
        self.var_asunto = tk.StringVar(value=c["asunto"])
        self.var_proveedor = tk.StringVar(value=c["proveedor"])
        self.var_servidor = tk.StringVar(value=c["servidor"])
        self.var_puerto = tk.StringVar(value=str(c["puerto"]))
        self.var_remitente = tk.StringVar(value=c["remitente"])
        self.var_password = tk.StringVar(value=c["password"])
        self.var_recordar = tk.BooleanVar(value=c["recordar_password"])
        self.var_sesion = tk.StringVar()
        self.var_cuenta = tk.StringVar()
        self.var_marcados = tk.StringVar()
        self.var_progreso = tk.StringVar()
        self.var_todos = tk.BooleanVar()

        self._construir()
        self.pintar_cuenta()
        self.refrescar(avisar=False)
        self.after(3000, self._vigilar_carpeta)
        self.after(200, self._vaciar_cola)
        self.protocol("WM_DELETE_WINDOW", self.cerrar)

    # ---- ventana principal

    def _construir(self):
        estilo = ttk.Style()
        estilo.configure("Treeview", rowheight=26)
        estilo.configure("Grande.TButton", font=("Segoe UI", 13, "bold"), padding=(10, 16))
        estilo.configure("Gris.TLabel", foreground=GRIS)

        raiz = ttk.Frame(self, padding=14)
        raiz.pack(fill="both", expand=True)
        raiz.columnconfigure(1, weight=1)
        raiz.rowconfigure(1, weight=1)

        # --- arriba: carpeta y configuracion
        # (lo de la derecha se empaca primero para que nunca quede tapado)
        arriba = ttk.Frame(raiz)
        arriba.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        ttk.Button(arriba, text="Elegir carpeta",
                   command=self.elegir_carpeta).pack(side="left")
        self.btn_config = ttk.Button(arriba, text="⚙  Configuración",
                                     command=self.abrir_config)
        self.btn_config.pack(side="right")
        ttk.Label(arriba, textvariable=self.var_cuenta,
                  style="Gris.TLabel").pack(side="right", padx=12)
        self.var_carpeta_corta = tk.StringVar()
        ttk.Label(arriba, textvariable=self.var_carpeta_corta,
                  style="Gris.TLabel").pack(side="left", padx=10)

        # --- izquierda: lo que hay en la carpeta
        izq = ttk.Frame(raiz)
        izq.grid(row=1, column=0, sticky="ns", padx=(0, 14))
        izq.rowconfigure(1, weight=1)
        ttk.Label(izq, text="En la carpeta").grid(row=0, column=0, sticky="w", pady=(0, 4))
        self.lst_archivos = tk.Listbox(izq, width=29, activestyle="none", relief="solid",
                                       borderwidth=1, highlightthickness=0,
                                       font=("Segoe UI", 9))
        self.lst_archivos.grid(row=1, column=0, sticky="ns")
        barra_izq = ttk.Scrollbar(izq, orient="vertical", command=self.lst_archivos.yview)
        barra_izq.grid(row=1, column=1, sticky="ns")
        self.lst_archivos.configure(yscrollcommand=barra_izq.set)

        # --- centro: la gente
        centro = ttk.Frame(raiz)
        centro.grid(row=1, column=1, sticky="nsew")
        centro.columnconfigure(0, weight=1)
        centro.rowconfigure(1, weight=1)
        ttk.Checkbutton(centro, text="Seleccionar todos", variable=self.var_todos,
                        command=self.marcar_todos).grid(row=0, column=0, sticky="w",
                                                        pady=(0, 4))
        self.tree = ttk.Treeview(centro, columns=("nombre", "cedula", "correo", "estado"),
                                 show="tree headings", selectmode="browse")
        self.tree.column("#0", width=42, minwidth=42, stretch=False, anchor="center")
        for clave, titulo, ancho in (("nombre", "Nombre", 165), ("cedula", "Cédula", 85),
                                     ("correo", "Correo", 175), ("estado", "Estado", 125)):
            self.tree.heading(clave, text=titulo, anchor="w")
            self.tree.column(clave, width=ancho, minwidth=60, anchor="w",
                             stretch=clave in ("nombre", "correo"))
        self.tree.tag_configure("falta", foreground=GRIS)
        self.tree.tag_configure("aviso", foreground="#b26a00")
        self.tree.tag_configure("enviando", foreground=AZUL)
        self.tree.tag_configure("hecho", foreground=VERDE)
        self.tree.tag_configure("fallo", foreground=ROJO)
        self.tree.grid(row=1, column=0, sticky="nsew")
        barra_tabla = ttk.Scrollbar(centro, orient="vertical", command=self.tree.yview)
        barra_tabla.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=barra_tabla.set)
        self.tree.bind("<Button-1>", self.clic_tabla)
        self.tree.bind("<Double-1>", self.doble_clic)
        self.tree.bind("<Button-3>", self.menu_tabla)

        pie = ttk.Frame(centro)
        pie.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(pie, text="+ Agregar persona", command=self.agregar).pack(side="left")
        ttk.Label(pie, text="Doble clic en una persona para corregirla.",
                  style="Gris.TLabel").pack(side="left", padx=10)

        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Corregir", command=self.editar)
        self.menu.add_command(label="Quitar de la lista", command=self.quitar)

        # --- derecha: enviar, imprimir y como va
        der = ttk.Frame(raiz)
        der.grid(row=1, column=2, sticky="n", padx=(16, 0))
        self.btn_enviar = ttk.Button(der, text="Enviar", style="Grande.TButton",
                                     width=12, command=self.enviar)
        self.btn_enviar.pack(fill="x")
        self.btn_imprimir = ttk.Button(der, text="Imprimir", style="Grande.TButton",
                                       width=12, command=self.imprimir)
        self.btn_imprimir.pack(fill="x", pady=(12, 0))
        ttk.Label(der, textvariable=self.var_marcados,
                  style="Gris.TLabel").pack(pady=(12, 0))
        self.barra = ttk.Progressbar(der, mode="determinate")
        self.barra.pack(fill="x", pady=(18, 6))
        ttk.Label(der, textvariable=self.var_progreso, wraplength=180,
                  justify="center").pack()

    # ---- utilidades

    def log(self, texto):
        """Queda en envios.log, por si hay que revisar algo despues."""
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
            elif tipo == "fila":
                i, estado = dato
                if i < len(self.empleados):
                    self.empleados[i]["estado"] = estado
                    if estado == "enviado":
                        self.empleados[i]["marcado"] = False  # para no enviarle dos veces
                    self.pintar_fila(i)
                    self._contar()
                    if estado == "enviando":
                        self.tree.see(str(i))
            elif tipo == "progreso":
                hechos, total, texto = dato
                self.barra.configure(maximum=max(total, 1), value=hechos)
                self.var_progreso.set(texto)
            elif tipo == "sesion":
                self.sesion_gmail = dato
                self.pintar_cuenta()
            elif tipo == "error":
                messagebox.showerror("Recibos", dato, parent=self.ventana_config or self)
            elif tipo == "fin":
                self._ocupado(False)
                if dato:
                    self.var_progreso.set(dato)
        self.after(200, self._vaciar_cola)

    def _ocupado(self, si):
        self.trabajando = si
        estado = "disabled" if si else "normal"
        for boton in (self.btn_enviar, self.btn_imprimir, self.btn_config, self.btn_gmail):
            if boton is not None:
                try:
                    boton.configure(state=estado)
                except tk.TclError:
                    pass

    def leer_config(self):
        if self.ventana_config is not None:
            self.config_datos["cuerpo"] = self.txt_cuerpo.get("1.0", "end-1c")
        puerto = self.var_puerto.get().strip()
        self.config_datos.update({
            "carpeta": self.var_carpeta.get().strip(),
            "proveedor": self.var_proveedor.get(),
            "servidor": self.var_servidor.get().strip(),
            "puerto": int(puerto) if puerto.isdigit() else 587,
            "remitente": self.var_remitente.get().strip(),
            "password": self.var_password.get(),
            "recordar_password": bool(self.var_recordar.get()),
            "asunto": self.var_asunto.get(),
        })
        return self.config_datos

    def cerrar(self):
        if self.trabajando and not messagebox.askyesno(
                "Recibos", "Todavía se está enviando. ¿Cerrar de todas formas?"):
            return
        guardar_config(self.leer_config())
        guardar_empleados(self.empleados)
        self.destroy()

    # ---- carpeta

    def elegir_carpeta(self):
        inicio = self.var_carpeta.get() or str(Path.home())
        carpeta = filedialog.askdirectory(title="Carpeta donde se meten los RAR de los recibos",
                                          initialdir=inicio)
        if carpeta:
            self.var_carpeta.set(os.path.normpath(carpeta))
            guardar_config(self.leer_config())
            self.refrescar(avisar=True)

    def _vigilar_carpeta(self):
        """Cada 3 segundos mira si llego algo nuevo a la carpeta (un RAR)."""
        if not self.trabajando and not self.refrescando:
            carpeta = self.var_carpeta.get().strip()
            if self.hay_pendientes or firma_carpeta(carpeta) != self.firma:
                self.refrescar(avisar=False)
        self.after(3000, self._vigilar_carpeta)

    def refrescar(self, avisar=True):
        if self.refrescando:
            return
        self.refrescando = True
        try:
            carpeta = self.var_carpeta.get().strip()
            errores = []
            self.hay_pendientes = False
            if carpeta and os.path.isdir(carpeta):
                self.configure(cursor="watch")
                self.update_idletasks()
                try:
                    abiertos, errores, self.hay_pendientes = extraer_rars_nuevos(
                        carpeta, self.rars_fallidos)
                except Exception as e:
                    abiertos, errores = [], [("la carpeta", str(e))]
                finally:
                    self.configure(cursor="")
                for nombre in abiertos:
                    self.log("Abierto %s" % nombre)
            indice = indexar_recibos(carpeta)
            emparejar(self.empleados, indice)
            self.firma = firma_carpeta(carpeta)
            self.pintar_archivos(carpeta)
            self.pintar_tabla()
            for nombre, error in errores:
                self.log("No se pudo abrir %s: %s" % (nombre, error))
                messagebox.showerror("Recibos", "No se pudo abrir %s\n\n%s" % (nombre, error))
            if (avisar and self.empleados and indice
                    and not any(e["pdfs"] for e in self.empleados)):
                messagebox.showwarning(
                    "Recibos",
                    "Hay recibos en la carpeta pero ninguno coincide con las cédulas "
                    "de la lista.\nRevisa que las cédulas estén bien escritas.")
        finally:
            self.refrescando = False

    def pintar_archivos(self, carpeta):
        # arriba se ve solo el final de la ruta, si es muy larga
        partes = Path(carpeta).parts if carpeta else ()
        corta = carpeta if len(carpeta) <= 45 else os.path.join("...", *partes[-2:])
        self.var_carpeta_corta.set(corta)
        lista = self.lst_archivos
        lista.delete(0, "end")
        if not carpeta or not os.path.isdir(carpeta):
            lista.insert("end", "  Dale a Elegir carpeta.")
            return
        ruta = Path(carpeta)
        try:
            comprimidos = sorted(p.name for p in ruta.iterdir()
                                 if p.is_file() and p.suffix.lower() in COMPRIMIDOS)
            pdfs = sorted(p.name for p in ruta.rglob("*.pdf"))
        except OSError:
            comprimidos, pdfs = [], []
        for nombre in comprimidos:
            lista.insert("end", "  " + nombre)
            lista.itemconfig("end", foreground=AZUL)
        if pdfs:
            lista.insert("end", "  %d recibos:" % len(pdfs))
            lista.itemconfig("end", foreground=GRIS)
            for nombre in pdfs:
                lista.insert("end", "  " + nombre)
        if not comprimidos and not pdfs:
            lista.insert("end", "  Todavía no hay recibos.")
            lista.insert("end", "  Copia aquí el RAR.")

    # ---- tabla

    def pintar_fila(self, i):
        e = self.empleados[i]
        if not e["pdfs"]:
            casilla, texto, color = "falta", "Sin recibo", "falta"
        else:
            casilla = "marcado" if e["marcado"] else "libre"
            if e.get("estado") in ESTADOS:
                texto, color = ESTADOS[e["estado"]]
            elif "@" not in e["correo"]:
                texto, color = "Falta el correo", "aviso"
            else:
                texto, color = "Listo", ""
        datos = {"image": self.img[casilla], "tags": (color,) if color else (),
                 "values": (e["nombre"], e["cedula"], e["correo"], texto)}
        iid = str(i)
        if self.tree.exists(iid):
            self.tree.item(iid, **datos)
        else:
            self.tree.insert("", "end", iid=iid, **datos)

    def pintar_tabla(self):
        seleccion = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        for i in range(len(self.empleados)):
            self.pintar_fila(i)
        for iid in seleccion:
            if self.tree.exists(iid):
                self.tree.selection_set(iid)
        self._contar()

    def _contar(self):
        con_recibo = [e for e in self.empleados if e["pdfs"]]
        marcados = sum(1 for e in con_recibo if e["marcado"])
        self.var_marcados.set("%d de %d seleccionados" % (marcados, len(con_recibo)))
        self.var_todos.set(bool(con_recibo) and marcados == len(con_recibo))

    def clic_tabla(self, evento):
        if self.tree.identify_column(evento.x) != "#0":
            return None
        if self.tree.identify_region(evento.x, evento.y) == "heading":
            self.var_todos.set(not self.var_todos.get())
            self.marcar_todos()
            return "break"
        fila = self.tree.identify_row(evento.y)
        if fila and not self.trabajando:
            e = self.empleados[int(fila)]
            if e["pdfs"]:
                e["marcado"] = not e["marcado"]
                self.pintar_fila(int(fila))
                self._contar()
        return "break"

    def marcar_todos(self):
        if not self.trabajando:
            valor = self.var_todos.get()
            for e in self.empleados:
                e["marcado"] = valor and bool(e["pdfs"])
        self.pintar_tabla()

    def doble_clic(self, evento):
        if self.tree.identify_column(evento.x) == "#0":
            return "break"
        fila = self.tree.identify_row(evento.y)
        if fila:
            self.tree.selection_set(fila)
            self.editar()
        return "break"

    def menu_tabla(self, evento):
        fila = self.tree.identify_row(evento.y)
        if fila:
            self.tree.selection_set(fila)
            self.menu.tk_popup(evento.x_root, evento.y_root)

    def _reemparejar(self):
        emparejar(self.empleados, indexar_recibos(self.var_carpeta.get().strip()))
        self.pintar_tabla()

    def agregar(self):
        if self.trabajando:
            return
        datos = pedir_datos(self, "Agregar persona")
        if datos:
            datos.update({"marcado": False, "pdfs": [], "estado": ""})
            self.empleados.append(datos)
            guardar_empleados(self.empleados)
            self._reemparejar()
            self.tree.see(str(len(self.empleados) - 1))

    def editar(self):
        seleccion = self.tree.selection()
        if not seleccion or self.trabajando:
            return
        i = int(seleccion[0])
        datos = pedir_datos(self, "Corregir persona", self.empleados[i])
        if datos:
            self.empleados[i].update(datos)
            guardar_empleados(self.empleados)
            self._reemparejar()

    def quitar(self):
        seleccion = self.tree.selection()
        if not seleccion or self.trabajando:
            return
        i = int(seleccion[0])
        if not messagebox.askyesno("Recibos", "¿Quitar a %s de la lista?"
                                   % self.empleados[i]["nombre"]):
            return
        del self.empleados[i]
        guardar_empleados(self.empleados)
        self.pintar_tabla()

    # ---- configuracion

    def abrir_config(self):
        if self.trabajando:
            return
        if self.ventana_config is not None:
            self.ventana_config.lift()
            return
        win = tk.Toplevel(self)
        self.ventana_config = win
        win.title("Configuración")
        win.transient(self)
        win.resizable(False, False)
        win.protocol("WM_DELETE_WINDOW", self.cerrar_config)
        marco = ttk.Frame(win, padding=16)
        marco.pack(fill="both", expand=True)
        marco.columnconfigure(0, weight=1)

        # cuenta de Gmail
        f1 = ttk.LabelFrame(marco, text="Correo que envía los recibos", padding=12)
        f1.grid(row=0, column=0, sticky="ew")
        self.btn_gmail = ttk.Button(f1, command=self.iniciar_gmail)
        self.btn_gmail.pack(side="left")
        ttk.Label(f1, textvariable=self.var_sesion).pack(side="left", padx=12)
        self.btn_salir = ttk.Button(f1, text="Cerrar sesión", command=self.cerrar_sesion_gmail)
        self.btn_salir.pack(side="right")

        # mensaje
        f2 = ttk.LabelFrame(marco, text="Mensaje", padding=12)
        f2.grid(row=1, column=0, sticky="ew", pady=(14, 0))
        f2.columnconfigure(1, weight=1)
        ttk.Label(f2, text="Asunto:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(f2, textvariable=self.var_asunto, width=56).grid(row=0, column=1,
                                                                  sticky="ew")
        self.txt_cuerpo = tk.Text(f2, height=5, width=56, wrap="word", relief="solid",
                                  borderwidth=1, font=("Segoe UI", 10))
        self.txt_cuerpo.insert("1.0", self.config_datos["cuerpo"])
        self.txt_cuerpo.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Label(f2, style="Gris.TLabel",
                  text="Donde dice {nombre} o {cedula} se pone el de cada persona.").grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # lista de personas
        f3 = ttk.LabelFrame(marco, text="Lista de personas", padding=12)
        f3.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        ttk.Button(f3, text="Abrir lista en Excel", command=self.abrir_csv).pack(side="left")
        ttk.Label(f3, style="Gris.TLabel",
                  text="Para cargar muchas personas de una vez.").pack(side="left", padx=12)

        # otro correo (escondido, para Outlook, Yahoo o sin el boton de Gmail)
        self.btn_otro = ttk.Button(marco, style="Toolbutton", command=self._alternar_otro)
        self.btn_otro.grid(row=3, column=0, sticky="w", pady=(14, 0))
        self.f_otro = ttk.Frame(marco, padding=(18, 6, 0, 0))
        self.f_otro.grid(row=4, column=0, sticky="ew")
        self._llenar_otro(self.f_otro)
        self._mostrar_otro(not self.sesion_gmail and bool(self.var_remitente.get().strip()))

        ttk.Button(marco, text="Listo", command=self.cerrar_config).grid(
            row=5, column=0, sticky="e", pady=(16, 0))
        self.pintar_cuenta()
        win.grab_set()

    def _llenar_otro(self, f):
        f.columnconfigure(1, weight=1)
        self.campos_manuales = []

        def manual(widget, normal="normal"):
            self.campos_manuales.append((widget, normal))
            return widget

        ttk.Label(f, text="Correo:").grid(row=0, column=0, sticky="e", padx=(0, 6), pady=3)
        manual(ttk.Entry(f, textvariable=self.var_remitente, width=34)).grid(
            row=0, column=1, sticky="ew", pady=3)
        ttk.Label(f, text="Contraseña de aplicación:").grid(row=1, column=0, sticky="e",
                                                           padx=(0, 6), pady=3)
        manual(ttk.Entry(f, textvariable=self.var_password, show="*", width=34)).grid(
            row=1, column=1, sticky="ew", pady=3)
        ttk.Label(f, text="Servicio:").grid(row=2, column=0, sticky="e", padx=(0, 6), pady=3)
        combo = manual(ttk.Combobox(f, textvariable=self.var_proveedor, state="readonly",
                                    values=list(PROVEEDORES), width=24), "readonly")
        combo.grid(row=2, column=1, sticky="w", pady=3)
        combo.bind("<<ComboboxSelected>>", self.cambiar_proveedor)
        ttk.Label(f, text="Servidor:").grid(row=3, column=0, sticky="e", padx=(0, 6), pady=3)
        caja = ttk.Frame(f)
        caja.grid(row=3, column=1, sticky="w", pady=3)
        manual(ttk.Entry(caja, textvariable=self.var_servidor, width=24)).pack(side="left")
        ttk.Label(caja, text="  puerto:").pack(side="left")
        manual(ttk.Entry(caja, textvariable=self.var_puerto, width=6)).pack(side="left")
        abajo = ttk.Frame(f)
        abajo.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        manual(ttk.Checkbutton(abajo, text="Recordar la contraseña",
                               variable=self.var_recordar)).pack(side="left")
        manual(ttk.Button(abajo, text="Probar conexión",
                          command=self.probar_conexion)).pack(side="right")
        ttk.Label(f, style="Gris.TLabel",
                  text="Si hay sesión de Gmail se usa Gmail, y esto queda apagado.").grid(
            row=5, column=0, columnspan=2, sticky="w", pady=(6, 0))

    def _mostrar_otro(self, abrir):
        self.otro_abierto = abrir
        if abrir:
            self.f_otro.grid()
        else:
            self.f_otro.grid_remove()
        self.btn_otro.configure(text=("▾ " if abrir else "▸ ") + "Usar otro correo "
                                "(Outlook, Yahoo...) con contraseña de aplicación")

    def _alternar_otro(self):
        self._mostrar_otro(not self.otro_abierto)

    def cerrar_config(self):
        win = self.ventana_config
        if win is None:
            return
        guardar_config(self.leer_config())
        self.ventana_config = None
        self.btn_gmail = self.btn_salir = None
        self.campos_manuales = []
        win.grab_release()
        win.destroy()
        self.pintar_cuenta()

    def cambiar_proveedor(self, *_):
        servidor, puerto = PROVEEDORES.get(self.var_proveedor.get(), ("", 587))
        if servidor:
            self.var_servidor.set(servidor)
            self.var_puerto.set(str(puerto))

    def abrir_csv(self):
        padre = self.ventana_config or self
        guardar_empleados(self.empleados)
        try:
            os.startfile(str(ARCHIVO_EMPLEADOS))
        except Exception as e:
            messagebox.showerror("Recibos", str(e), parent=padre)
            return
        messagebox.showinfo(
            "Recibos",
            "Escribe la lista en Excel (nombre, cédula y correo), guárdala y "
            "cierra Excel.\n\nDespués dale a Aceptar para cargarla.", parent=padre)
        self.empleados = cargar_empleados()
        self._reemparejar()

    # ---- sesion de Gmail

    def pintar_cuenta(self):
        if self.sesion_gmail:
            correo = self.sesion_gmail.get("correo") or "tu cuenta de Gmail"
            self.var_sesion.set("Conectado como %s" % correo)
            self.var_cuenta.set("Se envía desde %s" % correo)
        else:
            self.var_sesion.set("No has iniciado sesión.")
            remitente = self.var_remitente.get().strip()
            if remitente and self.var_password.get():
                self.var_cuenta.set("Se envía desde %s" % remitente)
            else:
                self.var_cuenta.set("⚠ Falta configurar el correo")
        if self.btn_gmail is not None:
            self.btn_gmail.configure(text="Cambiar de cuenta" if self.sesion_gmail
                                     else "Iniciar sesión con Gmail")
            self.btn_salir.configure(state="normal" if self.sesion_gmail else "disabled")
        for widget, normal in self.campos_manuales:
            widget.configure(state="disabled" if self.sesion_gmail else normal)

    def iniciar_gmail(self):
        if self.trabajando:
            return
        if not buscar_client_secret():
            messagebox.showwarning(
                "Recibos",
                "Este programa todavía no tiene configurado el inicio de sesión "
                "con Gmail (falta el archivo client_secret.json).\n\n"
                "Mira el README, parte 'Iniciar sesión con Gmail'. Mientras "
                "tanto se puede usar otro correo con contraseña de aplicación.",
                parent=self.ventana_config or self)
            return
        self._ocupado(True)
        self.var_sesion.set("Termina de iniciar sesión en el navegador...")
        self.log("Se abrio el navegador para iniciar sesion con Gmail.")
        threading.Thread(target=self._hilo_login, daemon=True).start()

    def _hilo_login(self):
        try:
            sesion = gmail_login()
        except Exception as e:
            self.mensajes.put(("log", "No se pudo iniciar sesion: %s" % detalle(e)))
            self.mensajes.put(("sesion", self.sesion_gmail))
            self.mensajes.put(("fin", ""))
            self.mensajes.put(("error", "No se pudo iniciar sesión con Gmail:\n%s"
                               % detalle(e)))
            return
        self.mensajes.put(("sesion", sesion))
        self.mensajes.put(("log", "Sesion iniciada como %s"
                           % (sesion.get("correo") or "Gmail")))
        self.mensajes.put(("fin", ""))

    def cerrar_sesion_gmail(self):
        if not self.sesion_gmail or self.trabajando:
            return
        if not messagebox.askyesno("Recibos",
                                   "¿Cerrar la sesión de Gmail en esta computadora?",
                                   parent=self.ventana_config or self):
            return
        threading.Thread(target=revocar_gmail,
                         args=(self.sesion_gmail.get("refresh_token"),),
                         daemon=True).start()
        borrar_sesion_gmail()
        self.sesion_gmail = None
        self.pintar_cuenta()
        self.log("Sesion de Gmail cerrada.")

    # ---- envio

    def _falta_smtp(self, cfg):
        if not cfg["remitente"] or "@" not in cfg["remitente"]:
            return "Falta poner el correo que envía."
        if not cfg["password"]:
            return "Falta la contraseña de aplicación."
        if not cfg["servidor"]:
            return "Falta el servidor de correo."
        return None

    def conectar(self, cfg):
        if cfg["puerto"] == 465:
            servidor = smtplib.SMTP_SSL(cfg["servidor"], cfg["puerto"], timeout=40)
        else:
            servidor = smtplib.SMTP(cfg["servidor"], cfg["puerto"], timeout=40)
            servidor.starttls(context=ssl.create_default_context())
        servidor.login(cfg["remitente"], cfg["password"])
        return servidor

    def probar_conexion(self):
        padre = self.ventana_config or self
        cfg = self.leer_config()
        falta = self._falta_smtp(cfg)
        if falta:
            messagebox.showwarning("Recibos", falta, parent=padre)
            return
        guardar_config(cfg)
        try:
            self.conectar(cfg).quit()
        except Exception as e:
            self.log("Conexion fallida: %s" % e)
            messagebox.showerror(
                "Recibos",
                "No se pudo entrar al correo:\n%s\n\nRecuerda que va la contraseña "
                "de aplicación, no la contraseña normal." % e, parent=padre)
            return
        self.log("Conexion correcta con %s" % cfg["servidor"])
        self.pintar_cuenta()
        messagebox.showinfo("Recibos", "Todo bien, el correo funciona.", parent=padre)

    def enviar(self):
        if self.trabajando:
            return
        cfg = self.leer_config()
        sesion = self.sesion_gmail
        if not sesion and self._falta_smtp(cfg):
            messagebox.showinfo("Recibos", "Antes de enviar hay que poner el correo.\n\n"
                                "En Configuración, dale a Iniciar sesión con Gmail.")
            self.abrir_config()
            return
        guardar_config(cfg)
        remitente = (sesion.get("correo") or "tu cuenta de Gmail") if sesion \
            else cfg["remitente"]
        elegidos = [(i, e) for i, e in enumerate(self.empleados)
                    if e["marcado"] and e["pdfs"]]
        if not elegidos:
            messagebox.showinfo("Recibos", "No hay nadie seleccionado.\n\nMarca las "
                                "casillas verdes, o dale a Seleccionar todos.")
            return
        sin_correo = [e["nombre"] for _, e in elegidos if "@" not in e["correo"]]
        listos = [(i, e) for i, e in elegidos if "@" in e["correo"]]
        if not listos:
            messagebox.showwarning("Recibos", "Ninguno de los seleccionados tiene correo.")
            return
        aviso = "Se van a enviar %d recibos desde\n%s" % (len(listos), remitente)
        if sin_correo:
            aviso += "\n\nNo tienen correo y se saltan: %s" % ", ".join(sin_correo[:8])
        if not messagebox.askyesno("Recibos", aviso + "\n\n¿Enviar?"):
            return
        self._ocupado(True)
        threading.Thread(target=self._hilo_enviar, args=(dict(cfg), sesion, listos),
                         daemon=True).start()

    def _hilo_enviar(self, cfg, sesion, listos):
        total = len(listos)
        enviados, fallidos = 0, []
        servidor = token = None
        self.mensajes.put(("progreso", (0, total, "Conectando...")))
        try:
            if sesion:
                remitente = sesion.get("correo", "")
                token = gmail_access_token(sesion)
            else:
                remitente = cfg["remitente"]
                servidor = self.conectar(cfg)
        except Exception as e:
            texto = detalle(e)
            self.mensajes.put(("log", "ERROR de conexion: %s" % texto))
            ayuda = ("En Configuración, cierra sesión y vuelve a iniciar sesión con Gmail."
                     if sesion else "Revisa el correo y la contraseña en Configuración.")
            self.mensajes.put(("progreso", (0, total, "")))
            self.mensajes.put(("fin", "✖ No se pudo entrar al correo"))
            self.mensajes.put(("error", "No se pudo entrar al correo:\n%s\n\n%s"
                               % (texto, ayuda)))
            return

        for n, (i, emp) in enumerate(listos, 1):
            self.mensajes.put(("fila", (i, "enviando")))
            self.mensajes.put(("progreso", (n - 1, total, "Enviando %d de %d..." % (n, total))))
            try:
                msg = self._armar_correo(cfg, remitente, emp)
                if servidor:
                    servidor.send_message(msg)
                else:
                    try:
                        gmail_enviar(token, msg)
                    except urllib.error.HTTPError as e:
                        if e.code != 401:
                            raise
                        token = gmail_access_token(sesion)  # vencio la hora
                        gmail_enviar(token, msg)
                enviados += 1
                self.mensajes.put(("fila", (i, "enviado")))
                self.mensajes.put(("log", "OK  %s -> %s" % (emp["nombre"], emp["correo"])))
            except Exception as e:
                texto = detalle(e)
                fallidos.append("%s: %s" % (emp["nombre"], texto))
                self.mensajes.put(("fila", (i, "fallo")))
                self.mensajes.put(("log", "FALLO  %s -> %s : %s" % (
                    emp["nombre"], emp["correo"], texto)))
        if servidor:
            try:
                servidor.quit()
            except Exception:
                pass
        self.mensajes.put(("progreso", (total, total, "")))
        self.mensajes.put(("log", "Terminado: %d enviados, %d con fallo"
                           % (enviados, len(fallidos))))
        if fallidos:
            self.mensajes.put(("fin", "✔ %d enviados\n✖ %d no se enviaron"
                               % (enviados, len(fallidos))))
            self.mensajes.put(("error", "%d no se pudieron enviar:\n\n%s"
                               % (len(fallidos), "\n".join(fallidos[:6]))))
        else:
            self.mensajes.put(("fin", "✔ Listo, %d enviados" % enviados))

    def _armar_correo(self, cfg, remitente, emp):
        msg = EmailMessage()
        if remitente:
            msg["From"] = remitente
        msg["To"] = emp["correo"]
        msg["Subject"] = self._plantilla(cfg["asunto"], emp)
        msg.set_content(self._plantilla(cfg["cuerpo"], emp))
        for pdf in emp["pdfs"]:
            with open(pdf, "rb") as f:
                msg.add_attachment(f.read(), maintype="application",
                                   subtype="pdf", filename=pdf.name)
        return msg

    @staticmethod
    def _plantilla(texto, emp):
        return (texto.replace("{nombre}", emp["nombre"])
                     .replace("{cedula}", emp["cedula"]))

    # ---- impresion

    def imprimir(self):
        if self.trabajando:
            return
        elegidos = [(i, e) for i, e in enumerate(self.empleados)
                    if e["marcado"] and e["pdfs"]]
        if not elegidos:
            messagebox.showinfo("Recibos", "No hay nadie seleccionado.\n\nMarca las "
                                "casillas verdes, o dale a Seleccionar todos.")
            return
        if not messagebox.askyesno(
                "Recibos",
                "Se van a imprimir %d recibos en la impresora que Windows tiene "
                "puesta por defecto.\n\n¿Imprimir?" % len(elegidos)):
            return
        self._ocupado(True)
        threading.Thread(target=self._hilo_imprimir, args=(elegidos,), daemon=True).start()

    def _hilo_imprimir(self, elegidos):
        total = len(elegidos)
        ok = malos = 0
        for n, (i, emp) in enumerate(elegidos, 1):
            self.mensajes.put(("progreso", (n - 1, total,
                                            "Imprimiendo %d de %d..." % (n, total))))
            try:
                for pdf in emp["pdfs"]:
                    os.startfile(str(pdf), "print")
                ok += 1
                self.mensajes.put(("fila", (i, "impreso")))
                self.mensajes.put(("log", "Impreso: %s" % emp["nombre"]))
            except Exception as e:
                malos += 1
                self.mensajes.put(("fila", (i, "noimpreso")))
                self.mensajes.put(("log", "No se pudo imprimir %s : %s" % (emp["nombre"], e)))
            time.sleep(1.5)
        self.mensajes.put(("progreso", (total, total, "")))
        if malos:
            self.mensajes.put(("fin", "✔ %d impresos\n✖ %d no se imprimieron" % (ok, malos)))
            self.mensajes.put((
                "error",
                "No se pudieron imprimir %d recibos.\n\nSi no sale nada, abre un PDF con "
                "clic derecho > Abrir con > Adobe Reader, y ponlo como programa por "
                "defecto para los PDF." % malos))
        else:
            self.mensajes.put(("fin", "✔ %d mandados a la impresora" % ok))


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
