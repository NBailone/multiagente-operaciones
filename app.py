"""
Sistema de Automatización de Operaciones — Interfaz Profesional
================================================================
Aplicación de escritorio para la gestión de agentes operativos:
  1. Agente de Impresión Documental
  2. Agente de Procesamiento de Contenedores
  3. Agente de Despacho de Correos

Construido con CustomTkinter sobre Tkinter.
"""

import sys
import os
import re
import threading
import queue
import time
import traceback
import json
import subprocess
import string
import imaplib
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import email
import shutil
import tempfile
import dotenv
from copy import copy
import zipfile
import xml.etree.ElementTree as ET
from PIL import Image

# ── Auto-install UI dependencies ──────────────────────────────────────────
def _instalar_deps_ui():
    deps = [
        ("customtkinter", "customtkinter"),
        ("openpyxl", "openpyxl"),
        ("xlrd", "xlrd"),
        ("win32com", "pywin32"),
    ]
    for mod, pip_name in deps:
        try:
            __import__(mod)
        except ImportError:
            print(f"[!] Instalando {pip_name}...")
            try:
                subprocess.check_call(
                    [sys.executable, "-m", "pip", "install", pip_name],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
                # pywin32 necesita post-instalación para registrar las DLLs COM
                if pip_name == "pywin32":
                    try:
                        subprocess.check_call(
                            [sys.executable, "-m", "pywin32_postinstall", "-install"],
                            stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL,
                        )
                    except Exception:
                        pass  # no es fatal si falla
                print(f"[OK] {pip_name} instalado.")
            except Exception as e:
                print(f"[X] No se pudo instalar {pip_name}: {e}")

if not getattr(sys, 'frozen', False):
    _instalar_deps_ui()

import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
import xlrd

from constants import Palette, FONT_FAMILY, FONT_MONO, FONT_LEVEL_SCALES, FONT_BASE_SIZES
from constants import IMAP_SERVER, PUERTO_IMAP
from constants import DESTINATARIOS_GRUPAL, DESTINATARIOS_INDIVIDUAL

from utils import (buscar_archivo_en_pendrive, formatear_fecha_excel,
                    adjuntar_archivo, preguntar_reintentar,
                    celda_es_mergeada, primera_fila_libre, ya_existe_en_hoja,
                    buscar_bl_por_carpeta_xlsx, buscar_bl_por_carpeta_xls)

from panels import (
    ImpresionMixin, PlanillasMixin, CorreosMixin, DescargaMixin,
    SuperAutoMixin, BackupMixin, ControlMixin, AjustesMixin,
)

# ── .env loading for secret key ─────────────────────────────────────────
_frozen = getattr(sys, 'frozen', False)
_base_dir = os.path.dirname(sys.executable) if _frozen else os.path.dirname(__file__)
_dotenv_path = os.path.join(_base_dir, ".env")

dotenv.load_dotenv(_dotenv_path)

def _ensure_secret_key(dotenv_path: str) -> None:
    key = os.environ.get("MULTIAGENTE_SECRET_KEY")
    if key:
        return
    import secrets
    key = secrets.token_hex(32)
    try:
        with open(dotenv_path, "w", encoding="utf-8") as f:
            f.write(f"MULTIAGENTE_SECRET_KEY={key}\n")
    except OSError:
        print(f"[WARN] Could not write {dotenv_path} — key in memory only for this session.")
    os.environ["MULTIAGENTE_SECRET_KEY"] = key

_ensure_secret_key(_dotenv_path)

# ── Redirector de stdout al widget de log ───────────────────────────────
class OutputRedirector:
    """Captura stdout/stderr y los redirige al widget de log con colores."""
    def __init__(self, log_widget, app):
        self.log_widget = log_widget
        self.app = app
        self.buffer = ""

    def write(self, text):
        self.buffer += text
        if "\n" in self.buffer:
            lines = self.buffer.split("\n")
            self.buffer = lines.pop()
            for line in lines:
                if line.strip():
                    self.app.emit_log(line.strip())
        # flush on \n line endings

    def flush(self):
        if self.buffer.strip():
            self.app.emit_log(self.buffer.strip())
            self.buffer = ""

# ── Ventana Principal ───────────────────────────────────────────────────
class App(ctk.CTk, ImpresionMixin, PlanillasMixin, CorreosMixin, DescargaMixin,
                 SuperAutoMixin, BackupMixin, ControlMixin, AjustesMixin):
    def __init__(self):
        super().__init__()

        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(__file__)
        self.config_file = os.path.join(base_dir, "ui_config.json")
        self.config = {}
        self._master_pw_cache = ""
        self._cargar_config()

        # ── Configuración de ventana (antes de mostrar) ───────────────
        self.title("Sistema de Automatización de Operaciones")
        self.minsize(900, 550)
        self.configure(fg_color=Palette.BG_MAIN)

        # Restaurar geometría guardada o usar 80% de la pantalla
        if "window_geo" in self.config:
            self.geometry(self.config["window_geo"])
        else:
            self._centrar_ventana()

        # Forzar centrado si la ventana quedó fuera de pantalla (ej. monitor más chico)
        self.update_idletasks()
        self.after(100, self._asegurar_visible)

        # Icono (si existe)
        icon_dir = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
        icon_path = os.path.join(icon_dir, "icono.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        # ── Contraseña maestra al iniciar ─────────────────────────────
        # Se ejecuta vía after() al final de __init__ para evitar conflictos
        pw = self._cfg_obtener("seguridad", "password", "")
        self._pw_inicio_valida = not pw  # si no hay pass, ya está validado

        # ── Cola de logs thread-safe ──────────────────────────────────
        self.log_queue = queue.Queue()
        self._poll_log_queue()

        # ── Estados ───────────────────────────────────────────────────
        self.tarea_activa = False
        self._cancelar_tarea = threading.Event()  # flag para cancelar tareas en curso
        self._log_ctx = threading.local()  # per-thread log panel context
        self.datos_planillas = []  # Resultados del agente 2
        self._mail_data = {}       # {item_id: {mid, subject, date, checked, downloaded}}
        self.panel_actual = ""      # Panel activo actual
        self._panel_frames = {}    # nombre -> CTkFrame, persistente entre cambios de panel
        self._resultados_pendientes = None  # Resultados para popup de resumen
        self._excel_com_ok = True
        self._super_auto = False          # Toggle de súper automatización
        self._super_guarda = ""           # Guarda elegido para súper auto
        self._ultimas_carpetas = []       # Carpetas descargadas en la última tanda
        self._backup_resultados = None    # Resultados del último backup
        self.logs_por_panel = {    # Historial de consola por agente
            "impresion": [],
            "planillas": [],
            "descargar": [],
            "correos": [],
            "backup": [],
            "cargar-datos": [],
        }
        # T7/T8: almacén de rutas CONTENEDOR + k por fila TreeView
        self._cargar_datos_rutas = {}
        self._cargar_datos_comparacion = {}  # iid → dict para popup comparación
        self._cargar_datos_idx = 0
        self._cf_auto_var = ctk.BooleanVar(value=False)  # Control Final auto mode
        self._ct_auto_var = ctk.BooleanVar(value=False)  # Control Tickets auto mode

        # ── Construir UI ──────────────────────────────────────────────
        self._crear_sidebar()
        self._crear_area_principal()

        # Seleccionar panel inicial
        self._cambiar_panel("descargar")

        # Bind de teclas
        self.bind("<Escape>", lambda e: self._confirmar_salida())
        self.protocol("WM_DELETE_WINDOW", self._confirmar_salida)
        # Destroy popups on minimize to prevent grab deadlock
        self._open_popups = set()
        self.bind("<Unmap>", self._on_main_unmap)

        # Verificar contraseña maestra al iniciar, luego diagnóstico
        self.after(300, self._verificar_inicio)

    def _centrar_ventana(self):
        """Centra la ventana en la pantalla actual al 80% del tamaño."""
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = min(int(sw * 0.8), 1280)
        h = min(int(sh * 0.8), 800)
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _asegurar_visible(self):
        """Si la ventana está fuera del área visible, la re-centra."""
        try:
            x = self.winfo_x()
            y = self.winfo_y()
            w = self.winfo_width()
            h = self.winfo_height()
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            ajustar = False
            if x + w < 100:  # muy a la izquierda o invisible
                x = max(0, (sw - w) // 2)
                ajustar = True
            if y + h < 50:  # muy arriba
                y = max(0, (sh - h) // 2)
                ajustar = True
            if x > sw - 50:  # muy a la derecha
                x = max(0, (sw - w) // 2)
                ajustar = True
            if ajustar:
                self.geometry(f"+{x}+{y}")
        except Exception:
            pass

    def _register_popup(self, popup):
        """Track a Toplevel popup so it's destroyed on minimize."""
        self._open_popups.add(popup)
        popup.bind("<Destroy>", lambda e, p=popup: self._open_popups.discard(p))

    def _on_main_unmap(self, event):
        """When main window is minimized, destroy all open popups to release grabs."""
        for popup in list(self._open_popups):
            try:
                popup.grab_release()
            except Exception:
                pass
            try:
                popup.destroy()
            except Exception:
                pass
        self._open_popups.clear()
        # Also destroy any untracked Toplevel children (safety net)
        for w in self.winfo_children():
            if isinstance(w, ctk.CTkToplevel):
                try:
                    w.grab_release()
                except Exception:
                    pass
                try:
                    w.destroy()
                except Exception:
                    pass

    def _verificar_inicio(self):
        if not self._pw_inicio_valida:
            self._mostrar_dialogo_password()

    def _mostrar_dialogo_password(self):
        pw_config = self._cfg_obtener("seguridad", "password", "")
        dlg = ctk.CTkToplevel(self)
        dlg.title("Acceso Restringido")
        dlg.geometry("400x215")
        dlg.configure(fg_color=Palette.BG_CARD)
        dlg.transient(self)
        dlg.lift()
        dlg.grab_set()
        dlg.minsize(400, 215)
        dlg.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 400) // 2
        py = self.winfo_y() + (self.winfo_height() - 215) // 2
        dlg.geometry(f"400x215+{px}+{py}")

        ctk.CTkLabel(
            dlg, text="Ingresá la contraseña para abrir la aplicación:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(pady=(16, 6))

        entry_pw = ctk.CTkEntry(
            dlg, width=280, height=36, show="*",
            font=ctk.CTkFont(family=FONT_FAMILY, size=14),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=6,
        )
        entry_pw.pack(pady=(0, 4))
        entry_pw.focus_set()
        dlg.after(50, lambda: entry_pw.focus_force())

        def verificar():
            if entry_pw.get() == pw_config:
                self._master_pw_cache = entry_pw.get()
                dlg.destroy()
                self._pw_inicio_valida = True
            else:
                lbl_status.configure(text="Contraseña incorrecta. Intentá de nuevo.", text_color=Palette.ERROR)
                entry_pw.delete(0, "end")
                entry_pw.focus_set()

        entry_pw.bind("<Return>", lambda e: verificar())

        # Botón de recuperación (entre la entry y los botones)
        def recuperar():
            self._recuperar_password(dlg, lbl_status)

        ctk.CTkButton(
            dlg, text="🔑  Recuperar Clave",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.WARNING, hover_color=Palette.WARNING_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=30, width=180,
            command=recuperar,
        ).pack(pady=(4, 2))

        # Un solo label para errores y recuperación
        lbl_status = ctk.CTkLabel(
            dlg, text="",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_SECONDARY,
        )
        lbl_status.pack(pady=(0, 2))

        btn_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_frame.pack(pady=(0, 4))
        ctk.CTkButton(
            btn_frame, text="Ingresar", width=110, height=36,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6,
            command=verificar,
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            btn_frame, text="Salir", width=110, height=36,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            fg_color=Palette.BG_HOVER, hover_color=Palette.ERROR,
            text_color=Palette.TEXT_SECONDARY, corner_radius=6,
            command=lambda: (dlg.destroy(), self.destroy()),
        ).pack(side="left", padx=5)

    # ── Sidebar ──────────────────────────────────────────────────────────
    def _crear_sidebar(self):
        self.sidebar = ctk.CTkFrame(
            self, width=230, corner_radius=0, fg_color=Palette.BG_SIDEBAR
        )
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # ── Logo / Título ────────────────────────────────────────────
        header = ctk.CTkFrame(
            self.sidebar, fg_color="transparent", height=80
        )
        header.pack(fill="x", padx=14, pady=(4, 0))
        header.pack_propagate(False)

        # Cargar icono de la aplicación
        try:
            icon_dir = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
            icon_path = os.path.join(icon_dir, "icono.ico")
            if os.path.exists(icon_path):
                pil_image = Image.open(icon_path)
                pil_image = pil_image.resize((72, 72), Image.LANCZOS)
                self._logo_icon = ctk.CTkImage(pil_image, size=(72, 72))
                ctk.CTkLabel(header, image=self._logo_icon, text="").pack(expand=True)
            else:
                ctk.CTkLabel(
                    header, text="MULTIAGENTE",
                    font=ctk.CTkFont(family=FONT_FAMILY, size=20, weight="bold"),
                    text_color=Palette.TEXT_PRIMARY,
                ).pack(anchor="w")
        except Exception:
            ctk.CTkLabel(
                header, text="MULTIAGENTE",
                font=ctk.CTkFont(family=FONT_FAMILY, size=20, weight="bold"),
                text_color=Palette.TEXT_PRIMARY,
            ).pack(anchor="w")

        ctk.CTkFrame(
            self.sidebar, fg_color=Palette.DIVIDER, height=1
        ).pack(fill="x", padx=14, pady=(4, 4))

        # ── Zona de navegación con scroll ─────────────────────────────
        self._nav_scroll = ctk.CTkScrollableFrame(
            self.sidebar, fg_color="transparent",
            scrollbar_button_color=Palette.TEXT_MUTED,
            scrollbar_button_hover_color=Palette.TEXT_SECONDARY,
        )
        self._nav_scroll.pack(fill="both", expand=True, padx=0, pady=0)

        nav_label = ctk.CTkLabel(
            self._nav_scroll,
            text="MÓDULOS",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            text_color=Palette.TEXT_MUTED,
        )
        nav_label.pack(anchor="w", padx=14, pady=(6, 4))

        # ── Iconos Lucide (compartidos por todos los paneles) ─────────
        _icon_dir = os.path.join(os.path.dirname(__file__), "assets", "icons")
        _icon_size = (20, 20)
        self._icons = {}
        for _name in ["inbox", "file-text", "table", "clipboard-list", "send", "save",
                       "settings", "mail-search", "search", "mail", "cloud-download",
                       "refresh-cw", "bookmark", "printer", "play", "shield", "pencil",
                       "check", "x", "ticket", "zap", "folder-open", "hard-drive",
                       "folder", "banknote", "lock", "bot", "palette", "eye"]:
            self._icons[_name] = ctk.CTkImage(
                Image.open(os.path.join(_icon_dir, f"{_name}.png")), size=_icon_size
            )

        self.btn_descargar = self._crear_btn_nav(
            "Descargar Mails", "inbox", 0, lambda: self._cambiar_panel("descargar"),
            image=self._icons["inbox"],
        )
        self.btn_impresion = self._crear_btn_nav(
            "Impresión Documental", "file-text", 1, lambda: self._cambiar_panel("impresion"),
            image=self._icons["file-text"],
        )
        self.btn_planillas = self._crear_btn_nav(
            "Completar Planillas", "table", 2, lambda: self._cambiar_panel("planillas"),
            image=self._icons["table"],
        )
        self.btn_cargar_datos = self._crear_btn_nav(
            "Controlar Datos", "clipboard-list", 3, lambda: self._cambiar_panel("cargar-datos"),
            image=self._icons["clipboard-list"],
        )
        self.btn_correos = self._crear_btn_nav(
            "Enviar Correos", "send", 4, lambda: self._cambiar_panel("correos"),
            image=self._icons["send"],
        )
        self.btn_backup = self._crear_btn_nav(
            "Backup", "save", 5, lambda: self._cambiar_panel("backup"),
            image=self._icons["save"],
        )

        self.btn_ajustes = self._crear_btn_nav(
            "Ajustes", "settings", 6, lambda: self._cambiar_panel("ajustes"),
            image=self._icons["settings"],
        )

        # ── Separador ────────────────────────────────────────────────
        ctk.CTkFrame(
            self._nav_scroll, fg_color=Palette.DIVIDER, height=1
        ).pack(fill="x", padx=14, pady=(16, 10))

        # ── Súper Auto Toggle ────────────────────────────────────────

        ctk.CTkLabel(
            self._nav_scroll, text="AUTOMATIZACIÓN",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            text_color=Palette.TEXT_MUTED,
        ).pack(anchor="w", padx=14, pady=(0, 4))

        self._super_switch = ctk.CTkSwitch(
            self._nav_scroll, text="⚡ Súper Auto",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13),
            progress_color=Palette.ACCENT,
            button_color=Palette.TEXT_SECONDARY,
            button_hover_color=Palette.TEXT_PRIMARY,
            command=self._super_toggle,
        )
        self._super_switch.pack(anchor="w", padx=12, pady=(2, 2))

        self._super_lbl_guarda = ctk.CTkLabel(
            self._nav_scroll, text="",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_MUTED,
        )
        self._super_lbl_guarda.pack(anchor="w", padx=30, pady=(0, 6))

        # ── Botón Salir ──────────────────────────────────────────────
        self.btn_salir = ctk.CTkButton(
            self.sidebar,
            text="Salir del Sistema",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color="transparent",
            hover_color=Palette.ERROR_BG,
            text_color=Palette.TEXT_PRIMARY,
            border_color=Palette.BORDER,
            border_width=1,
            corner_radius=6,
            height=36,
            command=self._confirmar_salida,
        )
        self.btn_salir.pack(side="bottom", fill="x", padx=10, pady=10)

    def _crear_btn_nav(self, texto, icono, idx, comando, image=None):
        btn = ctk.CTkButton(
            self._nav_scroll,
            text=f"  {texto}",
            image=image, compound="left" if image else "none",
            font=ctk.CTkFont(family=FONT_FAMILY, size=14),
            fg_color="transparent",
            hover_color=Palette.BG_HOVER,
            text_color=Palette.WHITE,
            anchor="w",
            corner_radius=8,
            height=46,
            command=comando,
        )
        btn.pack(fill="x", padx=8, pady=2)
        btn._nav_idx = idx
        return btn

    def _marcar_nav_activo(self, idx):
        botones = [self.btn_descargar, self.btn_impresion, self.btn_planillas, self.btn_cargar_datos, self.btn_correos, self.btn_backup, self.btn_ajustes]
        for i, b in enumerate(botones):
            if i == idx:
                b.configure(
                    fg_color=Palette.ACCENT,
                    hover_color=Palette.ACCENT_HOVER,
                    text_color=Palette.WHITE,
                )
            else:
                b.configure(
                    fg_color="transparent",
                    hover_color=Palette.BG_HOVER,
                    text_color=Palette.WHITE,
                )

    # ── Área Principal ───────────────────────────────────────────────────
    def _crear_area_principal(self):
        self.main_area = ctk.CTkFrame(
            self, corner_radius=0, fg_color=Palette.BG_MAIN
        )
        self.main_area.pack(side="left", fill="both", expand=True)

        # ── Barra superior ───────────────────────────────────────────
        self.top_bar = ctk.CTkFrame(
            self.main_area, fg_color="transparent", height=56
        )
        self.top_bar.pack(fill="x", padx=16, pady=(16, 0))
        self.top_bar.pack_propagate(False)

        self.lbl_titulo_panel = ctk.CTkLabel(
            self.top_bar,
            text="",
            font=ctk.CTkFont(family=FONT_FAMILY, size=20, weight="bold"),
            text_color=Palette.TEXT_PRIMARY,
        )
        self.lbl_titulo_panel.pack(side="left")

        # ── Controles OCR (mostrar solo en panel Control de Datos) ────
        self._ocr_topbar_frame = ctk.CTkFrame(
            self.top_bar, fg_color="transparent"
        )
        # Se packea/empaca dinámicamente en _cambiar_panel

        ctk.CTkLabel(
            self._ocr_topbar_frame,
            text="Modelo OCR:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(side="left")

        self._ocr_method_menu = ctk.CTkOptionMenu(
            self._ocr_topbar_frame,
            values=["Local", "API Visión"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=Palette.BG_INPUT,
            button_color=Palette.ACCENT,
            button_hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.TEXT_PRIMARY,
            dropdown_fg_color=Palette.BG_CARD,
            dropdown_hover_color=Palette.BG_HOVER,
            dropdown_text_color=Palette.TEXT_PRIMARY,
            corner_radius=4, width=130, height=28,
            command=self._on_ocr_method_change,
        )
        self._ocr_method_menu.pack(side="left", padx=(4, 16))

        ctk.CTkLabel(
            self._ocr_topbar_frame,
            text="Modelo:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(side="left")

        self._modelo_vision_menu = ctk.CTkOptionMenu(
            self._ocr_topbar_frame,
            values=[],
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=Palette.BG_INPUT,
            button_color=Palette.ACCENT,
            button_hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.TEXT_PRIMARY,
            dropdown_fg_color=Palette.BG_CARD,
            dropdown_hover_color=Palette.BG_HOVER,
            dropdown_text_color=Palette.TEXT_PRIMARY,
            corner_radius=4, width=200, height=28,
            state="disabled",
            command=self._on_modelo_vision_change,
        )
        self._modelo_vision_menu.pack(side="left")

        # Oculto por defecto — se muestra en _cambiar_panel para cargar-datos
        self._ocr_topbar_frame.pack_forget()

        self.lbl_badge = ctk.CTkLabel(
            self.top_bar,
            text="",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
            text_color=Palette.WHITE,
            fg_color=Palette.ACCENT,
            corner_radius=4,
            padx=8,
            pady=2,
        )
        # badge shown conditionally

        # ── Contenedor de paneles (Resizable PanedWindow) ────────────
        self.paned_window = tk.PanedWindow(
            self.main_area, orient="vertical",
            bg=Palette.BORDER, bd=0, sashwidth=4,
            relief="flat"
        )
        self.paned_window.pack(fill="both", expand=True, padx=16, pady=(8, 16))

        self.panel_container = ctk.CTkFrame(
            self.paned_window, fg_color="transparent"
        )
        self.paned_window.add(self.panel_container, minsize=200, stretch="always")

        # ── Área de log (inferior, compartida) ────────────────────────
        self.log_container = ctk.CTkFrame(self.paned_window, fg_color="transparent")
        self.paned_window.add(self.log_container, minsize=100, stretch="never")
        self._crear_panel_log(self.log_container)

        # Restaurar sash del panel inicial
        if "sash_planillas" in self.config:
            self.after(200, lambda: self.paned_window.sash_place(0, 0, self.config["sash_planillas"]))

    def _crear_panel_log(self, parent):
        log_frame = ctk.CTkFrame(
            parent, fg_color=Palette.BG_LOG, corner_radius=8,
            border_width=1, border_color=Palette.BORDER
        )
        log_frame.pack(fill="both", expand=True, padx=0, pady=(8, 0))

        # Header del log
        log_header = ctk.CTkFrame(log_frame, fg_color="transparent", height=28)
        log_header.pack(fill="x", padx=12, pady=(6, 0))
        log_header.pack_propagate(False)

        ctk.CTkLabel(
            log_header,
            text="CONSOLA DE SALIDA",
            font=ctk.CTkFont(family=FONT_FAMILY, size=15, weight="bold"),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(side="left")

        self.lbl_log_count = ctk.CTkLabel(
            log_header,
            text="",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13),
            text_color=Palette.TEXT_SECONDARY,
        )
        self.lbl_log_count.pack(side="right")

        self.btn_clear_log = ctk.CTkButton(
            log_header,
            text="Limpiar",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color="transparent",
            hover_color=Palette.BG_HOVER,
            text_color=Palette.TEXT_PRIMARY,
            width=60,
            height=28,
            command=self._limpiar_log,
        )
        self.btn_clear_log.pack(side="right", padx=(0, 6))

        # Text widget con scroll
        self.log_text = ctk.CTkTextbox(
            log_frame,
            font=ctk.CTkFont(family=FONT_MONO, size=13),
            fg_color=Palette.BG_LOG,
            text_color=Palette.TEXT_SECONDARY,
            border_width=0,
            corner_radius=0,
            wrap="word",
        )
        self.log_text.pack(fill="both", expand=True, padx=8, pady=(4, 8))
        self.log_text.configure(state="disabled")

        # Configurar tags de color
        self.log_text.tag_config("error", foreground=Palette.ERROR)
        self.log_text.tag_config("success", foreground=Palette.SUCCESS)
        # warning eliminado — se usa color por defecto
        self.log_text.tag_config("info", foreground=Palette.INFO)

        self._log_line_count = 0
        self._log_tags = {}

    # ── Navegación de paneles ───────────────────────────────────────────
    def _cambiar_panel(self, nombre):
        """Navega al panel indicado sin verificar tareas activas."""
        # Verificar contraseña maestra para Ajustes (se hace en _forzado)
        self._cambiar_panel_forzado(nombre)

    def _cambiar_panel_forzado(self, nombre):
        # Verificar contraseña maestra para Ajustes
        if nombre == "ajustes" and not self._verificar_password_maestra():
            return

        # Guardar log del panel actual ANTES de ocultarlo
        if self.panel_actual and self.panel_actual in self.logs_por_panel:
            self.logs_por_panel[self.panel_actual] = self._capturar_lineas_log()

        # Ocultar el frame del panel actual (si existe)
        if self.panel_actual and self.panel_actual in self._panel_frames:
            self._panel_frames[self.panel_actual].pack_forget()

        # Construir el frame del panel destino si es primera vez
        if nombre not in self._panel_frames:
            if nombre == "impresion":
                self._panel_impresion()
            elif nombre == "planillas":
                self._panel_planillas()
            elif nombre == "descargar":
                self._panel_descargar()
            elif nombre == "correos":
                self._panel_correos()
            elif nombre == "backup":
                self._panel_backup()
            elif nombre == "ajustes":
                self._panel_ajustes()

        # "cargar-datos" se reintenta siempre — el guard interno detecta
        # si ya está completo (frame + tree_coordinacion) y sale rápido.
        if nombre == "cargar-datos":
            self._panel_cargar_datos()

        # Mostrar el frame del panel destino
        frame = self._panel_frames.get(nombre)
        if frame:
            frame.pack(in_=self.panel_container, fill="both", expand=True)

        # Reescanear carpetas al volver al panel de impresión
        if nombre == "impresion":
            self.after(50, self._imp_escanear_carpetas)

        # Actualizar título y navegación
        idx_map = {"descargar": 0, "impresion": 1, "planillas": 2, "cargar-datos": 3, "correos": 4, "backup": 5, "ajustes": 6}
        idx = idx_map.get(nombre, 1)
        self._marcar_nav_activo(idx)

        titles = {
            "impresion": "Impresión Documental",
            "planillas": "Completar Planillas",
            "descargar": "Descargar Mails",
            "correos": "Enviar Correos",
            "backup": "Backup",
            "cargar-datos": "Control de Datos",
            "ajustes": "Configuración del Sistema",
        }
        self.lbl_titulo_panel.configure(text=titles.get(nombre, ""))

        # Mostrar/ocultar controles OCR según panel
        if nombre == "cargar-datos":
            if not self._ocr_topbar_frame.winfo_ismapped():
                self._ocr_topbar_frame.pack(side="right")
                # Sincronizar método OCR desde config
                metodo_ocr = self.config.get("ocr_method", "api_vision")
                self._ocr_method_menu.set("API Visión" if metodo_ocr == "api_vision" else "Local")
                self._actualizar_modelos_vision()
                # Sincronizar state del modelo según método actual
                es_api = metodo_ocr == "api_vision"
                self._modelo_vision_menu.configure(state="normal" if es_api else "disabled")
        else:
            self._ocr_topbar_frame.pack_forget()

        if nombre == "planillas":
            self._actualizar_titulo_precintos()

        # Mostrar/ocultar consola según panel
        if nombre == "ajustes":
            self.log_container.pack_forget()
            # Sincronizar dropdown de modelo desde config actual
            if hasattr(self, '_ent_vision_model') and self._ent_vision_model.winfo_exists():
                import procesar_tickets
                modelo_default = self.config.get("api_vision", {}).get("model",
                                    procesar_tickets.MODELO_VISION_DEFAULT)
                modelos_guardados = self.config.get("api_vision", {}).get("custom_models", [])
                modelos = modelos_guardados if modelos_guardados else list(procesar_tickets.MODELOS_VISION)
                if modelo_default in modelos:
                    self._ent_vision_model.set(modelo_default)
                elif modelos:
                    self._ent_vision_model.set(modelos[0])
        else:
            if not self.log_container.winfo_ismapped():
                self.paned_window.add(self.log_container, minsize=100, stretch="never")

        # Guardar sash del panel anterior, restaurar sash del nuevo
        panel_anterior = self.panel_actual
        self.panel_actual = nombre
        if hasattr(self, 'paned_window') and self.paned_window.winfo_exists():
            try:
                sash_y = self.paned_window.sash_coord(0)[1]
                if panel_anterior and sash_y > 0:
                    self.config[f"sash_{panel_anterior}"] = sash_y
            except Exception:
                pass
            if f"sash_{nombre}" in self.config:
                self.after(150, lambda n=nombre: self._restaurar_sash(n))

        # Restaurar la consola guardada para este panel
        self._restaurar_log(nombre)

    def _restaurar_sash(self, nombre):
        try:
            self.paned_window.sash_place(0, 0, self.config.get(f"sash_{nombre}", 500))
        except Exception:
            pass

    # ═══════════════════════════════════════════════════════════════════
    # PANEL 1: IMPRESIÓN DOCUMENTAL
    # ═══════════════════════════════════════════════════════════════════

    # ═══════════════════════════════════════════════════════════════════

    # ═══════════════════════════════════════════════════════════════════


    def _mostrar_resumen(self, resultados):
        """Popup lindo con resumen de lo procesado (SOBRES, COBRO, PC)."""
        popup = ctk.CTkToplevel(self)
        popup.title("Resultado del Análisis")
        popup.geometry("440x340")
        popup.configure(fg_color=Palette.BG_CARD)
        popup.transient(self)
        popup.lift()

        # Centrar sobre la ventana principal
        popup.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 440) // 2
        py = self.winfo_y() + (self.winfo_height() - 340) // 2
        popup.geometry(f"440x340+{px}+{py}")

        # Título
        ctk.CTkLabel(
            popup, text="ANÁLISIS COMPLETADO",
            font=ctk.CTkFont(family=FONT_FAMILY, size=18, weight="bold"),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(pady=(24, 20))

        # Items
        items = [
            ("SOBRES", resultados.get("sobres", {})),
            ("COBRO", resultados.get("cobro", {})),
            ("PC (Precintos)", resultados.get("pc", {})),
        ]

        for nombre, res in items:
            ok = res.get("ok", False)
            detalle = res.get("detalle", "No procesado")
            icono = "✓" if ok else "✗"
            color = Palette.SUCCESS if ok else Palette.ERROR

            row_frame = ctk.CTkFrame(popup, fg_color="transparent")
            row_frame.pack(fill="x", padx=32, pady=4)

            ctk.CTkLabel(
                row_frame, text=icono,
                font=ctk.CTkFont(family=FONT_FAMILY, size=20, weight="bold"),
                text_color=color, width=30,
            ).pack(side="left")

            ctk.CTkLabel(
                row_frame, text=nombre,
                font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
                text_color=Palette.TEXT_PRIMARY, width=120, anchor="w",
            ).pack(side="left")

            ctk.CTkLabel(
                row_frame, text=detalle,
                font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                text_color=Palette.TEXT_SECONDARY,
            ).pack(side="left")

        # Botón cerrar
        ctk.CTkButton(
            popup, text="Cerrar", width=120, height=36,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6,
            command=popup.destroy,
        ).pack(pady=(24, 16))

    def _abrir_excel_seguro(self, ruta, solo_lectura=False):
        """Abre un archivo Excel. Si está abierto, pide cerrarlo y reintenta."""
        nombre = os.path.basename(ruta)
        while True:
            try:
                return openpyxl.load_workbook(ruta, data_only=False, read_only=solo_lectura)
            except PermissionError:
                if not preguntar_reintentar(nombre, parent=self):
                    raise

    def _guardar_excel_seguro(self, wb, ruta):
        """Guarda un archivo Excel. Si está abierto, pide cerrarlo y reintenta."""
        nombre = os.path.basename(ruta)
        while True:
            try:
                wb.save(ruta)
                return
            except PermissionError:
                if not preguntar_reintentar(nombre, parent=self):
                    raise

    # ═══════════════════════════════════════════════════════════════════
    # SISTEMA DE LOG
    # ═══════════════════════════════════════════════════════════════════
    def _log(self, mensaje):
        """Agrega un mensaje al log desde cualquier hilo, etiquetado con el panel de origen."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        panel = getattr(self._log_ctx, "panel", self.panel_actual) or "general"
        self.log_queue.put(("_LOG_", panel, f"[{timestamp}] {mensaje}"))

    def _set_log_panel(self, panel):
        """Set the panel context for the current thread. All _log() calls from this thread
        will be tagged with this panel name."""
        self._log_ctx.panel = panel

    def _log_warning(self, mensaje):
        """Advertencia suprimida — no se muestra en el log."""
        pass

    def _log_error(self, mensaje):
        """Agrega un mensaje de error al log (hilo principal)."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.emit_log(f"[{timestamp}] ERROR: {mensaje}")

    def emit_log(self, texto):
        """Inserta texto coloreado en el widget de log (hilo principal)."""
        self.log_text.configure(state="normal")
        if self._log_line_count > 0:
            self.log_text.insert("end", "\n")
            
        tag = None
        txt_up = texto.upper()
        if "ERROR" in txt_up or "CRÍTICO" in txt_up or "NO SE PUDO" in txt_up:
            tag = "error"
        elif "ÉXITO" in txt_up or "FINALIZADO" in txt_up or "CORRECTAMENTE" in txt_up or "COMPLETADO" in txt_up:
            tag = "success"
        # ADVERTENCIA, FALTA, NO SE ENCONTRÓ → color por defecto (eliminado el amarillo)
        elif "ESCANEANDO" in txt_up or "PREPARANDO" in txt_up or "CONECTANDO" in txt_up:
            tag = "info"

        if tag:
            self.log_text.insert("end", texto, tag)
        else:
            self.log_text.insert("end", texto)
            
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self._log_line_count += 1

        # Actualizar contador
        self.lbl_log_count.configure(text=f"{self._log_line_count} líneas")

        # Limitar líneas a 500
        if self._log_line_count > 500:
            self._trim_log()

    def _poll_log_queue(self):
        """Procesa la cola de mensajes de log periódicamente.
        
        Soporta strings (emit_log directo) y tuplas:
        - ("_LOG_", panel, text) → log etiquetado por panel
        - (callable, mensaje) → llama callable(mensaje) en hilo principal
        - ("_OCR_RESULT_", data) → resultado de OCR — pobla TreeView (T7)
        - ("_OCR_DONE_", None) → señal de fin de OCR (T7+)
        
        Procesa max 20 msg/tick para no congelar el main thread.
        """
        try:
            _processed = 0
            while _processed < 20:
                msg = self.log_queue.get_nowait()
                if isinstance(msg, tuple):
                    if len(msg) == 3 and msg[0] == "_LOG_":
                        # Tagged log message: ("_LOG_", panel, text)
                        panel, text = msg[1], msg[2]
                        # Always store in the panel's persistent buffer
                        self.logs_por_panel.setdefault(panel, []).append(text)
                        # Only write to visible textbox if this panel is active
                        if panel == self.panel_actual:
                            self.emit_log(text)
                    elif len(msg) == 2:
                        if callable(msg[0]):
                            msg[0](msg[1])
                        elif msg[0] == "_OCR_RESULT_":
                            self._procesar_resultado_ocr(msg[1])
                        elif msg[0] == "_OCR_DONE_":
                            self._cargar_datos_done()
                        elif msg[0] == "_CONTROL_FINAL_RESULT_":
                            self._procesar_resultado_control_final(msg[1])
                        elif msg[0] == "_TAREA_COMPLETA_":
                            self._finalizar_tarea()
                else:
                    # Plain string fallback — tag with current panel
                    self.logs_por_panel.setdefault(self.panel_actual, []).append(msg)
                    self.emit_log(msg)
                _processed += 1
        except queue.Empty:
            pass
        self.after(100, self._poll_log_queue)

    @staticmethod
    def _fmt_kg(v):
        """Format kg value: no .0, with thousands separator. E.g. 15690 → '15.690'"""
        try:
            n = float(str(v).replace(',', '.'))
            return f"{int(n):,}".replace(",", ".")
        except (ValueError, TypeError):
            return str(v) if v else "0"

    @staticmethod
    def _fmt_dni(v):
        """Format DNI with thousands separator. E.g. 27981220 → '27.981.220'"""
        try:
            digits = ''.join(c for c in str(v) if c.isdigit())
            n = int(digits)
            return f"{n:,}".replace(",", ".")
        except (ValueError, TypeError):
            return str(v) if v else ""

    def _procesar_resultado_ocr(self, data):
        """Inserta una fila por ticket con color verde/rojo según match.
        Almacena datos de comparación para abrir detalle al hacer click."""
        ticket = data.get("ticket", {})
        cont_data = data.get("contenedor")
        ruta_match = data.get("match")
        shared_excels = data.get("shared_excels", [])
        shared_neto_sum = data.get("shared_neto_sum", 0)

        # Valores OCR
        archivo   = ticket.get("archivo", "")
        patente   = ticket.get("patente", "")
        semi      = ticket.get("semi", "")
        conductor = ticket.get("conductor", "")
        dni       = ticket.get("dni", "")
        neto_ocr  = ticket.get("neto", 0)
        tara_ocr  = ticket.get("tara", 0)
        contenedor_str = ticket.get("contenedor", "")
        contenedor_display = contenedor_str if contenedor_str else "-"
        permiso   = ticket.get("permiso", "")

        import re as _re

        def _normalizar_pat(p):
            """Normalización permisiva para MATCHING (buscar fila de Excel).
            L→I, 0→O, 1→I, 8→B, 5→S para tolerar errores comunes de OCR."""
            p = p.upper().strip()
            p = p.replace('0', 'O').replace('1', 'I').replace('L', 'I').replace('8', 'B').replace('5', 'S')
            p = _re.sub(r'[^A-Z0-9]', '', p)
            return p

        def _normalizar_txt(t):
            return str(t).upper().strip()

        def _comparar_num(a, b):
            try:
                return abs(float(a) - float(b)) < 1
            except (ValueError, TypeError):
                return False

        # Buscar camión matching en cont_data
        camion_match = None
        k_idx = None
        if cont_data and cont_data.get("camiones"):
            patente_norm = _normalizar_pat(patente)
            semi_norm    = _normalizar_pat(semi) if semi else ""
            dni_norm     = _re.sub(r'\D', '', str(dni).strip()) if dni else ""
            cond_norm    = _normalizar_txt(conductor) if conductor else ""
            cont_norm    = _normalizar_txt(contenedor_str) if contenedor_str else ""

            self._log(f"[Control Datos] Ticket: patente='{patente}' semi='{semi}' dni='{dni}' cond='{conductor}' cont='{contenedor_str}'")
            for camion in cont_data["camiones"]:
                cam_pat  = _normalizar_pat(camion.get("patente_camion", ""))
                cam_semi = _normalizar_pat(camion.get("patente_semi", ""))
                cam_dni  = _re.sub(r'\D', '', str(camion.get("dni", "")).strip())
                cam_cond = _normalizar_txt(camion.get("conductor", ""))
                cam_cont = _normalizar_txt(str(camion.get("contenedor", "")))

                def _nombres_equivalentes(a, b):
                    """True si al menos 2 palabras coinciden entre ambos nombres
                    (sin importar acentos, mayúsculas, puntuación, ni palabras extra).
                    
                    Ej: "RETA JORGE SEBASTIAN" y "RETA JORGE SEBASTIÁN" → 3/3 ✅
                        "CLAUDIO PAEZ" y "CLAUDIO CESAR PAEZ" → 2/3 ≥ 2 ✅
                        "JUAN PEREZ" y "CLAUDIO PAEZ" → 0 ❌
                    """
                    if not a or not b:
                        return False
                    import unicodedata
                    def _limpiar(s):
                        # Quitar acentos, puntuación, mayúsculas
                        s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode('ascii')
                        return re.sub(r'[^\w\s]', ' ', s).strip().upper()
                    a_set = set(_limpiar(a).split())
                    b_set = set(_limpiar(b).split())
                    if not a_set or not b_set:
                        return False
                    return len(a_set & b_set) >= 2

                # Intentar match por cualquier campo
                razon = ""
                if patente_norm and cam_pat and patente_norm == cam_pat:
                    razon = f"patente_camion ({patente})"
                elif semi_norm and cam_semi and semi_norm == cam_semi:
                    razon = f"patente_semi ({semi})"
                elif dni_norm and cam_dni and dni_norm == cam_dni:
                    razon = "DNI"
                elif cond_norm and cam_cond and _nombres_equivalentes(cond_norm, cam_cond):
                    razon = f"conductor ({conductor})"
                elif cont_norm and cam_cont and cont_norm == cam_cont:
                    razon = f"contenedor ({contenedor_str})"

                if razon:
                    camion_match = camion
                    k_idx = camion.get("k")
                    self._log(f"[Control Datos] ✅ Match por {razon}")
                    break

            if not camion_match:
                # Log de diagnóstico: qué campos no matchearon
                for camion in cont_data["camiones"]:
                    self._log(f"[Control Datos]   vs camión: '{camion.get('patente_camion','')}' "
                              f"semi:'{camion.get('patente_semi','')}' "
                              f"dni:'{camion.get('dni','')}' "
                              f"cond:'{camion.get('conductor','')}'")

        tiene_match = (camion_match is not None and ruta_match is not None)

        if not tiene_match:
            iid = f"ocr_{self._cargar_datos_idx}"
            self._cargar_datos_idx += 1
            estado = "⚠ No encontrado"
            valores = (data.get("cliente", archivo), patente, semi, conductor,
                       self._fmt_dni(dni), self._fmt_kg(neto_ocr), self._fmt_kg(tara_ocr),
                       contenedor_display, permiso, estado)
            try:
                self.tree_carga.insert("", "end", values=valores, iid=iid,
                                       tags=("tag_no_match",))
            except Exception:
                pass
            self._log(f"[Control Datos] {archivo}: ⚠ Sin match")
            return

        # Comparación campo-por-campo (normalización de formato, NO de caracteres)
        def _raw_txt(v):
            # Saca espacios/puntuación PERO preserva I≠L, 0≠O, 1≠I, etc.
            return _re.sub(r'[^A-Z0-9]', '', str(v).upper().strip())

        def _nombres_equivalentes(a, b):
            """True si al menos 2 palabras coinciden (sin acentos, puntuación, mayúsculas)."""
            if not a or not b:
                return False
            import unicodedata
            def _limpiar(s):
                s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode('ascii')
                return _re.sub(r"[^\w\s]", " ", s).strip().upper()
            a_set = set(_limpiar(a).split())
            b_set = set(_limpiar(b).split())
            if not a_set or not b_set:
                return False
            return len(a_set & b_set) >= 2

        cam_pat    = camion_match.get("patente_camion", "")
        cam_semi   = camion_match.get("patente_semi", "")
        cam_cond   = camion_match.get("conductor", "")
        cam_dni    = camion_match.get("dni", "")
        peso_carga = camion_match.get("peso_carga", 0)
        tara_cont  = camion_match.get("tara_cont", 0)
        pe_val     = data.get("pe", "") or ""

        ok_patente   = _raw_txt(patente) == _raw_txt(cam_pat)
        ok_semi      = _raw_txt(semi) == _raw_txt(cam_semi)
        ok_conductor = _nombres_equivalentes(conductor, cam_cond)
        ok_dni       = _raw_txt(dni) == _raw_txt(cam_dni)
        # Shared trips: compare against SUM of Neto from all Excels
        if shared_excels:
            ok_neto = _comparar_num(neto_ocr, shared_neto_sum)
        else:
            ok_neto = _comparar_num(neto_ocr, peso_carga)
        ok_tara      = _comparar_num(tara_ocr, tara_cont)
        ok_contenedor      = _raw_txt(contenedor_str) == _raw_txt(camion_match.get("contenedor", ""))
        ok_permiso   = _raw_txt(permiso) == _raw_txt(pe_val) if pe_val else True

        # Estado general: TODO verde si todos los campos coinciden
        all_ok = (ok_patente and ok_semi and ok_conductor and ok_dni
                  and ok_neto and ok_tara and ok_contenedor and ok_permiso)
        if all_ok:
            estado = "✅ Ok"
            tag = "tag_ok"
        else:
            estado = "❌ Diferencia"
            tag = "tag_mismatch"

        iid = f"ocr_{self._cargar_datos_idx}"
        self._cargar_datos_idx += 1
        nombre_contenedor = os.path.basename(ruta_match)
        valores = (
            data.get("cliente", archivo),
            patente, semi, conductor, self._fmt_dni(dni),
            self._fmt_kg(neto_ocr), self._fmt_kg(tara_ocr),
            contenedor_display,
            permiso, estado,
        )
        try:
            self.tree_carga.insert("", "end", values=valores, iid=iid, tags=(tag,))
            self._cargar_datos_rutas[iid] = (ruta_match, k_idx)
        except Exception:
            pass

        # Guardar datos para popup de comparación al hacer doble-click
        self._cargar_datos_comparacion[iid] = {
            "archivo": archivo,
            "ticket": {
                "Patente": patente,
                "Semirremolque": semi,
                "Conductor": conductor,
                "DNI": dni,
                "Neto (kg)": f"{neto_ocr:.0f}",
                "Tara (kg)": f"{tara_ocr:.0f}",
                "Contenedor": contenedor_display,
                "Permiso": permiso,
            },
            "contenedor": {
                "Patente": cam_pat,
                "Semirremolque": cam_semi,
                "Conductor": cam_cond,
                "DNI": cam_dni,
                "Neto (kg)": f"{peso_carga:.0f}",
                "Tara (kg)": f"{tara_cont:.0f}",
                "Contenedor": camion_match.get("contenedor", "") or "-",
                "Permiso": pe_val,
            },
            "ok": {
                "Patente": ok_patente,
                "Semirremolque": ok_semi,
                "Conductor": ok_conductor,
                "DNI": ok_dni,
                "Neto (kg)": ok_neto,
                "Tara (kg)": ok_tara,
                "Contenedor": ok_contenedor,
                "Permiso": ok_permiso,
            },
            "shared_excels": shared_excels,
            "shared_neto_sum": shared_neto_sum,
        }

        # Log
        if not all_ok:
            diffs = []
            if not ok_neto:
                diffs.append(f"neto OCR={neto_ocr:.0f} vs CONT={peso_carga:.0f}")
            if not ok_tara:
                diffs.append(f"tara OCR={tara_ocr:.0f} vs CONT={tara_cont:.0f}")
            self._log(f"[Control Datos] {archivo}: Diferencias — {'; '.join(diffs)}")
        else:
            self._log(f"[Control Datos] {archivo}: ✅ Todo OK")

    def _abrir_comparacion(self, event=None):
        """Abre popup con tabla vertical: Ticket | Excel, verde/rojo por campo."""
        sel = self.tree_carga.selection()
        if not sel:
            return
        iid = sel[0]
        datos = self._cargar_datos_comparacion.get(iid)
        if not datos:
            return

        def _build_popup(level):
            fsizes = self._get_font_sizes(level)
            geom = self._get_popup_geometry(520, 420, level)
            ancho = int(geom.split("x")[0])

            # Check if shared trip
            shared_excels = datos.get("shared_excels", [])
            shared_neto_sum = datos.get("shared_neto_sum", 0)
            is_shared = len(shared_excels) > 1

            # ── Number formatters ──
            _fmt_kg = self._fmt_kg
            _fmt_dni = self._fmt_dni

            if is_shared:
                geom = self._get_popup_geometry(720, 420, level)
                ancho = int(geom.split("x")[0])

            dlg = ctk.CTkToplevel(self)
            dlg.title(f"Comparación — {datos['archivo']}" + (" (Compartido)" if is_shared else ""))
            dlg.transient(self)
            dlg.grab_set()
            dlg.resizable(False, False)

            # ── Top bar: header + font level override ──────────────────
            top_bar = ctk.CTkFrame(dlg, fg_color=Palette.BG_SIDEBAR, corner_radius=6, height=36)
            top_bar.pack(fill="x", padx=12, pady=(12, 0))
            top_bar.pack_propagate(False)

            # Grid header: align with body columns
            top_bar.grid_columnconfigure(0, minsize=120, weight=0)

            if is_shared:
                # 4 columns: Campo | Ticket | Excel 1 | Excel 2
                headers = ["Campo", "Ticket (OCR)"]
                for i, ex in enumerate(shared_excels):
                    headers.append(f"Excel {i+1}")
                for col_i in range(1, len(headers)):
                    top_bar.grid_columnconfigure(col_i, weight=1)
                for col_i, txt in enumerate(headers):
                    lbl = ctk.CTkLabel(
                        top_bar, text=txt, width=120 if col_i == 0 else 160,
                        font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["header"], weight="bold"),
                        text_color=Palette.TEXT_SECONDARY,
                    )
                    lbl.grid(row=0, column=col_i, sticky="ew", padx=4, pady=4)
            else:
                # Normal 3 columns: Campo | Ticket | Contenedor
                for col_i in range(1, 3):
                    top_bar.grid_columnconfigure(col_i, weight=1)
                for col_i, txt in enumerate(["Campo", "Ticket (OCR)", "Contenedor (Excel)"]):
                    lbl = ctk.CTkLabel(
                        top_bar, text=txt, width=120 if col_i == 0 else 160,
                        font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["header"], weight="bold"),
                        text_color=Palette.TEXT_SECONDARY,
                    )
                    lbl.grid(row=0, column=col_i, sticky="ew", padx=4, pady=4)

            # Font level override selector
            override_menu = ctk.CTkOptionMenu(
                top_bar, values=["1", "2", "3"],
                font=ctk.CTkFont(family=FONT_FAMILY, size=10),
                fg_color=Palette.BG_INPUT, button_color=Palette.ACCENT,
                width=50, height=26,
                command=lambda v: (dlg.destroy(), _build_popup(int(v))),
            )
            override_menu.set(str(level))
            override_menu.pack(side="right", padx=(0, 8))

            # Cuerpo: grid para columnas alineadas
            body = ctk.CTkFrame(dlg, fg_color="transparent")
            body.pack(fill="both", expand=True, padx=12, pady=8)
            body.grid_columnconfigure(0, minsize=120, weight=0)  # Campo — fijo
            for col_i in range(1, 4 if is_shared else 3):
                body.grid_columnconfigure(col_i, weight=1)

            campos = [("Camion", "Patente"), ("Semirremolque", "Semirremolque"), ("Conductor", "Conductor"),
                       ("DNI", "DNI"), ("Neto (kg)", "Neto (kg)"), ("Tara (kg)", "Tara (kg)"),
                       ("Contenedor", "Contenedor"), ("Permiso", "Permiso")]

            for row_i, (label, key) in enumerate(campos):
                val_ticket = datos["ticket"].get(key, "")
                val_cont   = datos["contenedor"].get(key, "")
                ok         = datos["ok"].get(key, False)

                bg = "#C8FFC8" if ok else "#FFC8C8"
                fg = "#006400" if ok else "#8B0000"

                ctk.CTkLabel(
                    body, text=label, width=120, anchor="w",
                    font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["data"], weight="bold"),
                    text_color=Palette.TEXT_PRIMARY,
                ).grid(row=row_i, column=0, sticky="w", padx=(4, 0), pady=2)

                if is_shared:
                    # Shared: Ticket (col 1) + Excel 1 (col 2) + Excel 2 (col 3)
                    val_ticket_neto = datos["ticket"].get("Neto (kg)", 0)

                    # ── Column 1: Ticket value (formatted) ──
                    if key in ("Neto (kg)", "Tara (kg)"):
                        display_ticket = _fmt_kg(val_ticket)
                    elif key == "DNI":
                        display_ticket = _fmt_dni(val_ticket)
                    else:
                        display_ticket = val_ticket
                    ctk.CTkLabel(
                        body, text=display_ticket, anchor="center",
                        font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["data"]),
                        fg_color=bg, text_color=fg, corner_radius=4,
                    ).grid(row=row_i, column=1, sticky="ew", padx=4, pady=2)

                    # ── Columns 2..N: each Excel ──
                    for col_i, ex in enumerate(shared_excels, start=2):
                        ex_camion = ex.get("camion", {})
                        ex_data = ex.get("data", {})
                        ex_pe = ex.get("pe", "")

                        if key == "Neto (kg)":
                            # Neto: show individual + total
                            neto_individual = ex_camion.get("peso_carga", 0)
                            val = f"{_fmt_kg(neto_individual)} ({_fmt_kg(shared_neto_sum)})"
                            try:
                                ticket_neto = float(str(val_ticket_neto).replace('.', '').replace(',', '.'))
                                sums_ok = abs(shared_neto_sum - ticket_neto) < 1
                            except (ValueError, AttributeError):
                                sums_ok = False
                            cell_bg = "#C8FFC8" if sums_ok else "#FFC8C8"
                            cell_fg = "#006400" if sums_ok else "#8B0000"
                        elif key == "Tara (kg)":
                            val = _fmt_kg(ex_camion.get("tara_cont", ""))
                            cell_bg, cell_fg = bg, fg
                        elif key == "Patente":
                            val = str(ex_camion.get("patente_camion", ""))
                            cell_bg, cell_fg = bg, fg
                        elif key == "Semirremolque":
                            val = str(ex_camion.get("patente_semi", ""))
                            cell_bg, cell_fg = bg, fg
                        elif key == "Conductor":
                            val = str(ex_camion.get("conductor", ""))
                            cell_bg, cell_fg = bg, fg
                        elif key == "DNI":
                            val = _fmt_dni(ex_camion.get("dni", ""))
                            cell_bg, cell_fg = bg, fg
                        elif key == "Contenedor":
                            val = str(ex_camion.get("contenedor", ex_data.get("contenedor", "")))
                            cell_bg, cell_fg = bg, fg
                        elif key == "Permiso":
                            val = str(ex_pe)
                            cell_bg, cell_fg = bg, fg
                        else:
                            val = ""
                            cell_bg, cell_fg = bg, fg

                        lbl = ctk.CTkLabel(
                            body, text=val, anchor="center",
                            font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["data"]),
                            fg_color=cell_bg, text_color=cell_fg, corner_radius=4,
                        )
                        lbl.grid(row=row_i, column=col_i, sticky="ew", padx=4, pady=2)
                else:
                    # Normal: Ticket + single Excel — format numbers
                    if key in ("Neto (kg)", "Tara (kg)"):
                        display_ticket = _fmt_kg(val_ticket)
                        display_cont = _fmt_kg(val_cont)
                    elif key == "DNI":
                        display_ticket = _fmt_dni(val_ticket)
                        display_cont = _fmt_dni(val_cont)
                    else:
                        display_ticket = val_ticket
                        display_cont = val_cont
                    for col_i, val in enumerate([display_ticket, display_cont], start=1):
                        lbl = ctk.CTkLabel(
                            body, text=val, anchor="center",
                            font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["data"]),
                            fg_color=bg, text_color=fg, corner_radius=4,
                        )
                        lbl.grid(row=row_i, column=col_i, sticky="ew", padx=4, pady=2)

            # Leyenda al pie
            leyenda = ctk.CTkFrame(dlg, fg_color="transparent")
            leyenda.pack(fill="x", pady=(0, 12))
            inner = ctk.CTkFrame(leyenda, fg_color="transparent")
            inner.pack(anchor="center")
            ctk.CTkLabel(
                inner, text="", width=14, height=14,
                fg_color="#C8FFC8", corner_radius=7,
            ).pack(side="left", padx=(4, 2))
            ctk.CTkLabel(
                inner, text="Coincide",
                font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["legend"]),
                text_color=Palette.TEXT_MUTED,
            ).pack(side="left", padx=(0, 14))
            ctk.CTkLabel(
                inner, text="", width=14, height=14,
                fg_color="#FFC8C8", corner_radius=7,
            ).pack(side="left", padx=(4, 2))
            ctk.CTkLabel(
                inner, text="Difiere",
                font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["legend"]),
                text_color=Palette.TEXT_MUTED,
            ).pack(side="left", padx=(0, 14))

            # Ajustar ventana al contenido real
            dlg.update_idletasks()
            alto = dlg.winfo_reqheight()
            x = (dlg.winfo_screenwidth() - ancho) // 2
            y = (dlg.winfo_screenheight() - alto) // 2
            dlg.geometry(f"{ancho}x{alto}+{x}+{y}")

        _build_popup(self.config.get("font_level", 1))

    def _limpiar_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self._log_line_count = 0
        self.lbl_log_count.configure(text="")
        # También limpiar el historial guardado del panel actual
        if self.panel_actual in self.logs_por_panel:
            self.logs_por_panel[self.panel_actual] = []

    def _capturar_lineas_log(self):
        """Devuelve una lista con todas las líneas de texto del log (sin tags)."""
        contenido = self.log_text.get("1.0", "end-1c")
        if not contenido.strip():
            return []
        return [l for l in contenido.split("\n") if l.strip()]

    def _restaurar_log(self, nombre):
        """Limpia el widget y restaura las líneas guardadas para el panel dado."""
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")

        lineas = self.logs_por_panel.get(nombre, [])
        if lineas:
            for i, linea in enumerate(lineas):
                if i > 0:
                    self.log_text.insert("end", "\n")
                # Detectar tag para colorear
                txt_up = linea.upper()
                tag = None
                if any(kw in txt_up for kw in ("ERROR", "CRÍTICO", "NO SE PUDO")):
                    tag = "error"
                elif any(kw in txt_up for kw in ("ÉXITO", "FINALIZADO", "CORRECTAMENTE", "COMPLETADO", "GUARDADA")):
                    tag = "success"
                elif any(kw in txt_up for kw in ("ADVERTENCIA", "FALTA", "NO SE ENCONTRÓ")):
                    tag = "warning"
                elif any(kw in txt_up for kw in ("ESCANEANDO", "PREPARANDO", "CONECTANDO", "SUBIENDO", "LEYENDO", "COPIADO")):
                    tag = "info"
                if tag:
                    self.log_text.insert("end", linea, tag)
                else:
                    self.log_text.insert("end", linea)
            self._log_line_count = len(lineas)
            self.lbl_log_count.configure(text=f"{self._log_line_count} líneas")
        else:
            self._log_line_count = 0
            self.lbl_log_count.configure(text="")

        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _trim_log(self):
        """Elimina las líneas más antiguas para mantener un máximo de 500."""
        self.log_text.configure(state="normal")
        # Obtener todo, descartar primeras 200 líneas
        contenido = self.log_text.get("1.0", "end-1c")
        lineas = contenido.split("\n")
        if len(lineas) > 500:
            keep = lineas[-300:]  # Mantener últimas 300
            self.log_text.delete("1.0", "end")
            self.log_text.insert("1.0", "\n".join(keep))
            self._log_line_count = len(keep)
            self.lbl_log_count.configure(text=f"{self._log_line_count} líneas")
        self.log_text.configure(state="disabled")

    # ═══════════════════════════════════════════════════════════════════
    # BARRA DE ESTADO (Desactivada)
    # ═══════════════════════════════════════════════════════════════════
    def _set_status(self, texto):
        pass

    # ═══════════════════════════════════════════════════════════════════
    # CONFIG HELPERS
    # ═══════════════════════════════════════════════════════════════════
    def _cfg_obtener(self, seccion, clave, default):
        val = self.config.get(seccion, {}).get(clave, default)
        if seccion == "seguridad" and clave == "password" and val != default:
            val = self._decrypt_val(val, os.environ["MULTIAGENTE_SECRET_KEY"])
            if val.startswith("enc::"):
                return default
        return val

    def _cfg_obtener_correo(self, clave, default):
        val = self._cfg_obtener("correo", clave, default)
        if clave == "password" and val != default:
            key = os.environ["MULTIAGENTE_SECRET_KEY"]
            val = self._decrypt_val(val, key)
            if val.startswith("enc::"):
                return default
        return val

    def _cfg_obtener_docs(self, clave, default):
        return self._cfg_obtener("documentos", clave, default)

    def _cfg_obtener_rutas(self, clave, default):
        return self._cfg_obtener("rutas", clave, default)

    def _resolver_ruta(self, clave, default):
        r"""Resuelve una ruta configurable: 'Desktop' → %USERPROFILE%\Desktop, sino se usa tal cual."""
        valor = self._cfg_obtener_rutas(clave, default)
        if valor.lower() == "desktop":
            nombre_esc = self._cfg_obtener_rutas("escritorio_nombre", "Desktop")
            return os.path.join(os.path.expanduser("~"), nombre_esc)
        return valor

    def _get_font_sizes(self, level=None):
        """Return scaled font sizes for the given level."""
        if level is None:
            level = self.config.get("font_level", 1)
        scale = FONT_LEVEL_SCALES.get(level, 1.0)
        return {
            "data": int(FONT_BASE_SIZES["data"] * scale),
            "header": int(FONT_BASE_SIZES["header"] * scale),
            "legend": int(FONT_BASE_SIZES["legend"] * scale),
        }

    def _get_popup_geometry(self, base_w, base_h, level=None):
        """Return scaled geometry string for popups."""
        if level is None:
            level = self.config.get("font_level", 1)
        scale = FONT_LEVEL_SCALES.get(level, 1.0)
        w = int(base_w * scale)
        h = int(base_h * scale)
        return f"{w}x{h}"

    def _calc_popup_height(self, num_rows, level=None):
        """Calculate popup height based on number of data rows."""
        if level is None:
            level = self.config.get("font_level", 1)
        scale = FONT_LEVEL_SCALES.get(level, 1.0)
        top_bar = 48          # header bar + padding
        row_h = int(32 * scale)  # each data row
        legend = 42           # legend at bottom
        scroll_pad = 16       # scroll frame padding
        window_pad = 24       # top + bottom window padding
        h = top_bar + (num_rows * row_h) + legend + scroll_pad + window_pad
        return h

    def _recuperar_password(self, parent_dlg, lbl_status=None):
        """Envía la contraseña por SMTP directamente al correo configurado."""
        import smtplib
        from email.mime.text import MIMEText

        usuario = self._cfg_obtener_correo("usuario", "")
        pw_email = self._cfg_obtener_correo("password", "")
        pw_app = self._cfg_obtener("seguridad", "password", "")
        imap_srv = self._cfg_obtener_correo("imap_server", "imap.gmail.com")

        if not usuario or not pw_email:
            if lbl_status:
                self.after(0, lambda: lbl_status.configure(
                    text="No hay correo configurado en Ajustes > Correo.",
                    text_color=Palette.ERROR))
            return

        if not pw_app:
            if lbl_status:
                self.after(0, lambda: lbl_status.configure(
                    text="No hay contraseña maestra configurada.",
                    text_color=Palette.WARNING))
            return

        # Derivar SMTP del IMAP: imap.dominio.com → smtp.dominio.com
        if "gmail" in imap_srv.lower():
            smtp_srv = "smtp.gmail.com"
        else:
            smtp_srv = "smtp." + imap_srv.removeprefix("imap.")

        if lbl_status:
            self.after(0, lambda: lbl_status.configure(
                text=f"Enviando a {usuario} vía {smtp_srv}...",
                text_color=Palette.INFO))

        def _enviar():
            ultimo_error = ""
            for puerto in (587, 25, 465):
                try:
                    msg = MIMEText(f"Tu contraseña maestra del Sistema de Automatización es:\n\n"
                                   f"  {pw_app}\n\n"
                                   f"Usala para ingresar a la aplicación.")
                    msg["Subject"] = "Clave de Recuperación"
                    msg["From"] = usuario
                    msg["To"] = usuario

                    if puerto == 465:
                        server = smtplib.SMTP_SSL(smtp_srv, puerto, timeout=10)
                    else:
                        server = smtplib.SMTP(smtp_srv, puerto, timeout=10)
                        server.ehlo()
                        try:
                            server.starttls()
                            server.ehlo()
                        except Exception:
                            pass  # algunos servidores no soportan STARTTLS
                    server.login(usuario, pw_email)
                    server.send_message(msg)
                    server.quit()

                    self.after(0, lambda: lbl_status.configure(
                        text=f"Enviado a {usuario}",
                        text_color=Palette.SUCCESS)) if lbl_status else \
                        self.after(0, lambda: messagebox.showinfo("Contraseña enviada",
                            f"La contraseña fue enviada a:\n{usuario}\n\n"
                            f"Revisá tu bandeja de entrada."))
                    return
                except Exception as e:
                    ultimo_error = str(e)[:100]
                    continue

            self.after(0, lambda: lbl_status.configure(
                text=f"Error: {ultimo_error}",
                text_color=Palette.ERROR)) if lbl_status else \
                self.after(0, lambda: messagebox.showerror("Error al enviar",
                    f"No se pudo enviar en ningún puerto:\n\n{ultimo_error}\n\n"
                    f"SMTP: {smtp_srv}\nPuertos: 587, 25, 465\n"
                    f"Usuario: {usuario}\n\n"
                    f"Verificá Ajustes > Correo."))

        t = threading.Thread(target=_enviar, daemon=True)
        t.start()

    def _verificar_password_maestra(self):
        """Si hay contraseña maestra configurada, la pide. Retorna True si pasa."""
        pw_config = self._cfg_obtener("seguridad", "password", "")
        if not pw_config:
            return True
        return self._mostrar_dialogo_password_ajustes(pw_config)

    def _mostrar_dialogo_password_ajustes(self, pw_config):
        """Diálogo CTk para pedir contraseña al entrar a Ajustes (bloqueante)."""
        resultado = {"ok": False}

        dlg = ctk.CTkToplevel(self)
        dlg.title("Acceso Restringido")
        dlg.geometry("380x190")
        dlg.configure(fg_color=Palette.BG_CARD)
        dlg.transient(self)
        dlg.lift()
        dlg.grab_set()
        dlg.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 380) // 2
        py = self.winfo_y() + (self.winfo_height() - 190) // 2
        dlg.geometry(f"380x190+{px}+{py}")

        ctk.CTkLabel(
            dlg, text="Ingresá la contraseña para acceder a Ajustes:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(pady=(24, 12))

        entry_pw = ctk.CTkEntry(
            dlg, width=260, height=36, show="*",
            font=ctk.CTkFont(family=FONT_FAMILY, size=14),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=6,
        )
        entry_pw.pack(pady=(0, 12))
        entry_pw.focus_set()
        dlg.after(50, lambda: entry_pw.focus_force())

        lbl_error = ctk.CTkLabel(
            dlg, text="",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.ERROR,
        )
        lbl_error.pack()

        def verificar():
            if entry_pw.get() == pw_config:
                resultado["ok"] = True
                self._master_pw_cache = entry_pw.get()
                dlg.destroy()
            else:
                lbl_error.configure(text="Contraseña incorrecta. Intentá de nuevo.")
                entry_pw.delete(0, "end")
                entry_pw.focus_set()

        entry_pw.bind("<Return>", lambda e: verificar())
        btn_frame = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_frame.pack(pady=(12, 0))
        ctk.CTkButton(
            btn_frame, text="Ingresar", width=100, height=32,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6,
            command=verificar,
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            btn_frame, text="Cancelar", width=100, height=32,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.BG_HOVER, hover_color=Palette.ERROR,
            text_color=Palette.TEXT_SECONDARY, corner_radius=6,
            command=dlg.destroy,
        ).pack(side="left", padx=4)

        dlg.wait_window()
        return resultado["ok"]


    def _cargar_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    self.config = json.load(f)
            except Exception as e:
                self._log(f"ERROR al cargar configuración: {e}")

    def _guardar_config(self):
        try:
            self.config["window_geo"] = self.geometry()
            if hasattr(self, 'paned_window') and self.paned_window.winfo_exists():
                sash = self.paned_window.sash_coord(0)
                if self.panel_actual and sash[1] > 0:
                    self.config[f"sash_{self.panel_actual}"] = sash[1]
            # Guardar anchos de columna del árbol de mails
            if hasattr(self, '_mail_tree') and self._mail_tree.winfo_exists():
                self.config["mail_tree_columns"] = {
                    col: self._mail_tree.column(col, "width")
                    for col in ("sel", "asunto", "fecha", "adjuntos", "carpeta")
                }
            # Guardar anchos de columna de Control de Tickets
            if hasattr(self, 'tree_carga') and self.tree_carga.winfo_exists():
                self.config["tree_carga_columns"] = {
                    col: self.tree_carga.column(col, "width")
                    for col in ("cliente", "camion", "semi", "conductor", "dni",
                                "neto", "tara", "contenedor", "permiso", "estado")
                }
            # Guardar anchos de columna de Coordinacion
            if hasattr(self, 'tree_coordinacion') and self.tree_coordinacion.winfo_exists():
                self.config["tree_coordinacion_columns"] = {
                    col: self.tree_coordinacion.column(col, "width")
                    for col in ("carpeta", "giro", "cliente", "destino",
                                "buque", "viaje", "booking", "pto_descarga", "pto_final",
                                "fecha_of_pe", "fecha_carga", "peso_flexi", "estado")
                }
            # Guardar anchos de columna de Control Final
            if hasattr(self, 'tree_control_final') and self.tree_control_final.winfo_exists():
                self.config["tree_control_final_columns"] = {
                    col: self.tree_control_final.column(col, "width")
                    for col in ("cliente", "camion", "semi", "conductor", "dni",
                                "neto", "tara", "contenedor", "permiso", "estado", "salida_aduana")
                }
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            self._log(f"ERROR al guardar configuración: {e}")

    def _encrypt_val(self, value, key):
        if not value:
            return ""
        if value.startswith("enc::"):
            return value
        import hashlib
        import base64
        try:
            salt = os.urandom(16)
            derived = hashlib.pbkdf2_hmac('sha256', key.encode('utf-8'), salt, 10000, dklen=len(value))
            encrypted = bytes(a ^ b for a, b in zip(value.encode('utf-8'), derived))
            payload = base64.b64encode(salt + encrypted).decode('utf-8')
            return f"enc::{payload}"
        except Exception as e:
            self._log(f"ERROR al encriptar: {e}")
            return value

    def _decrypt_val(self, value, key):
        if not value or not value.startswith("enc::"):
            return value
        import hashlib
        import base64
        try:
            payload = value[5:]
            data = base64.b64decode(payload.encode('utf-8'))
            if len(data) < 16:
                return value
            salt = data[:16]
            encrypted = data[16:]
            derived = hashlib.pbkdf2_hmac('sha256', key.encode('utf-8'), salt, 10000, dklen=len(encrypted))
            decrypted = bytes(a ^ b for a, b in zip(encrypted, derived))
            return decrypted.decode('utf-8')
        except Exception as e:
            try:
                self._log(f"ERROR al desencriptar: {e}")
            except AttributeError:
                print(f"[WARN] Decrypt failed (log_queue not ready): {e}")
            return value

    def _clave_encriptacion(self):
        """Clave de encriptación desde variable de entorno, con fallback."""
        return os.environ.get("MULTIAGENTE_SECRET_KEY", "api_vision")

    def _decrypt_api_key(self, raw):
        """Desencripta API key probando clave actual, con fallback a la vieja."""
        result = self._decrypt_val(raw, self._clave_encriptacion())
        if result and result.startswith("enc::"):
            result = self._decrypt_val(raw, "api_vision")
        return result

    # ═══════════════════════════════════════════════════════════════════
    # CONFIRMACIÓN DE SALIDA
    # ═══════════════════════════════════════════════════════════════════
    def _confirmar_salida(self):
        """Cierra la aplicación con confirmación genérica."""
        ok = messagebox.askyesno(
            "Salir",
            "¿Está seguro de que desea salir del sistema?"
        )
        if not ok:
            return
        self._guardar_config()
        self.destroy()

# ── Punto de entrada ─────────────────────────────────────────────────────
def main():
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    ctk.set_widget_scaling(1.0)
    ctk.set_window_scaling(1.0)

    # Forzar escala DPI para pantallas HiDPI
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = App()
    app.mainloop()

if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    main()
