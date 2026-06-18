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
import dotenv
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
class App(ctk.CTk):
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

        # ── Construir UI ──────────────────────────────────────────────
        self._crear_sidebar()
        self._crear_area_principal()

        # Seleccionar panel inicial
        self._cambiar_panel("descargar")

        # Bind de teclas
        self.bind("<Escape>", lambda e: self._confirmar_salida())
        self.protocol("WM_DELETE_WINDOW", self._confirmar_salida)

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

        self.btn_descargar = self._crear_btn_nav(
            "Descargar Mails", "📥", 0, lambda: self._cambiar_panel("descargar")
        )
        self.btn_impresion = self._crear_btn_nav(
            "Impresión Documental", "📄", 1, lambda: self._cambiar_panel("impresion")
        )
        self.btn_planillas = self._crear_btn_nav(
            "Completar Planillas", "📊", 2, lambda: self._cambiar_panel("planillas")
        )
        self.btn_cargar_datos = self._crear_btn_nav(
            "Controlar Datos", "📋", 3, lambda: self._cambiar_panel("cargar-datos")
        )
        self.btn_correos = self._crear_btn_nav(
            "Enviar Correos", "📤", 4, lambda: self._cambiar_panel("correos")
        )
        self.btn_backup = self._crear_btn_nav(
            "Backup", "💾", 5, lambda: self._cambiar_panel("backup")
        )

        self.btn_ajustes = self._crear_btn_nav(
            "Ajustes", "⚙", 6, lambda: self._cambiar_panel("ajustes")
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

    def _crear_btn_nav(self, texto, icono, idx, comando):
        btn = ctk.CTkButton(
            self._nav_scroll,
            text=f"  {icono}  {texto}",
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
    def _panel_impresion(self):
        if "impresion" in self._panel_frames:
            return
        self._imp_carpetas_vars = {}
        self._imp_select_all_var = ctk.BooleanVar(value=True)

        frame = ctk.CTkFrame(self.panel_container, fg_color="transparent")
        self._panel_frames["impresion"] = frame
        # frame.pack(fill="both", expand=True) — packed by _cambiar_panel_forzado

        # ── Toolbar: solo Refrescar ──────────────────────────────────
        toolbar = ctk.CTkFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER, height=44
        )
        toolbar.pack(fill="x", pady=(0, 6))
        toolbar.pack_propagate(False)

        self._imp_btn_refresh = ctk.CTkButton(
            toolbar,
            text="🔄  Refrescar",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_HOVER, hover_color=Palette.ACCENT_DIM,
            text_color=Palette.TEXT_PRIMARY, corner_radius=6, height=34,
            command=self._imp_escanear_carpetas,
        )
        self._imp_btn_refresh.pack(side="left", padx=4, pady=4)

        self._imp_btn_dorsos = ctk.CTkButton(
            toolbar,
            text="📄 Dorsos",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_HOVER, hover_color=Palette.ACCENT_DIM,
            text_color=Palette.TEXT_PRIMARY, corner_radius=6, height=34,
            command=self._imp_popup_dorsos,
        )
        self._imp_btn_dorsos.pack(side="left", padx=4, pady=4)

        self._imp_btn_sellos_mic = ctk.CTkButton(
            toolbar, text="🔖 Sellos MIC",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_HOVER, hover_color=Palette.ACCENT_DIM,
            text_color=Palette.TEXT_PRIMARY, corner_radius=6, height=34,
            command=self._imp_popup_sellos_mic,
        )
        self._imp_btn_sellos_mic.pack(side="left", padx=4, pady=4)

        self._imp_btn_sellos_crt = ctk.CTkButton(
            toolbar, text="🔖 Sellos CRT",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_HOVER, hover_color=Palette.ACCENT_DIM,
            text_color=Palette.TEXT_PRIMARY, corner_radius=6, height=34,
            command=self._imp_popup_sellos_crt,
        )
        self._imp_btn_sellos_crt.pack(side="left", padx=4, pady=4)

        self._imp_lbl_estado = ctk.CTkLabel(
            toolbar,
            text="", font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_MUTED,
        )
        self._imp_lbl_estado.pack(side="right", padx=12)

        # ── Split izquierda-derecha ──────────────────────────────────
        split_frame = ctk.CTkFrame(frame, fg_color="transparent")
        split_frame.pack(fill="both", expand=True)

        # Columna izquierda: lista de carpetas
        left_frame = ctk.CTkFrame(
            split_frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER
        )
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 4))

        ctk.CTkLabel(
            left_frame, text="CARPETAS DE CARGA",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(anchor="w", padx=14, pady=(10, 2))

        self._imp_scroll_carpetas = ctk.CTkScrollableFrame(
            left_frame, fg_color="transparent"
        )
        self._imp_scroll_carpetas.pack(fill="both", expand=True, padx=4, pady=(0, 6))

        # Columna derecha: opciones de impresión
        right_frame = ctk.CTkFrame(
            split_frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER, width=310
        )
        right_frame.pack(side="right", fill="y", padx=(4, 0))
        right_frame.pack_propagate(False)

        ctk.CTkLabel(
            right_frame, text="OPCIONES DE IMPRESIÓN",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(anchor="w", padx=12, pady=(8, 6))

        opciones = [
            ("sobre", "Sobre (Planilla de Carga)", 1),
            ("permiso", "Permiso de Exportación", 2),
            ("hoja_ruta", "Hoja de Ruta", 2),
            ("servicio_ata", "Servicio ATA / Recibo ATA", 0),
        ]
        def _verificar_boton_imprimir(*args):
            """Activa el botón IMPRIMIR solo si hay carpetas y opciones seleccionadas."""
            hay_carpetas = any(v.get() for v in self._imp_carpetas_vars.values())
            hay_opciones = any(v.get() for v in self._imp_opciones_vars.values())
            if hay_carpetas and hay_opciones:
                self._imp_btn_imprimir.configure(state="normal", fg_color=Palette.ACCENT)
            else:
                self._imp_btn_imprimir.configure(state="disabled", fg_color=Palette.ACCENT_DIM)

        self._imp_verificar_btn = _verificar_boton_imprimir

        self._imp_opciones_vars = {}
        self._imp_opciones_copias = {}
        for key, label, default_copias in opciones:
            row = ctk.CTkFrame(right_frame, fg_color="transparent")
            row.pack(fill="x", padx=12, pady=1)
            var = ctk.BooleanVar(value=True)
            var.trace_add("write", _verificar_boton_imprimir)
            cb = ctk.CTkCheckBox(
                row, text=label,
                font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                variable=var,
                fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
                border_color=Palette.BORDER, checkmark_color=Palette.WHITE,
                text_color=Palette.TEXT_PRIMARY,
            )
            cb.pack(side="left")
            if key != "servicio_ata":
                entry = ctk.CTkEntry(
                    row, width=45, height=26,
                    font=ctk.CTkFont(family=FONT_FAMILY, size=11),
                    fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
                    text_color=Palette.TEXT_PRIMARY, corner_radius=4,
                )
                entry.insert(0, str(default_copias))
                entry.pack(side="right", padx=(0, 4))
                ctk.CTkLabel(
                    row, text="copias",
                    font=ctk.CTkFont(family=FONT_FAMILY, size=10),
                    text_color=Palette.TEXT_MUTED,
                ).pack(side="right")
                self._imp_opciones_copias[key] = entry
            else:
                self._imp_opciones_copias[key] = None
            self._imp_opciones_vars[key] = var

        # Botón IMPRIMIR (arranca activo porque hay carpetas y opciones por defecto)
        self._imp_btn_imprimir = ctk.CTkButton(
            right_frame,
            text="🖨  IMPRIMIR",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=36,
            command=self._imp_ejecutar_desde_panel,
        )
        self._imp_btn_imprimir.pack(fill="x", padx=12, pady=(10, 8))

        # Escanear carpetas después de que la interfaz sea visible
        self._imp_lbl_estado.configure(text="🔍 Buscando carpetas...")
        self.after(100, self._imp_escanear_carpetas)

    def _imp_escanear_carpetas(self):
        """Escanea el Escritorio y lista las carpetas de carga en la columna izquierda."""
        for w in self._imp_scroll_carpetas.winfo_children():
            w.destroy()
        self._imp_carpetas_vars.clear()

        # Mostrar indicador de búsqueda en la toolbar
        self._imp_lbl_estado.configure(text="🔍 Buscando carpetas...")
        self.update_idletasks()

        escritorio = self._resolver_ruta("planillas_carga", "Desktop")
        carpetas = []
        if os.path.exists(escritorio):
            for item in sorted(os.listdir(escritorio)):
                ruta = os.path.join(escritorio, item)
                if os.path.isdir(ruta) and not item.startswith(".") and item.upper() not in ("RECYCLED", "RECYCLER"):
                    try:
                        archivos = os.listdir(ruta)
                    except Exception:
                        continue
                    excels = [a for a in archivos if "CONTENEDORES" in a.upper() and (a.upper().endswith(".XLSX") or a.upper().endswith(".XLS"))]
                    if excels:
                        match_frac = re.search(r"(F(?:RACCION)?\s*\d+)", item, re.IGNORECASE)
                        frac = match_frac.group(1).upper() if match_frac else ""
                        carpetas.append((ruta, item, frac))

        if not carpetas:
            ctk.CTkLabel(
                self._imp_scroll_carpetas,
                text="No se encontraron carpetas\nde carga en el Escritorio",
                font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                text_color=Palette.TEXT_MUTED, justify="center",
            ).pack(pady=40)
            self._imp_lbl_estado.configure(text="Sin carpetas")
            return

        self._imp_lbl_estado.configure(text=f"{len(carpetas)} carpetas")

        # Select All
        cb_all = ctk.CTkCheckBox(
            self._imp_scroll_carpetas,
            text="SELECCIONAR TODAS",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            variable=self._imp_select_all_var,
            command=self._imp_toggle_all,
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            border_color=Palette.BORDER, checkmark_color=Palette.WHITE,
            text_color=Palette.TEXT_PRIMARY,
        )
        cb_all.pack(anchor="w", padx=12, pady=(8, 4))

        ctk.CTkFrame(self._imp_scroll_carpetas, fg_color=Palette.DIVIDER, height=1).pack(fill="x", padx=8, pady=(0, 4))

        for ruta, nombre, frac in carpetas:
            var = ctk.BooleanVar(value=True)
            if hasattr(self, '_imp_verificar_btn'):
                var.trace_add("write", self._imp_verificar_btn)
            label = nombre
            if frac:
                label += f"  [{frac}]"
            cb = ctk.CTkCheckBox(
                self._imp_scroll_carpetas, text=label,
                font=ctk.CTkFont(family=FONT_FAMILY, size=11),
                variable=var,
                fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
                border_color=Palette.BORDER, checkmark_color=Palette.WHITE,
                text_color=Palette.TEXT_PRIMARY,
            )
            cb.pack(anchor="w", padx=16, pady=1)
            self._imp_carpetas_vars[ruta] = var

        # Verificar estado del botón IMPRIMIR
        if hasattr(self, '_imp_verificar_btn'):
            self._imp_verificar_btn()

    def _imp_toggle_all(self):
        estado = self._imp_select_all_var.get()
        for var in self._imp_carpetas_vars.values():
            var.set(estado)

    def _imp_diagnosticar(self):
        """Analiza las carpetas seleccionadas y muestra los archivos encontrados SIN imprimir."""
        seleccionadas = [ruta for ruta, var in self._imp_carpetas_vars.items() if var.get()]
        if not seleccionadas:
            messagebox.showinfo("Sin selección", "Seleccione al menos una carpeta para diagnosticar.")
            return
        if self.tarea_activa:
            messagebox.showwarning("Agente ocupado", "Hay una tarea en ejecución.")
            return

        self.tarea_activa = True
        self._imp_btn_imprimir.configure(state="disabled")
        self._imp_btn_refresh.configure(state="disabled")
        self._limpiar_log()

        opciones = {k: v.get() for k, v in self._imp_opciones_vars.items()}
        self._log("══════════ DIAGNÓSTICO DE IMPRESIÓN ══════════")
        self._log(f"Carpetas seleccionadas: {len(seleccionadas)}")
        self._log(f"Opciones activas: {', '.join(k for k, v in opciones.items() if v)}")
        self._log("")

        t = threading.Thread(target=self._imp_diag_worker, args=(seleccionadas, opciones), daemon=True)
        t.start()

    def _imp_diag_worker(self, seleccionadas, opciones):
        """Worker de diagnóstico: lista archivos encontrados sin imprimir."""
        self._set_log_panel("impresion")
        anio = datetime.now().strftime("%y")
        prefijo_permiso = f"{anio}069EC"

        for ruta in seleccionadas:
            nombre = os.path.basename(ruta)
            self._log(f"📂 {nombre}")
            self._log(f"   Ruta: {ruta}")
            try:
                archivos = sorted(os.listdir(ruta))
            except Exception as e:
                self._log(f"   ❌ Error al leer carpeta: {e}")
                continue

            # Clasificar archivos
            sobres = [a for a in archivos if "CONTENEDORES" in a.upper() and (a.upper().endswith(".XLSX") or a.upper().endswith(".XLS"))]
            permisos = [a for a in archivos if a.upper().startswith(prefijo_permiso) and a.upper().endswith(".PDF")]
            # Buscar Hoja de Ruta: flexible, que contenga "HOJA" y "RUTA" (no requiere "DE")
            hojas_ruta = [a for a in archivos if "HOJA" in a.upper() and "RUTA" in a.upper() and (a.upper().endswith(".XLSX") or a.upper().endswith(".XLS"))]

            # Buscar también con patrón más amplio si no se encontró
            if not permisos:
                # Buscar cualquier PDF que empiece con el año
                permisos = [a for a in archivos if a.upper().endswith(".PDF") and a.upper()[:2] == anio]

            self._log(f"   ┌─ SOBRE (Excel CONTENEDORES):")
            if sobres:
                for a in sobres:
                    self._log(f"   │  ✓ {a}")
            else:
                self._log(f"   │  ❌ No encontrado (busca: *CONTENEDORES*.xlsx)")

            self._log(f"   ├─ PERMISO EXPORTACIÓN (PDF {prefijo_permiso}*):")
            if permisos:
                for a in permisos:
                    self._log(f"   │  ✓ {a}")
            else:
                self._log(f"   │  ❌ No encontrado (busca: {prefijo_permiso}*.PDF)")
                # Mostrar todos los PDF para ayudar a diagnosticar
                todos_pdf = [a for a in archivos if a.upper().endswith(".PDF")]
                if todos_pdf:
                    self._log(f"   │  ⚠ PDFs en la carpeta: {todos_pdf}")

            self._log(f"   ├─ HOJA DE RUTA (Excel HOJA DE RUTA*):")
            if hojas_ruta:
                for a in hojas_ruta:
                    self._log(f"   │  ✓ {a}")
            else:
                self._log(f"   │  ❌ No encontrado (busca: HOJA DE RUTA*.xlsx)")

            # Detectar hojas Recibo ATA: nombres exactos "Recibo Ata", "Recibo Ata 2", ... "Recibo Ata 8"
            hojas_ata = []
            todas_hojas = []
            nombres_ata = [f"RECIBO ATA"] + [f"RECIBO ATA {i}" for i in range(2, 9)]
            if sobres:
                for sobre in sobres:
                    ruta_excel = os.path.join(ruta, sobre)
                    try:
                        if sobre.lower().endswith(".xlsx"):
                            wb = openpyxl.load_workbook(ruta_excel, read_only=True)
                            todas_hojas = list(wb.sheetnames)
                            for sn in wb.sheetnames:
                                if sn.upper() in nombres_ata:
                                    hojas_ata.append((sobre, sn))
                            wb.close()
                        elif sobre.lower().endswith(".xls"):
                            book = xlrd.open_workbook(ruta_excel)
                            todas_hojas = list(book.sheet_names())
                            for sn in book.sheet_names():
                                if sn.upper() in nombres_ata:
                                    hojas_ata.append((sobre, sn))
                    except Exception as e:
                        self._log(f"   │  ⚠ Error leyendo hojas de {sobre}: {e}")

            self._log(f"   ├─ RECIBOS ATA (hojas dentro del Excel, {len(hojas_ata)} choferes):")
            if hojas_ata:
                for archivo, hoja in hojas_ata:
                    self._log(f"   │  ✓ {archivo} → Hoja: \"{hoja}\"")
            else:
                self._log(f"   │  ❌ No encontrado. Hojas disponibles: {todas_hojas}")

            # Mostrar TODOS los archivos reales de la carpeta
            self._log(f"   └─ TODOS los archivos ({len(archivos)}):")
            for a in archivos:
                self._log(f"      · {a}")
            self._log("")

        self._log("══════════ DIAGNÓSTICO COMPLETADO ══════════")
        self.after(0, self._imp_diag_done)

    def _imp_diag_done(self):
        self.tarea_activa = False
        try:
            self._imp_btn_imprimir.configure(state="normal")
            self._imp_btn_refresh.configure(state="normal")
        except (AttributeError, Exception):
            pass

    def _imp_popup_dorsos(self):
        """Popup para imprimir dorsos con cantidad de hojas editable."""
        popup = ctk.CTkToplevel(self)
        popup.title("Impresión de Dorsos")
        popup.geometry("400x330")
        popup.configure(fg_color=Palette.BG_CARD)
        popup.transient(self)
        popup.lift()

        popup.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 400) // 2
        py = self.winfo_y() + (self.winfo_height() - 330) // 2
        popup.geometry(f"400x330+{px}+{py}")

        ctk.CTkLabel(
            popup, text="IMPRESIÓN DE DORSOS",
            font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(pady=(20, 16))

        var_mic = ctk.BooleanVar(value=False)
        var_crt = ctk.BooleanVar(value=False)
        var_pe = ctk.BooleanVar(value=False)

        def _actualizar_boton():
            if var_mic.get() or var_crt.get() or var_pe.get():
                btn_imprimir.configure(state="normal", fg_color=Palette.ACCENT)
            else:
                btn_imprimir.configure(state="disabled", fg_color=Palette.ACCENT_DIM)

        def _crear_fila(label, var, default_cant):
            frame = ctk.CTkFrame(popup, fg_color="transparent")
            frame.pack(fill="x", padx=30, pady=4)
            cb = ctk.CTkCheckBox(
                frame, text=label,
                font=ctk.CTkFont(family=FONT_FAMILY, size=13),
                variable=var, command=_actualizar_boton,
                fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
                border_color=Palette.BORDER, checkmark_color=Palette.WHITE,
                text_color=Palette.TEXT_PRIMARY,
            )
            cb.pack(side="left")
            entry = ctk.CTkEntry(
                frame, width=60, height=28,
                font=ctk.CTkFont(family=FONT_FAMILY, size=13),
                fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
                text_color=Palette.TEXT_PRIMARY, corner_radius=4,
            )
            entry.insert(0, str(default_cant))
            entry.pack(side="right")
            ctk.CTkLabel(
                frame, text="hojas",
                font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                text_color=Palette.TEXT_MUTED,
            ).pack(side="right", padx=(0, 6))
            return entry

        def_mic = self._cfg_obtener_docs("dorso_mic", 15)
        def_crt = self._cfg_obtener_docs("dorso_crt", 4)
        def_pe  = self._cfg_obtener_docs("dorso_pe", 2)
        entry_mic = _crear_fila("Dorso MIC", var_mic, def_mic)
        entry_crt = _crear_fila("Dorso CRT", var_crt, def_crt)
        entry_pe = _crear_fila("Dorso PE", var_pe, def_pe)

        def _imprimir_dorsos():
            base = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
            archivos = {
                "MIC": os.path.join(base, "DORSO MIC.pdf"),
                "CRT": os.path.join(base, "DORSO CRT.pdf"),
                "PE":  os.path.join(base, "dorso PE.pdf"),
            }
            # Leer estado de tkinter en hilo principal antes de destruir el popup
            tipos = []
            for tipo, var, entry in [("MIC", var_mic, entry_mic),
                                      ("CRT", var_crt, entry_crt),
                                      ("PE", var_pe, entry_pe)]:
                if var.get():
                    try:
                        n = int(entry.get())
                    except ValueError:
                        n = (self._cfg_obtener_docs("dorso_mic", 15) if tipo == "MIC" else
                             self._cfg_obtener_docs("dorso_crt", 4) if tipo == "CRT" else
                             self._cfg_obtener_docs("dorso_pe", 2))
                    tipos.append((tipo, n, archivos[tipo]))
            popup.destroy()

            def _esperar_cola_vacia():
                # Espera 3s para que el visor PDF envíe las últimas copias al spooler
                time.sleep(3.0)
                try:
                    import win32print
                    printer = win32print.GetDefaultPrinter()
                    deadline = time.time() + 120
                    while time.time() < deadline:
                        h = win32print.OpenPrinter(printer)
                        try:
                            jobs = win32print.EnumJobs(h, 0, 100, 1)
                        finally:
                            win32print.ClosePrinter(h)
                        if not jobs:
                            return
                        time.sleep(1.5)
                except Exception:
                    time.sleep(10.0)

            def _worker():
                self._set_log_panel("impresion")
                total = 0
                for idx, (tipo, n, ruta) in enumerate(tipos):
                    if not os.path.exists(ruta):
                        self._log(f"⚠ Dorso {tipo}: archivo no encontrado ({ruta})")
                        continue
                    for copia in range(n):
                        os.startfile(ruta, "print")
                        time.sleep(0.5)
                    total += n
                    self._log(f"Dorso {tipo}: {n} copias enviadas → {os.path.basename(ruta)}")
                    # Si hay más tipos después, esperar que la cola se vacíe antes de continuar
                    if idx < len(tipos) - 1:
                        self._log(f"  ↳ Esperando cola de impresión antes de siguiente dorso...")
                        _esperar_cola_vacia()
                self._log(f"Total dorsos: {total} hojas")

            threading.Thread(target=_worker, daemon=True).start()

        btn_imprimir = ctk.CTkButton(
            popup, text="🖨  IMPRIMIR DORSOS",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=40,
            command=_imprimir_dorsos,
        )
        btn_imprimir.pack(fill="x", padx=30, pady=(20, 12))

    def _imp_popup_sellos_mic(self):
        """Popup para imprimir Sellos MIC: guardas = nombres de hojas del Excel."""
        # Usar la misma lógica de búsqueda que CRT
        base = self._cfg_obtener_rutas("mic_sellos", os.path.join("TRABAJO", "01_PLANILLAS"))
        ruta_mic = buscar_archivo_en_pendrive("FECHA MIC Y SELLOS.xlsx", base)
        if not ruta_mic:
            ruta_mic = buscar_archivo_en_pendrive("FECHA MIC Y SELLOS.xls", base)

        guardas = []
        if ruta_mic:
            try:
                wb = self._abrir_excel_seguro(ruta_mic)
                guardas = list(wb.sheetnames)
                wb.close()
            except Exception:
                pass

        if not guardas:
            messagebox.showerror("Archivo no encontrado",
                f"No se encontró FECHA MIC Y SELLOS.\n\nRuta configurada: {base}\n"
                f"Se buscó en D-Z y en la carpeta del programa.")
            return

        fecha_hoy = datetime.now().strftime("%Y-%m-%d")
        ultimo_guarda_dia = self._cfg_obtener("super_auto", f"mic_guarda_{fecha_hoy}", "")

        popup = ctk.CTkToplevel(self)
        popup.title("Sellos MIC")
        popup.geometry("400x270")
        popup.configure(fg_color=Palette.BG_CARD)
        popup.transient(self)
        popup.lift()
        popup.grab_set()
        popup.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 400) // 2
        py = self.winfo_y() + (self.winfo_height() - 270) // 2
        popup.geometry(f"400x270+{px}+{py}")

        ctk.CTkLabel(popup, text="SELLOS MIC",
                     font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
                     text_color=Palette.TEXT_PRIMARY).pack(pady=(20, 16))

        ctk.CTkLabel(popup, text="Seleccioná el guarda:",
                     font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                     text_color=Palette.TEXT_SECONDARY).pack(anchor="w", padx=30, pady=(0, 4))
        guarda_var = ctk.StringVar(value=ultimo_guarda_dia if ultimo_guarda_dia in guardas else (guardas[0] if guardas else ""))
        ctk.CTkOptionMenu(popup, variable=guarda_var, values=guardas,
                          font=ctk.CTkFont(family=FONT_FAMILY, size=13),
                          fg_color=Palette.BG_INPUT, button_color=Palette.ACCENT,
                          button_hover_color=Palette.ACCENT_HOVER,
                          text_color=Palette.TEXT_PRIMARY, corner_radius=6,
                          width=240, height=34).pack(pady=(0, 12))

        ctk.CTkLabel(popup, text="Cantidad de copias:",
                     font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                     text_color=Palette.TEXT_SECONDARY).pack(anchor="w", padx=30, pady=(0, 4))
        copia_entry = ctk.CTkEntry(popup, width=80, height=30,
                                   font=ctk.CTkFont(family=FONT_FAMILY, size=13),
                                   fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
                                   text_color=Palette.TEXT_PRIMARY, corner_radius=4)
        copia_entry.insert(0, "1")
        copia_entry.pack(pady=(0, 12))

        def _imprimir():
            guarda = guarda_var.get()
            try: copias = int(copia_entry.get().strip())
            except ValueError: copias = 1
            self.config.setdefault("super_auto", {})[f"mic_guarda_{fecha_hoy}"] = guarda
            self._guardar_config()
            popup.destroy()
            self._imp_sellos_mic_worker(guarda, copias, ruta_mic)

        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=(8, 0))
        ctk.CTkButton(btn_frame, text="🖨  Imprimir", width=120, height=34,
                      font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
                      fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
                      text_color=Palette.WHITE, corner_radius=6,
                      command=_imprimir).pack(side="left", padx=6)
        ctk.CTkButton(btn_frame, text="Cancelar", width=100, height=34,
                      font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                      fg_color=Palette.BG_HOVER, hover_color=Palette.ERROR,
                      text_color=Palette.TEXT_SECONDARY, corner_radius=6,
                      command=popup.destroy).pack(side="left", padx=6)

    def _imp_sellos_mic_worker(self, guarda, copias, ruta_mic=None):
        """Busca e imprime la hoja del guarda en FECHA MIC Y SELLOS."""
        if not ruta_mic:
            base = self._cfg_obtener_rutas("mic_sellos", os.path.join("TRABAJO", "01_PLANILLAS"))
            ruta_mic = buscar_archivo_en_pendrive("FECHA MIC Y SELLOS.xlsx", base)
            if not ruta_mic:
                ruta_mic = buscar_archivo_en_pendrive("FECHA MIC Y SELLOS.xls", base)
        if not ruta_mic:
            self._log(f"⚠ Sellos MIC: no se encontró FECHA MIC Y SELLOS")
            return

        def _worker():
            self._set_log_panel("impresion")
            try:
                wb = self._abrir_excel_seguro(ruta_mic)
                hoja_encontrada = None
                for sn in wb.sheetnames:
                    if guarda.upper() in sn.upper():
                        hoja_encontrada = sn; break
                if not hoja_encontrada:
                    self.log_queue.put(f"[...] ⚠ Sellos MIC: guarda '{guarda}' no hallado en ninguna hoja")
                    wb.close(); return
                wb.close()
                impresora = self._detectar_impresoras()[0] if self._detectar_impresoras() else "Default"
                self.log_queue.put(f"[...] 🖨 Sellos MIC: {guarda} → hoja '{hoja_encontrada}' ({copias} copias)")
                self._imp_enviar(ruta_mic, impresora, f"  Sellos MIC - {guarda}", hojas=[hoja_encontrada], copias=copias)
            except Exception as e:
                self.log_queue.put(f"[...] ⚠ Error Sellos MIC: {e}")
        t = threading.Thread(target=_worker, daemon=True)
        t.start()

    def _imp_popup_sellos_crt(self):
        """Popup para imprimir Sellos CRT: cantidad de copias."""
        popup = ctk.CTkToplevel(self)
        popup.title("Sellos CRT")
        popup.geometry("360x180")
        popup.configure(fg_color=Palette.BG_CARD)
        popup.transient(self)
        popup.lift()
        popup.grab_set()
        popup.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 360) // 2
        py = self.winfo_y() + (self.winfo_height() - 180) // 2
        popup.geometry(f"360x180+{px}+{py}")

        ctk.CTkLabel(popup, text="SELLOS CRT",
                     font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
                     text_color=Palette.TEXT_PRIMARY).pack(pady=(20, 16))

        ctk.CTkLabel(popup, text="Cantidad de copias:",
                     font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                     text_color=Palette.TEXT_SECONDARY).pack(pady=(0, 4))
        copia_entry = ctk.CTkEntry(popup, width=80, height=30,
                                   font=ctk.CTkFont(family=FONT_FAMILY, size=13),
                                   fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
                                   text_color=Palette.TEXT_PRIMARY, corner_radius=4)
        copia_entry.insert(0, "1")
        copia_entry.pack(pady=(0, 12))

        def _imprimir():
            try: copias = int(copia_entry.get().strip())
            except ValueError: copias = 1
            popup.destroy()
            self._imp_sellos_crt_worker(copias)

        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=(4, 0))
        ctk.CTkButton(btn_frame, text="🖨  Imprimir", width=120, height=34,
                      font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
                      fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
                      text_color=Palette.WHITE, corner_radius=6,
                      command=_imprimir).pack(side="left", padx=6)
        ctk.CTkButton(btn_frame, text="Cancelar", width=100, height=34,
                      font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                      fg_color=Palette.BG_HOVER, hover_color=Palette.ERROR,
                      text_color=Palette.TEXT_SECONDARY, corner_radius=6,
                      command=popup.destroy).pack(side="left", padx=6)

    def _imp_sellos_crt_worker(self, copias):
        """Imprime hoja 'Hoja1' de FECHA CRT Y ORIGINAL."""
        base = self._cfg_obtener_rutas("crt_original", os.path.join("TRABAJO", "01_PLANILLAS"))
        ruta_crt = buscar_archivo_en_pendrive("FECHA CRT Y ORIGINAL.xlsx", base)
        if not ruta_crt:
            ruta_crt = buscar_archivo_en_pendrive("FECHA CRT Y ORIGINAL.xls", base)
        if not ruta_crt:
            self._log(f"⚠ Sellos CRT: no se encontró FECHA CRT Y ORIGINAL en {base} (D-Z)")
            return

        def _worker():
            self._set_log_panel("impresion")
            try:
                wb = self._abrir_excel_seguro(ruta_crt)
                hoja = wb.sheetnames[0]
                wb.close()
                impresora = self._detectar_impresoras()[0] if self._detectar_impresoras() else "Default"
                self.log_queue.put(f"[...] 🖨 Sellos CRT: hoja '{hoja}' ({copias} copias)")
                self._imp_enviar(ruta_crt, impresora, f"  Sellos CRT", hojas=[hoja], copias=copias)
            except Exception as e:
                self.log_queue.put(f"[...] ⚠ Error Sellos CRT: {e}")
        t = threading.Thread(target=_worker, daemon=True)
        t.start()

    def _imp_ejecutar_desde_panel(self):
        seleccionadas = [ruta for ruta, var in self._imp_carpetas_vars.items() if var.get()]
        if not seleccionadas:
            messagebox.showinfo("Sin selección", "Seleccione al menos una carpeta para imprimir.")
            return
        self._imp_ejecutar(seleccionadas)

    def _detectar_impresoras(self):
        """Devuelve lista de impresoras disponibles con nombre completo (Name on Port:)."""
        try:
            import subprocess
            result = subprocess.run(
                ["powershell", "-Command",
                 "Get-Printer | ForEach-Object { $_.Name + ' on ' + $_.PortName + ':' }"],
                capture_output=True, text=True, timeout=10
            )
            if result.returncode == 0 and result.stdout.strip():
                printers = [p.strip() for p in result.stdout.strip().split("\n") if p.strip()]
                # Poner Microsoft Print to PDF primero si existe
                pdf = [p for p in printers if "PDF" in p.upper()]
                otros = [p for p in printers if "PDF" not in p.upper()]
                return pdf + otros
        except Exception:
            pass
        return ["Microsoft Print to PDF (simulación)"]

    def _detectar_tipo_carpeta(self, nombre_carpeta: str) -> str:
        """Return transport type from folder name: TERRESTRE, ISO, or FLEXI.

        Folder format: DD_MM_YYYY_CANT_TIPO_PE_CARPETA_DEST_[SUFFIX]
        The type is at index 4 (0-indexed). Old-format folders default to TERRESTRE.
        """
        partes = nombre_carpeta.split("_")
        if len(partes) < 5:
            return "TERRESTRE"
        tipo = partes[4]
        if tipo not in ("TERRESTRE", "ISO", "FLEXI"):
            return "TERRESTRE"
        return tipo

    def _imp_ejecutar(self, seleccionadas):
        """Ejecuta la cola de impresión."""
        if self.tarea_activa:
            messagebox.showwarning("Agente ocupado", "Hay una tarea en ejecución.")
            return

        self.tarea_activa = True
        self._imp_btn_imprimir.configure(text="⏳  Imprimiendo...", state="disabled")
        self._imp_btn_refresh.configure(state="disabled")
        self._limpiar_log()

        opciones = {k: v.get() for k, v in self._imp_opciones_vars.items()}
        # Leer copias de los entries (usar defaults si el entry no existe)
        copias = {}
        for key, ent in self._imp_opciones_copias.items():
            if ent is not None:
                try:
                    copias[key] = int(ent.get().strip())
                except ValueError:
                    copias[key] = 2
            else:
                copias[key] = 0
        impresora = self._detectar_impresoras()[0] if self._detectar_impresoras() else "Default"
        self._log(f"IMPRESORA: {impresora}")
        self._log(f"Carpetas a procesar: {len(seleccionadas)}")
        self._log(f"Opciones: {', '.join(k for k, v in opciones.items() if v)}")
        self._log("─" * 40)

        t = threading.Thread(
            target=self._imp_worker,
            args=(seleccionadas, opciones, impresora, copias),
            daemon=True,
        )
        t.start()

    def _imp_worker(self, seleccionadas, opciones, impresora, copias=None):
        """Hilo de fondo: procesa cada carpeta y lanza las impresiones."""
        self._set_log_panel("impresion")
        if copias is None:
            copias = {}
        anio = datetime.now().strftime("%y")
        # El permiso de exportación empieza con el año + 069EC (ej: 26069EC...)
        prefijo_permiso = f"{anio}069EC"
        total_ok = 0

        try:
            for ruta in seleccionadas:
                nombre = os.path.basename(ruta)
                self._log(f"Procesando: {nombre}")
                self._log(f"  Carpeta: {ruta}")
                archivos = sorted(os.listdir(ruta))

                # Listar todo lo que hay en la carpeta
                self._log(f"  Archivos en la carpeta ({len(archivos)}):")
                for a in archivos:
                    self._log(f"    · {a}")

                # Detectar archivos por tipo
                sobres = [a for a in archivos if "CONTENEDORES" in a.upper() and (a.upper().endswith(".XLSX") or a.upper().endswith(".XLS"))]
                permisos = [a for a in archivos if a.upper().startswith(prefijo_permiso) and a.upper().endswith(".PDF")]
                # Buscar Hoja de Ruta: flexible, que contenga "HOJA" y "RUTA" (no requiere "DE")
                hojas_ruta = [a for a in archivos if "HOJA" in a.upper() and "RUTA" in a.upper() and (a.upper().endswith(".XLSX") or a.upper().endswith(".XLS"))]

                # Detectar hojas Recibo ATA: nombres exactos
                nombres_ata = ["RECIBO ATA"] + [f"RECIBO ATA {i}" for i in range(2, 9)]
                hojas_ata = []
                if sobres:
                    for sobre in sobres:
                        ruta_excel = os.path.join(ruta, sobre)
                        try:
                            if sobre.lower().endswith(".xlsx"):
                                wb = openpyxl.load_workbook(ruta_excel, read_only=True)
                                for sn in wb.sheetnames:
                                    if sn.upper() in nombres_ata:
                                        hojas_ata.append((sobre, sn))
                                wb.close()
                            elif sobre.lower().endswith(".xls"):
                                book = xlrd.open_workbook(ruta_excel)
                                for sn in book.sheet_names():
                                    if sn.upper() in nombres_ata:
                                        hojas_ata.append((sobre, sn))
                        except Exception:
                            pass
                cant_choferes = len(hojas_ata) if hojas_ata else 1

                self._log(f"  Detectado: Sobre={sobres}, Permiso={permisos}, HojaRuta={hojas_ruta}, RecibosATA(hojas)={hojas_ata}, Choferes={cant_choferes}")

                # 1. Sobre: imprimir SOLO la hoja "SOBRE"
                if opciones.get("sobre", True):
                    for a in sobres:
                        hojas_sobre = self._imp_hojas_sobre(os.path.join(ruta, a))
                        self._imp_enviar(os.path.join(ruta, a), impresora, f"Sobre: {a}", hojas=hojas_sobre)
                        total_ok += 1

                # 2. Permiso de Exportación (PDF)
                if opciones.get("permiso"):
                    if permisos:
                        n_copias_permiso = copias.get("permiso", 2) or self._cfg_obtener_docs("permiso_exportacion", 2)
                        for a in permisos:
                            for copia_num in range(n_copias_permiso):
                                self._imp_enviar(os.path.join(ruta, a), impresora, f"Permiso Exp. (copia {copia_num+1}/{n_copias_permiso}): {a}")
                                total_ok += 1
                    else:
                        self._log(f"  ⚠ No se encontró Permiso de Exportación (busca: {prefijo_permiso}*.PDF)")

                # 3. Hoja de Ruta (Excel)
                if opciones.get("hoja_ruta"):
                    if hojas_ruta:
                        n_copias_hr = copias.get("hoja_ruta", 2) or self._cfg_obtener_docs("hoja_ruta", 2)
                        for a in hojas_ruta:
                            self._imp_enviar(os.path.join(ruta, a), impresora, f"Hoja Ruta ({n_copias_hr} copias): {a}", copias=n_copias_hr)
                            total_ok += 1
                    else:
                        self._log(f"  ⚠ No se encontró Hoja de Ruta (busca: HOJA DE RUTA*.xls)")

                # 4. Servicio ATA / Recibo ATA: imprimir SOLO las hojas exactas
                if opciones.get("servicio_ata"):
                    tipo = self._detectar_tipo_carpeta(nombre)
                    if tipo in ("ISO", "FLEXI"):
                        self._log(f"  ⏭ Saltando Recibo ATA: carpeta marítima ({tipo})")
                    else:
                        if hojas_ata:
                            for archivo, hoja in hojas_ata:
                                self._imp_enviar(os.path.join(ruta, archivo), impresora, f"Recibo ATA: {archivo} → {hoja}", hojas=[hoja])
                                total_ok += 1
                        else:
                            self._log(f"  ⚠ No se encontraron hojas Recibo Ata en el Excel")

            self._log(f"COMPLETADO: Impresión finalizada — {total_ok} documentos enviados.")
        except Exception as e:
            self._log(f"ERROR en impresión: {e}")
            traceback.print_exc()
        finally:
            self.after(0, self._imp_done)

    def _imp_hojas_sobre(self, ruta_excel):
        """Busca la hoja 'SOBRE'. Si no existe, devuelve la primera hoja."""
        try:
            if ruta_excel.lower().endswith(".xlsx"):
                wb = openpyxl.load_workbook(ruta_excel, read_only=True)
                nombres = list(wb.sheetnames)
                for sn in nombres:
                    if sn.upper() == "SOBRE":
                        wb.close()
                        return [sn]
                wb.close()
                return [nombres[0]] if nombres else None
            elif ruta_excel.lower().endswith(".xls"):
                book = xlrd.open_workbook(ruta_excel)
                nombres = list(book.sheet_names())
                for sn in nombres:
                    if sn.upper() == "SOBRE":
                        return [sn]
                return [nombres[0]] if nombres else None
        except Exception:
            pass
        return None  # error, imprimir todo

    def _imp_enviar(self, ruta_archivo, impresora, descripcion, hojas=None, copias=1):
        """Envía un archivo a la impresora seleccionada. hojas=lista, copias=N."""
        nombre_archivo = os.path.basename(ruta_archivo)
        self._log(f"  {descripcion}")
        self._log(f"     Archivo: {nombre_archivo}")
        if hojas:
            self._log(f"     Hojas: {hojas}")
        try:
            if "simulación" in impresora.lower():
                self._log(f"     → SIMULADO (sin impresora)")
                return True

            self._log(f"     → Impresora: {impresora}")
            ext = os.path.splitext(ruta_archivo)[1].upper()
            es_excel = ext in (".XLSX", ".XLS")

            if es_excel and self._excel_com_ok:
                result = self._imp_excel_com(ruta_archivo, impresora, hojas, copias)
            else:
                if es_excel and not self._excel_com_ok:
                    self._log(f"     → Modo visible (COM no disponible, se abrirá Excel)")
                if "pdf" in impresora.lower():
                    self._log(f"     → Se abrirá ventana para guardar PDF")
                for i in range(copias):
                    os.startfile(ruta_archivo, "print")
                    if copias > 1:
                        self._log(f"     → Copia {i+1}/{copias}")
                    time.sleep(2.0)  # os.startfile es asíncrono, pausa entre copias
                result = True
            return result
        except Exception as e:
            self._log(f"     → ERROR: {e}")
            return False

    def _imp_excel_com(self, ruta, impresora, hojas=None, copias=1):
        """Imprime Excel con Python win32com (sin abrir ventana)."""
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                excel.ActivePrinter = impresora
            except Exception as e:
                self._log(f"     → ActivePrinter: {e}")

            wb = excel.Workbooks.Open(ruta)

            if hojas:
                for nombre_hoja in hojas:
                    try:
                        ws = wb.Worksheets(nombre_hoja)
                        # PrintOut(From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate)
                        ws.PrintOut(1, 9999, copias)
                        self._log(f"     → Hoja '{nombre_hoja}' enviada ({copias} copias)")
                    except Exception as e:
                        self._log(f"     → Error hoja '{nombre_hoja}': {e}")
            else:
                wb.PrintOut(1, 9999, copias)
                self._log(f"     → OK ({copias} copias)")

            wb.Close(SaveChanges=False)
            excel.Quit()
            return True
        except Exception as e:
            self._log(f"     → Error COM: {e}")
            return False
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _imp_done(self):
        self.tarea_activa = False
        try:
            self._imp_btn_imprimir.configure(
                text="🖨  IMPRIMIR", state="normal",
                fg_color=Palette.ACCENT,
            )
            self._imp_btn_refresh.configure(text="🔄  Refrescar", state="normal", command=self._imp_escanear_carpetas)
        except (AttributeError, Exception):
            pass

    # ═══════════════════════════════════════════════════════════════════
    # PANEL 2: PLANILLAS / SOBRES
    # ═══════════════════════════════════════════════════════════════════
    def _panel_planillas(self):
        if "planillas" in self._panel_frames:
            return
        frame = ctk.CTkFrame(self.panel_container, fg_color="transparent")
        self._panel_frames["planillas"] = frame
        # frame.pack(fill="both", expand=True) — packed by _cambiar_panel_forzado

        # ── Toolbar ──────────────────────────────────────────────────
        toolbar = ctk.CTkFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER, height=44
        )
        toolbar.pack(fill="x", pady=(0, 6))
        toolbar.pack_propagate(False)

        self.btn_ejecutar_planillas = ctk.CTkButton(
            toolbar,
            text="▶  Completar Planillas",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.ACCENT,
            hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE,
            corner_radius=6,
            height=34,
            width=200,
            command=self._popup_completar_planillas,
        )
        self.btn_ejecutar_planillas.pack(side="left", padx=4, pady=4)

        self.btn_agregar_guarda = ctk.CTkButton(
            toolbar,
            text="🛡  Agregar Guarda",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.SECONDARY,
            hover_color=Palette.SECONDARY_HOVER,
            text_color=Palette.WHITE,
            corner_radius=6,
            height=34,
            width=160,
            command=self._popup_agregar_guarda,
        )
        self.btn_agregar_guarda.pack(side="left", padx=4, pady=4)

        self.btn_editar_excels = ctk.CTkButton(
            toolbar, text="📝 Editar Excels",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_HOVER, hover_color=Palette.ACCENT_DIM,
            text_color=Palette.TEXT_PRIMARY, corner_radius=6, height=34,
            command=self._popup_editar_excels,
        )
        self.btn_editar_excels.pack(side="left", padx=4, pady=4)

        self.lbl_estado_planillas = ctk.CTkLabel(
            toolbar,
            text="Listo para analizar el Escritorio",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.TEXT_PRIMARY,
        )
        self.lbl_estado_planillas.pack(side="left", padx=(8, 0))

        self.progress_planillas = ctk.CTkProgressBar(
            toolbar, width=160, height=8, corner_radius=4,
            fg_color=Palette.BG_INPUT, progress_color=Palette.ACCENT,
        )
        self.progress_planillas.pack(side="right", padx=16)
        self.progress_planillas.set(0)

        self.btn_limpiar_planillas = ctk.CTkButton(
            toolbar,
            text="Limpiar",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color="transparent",
            hover_color=Palette.BG_HOVER,
            text_color=Palette.TEXT_PRIMARY,
            corner_radius=4, height=30, width=70,
            command=self._limpiar_planillas,
        )
        self.btn_limpiar_planillas.pack(side="right", padx=4)

        # ── Tabla de resultados ─────────────────────────────────────
        self._crear_tabla_planillas(frame)

        # ── Resumen ──────────────────────────────────────────────────
        resumen_frame = ctk.CTkFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER, height=48
        )
        resumen_frame.pack(fill="x", pady=(6, 0))
        resumen_frame.pack_propagate(False)

        self.lbl_resumen_planillas = ctk.CTkLabel(
            resumen_frame,
            text="Sin datos analizados",
            font=ctk.CTkFont(family=FONT_FAMILY, size=14),
            text_color=Palette.TEXT_MUTED,
        )
        self.lbl_resumen_planillas.pack(side="left", padx=16, pady=12)

    def _crear_tabla_planillas(self, parent):
        """Tabla estilo Treeview dentro de un frame oscuro."""
        tabla_frame = ctk.CTkFrame(
            parent, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER
        )
        tabla_frame.pack(fill="both", expand=True)

        # Treeview con estilo oscuro
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Contenedores.Treeview",
            background=Palette.BG_TABLE,
            foreground=Palette.TEXT_PRIMARY,
            fieldbackground=Palette.BG_TABLE,
            borderwidth=0,
            font=(FONT_FAMILY, 10),
            rowheight=28,
        )
        style.configure(
            "Contenedores.Treeview.Heading",
            background=Palette.BG_SIDEBAR,
            foreground=Palette.TEXT_SECONDARY,
            font=(FONT_FAMILY, 10, "bold"),
            borderwidth=0,
            padding=(8, 6),
        )
        style.map(
            "Contenedores.Treeview",
            background=[("selected", Palette.ACCENT_DIM)],
            foreground=[("selected", Palette.WHITE)],
        )

        columnas = (
            "fecha", "cant", "pe", "carpeta", "terminal",
            "transporte", "destinatario", "bl", "fraccion", "servicio"
        )
        headers = (
            "Fecha Carga", "Cant.", "P.E.", "Carpeta", "Terminal",
            "Transporte", "Destinatario", "B/L", "Fracción", "Servicio (ATA)"
        )
        anchos = (90, 50, 65, 65, 70, 120, 110, 70, 80, 100)

        self.tree_planillas = ttk.Treeview(
            tabla_frame, columns=columnas, show="headings",
            style="Contenedores.Treeview", selectmode="browse",
        )
        for col, hdr, w in zip(columnas, headers, anchos):
            self.tree_planillas.heading(col, text=hdr)
            self.tree_planillas.column(col, width=w, anchor="center", minwidth=w)

        # Scrollbar horizontal (arriba)
        scroll_planillas_x = ctk.CTkScrollbar(
            tabla_frame, orientation="horizontal",
            fg_color=Palette.BG_CARD, button_color=Palette.TEXT_MUTED,
            button_hover_color=Palette.TEXT_SECONDARY,
        )
        self.tree_planillas.configure(xscrollcommand=scroll_planillas_x.set)
        scroll_planillas_x.configure(command=self.tree_planillas.xview)
        scroll_planillas_x.pack(fill="x", padx=2, pady=(2, 0))

        # Scrollbar vertical
        scroll_y = ctk.CTkScrollbar(
            tabla_frame, orientation="vertical",
            fg_color=Palette.BG_CARD, button_color=Palette.TEXT_MUTED,
            button_hover_color=Palette.TEXT_SECONDARY,
        )
        scroll_y.pack(side="right", fill="y", padx=(0, 2), pady=2)

        self.tree_planillas.configure(yscrollcommand=scroll_y.set)
        scroll_y.configure(command=self.tree_planillas.yview)

        self.tree_planillas.pack(fill="both", expand=True, padx=2, pady=(0, 0))

        # Tags de color por empresa
        self.tree_planillas.tag_configure("vitapro", background="#003340", foreground="#87ceeb")
        self.tree_planillas.tag_configure("ewos", background="#003300", foreground="#00ff00")
        self.tree_planillas.tag_configure("dicoal", background="#4d2e00", foreground="#ffb74d")
        self.tree_planillas.tag_configure("nutreco", background="#4d4100", foreground="#ffee58")
        self.tree_planillas.tag_configure("cargill", background="#0d2a59", foreground="#66b2ff")
        self.tree_planillas.tag_configure("biomar", background="#1a4d2e", foreground="#69f0ae")

    # ═══════════════════════════════════════════════════════════════════
    # PANEL 3: CORREOS
    # ═══════════════════════════════════════════════════════════════════
    def _panel_correos(self):
        if "correos" in self._panel_frames:
            return
        frame = ctk.CTkFrame(self.panel_container, fg_color="transparent")
        self._panel_frames["correos"] = frame
        # frame.pack(fill="both", expand=True) — packed by _cambiar_panel_forzado

        # ── Toolbar ──────────────────────────────────────────────────
        toolbar = ctk.CTkFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER, height=56
        )
        toolbar.pack(fill="x", pady=(0, 6))
        toolbar.pack_propagate(False)

        self.btn_ejecutar_correos = ctk.CTkButton(
            toolbar,
            text="▶  Procesar y Despachar Correos",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.ACCENT,
            hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE,
            corner_radius=6,
            height=34,
            width=220,
            command=self._ejecutar_agente_correos,
        )
        self.btn_ejecutar_correos.pack(side="left", padx=4, pady=4)

        self.btn_elegir_correos = ctk.CTkButton(
            toolbar,
            text="📂  Elegir Correos",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.SECONDARY,
            hover_color=Palette.SECONDARY_HOVER,
            text_color=Palette.WHITE,
            corner_radius=6,
            height=34,
            width=155,
            command=self._elegir_carpetas_correos,
        )
        self.btn_elegir_correos.pack(side="left", padx=4, pady=4)

        self.lbl_estado_correos = ctk.CTkLabel(
            toolbar,
            text="Listo para despachar borradores",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.TEXT_PRIMARY,
        )
        self.lbl_estado_correos.pack(side="left", padx=(8, 0))

        self.progress_correos = ctk.CTkProgressBar(
            toolbar, width=160, height=8, corner_radius=4,
            fg_color=Palette.BG_INPUT, progress_color=Palette.ACCENT,
        )
        self.progress_correos.pack(side="right", padx=16)
        self.progress_correos.set(0)

        self.btn_limpiar_correos = ctk.CTkButton(
            toolbar,
            text="Limpiar",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color="transparent",
            hover_color=Palette.BG_HOVER,
            text_color=Palette.TEXT_PRIMARY,
            corner_radius=4, height=30, width=70,
            command=self._limpiar_correos,
        )
        self.btn_limpiar_correos.pack(side="right", padx=4)

        # ── Vista previa de correos ──────────────────────────────────
        preview_frame = ctk.CTkFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER
        )
        preview_frame.pack(fill="both", expand=True)

        # Treeview de correos
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Correos.Treeview",
            background=Palette.BG_TABLE,
            foreground=Palette.TEXT_PRIMARY,
            fieldbackground=Palette.BG_TABLE,
            borderwidth=0,
            font=(FONT_FAMILY, 10),
            rowheight=28,
        )
        style.map(
            "Correos.Treeview",
            background=[("selected", Palette.ACCENT_DIM)],
            foreground=[("selected", Palette.WHITE)],
        )
        style.configure(
            "Correos.Treeview.Heading",
            background=Palette.BG_SIDEBAR,
            foreground=Palette.TEXT_SECONDARY,
            font=(FONT_FAMILY, 10, "bold"),
            borderwidth=0,
            padding=(8, 6),
        )

        self.tree_correos = ttk.Treeview(
            preview_frame,
            columns=("tipo", "asunto", "destinatarios", "adjuntos"),
            show="headings",
            style="Correos.Treeview",
            selectmode="none",
        )
        self.tree_correos.heading("tipo", text="Tipo")
        self.tree_correos.heading("asunto", text="Asunto")
        self.tree_correos.heading("destinatarios", text="Destinatarios")
        self.tree_correos.heading("adjuntos", text="Adjuntos")
        self.tree_correos.column("tipo", width=60, minwidth=50, stretch=True, anchor="center")
        self.tree_correos.column("asunto", width=200, minwidth=100, stretch=True, anchor="w")
        self.tree_correos.column("destinatarios", width=100, minwidth=70, stretch=True, anchor="center")
        self.tree_correos.column("adjuntos", width=60, minwidth=50, stretch=True, anchor="center")

        # Scrollbar horizontal (arriba)
        scroll_correos_x = ctk.CTkScrollbar(
            preview_frame, orientation="horizontal",
            fg_color=Palette.BG_CARD, button_color=Palette.TEXT_MUTED,
            button_hover_color=Palette.TEXT_SECONDARY,
        )
        self.tree_correos.configure(xscrollcommand=scroll_correos_x.set)
        scroll_correos_x.configure(command=self.tree_correos.xview)
        scroll_correos_x.pack(fill="x", padx=2, pady=(2, 0))

        # Scrollbar vertical
        scroll_y2 = ctk.CTkScrollbar(
            preview_frame, orientation="vertical",
            fg_color=Palette.BG_CARD, button_color=Palette.TEXT_MUTED,
            button_hover_color=Palette.TEXT_SECONDARY,
        )
        scroll_y2.pack(side="right", fill="y", padx=(0, 2), pady=2)
        self.tree_correos.configure(yscrollcommand=scroll_y2.set)
        scroll_y2.configure(command=self.tree_correos.yview)
        self.tree_correos.pack(fill="both", expand=True, padx=2, pady=(0, 0))

        # Placeholder
        self.tree_correos.insert(
            "", "end",
            values=("—", "Sin correos procesados", "—", "—")
        )

    # ═══════════════════════════════════════════════════════════════════
    # PANEL: DESCARGAR MAILS (placeholder)
    # ═══════════════════════════════════════════════════════════════════
    def _panel_descargar(self):
        if "descargar" in self._panel_frames:
            return
        frame = ctk.CTkFrame(self.panel_container, fg_color="transparent")
        self._panel_frames["descargar"] = frame
        # frame.pack(fill="both", expand=True) — packed by _cambiar_panel_forzado

        # ── Toolbar 1: botones de búsqueda ────────────────────────────
        toolbar1 = ctk.CTkFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER, height=44
        )
        toolbar1.pack(fill="x", pady=(0, 4))
        toolbar1.pack_propagate(False)

        # Botón 1: Buscar y Descargar (modo automático, comportamiento original)
        self._mail_btn_buscar = ctk.CTkButton(
            toolbar1,
            text="📥  Buscar y Descargar",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=30, width=155,
            command=self._mail_ejecutar,
        )
        self._mail_btn_buscar.pack(side="left", padx=3, pady=3)

        ctk.CTkLabel(
            toolbar1, text="Últimos:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_SECONDARY,
        ).pack(side="left", padx=(1, 0))

        self._mail_entry_cantidad = ctk.CTkEntry(
            toolbar1, width=35, height=26,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=4,
        )
        self._mail_entry_cantidad.insert(0, self.config.get("descarga_mails", {}).get("papeles", "2"))
        self._mail_entry_cantidad.pack(side="left")

        ctk.CTkLabel(
            toolbar1, text="mails",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_MUTED,
        ).pack(side="left", padx=(1, 6))

        # Separador
        ctk.CTkLabel(
            toolbar1, text="│",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.BORDER,
        ).pack(side="left", padx=2)

        # Botón 2: Buscar con reglas (modo interactivo, el usuario elige qué descargar)
        self._mail_btn_buscar_reglas = ctk.CTkButton(
            toolbar1,
            text="🔍  Buscar",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.SECONDARY, hover_color=Palette.SECONDARY_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=30, width=90,
            command=self._mail_ejecutar_buscar,
        )
        self._mail_btn_buscar_reglas.pack(side="left", padx=3, pady=3)

        ctk.CTkLabel(
            toolbar1, text="Últimos:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_SECONDARY,
        ).pack(side="left", padx=(1, 0))

        self._mail_entry_cantidad_reglas = ctk.CTkEntry(
            toolbar1, width=35, height=26,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=4,
        )
        self._mail_entry_cantidad_reglas.insert(0, self.config.get("descarga_mails", {}).get("reglas", "4"))
        self._mail_entry_cantidad_reglas.pack(side="left")

        ctk.CTkLabel(
            toolbar1, text="mails",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_MUTED,
        ).pack(side="left", padx=(1, 6))

        # Separador
        ctk.CTkLabel(
            toolbar1, text="│",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.BORDER,
        ).pack(side="left", padx=2)

        # Botón 3: Mail sin filtros (últimos N sin filtrar)
        self._mail_btn_ultimos = ctk.CTkButton(
            toolbar1,
            text="📋  Mail sin filtros",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.SECONDARY, hover_color=Palette.SECONDARY_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=30, width=135,
            command=self._mail_ejecutar_ultimos,
        )
        self._mail_btn_ultimos.pack(side="left", padx=3, pady=3)

        self._mail_entry_sin_filtro = ctk.CTkEntry(
            toolbar1, width=40, height=28,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=4,
        )
        self._mail_entry_sin_filtro.insert(0, self.config.get("descarga_mails", {}).get("sin_filtro", "20"))
        self._mail_entry_sin_filtro.pack(side="left", padx=(4, 0))

        ctk.CTkLabel(
            toolbar1, text="mails",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_MUTED,
        ).pack(side="left", padx=(1, 4))

        # ── Toolbar 2: descarga de seleccionados + progreso ────────────
        toolbar2 = ctk.CTkFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER, height=36
        )
        toolbar2.pack(fill="x", pady=(0, 6))
        toolbar2.pack_propagate(False)

        self._mail_btn_descargar_sel = ctk.CTkButton(
            toolbar2,
            text="⬇  Descargar seleccionados",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=30, width=180,
            command=self._mail_descargar_seleccionados,
        )
        self._mail_btn_descargar_sel.pack(side="left", padx=8, pady=4)

        self._mail_progress = ctk.CTkProgressBar(
            toolbar2, width=120, height=8, corner_radius=4,
            fg_color=Palette.BG_INPUT, progress_color=Palette.ACCENT,
        )
        self._mail_progress.pack(side="right", padx=10)
        self._mail_progress.set(0)

        self._mail_lbl_estado = ctk.CTkLabel(
            toolbar2,
            text="Listo — Buscar mails de 'papeles' o Últimos 20",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_MUTED,
        )
        self._mail_lbl_estado.pack(side="right", padx=6)

        self.btn_limpiar_mail = ctk.CTkButton(
            toolbar2,
            text="Limpiar",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color="transparent",
            hover_color=Palette.BG_HOVER,
            text_color=Palette.TEXT_PRIMARY,
            corner_radius=4, height=30, width=70,
            command=self._limpiar_mail,
        )
        self.btn_limpiar_mail.pack(side="right", padx=4)

        # ── Tabla de mails ────────────────────────────────────────────
        tabla_frame = ctk.CTkFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER
        )
        tabla_frame.pack(fill="both", expand=True)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Mail.Treeview",
            background=Palette.BG_TABLE,
            foreground=Palette.TEXT_PRIMARY,
            fieldbackground=Palette.BG_TABLE,
            borderwidth=0, font=(FONT_FAMILY, 10), rowheight=28,
        )
        style.map(
            "Mail.Treeview",
            background=[("selected", Palette.ACCENT_DIM)],
            foreground=[("selected", Palette.WHITE)],
        )
        style.configure(
            "Mail.Treeview.Heading",
            background=Palette.BG_SIDEBAR,
            foreground=Palette.TEXT_SECONDARY,
            font=(FONT_FAMILY, 10, "bold"), borderwidth=0, padding=(8, 6),
        )

        self._mail_tree = ttk.Treeview(
            tabla_frame,
            columns=("sel", "asunto", "fecha", "adjuntos", "carpeta"),
            show="headings", style="Mail.Treeview", selectmode="browse",
        )
        self._mail_tree.heading("sel", text="✓")
        self._mail_tree.heading("asunto", text="Asunto")
        self._mail_tree.heading("fecha", text="Fecha")
        self._mail_tree.heading("adjuntos", text="Adjuntos")
        self._mail_tree.heading("carpeta", text="Carpeta destino")
        self._mail_tree.column("sel", width=35, anchor="center", stretch=False)
        self._mail_tree.column("asunto", width=280, anchor="w")
        self._mail_tree.column("fecha", width=130, anchor="center")
        self._mail_tree.column("adjuntos", width=70, anchor="center")
        self._mail_tree.column("carpeta", width=260, anchor="w")

        # Bind para toggle del checkbox al clickear en la columna "✓"
        self._mail_tree.bind("<ButtonRelease-1>", self._mail_toggle_check)

        scroll = ctk.CTkScrollbar(
            tabla_frame, orientation="vertical",
            fg_color=Palette.BG_CARD, button_color=Palette.TEXT_MUTED,
            button_hover_color=Palette.TEXT_SECONDARY,
        )
        scroll.pack(side="right", fill="y", padx=(0, 2), pady=2)
        self._mail_tree.configure(yscrollcommand=scroll.set)
        scroll.configure(command=self._mail_tree.yview)
        self._mail_tree.pack(fill="both", expand=True, padx=2, pady=2)

    # ═══════════════════════════════════════════════════════════════════
    # DESCARGAR MAILS: EJECUCIÓN
    # ═══════════════════════════════════════════════════════════════════

    def _mail_ejecutar(self):
        """Modo 0: Buscar y Descargar — busca mails con reglas y descarga los N más nuevos."""
        if self.tarea_activa:
            messagebox.showwarning("Agente ocupado", "Hay una tarea en ejecución.")
            return
        self.tarea_activa = True
        self._mail_btn_buscar.configure(text="⏳  Buscando...", state="disabled")
        self._mail_btn_buscar_reglas.configure(state="disabled")
        self._mail_btn_ultimos.configure(state="disabled")
        self._mail_btn_descargar_sel.configure(state="disabled")
        self._mail_progress.configure(mode="indeterminate")
        self._mail_progress.start()
        self._mail_lbl_estado.configure(text="Conectando al servidor IMAP...")
        self._limpiar_log()
        for row in self._mail_tree.get_children():
            self._mail_tree.delete(row)
        self._mail_data.clear()
        t = threading.Thread(target=self._mail_worker, daemon=True)
        t.start()

    def _mail_ejecutar_buscar(self):
        """Modo 1: Buscar con reglas — muestra resultados con checkboxes, el usuario elige."""
        if self.tarea_activa:
            messagebox.showwarning("Agente ocupado", "Hay una tarea en ejecución.")
            return
        try:
            cantidad = int(self._mail_entry_cantidad_reglas.get().strip())
        except ValueError:
            cantidad = 10
        self.tarea_activa = True
        self._mail_btn_buscar.configure(state="disabled")
        self._mail_btn_buscar_reglas.configure(text="⏳  Buscando...", state="disabled")
        self._mail_btn_ultimos.configure(state="disabled")
        self._mail_btn_descargar_sel.configure(state="disabled")
        self._mail_progress.configure(mode="indeterminate")
        self._mail_progress.start()
        self._mail_lbl_estado.configure(text="Buscando mails con reglas...")
        self._limpiar_log()
        for row in self._mail_tree.get_children():
            self._mail_tree.delete(row)
        self._mail_data.clear()
        t = threading.Thread(target=self._mail_buscar_worker, args=(cantidad, 1), daemon=True)
        t.start()

    def _mail_ejecutar_ultimos(self):
        """Modo 2: Últimos N sin filtro — busca los últimos N mails y muestra todos."""
        if self.tarea_activa:
            messagebox.showwarning("Agente ocupado", "Hay una tarea en ejecución.")
            return
        try:
            cantidad = int(self._mail_entry_sin_filtro.get().strip())
        except ValueError:
            cantidad = 20
        self.tarea_activa = True
        self._mail_btn_buscar.configure(state="disabled")
        self._mail_btn_buscar_reglas.configure(state="disabled")
        self._mail_btn_ultimos.configure(text="⏳  Buscando...", state="disabled")
        self._mail_btn_descargar_sel.configure(state="disabled")
        self._mail_progress.configure(mode="indeterminate")
        self._mail_progress.start()
        self._mail_lbl_estado.configure(text=f"Buscando últimos {cantidad} mails (sin filtro)...")
        self._limpiar_log()
        for row in self._mail_tree.get_children():
            self._mail_tree.delete(row)
        self._mail_data.clear()
        t = threading.Thread(target=self._mail_buscar_worker, args=(cantidad, 2), daemon=True)
        t.start()

    def _imap_conectar(self):
        """Crea y retorna una conexión IMAP con reintentos."""
        import imaplib
        srv = self._cfg_obtener_correo("imap_server", IMAP_SERVER)
        prt = self._cfg_obtener_correo("imap_puerto", PUERTO_IMAP)
        usr = self._cfg_obtener_correo("usuario", "")
        pwd = self._cfg_obtener_correo("password", "")
        for intento in range(3):
            try:
                mail = imaplib.IMAP4(srv, prt)
                mail.login(usr, pwd)
                mail.select("INBOX")
                return mail
            except Exception as e:
                if intento < 2:
                    self._log(f"  ⚠ Reintentando conexión... ({intento+2}/3)")
                    time.sleep(1.5)
                else:
                    raise e

    def _imap_reintentar(self, mail, operacion, *args, **kwargs):
        """Ejecuta una operación IMAP con reintento en caso de socket error."""
        import socket
        for intento in range(3):
            try:
                return operacion(*args, **kwargs)
            except (socket.error, OSError, ConnectionError, Exception) as e:
                msg = str(e).lower()
                if "eof" in msg or "socket" in msg or "connection" in msg or "timeout" in msg:
                    if intento < 2:
                        self._log(f"  ⚠ Error de conexión, reconectando... ({intento+2}/3)")
                        time.sleep(1)
                        try:
                            mail.logout()
                        except Exception:
                            pass
                        try:
                            mail = self._imap_conectar()
                        except Exception:
                            pass
                    else:
                        raise Exception(f"Error IMAP tras 3 intentos: {e}")
                else:
                    raise e

    def _mail_buscar_headers(self, mail, todos_ids, solo_papeles=True):
        """Helper: busca headers y devuelve [(fecha_dt, fecha_str, mid, asunto), ...].
        Usa un solo FETCH múltiple para evitar N viajes IMAP individuales."""
        from email.utils import parsedate_to_datetime
        resultados = []
        if not todos_ids:
            return resultados
        # Convertir IDs a string para fetch múltiple
        ids_comma = b','.join(mid if isinstance(mid, bytes) else str(mid).encode() for mid in todos_ids)
        status, data = mail.fetch(ids_comma, "(BODY.PEEK[HEADER.FIELDS (SUBJECT FROM DATE)])")
        if status != "OK":
            self._log(f"  ⚠ Error en fetch múltiple de headers, reintentando individual...")
            # Fallback: fetch individual por si el servidor no soporta fetch múltiple
            for mid in todos_ids:
                status2, data2 = mail.fetch(mid, "(BODY.PEEK[HEADER.FIELDS (SUBJECT FROM DATE)])")
                if status2 != "OK":
                    continue
                header = data2[0][1].decode("utf-8", errors="replace") if data2[0][1] else ""
                self._procesar_header(header, mid, solo_papeles, resultados, parsedate_to_datetime)
            return resultados
        for item in data:
            if item is None or len(item) < 2:
                continue
            header_bytes = item[1]
            if not header_bytes:
                continue
            header = header_bytes.decode("utf-8", errors="replace")
            # Extraer MID del primer campo (ej: b'1 (BODY[HEADER...')
            first_part = item[0]
            if isinstance(first_part, bytes):
                mid = first_part.split(b' ')[0]
            else:
                mid = first_part.split(' ')[0].encode()
            self._procesar_header(header, mid, solo_papeles, resultados, parsedate_to_datetime)
        return resultados

    def _procesar_header(self, header, mid, solo_papeles, resultados, parsedate_to_datetime):
        """Extrae datos de un header IMAP y los agrega a resultados si cumple filtros."""
        match_subj = re.search(r"Subject:\s*(.+)", header, re.IGNORECASE)
        match_from = re.search(r"From:\s*(.+)", header, re.IGNORECASE)
        if not match_subj:
            return
        asunto = match_subj.group(1).strip()
        remitente = match_from.group(1).strip() if match_from else ""
        if solo_papeles:
            if not asunto.lower().startswith("papeles"):
                return
            if "correo" not in remitente.lower():
                return
        match_date = re.search(r"Date:\s*(.+)", header, re.IGNORECASE)
        fecha_str = match_date.group(1).strip() if match_date else ""
        try:
            fecha_dt = parsedate_to_datetime(fecha_str)
        except Exception:
            fecha_dt = datetime.min
        resultados.append((fecha_dt, fecha_str, mid, asunto))

    def _mail_worker(self):
        """Modo 0: Descarga automática de los N mails más nuevos que cumplen las reglas."""
        self._set_log_panel("descargar")
        resultados = []
        try:
            try:
                cantidad = int(self._mail_entry_cantidad.get().strip())
            except ValueError:
                cantidad = 2
            self._log(f"Conectando a {self._cfg_obtener_correo('imap_server', IMAP_SERVER)}...")
            try:
                mail = self._imap_conectar()
            except Exception as e:
                self._log(f"ERROR: No se pudo conectar al servidor IMAP: {e}")
                return
            self._log(f"Conectado. Buscando los {cantidad} mails más nuevos de 'papeles'...")
            status, ids = mail.search(None, 'ALL')
            if status != "OK":
                self._log("Error en búsqueda IMAP")
                return
            todos_ids = ids[0].split()
            todos_ids = todos_ids[-20:]
            self._log("Buscando entre los últimos 20 mails...")
            mails_papeles = self._mail_buscar_headers(mail, todos_ids, solo_papeles=True)
            mails_papeles.sort(key=lambda x: x[0], reverse=True)
            seleccionados = mails_papeles[:cantidad]
            self._log(f"{len(mails_papeles)} mails 'papeles' encontrados. Descargando los {len(seleccionados)} más nuevos.")
            encontrados = 0
            escritorio = self._resolver_ruta("descarga_mails", "Desktop")
            for fecha_dt, fecha_str, mid, asunto in seleccionados:
                encontrados += 1
                self._log(f"  ✓ {asunto[:80]}")
                status, data = mail.fetch(mid, "(BODY[])")
                if status != "OK":
                    continue
                raw_email = data[0][1]
                import email as em_mod
                try:
                    msg = em_mod.message_from_bytes(raw_email, policy=em_mod.policy.default)
                except Exception:
                    try:
                        msg = em_mod.message_from_bytes(raw_email)
                    except Exception:
                        continue
                carpeta_temp = os.path.join(escritorio, f"_tmp_{mid.decode()}")
                os.makedirs(carpeta_temp, exist_ok=True)
                adjuntos = 0
                adj_nombres = []
                ruta_contenedores = None
                tiene_comparte = "COMPARTE" in asunto.upper() if asunto else False
                for part in msg.walk():
                    if part.get_content_disposition() == "attachment":
                        filename = part.get_filename()
                        if filename:
                            filepath = os.path.join(carpeta_temp, filename)
                            if os.path.exists(filepath):
                                base, ext = os.path.splitext(filename)
                                n = 2
                                while os.path.exists(filepath):
                                    filepath = os.path.join(carpeta_temp, f"{base}_{n}{ext}")
                                    n += 1
                                adj_nombres.append(os.path.basename(filepath))
                            else:
                                adj_nombres.append(filename)
                            with open(filepath, "wb") as f:
                                f.write(part.get_payload(decode=True))
                            adjuntos += 1
                            self._log(f"     → {os.path.basename(filepath)}")
                            if "CONTENEDORES" in filename.upper() and not ruta_contenedores:
                                ruta_contenedores = filepath
                            if "COMPARTE" in filename.upper():
                                tiene_comparte = True

                # Detectar si es mail "comparte" (2+ CONTENEDORES o 6 adjuntos o "comparte" en asunto/nombre)
                n_contenedores = sum(1 for n in adj_nombres if "CONTENEDORES" in n.upper())
                if tiene_comparte or n_contenedores >= 2 or adjuntos >= 6:
                    self._log(f"     🔀 Detectado mail COMPARTE ({n_contenedores} Contenedores, {adjuntos} adjuntos)")
                    carpetas_creadas = self._mail_procesar_comparte(carpeta_temp, adj_nombres, escritorio)
                    for nombre_carpeta, ruta_final in carpetas_creadas:
                        resultados.append((asunto, adj_nombres, nombre_carpeta))
                        self.after(0, lambda a=asunto, f=fecha_str, adj=adjuntos, c=os.path.basename(ruta_final):
                            self._mail_tree.insert("", "end", values=("✓", a[:80], f[:25], str(adj), c)))
                else:
                    nombre_carpeta = self._mail_nombre_carpeta(ruta_contenedores, carpeta_temp)
                    ruta_final = os.path.join(escritorio, nombre_carpeta)
                    if os.path.exists(ruta_final):
                        i = 2
                        while os.path.exists(ruta_final):
                            ruta_final = os.path.join(escritorio, f"{nombre_carpeta}_{i}")
                            i += 1
                    os.rename(carpeta_temp, ruta_final)
                    resultados.append((asunto, adj_nombres, nombre_carpeta))
                    self.after(0, lambda a=asunto, f=fecha_str, adj=adjuntos, c=os.path.basename(ruta_final):
                        self._mail_tree.insert("", "end", values=("✓", a[:80], f[:25], str(adj), c)))
            mail.logout()
            self._log(f"COMPLETADO: {encontrados} mails procesados. Archivos guardados en el Escritorio.")
        except Exception as e:
            self._log(f"ERROR: {e}")
            resultados = []
        finally:
            self.after(0, lambda: self._mail_done(resultados, modo=0))

    def _mail_buscar_worker(self, cantidad, modo):
        """Modo 1/2: Busca mails y los muestra en el tree con checkboxes (sin descargar).
           modo=1: solo 'papeles' de 'correo'
           modo=2: últimos mails sin filtro
        """
        self._set_log_panel("descargar")
        try:
            self._log(f"Conectando a {self._cfg_obtener_correo('imap_server', IMAP_SERVER)}...")
            try:
                mail = self._imap_conectar()
            except Exception as e:
                self._log(f"ERROR: No se pudo conectar al servidor IMAP: {e}")
                return
            status, ids = mail.search(None, 'ALL')
            if status != "OK":
                self._log("Error en búsqueda IMAP")
                return
            todos_ids = ids[0].split()
            if modo == 1:
                # Buscar entre los últimos 300 mails para encontrar los N con 'papeles'
                VENTANA = min(100, len(todos_ids))
                todos_ids = todos_ids[-VENTANA:]
                solo_papeles = True
                label = f"hasta {cantidad} con 'papeles'"
            else:
                todos_ids = todos_ids[-cantidad:]
                solo_papeles = False
                label = f"últimos {cantidad} sin filtro"
            self._log(f"Buscando {label} ({len(todos_ids)} resultados)...")
            mails_encontrados = self._mail_buscar_headers(mail, todos_ids, solo_papeles=solo_papeles)
            mails_encontrados.sort(key=lambda x: x[0], reverse=True)
            seleccionados = mails_encontrados[:cantidad]
            mail.logout()
            self._log(f"{len(seleccionados)} mails encontrados. Marcá los que quieras descargar y presioná 'Descargar seleccionados'.")

            # Poblar el tree en una sola pasada por el UI thread
            def poblar_tree():
                for fecha_dt, fecha_str, mid, asunto in seleccionados:
                    item_id = self._mail_tree.insert(
                        "", "end",
                        values=("☐", asunto[:80], fecha_str[:25], "—", "—"),
                    )
                    self._mail_data[item_id] = {
                        "mid": mid, "asunto": asunto,
                        "fecha_dt": fecha_dt, "fecha_str": fecha_str,
                        "checked": False, "downloaded": False,
                    }
            self.after(0, poblar_tree)
        except Exception as e:
            self._log(f"ERROR: {e}")
        finally:
            self.after(0, lambda: self._mail_done([], modo=modo))

    def _mail_toggle_check(self, event):
        region = self._mail_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self._mail_tree.identify_column(event.x)
        if column != "#1":
            return
        item = self._mail_tree.identify_row(event.y)
        if not item or item not in self._mail_data:
            return
        if self._mail_data[item].get("downloaded"):
            return
        current = self._mail_tree.set(item, "sel")
        new_val = "☑" if current.strip() == "☐" else "☐"
        self._mail_tree.set(item, "sel", new_val)
        self._mail_data[item]["checked"] = (new_val == "☑")

    def _mail_descargar_seleccionados(self):
        if self.tarea_activa:
            messagebox.showwarning("Agente ocupado", "Hay una tarea en ejecución.")
            return
        items = [i for i, d in self._mail_data.items() if d.get("checked") and not d.get("downloaded")]
        if not items:
            messagebox.showwarning("Sin selección", "Marcá al menos un mail (☑) para descargar.\n\nClick en la columna ✓ para marcar/desmarcar.")
            return
        self.tarea_activa = True
        self._mail_btn_buscar.configure(state="disabled")
        self._mail_btn_buscar_reglas.configure(state="disabled")
        self._mail_btn_ultimos.configure(state="disabled")
        self._mail_btn_descargar_sel.configure(text="⏳  Descargando...", state="disabled")
        self._mail_progress.configure(mode="indeterminate")
        self._mail_progress.start()
        self._mail_lbl_estado.configure(text=f"Descargando {len(items)} mail(s) seleccionados...")
        self._limpiar_log()
        t = threading.Thread(target=self._mail_descargar_worker, args=(items,), daemon=True)
        t.start()

    def _mail_descargar_worker(self, items_a_descargar):
        self._set_log_panel("descargar")
        resultados = []
        try:
            self._log(f"Conectando a {self._cfg_obtener_correo('imap_server', IMAP_SERVER)}...")
            try:
                mail = self._imap_conectar()
            except Exception as e:
                self._log(f"ERROR: No se pudo conectar al servidor IMAP: {e}")
                return
            escritorio = self._resolver_ruta("descarga_mails", "Desktop")
            import email as em_mod
            for item in items_a_descargar:
                data = self._mail_data[item]
                mid = data["mid"]
                asunto = data["asunto"]
                self._log(f"  Descargando: {asunto[:80]}")
                status, resp = mail.fetch(mid, "(BODY[])")
                if status != "OK":
                    self._log(f"     ⚠ Error al obtener el mail")
                    continue
                raw_email = resp[0][1]
                try:
                    msg = em_mod.message_from_bytes(raw_email, policy=em_mod.policy.default)
                except Exception:
                    try:
                        msg = em_mod.message_from_bytes(raw_email)
                    except Exception:
                        self._log(f"     ⚠ Error al parsear el mail")
                        continue
                carpeta_temp = os.path.join(escritorio, f"_tmp_{mid.decode()}")
                os.makedirs(carpeta_temp, exist_ok=True)
                adjuntos = 0
                adj_nombres = []
                ruta_contenedores = None
                tiene_comparte = "COMPARTE" in asunto.upper() if asunto else False
                for part in msg.walk():
                    if part.get_content_disposition() == "attachment":
                        filename = part.get_filename()
                        if filename:
                            filepath = os.path.join(carpeta_temp, filename)
                            if os.path.exists(filepath):
                                base, ext = os.path.splitext(filename)
                                n = 2
                                while os.path.exists(filepath):
                                    filepath = os.path.join(carpeta_temp, f"{base}_{n}{ext}")
                                    n += 1
                                adj_nombres.append(os.path.basename(filepath))
                            else:
                                adj_nombres.append(filename)
                            with open(filepath, "wb") as f:
                                f.write(part.get_payload(decode=True))
                            adjuntos += 1
                            self._log(f"     → {os.path.basename(filepath)}")
                            if "CONTENEDORES" in filename.upper() and not ruta_contenedores:
                                ruta_contenedores = filepath
                            if "COMPARTE" in filename.upper():
                                tiene_comparte = True

                n_contenedores = sum(1 for n in adj_nombres if "CONTENEDORES" in n.upper())
                if tiene_comparte or n_contenedores >= 2 or adjuntos >= 6:
                    self._log(f"     🔀 Detectado mail COMPARTE ({n_contenedores} Contenedores, {adjuntos} adjuntos)")
                    carpetas_creadas = self._mail_procesar_comparte(carpeta_temp, adj_nombres, escritorio)
                    for nombre_carpeta, ruta_final in carpetas_creadas:
                        resultados.append((asunto, adj_nombres, nombre_carpeta))
                        self._mail_data[item]["downloaded"] = True
                        self._mail_data[item]["checked"] = False
                        c_final = os.path.basename(ruta_final)
                        self.after(0, lambda i=item, a=asunto, adj=adjuntos, c=c_final:
                            self._mail_tree.set(i, "sel", "✓") or
                            self._mail_tree.set(i, "adjuntos", str(adj)) or
                            self._mail_tree.set(i, "carpeta", c))
                else:
                    nombre_carpeta = self._mail_nombre_carpeta(ruta_contenedores, carpeta_temp)
                    ruta_final = os.path.join(escritorio, nombre_carpeta)
                    if os.path.exists(ruta_final):
                        i = 2
                        while os.path.exists(ruta_final):
                            ruta_final = os.path.join(escritorio, f"{nombre_carpeta}_{i}")
                            i += 1
                    os.rename(carpeta_temp, ruta_final)
                    resultados.append((asunto, adj_nombres, nombre_carpeta))
                    self._mail_data[item]["downloaded"] = True
                    self._mail_data[item]["checked"] = False
                    c_final = os.path.basename(ruta_final)
                    self.after(0, lambda i=item, a=asunto, adj=adjuntos, c=c_final:
                        self._mail_tree.set(i, "sel", "✓") or
                        self._mail_tree.set(i, "adjuntos", str(adj)) or
                        self._mail_tree.set(i, "carpeta", c))
            mail.logout()
            self._log(f"COMPLETADO: {len(resultados)} mail(s) descargados al Escritorio.")
        except Exception as e:
            self._log(f"ERROR: {e}")
            resultados = []
        finally:
            self.after(0, lambda: self._mail_done(resultados, modo=3))

    def _mail_done(self, resultados, modo=0):
        self.tarea_activa = False
        try:
            self._mail_btn_buscar.configure(text="📥  Buscar y Descargar", state="normal")
            self._mail_btn_buscar_reglas.configure(text="🔍  Buscar", state="normal")
            self._mail_btn_ultimos.configure(text="📋  Mail sin filtros", state="normal")
            self._mail_btn_descargar_sel.configure(text="⬇  Descargar seleccionados", state="normal")
            self._mail_progress.stop()
            self._mail_progress.set(1)
        except (AttributeError, Exception):
            pass
        try:
            if modo == 0:
                self._mail_lbl_estado.configure(text="Descarga completada")
            elif modo in (1, 2):
                pendientes = sum(1 for d in self._mail_data.values() if d.get("checked") and not d.get("downloaded"))
                self._mail_lbl_estado.configure(
                    text=f"Listo — {pendientes} mail(s) marcados para descargar"
                )
            elif modo == 3:
                self._mail_lbl_estado.configure(text="Descarga de seleccionados completada")
        except (AttributeError, Exception):
            pass
        if resultados:
            if not self._super_auto:
                self._mail_popup_resumen(resultados)

        # Súper Auto: si está activo y hay resultados, disparar cadena
        if self._super_auto and resultados:
            self._log("⚡ Súper Auto: iniciando cadena automática...")
            self.after(500, lambda: self._super_ejecutar_cadena(resultados))

    def _clasificar_tipo_transporte(self, puerto_salida, peso_flexi):
        """Classify transport type from CONTENEDORES fields.

        Returns "TERRESTRE" | "ISO" | "FLEXI".
        - None / "" / "-" (after .strip()) on puerto_salida → TERRESTRE
        - set + peso in {None, "", 0, 0.0, "0"}                → ISO
        - set + peso > 0                                      → FLEXI
        """
        # Normalizar puerto
        if puerto_salida is None:
            return "TERRESTRE"
        puerto = str(puerto_salida).strip()
        if not puerto or puerto == "-":
            return "TERRESTRE"

        # Puerto está set → marítimo; decidir ISO vs FLEXI por peso
        if peso_flexi is None:
            return "ISO"
        peso_str = str(peso_flexi).strip()
        if not peso_str or peso_str == "0" or peso_str == "0.0":
            return "ISO"
        try:
            if float(peso_str) <= 0:
                return "ISO"
        except (ValueError, TypeError):
            # No se pudo parsear → tratar como ISO (conservador)
            return "ISO"
        return "FLEXI"

    def _mail_nombre_carpeta(self, ruta_contenedores, carpeta_temp):
        """Extrae datos del Excel CONTENEDORES y arma el nombre de carpeta."""
        if not ruta_contenedores:
            # Fallback: usar nombre genérico
            return os.path.basename(carpeta_temp).replace("_tmp_", "Papeles_")

        try:
            # Usar las mismas funciones de extracción del agente planillas
            if ruta_contenedores.lower().endswith(".xlsx"):
                datos = self._leer_xlsx_moderno(ruta_contenedores)
            else:
                datos = self._leer_xls_antiguo(ruta_contenedores)
            f_celda, pe_celda, carp_celda, dest_celda, bl_celda, trans_final, patente, precintos, guarda, *_ = datos
            puerto_salida = _[0] if _ else ""
            peso_flexi = _[1] if len(_) > 1 else ""
        except Exception as e:
            self._log(f"     ⚠ No se pudo leer el Excel: {e}")
            return os.path.basename(carpeta_temp).replace("_tmp_", "Papeles_")

        # Abreviar destinatario
        def abreviar(nombre, tipo):
            n = (nombre or "").upper()
            if "VITAPRO" in n: return "VITAPRO"
            if "EWOS" in n: return "EWOS"
            if "NUTRECO" in n: return "NUTRECO"
            if "DICOAL" in n: return "DICOAL"
            if "CARGILL" in n: return "CARGILL"
            if "BIOMAR" in n: return "BIOMAR"
            # Marítimo: nombre completo; terrestre: truncar como antes
            if tipo == "TERRESTRE":
                return (nombre or "")[:15]
            return nombre

        # Obtener P.E. completo
        pe_origen = pe_celda if pe_celda else ""
        pe_recortado = pe_origen[-5:].lstrip("0") if pe_origen else ""

        # Cantidad de contenedores desde el número al inicio del archivo
        nombre_archivo = os.path.basename(ruta_contenedores)
        match_cant = re.match(r"^(\d+)", nombre_archivo.strip())
        cant = match_cant.group(1) if match_cant else "1"

        # Fracción: último número en el nombre del archivo CONTENEDORES
        frac = ""
        match_frac = re.search(r"(?:Fracci[oó]n\s*)?(\d+)\s*\.(?:xls|xlsx)$", nombre_archivo, re.IGNORECASE)
        if match_frac:
            frac = f"F{match_frac.group(1)}"

        # Armar nombre: 16_05_2026_1_TERRESTRE_398W_560093_NUTRECO_F1
        # (o marítimo: 16_05_2026_1_ISO_398W_560093_NUTRECO_TRP)
        fecha_clean = f_celda.replace("/", "_") if f_celda else "sin_fecha"
        # Asegurar formato DD_MM_YYYY con leading zeros
        m = re.match(r"(\d{1,2})_(\d{1,2})_(\d{4})", fecha_clean)
        if m:
            fecha_clean = f"{int(m.group(1)):02d}_{int(m.group(2)):02d}_{m.group(3)}"

        # Clasificar tipo de transporte a partir de Puerto Salida / Peso Flexi
        tipo = self._clasificar_tipo_transporte(puerto_salida, peso_flexi)

        partes = [
            fecha_clean,
            cant,
            tipo,
            pe_recortado or "PE",
            str(carp_celda)[:10] if carp_celda else "0",
            abreviar(dest_celda or "", tipo),
        ]
        # Sufijo: terrestre → fracción; marítimo (ISO/FLEXI) → puerto de salida
        if tipo == "TERRESTRE":
            if frac:
                partes.append(frac)
        else:
            puerto_sufijo = str(puerto_salida).strip() if puerto_salida else ""
            if puerto_sufijo:
                partes.append(puerto_sufijo)

        return "_".join(partes)

    def _mail_procesar_comparte(self, carpeta_temp, adj_nombres, escritorio):
        """Procesa un correo 'comparte' con 2 Contenedores, 2 hojas de ruta y 2 PDFs.
        Crea 2 carpetas: COMPARTIDO y COMPARTIDO_CERRA."""
        import shutil as sh

        contenedores = []
        hojas_ruta = []
        pdfs = []
        otros = []
        for fn in adj_nombres:
            up = fn.upper()
            if "CONTENEDORES" in up and (up.endswith(".XLSX") or up.endswith(".XLS")):
                contenedores.append(fn)
            elif up.startswith("HOJA") and (up.endswith(".XLSX") or up.endswith(".XLS")):
                hojas_ruta.append(fn)
            elif up.endswith(".PDF"):
                pdfs.append(fn)
            else:
                otros.append(fn)

        if len(contenedores) != 2:
            self._log(f"  ⚠ Se esperaban 2 archivos Contenedores, se encontraron {len(contenedores)}")
            return []

        def _get_frac_num(filename):
            m = re.search(r"(\d+)\s*\.(?:xls|xlsx)$", filename, re.IGNORECASE)
            if not m:
                m = re.search(r"Fracci[oó]n\s*(\d+)", filename, re.IGNORECASE)
            return int(m.group(1)) if m else 0

        if _get_frac_num(contenedores[0]) > _get_frac_num(contenedores[1]):
            cierra_fn = contenedores[0]
            comparte_fn = contenedores[1]
        else:
            cierra_fn = contenedores[1]
            comparte_fn = contenedores[0]

        def _extraer_datos(ruta):
            """Extrae todos los datos de un archivo Contenedores."""
            try:
                if ruta.lower().endswith(".xlsx"):
                    datos = self._leer_xlsx_moderno(ruta)
                else:
                    datos = self._leer_xls_antiguo(ruta)
                f_celda, pe_celda, carp_celda, dest_celda = datos[0], datos[1], datos[2], datos[3]
                puerto = datos[9] if len(datos) > 9 else ""
                peso = datos[10] if len(datos) > 10 else ""
                pe = (pe_celda or "").strip()[-5:].lstrip("0")
                return f_celda, pe, carp_celda, dest_celda, puerto, peso
            except Exception as e:
                self._log(f"  ⚠ Error leyendo {os.path.basename(ruta)}: {e}")
                return "", "", "", "", "", ""

        def _abreviar_dest(dest, tipo):
            d = (dest or "").upper()
            for abr in ("VITAPRO", "EWOS", "NUTRECO", "DICOAL", "CARGILL", "BIOMAR"):
                if abr in d: return abr
            if tipo == "TERRESTRE":
                return (dest or "")[:15]
            return dest

        def _formatear_fecha(f_celda):
            """13/05/2026 → 13_05_2026 (con leading zeros)"""
            if not f_celda: return "sin_fecha"
            s = str(f_celda).strip()
            m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
            if m:
                return f"{int(m.group(1)):02d}_{int(m.group(2)):02d}_{m.group(3)}"
            return s.replace("/", "_")

        ruta_comparte = os.path.join(carpeta_temp, comparte_fn)
        ruta_cierra   = os.path.join(carpeta_temp, cierra_fn)

        f_a, pe_a, carp_a, dest_a, puerto_a, peso_a = _extraer_datos(ruta_comparte)
        f_b, pe_b, carp_b, dest_b, puerto_b, peso_b = _extraer_datos(ruta_cierra)

        if not pe_a or not pe_b:
            self._log("  ⚠ No se pudo extraer PE de los Contenedores")
            return []

        # Fracción desde el nombre del archivo
        def _extraer_frac(filename):
            # Último número antes de la extensión (sin requerir "Fracción")
            m = re.search(r"(\d+)\s*\.(?:xls|xlsx)$", filename, re.IGNORECASE)
            if not m:
                # Buscar después de "Fracción"
                m = re.search(r"Fracci[oó]n\s*(\d+)", filename, re.IGNORECASE)
            return f"F{m.group(1)}" if m else ""

        frac_a = _extraer_frac(comparte_fn)
        frac_b = _extraer_frac(cierra_fn)

        fecha_str = _formatear_fecha(f_a or f_b)
        match_cant = re.match(r"^(\d+)", os.path.basename(comparte_fn).strip())
        cant = match_cant.group(1) if match_cant else "1"

        # Leer PE de cada Hoja de Ruta (fila 4, col G = 7)
        pe_hojas = {}
        for fn_hoja in hojas_ruta:
            ruta_hoja = os.path.join(carpeta_temp, fn_hoja)
            try:
                if ruta_hoja.lower().endswith(".xls") and not ruta_hoja.lower().endswith(".xlsx"):
                    import xlrd
                    wb = xlrd.open_workbook(ruta_hoja)
                    ws = wb.sheet_by_index(0)
                    val = ws.cell_value(3, 6) if ws.nrows > 3 else ""  # fila 4=índice 3, col G=índice 6
                else:
                    wb = self._abrir_excel_seguro(ruta_hoja)
                    ws = wb.active
                    val = ws.cell(row=4, column=7).value
                    wb.close()
                if val:
                    pe_hoja = str(val).strip()[-5:].lstrip("0")
                    pe_hojas[fn_hoja] = pe_hoja
                    self._log(f"  📋 Hoja Ruta: {fn_hoja[:50]}... → PE={pe_hoja}")
            except Exception as e:
                self._log(f"  ⚠ Error leyendo hoja de ruta {fn_hoja}: {e}")

        hoja_a = hoja_b = None
        for fn_hoja, pe_h in pe_hojas.items():
            # Matching flexible: exacto, contenido, o substring compartido
            if pe_h == pe_a or pe_a in pe_h or pe_h in pe_a:
                hoja_a = fn_hoja
            elif pe_h == pe_b or pe_b in pe_h or pe_h in pe_b:
                hoja_b = fn_hoja

        # Si faltan hojas sin asignar, repartir las no usadas
        usadas = {hoja_a, hoja_b}
        for fn_hoja in pe_hojas:
            if fn_hoja not in usadas:
                if not hoja_a:
                    hoja_a = fn_hoja
                    usadas.add(fn_hoja)
                elif not hoja_b:
                    hoja_b = fn_hoja
                    usadas.add(fn_hoja)
                    self._log(f"  📋 Hoja Ruta asignada a CERRADO (fallback): {fn_hoja[:50]}...")

        pdf_a = next((fn for fn in pdfs if pe_a in fn.upper()), None)
        pdf_b = next((fn for fn in pdfs if pe_b in fn.upper()), None)
        usadas_pdf = {pdf_a, pdf_b}
        for fn in pdfs:
            if fn not in usadas_pdf:
                if not pdf_a:
                    pdf_a = fn
                    usadas_pdf.add(fn)
                elif not pdf_b:
                    pdf_b = fn
                    usadas_pdf.add(fn)

        def _mover(nombre_carpeta, archivos_fn):
            ruta_final = os.path.join(escritorio, nombre_carpeta)
            if os.path.exists(ruta_final):
                i = 2
                while os.path.exists(ruta_final):
                    ruta_final = os.path.join(escritorio, f"{nombre_carpeta}_{i}")
                    i += 1
            os.makedirs(ruta_final, exist_ok=True)
            for fn in archivos_fn:
                src = os.path.join(carpeta_temp, fn)
                if os.path.exists(src):
                    sh.move(src, os.path.join(ruta_final, fn))
            return ruta_final

        carpetas_creadas = []

        # Tipo de transporte independiente por archivo (A y B pueden ser terrestre / marítimo)
        tipo_a = self._clasificar_tipo_transporte(puerto_a, peso_a)
        tipo_b = self._clasificar_tipo_transporte(puerto_b, peso_b)

        def _append_suffix(parts, frac, tipo, puerto):
            if tipo == "TERRESTRE":
                if frac:
                    parts.append(frac)
            else:
                p = str(puerto).strip() if puerto else ""
                if p:
                    parts.append(p)

        # Carpeta A: COMPARTIDO
        pa = [fecha_str, cant, tipo_a, pe_a, str(carp_a)[:10] if carp_a else "0", _abreviar_dest(dest_a, tipo_a)]
        _append_suffix(pa, frac_a, tipo_a, puerto_a)
        pa.append("COMPARTIDO")
        archivos_a = [comparte_fn]
        if hoja_a: archivos_a.append(hoja_a)
        if pdf_a: archivos_a.append(pdf_a)
        ruta_a = _mover("_".join(pa), archivos_a)
        carpetas_creadas.append(("_".join(pa), ruta_a))
        self._log(f"  📁 {'/'.join(ruta_a.split(chr(92))[-2:])} ({len(archivos_a)} archivos)")

        # Carpeta B: COMPARTIDO_CERRA
        pb = [fecha_str, cant, tipo_b, pe_b, str(carp_b)[:10] if carp_b else "0", _abreviar_dest(dest_b, tipo_b)]
        _append_suffix(pb, frac_b, tipo_b, puerto_b)
        pb.append("COMPARTIDO_CERRADO")
        archivos_b = [cierra_fn]
        if hoja_b: archivos_b.append(hoja_b)
        if pdf_b: archivos_b.append(pdf_b)
        ruta_b = _mover("_".join(pb), archivos_b)
        carpetas_creadas.append(("_".join(pb), ruta_b))
        self._log(f"  📁 {'/'.join(ruta_b.split(chr(92))[-2:])} ({len(archivos_b)} archivos)")

        for fn in otros:
            src = os.path.join(carpeta_temp, fn)
            if os.path.exists(src):
                for _, rf in carpetas_creadas:
                    sh.copy2(src, os.path.join(rf, fn))
                os.remove(src)

        try:
            sh.rmtree(carpeta_temp)
        except Exception:
            pass

        return carpetas_creadas

    def _mail_popup_resumen(self, resultados):
        """Popup mostrando resumen de mails descargados y adjuntos."""
        popup = ctk.CTkToplevel(self)
        popup.title("Descarga de Mails — Completado")
        popup.geometry("500x380")
        popup.configure(fg_color=Palette.BG_CARD)
        popup.transient(self)
        popup.lift()
        popup.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 500) // 2
        py = self.winfo_y() + (self.winfo_height() - 380) // 2
        popup.geometry(f"500x380+{px}+{py}")

        ctk.CTkLabel(
            popup, text="✅ DESCARGAS COMPLETADAS",
            font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
            text_color=Palette.SUCCESS,
        ).pack(pady=(20, 8))

        ctk.CTkLabel(
            popup, text=f"{len(resultados)} mail(s) procesados — Archivos guardados en el Escritorio",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.TEXT_MUTED,
        ).pack(pady=(0, 12))

        # Frame scrolleable con los resultados
        scroll = ctk.CTkScrollableFrame(popup, fg_color=Palette.BG_MAIN, corner_radius=6)
        scroll.pack(fill="both", expand=True, padx=20, pady=(0, 12))

        for asunto, adjuntos, carpeta in resultados:
            frame_mail = ctk.CTkFrame(scroll, fg_color=Palette.BG_CARD, corner_radius=6,
                                       border_width=1, border_color=Palette.BORDER)
            frame_mail.pack(fill="x", pady=3)

            ctk.CTkLabel(
                frame_mail, text=asunto[:70],
                font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
                text_color=Palette.TEXT_PRIMARY, anchor="w",
            ).pack(anchor="w", padx=10, pady=(8, 2))

            ctk.CTkLabel(
                frame_mail, text=f"📁 {carpeta}  ·  {len(adjuntos)} adjunto(s)",
                font=ctk.CTkFont(family=FONT_FAMILY, size=11),
                text_color=Palette.TEXT_MUTED, anchor="w",
            ).pack(anchor="w", padx=10)

            if adjuntos:
                adj_texto = ", ".join(adjuntos)
                ctk.CTkLabel(
                    frame_mail, text=f"   📎 {adj_texto}",
                    font=ctk.CTkFont(family=FONT_FAMILY, size=11),
                    text_color=Palette.TEXT_SECONDARY, anchor="w",
                ).pack(anchor="w", padx=10, pady=(0, 8))

        ctk.CTkButton(
            popup, text="Cerrar", width=100, height=32,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6,
            command=popup.destroy,
        ).pack(pady=(0, 16))

    # ═══════════════════════════════════════════════════════════════════
    # SÚPER AUTO
    # ═══════════════════════════════════════════════════════════════════
    def _super_toggle(self):
        """Activa/desactiva el modo Súper Automatización."""
        if self._super_switch.get():
            guardas = self._cfg_obtener("valores", "guardas", ["Gonzalez", "Rodriguez", "Martinez", "Perez"])
            popup = ctk.CTkToplevel(self)
            popup.title("Súper Auto — Elegir Guarda")
            popup.geometry("380x180")
            popup.configure(fg_color=Palette.BG_CARD)
            popup.transient(self)
            popup.lift()
            popup.grab_set()
            popup.update_idletasks()
            px = self.winfo_x() + (self.winfo_width() - 380) // 2
            py = self.winfo_y() + (self.winfo_height() - 180) // 2
            popup.geometry(f"380x180+{px}+{py}")

            ctk.CTkLabel(popup, text="Seleccioná el guarda para Súper Auto:",
                         font=ctk.CTkFont(family=FONT_FAMILY, size=13),
                         text_color=Palette.TEXT_PRIMARY).pack(pady=(20, 10))
            guarda_var = ctk.StringVar(value=guardas[0] if guardas else "")
            ctk.CTkOptionMenu(popup, variable=guarda_var, values=guardas,
                              font=ctk.CTkFont(family=FONT_FAMILY, size=13),
                              fg_color=Palette.BG_INPUT, button_color=Palette.ACCENT,
                              button_hover_color=Palette.ACCENT_HOVER,
                              text_color=Palette.TEXT_PRIMARY, corner_radius=6,
                              width=220, height=34).pack(pady=(0, 12))

            def _confirmar():
                self._super_guarda = guarda_var.get()
                self._super_auto = True
                self._super_lbl_guarda.configure(text=f"Guarda: {self._super_guarda}")
                popup.destroy()

            def _cancelar():
                self._super_switch.deselect()
                self._super_auto = False
                self._super_guarda = ""
                self._super_lbl_guarda.configure(text="")
                popup.destroy()

            ctk.CTkButton(popup, text="Activar", width=100, height=32,
                          font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
                          fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
                          text_color=Palette.WHITE, corner_radius=6,
                          command=_confirmar).pack(side="left", padx=(100, 8))
            ctk.CTkButton(popup, text="Cancelar", width=100, height=32,
                          font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                          fg_color=Palette.BG_HOVER, hover_color=Palette.ERROR,
                          text_color=Palette.TEXT_SECONDARY, corner_radius=6,
                          command=_cancelar).pack(side="left")
        else:
            self._super_auto = False
            self._super_guarda = ""
            self._super_lbl_guarda.configure(text="")

    def _super_ejecutar_cadena(self, resultados):
        """Ejecuta la cadena: imprimir → guarda → planillas."""
        carpetas = []
        for _, _, nombre_carpeta in resultados:
            escritorio = self._resolver_ruta("descarga_mails", "Desktop")
            ruta = os.path.join(escritorio, nombre_carpeta)
            if os.path.isdir(ruta):
                carpetas.append(ruta)

        if not carpetas:
            self._log("⚡ Súper Auto: sin carpetas para procesar.")
            return

        self._log(f"⚡ Súper Auto: procesando {len(carpetas)} carpetas...")
        self._ultimas_carpetas = carpetas

        def _cadena():
            cfg = self._cfg_obtener("super_auto", "pasos", {})
            res_imp = self._super_imprimir(carpetas, cfg) if cfg.get("sobre", True) or cfg.get("permiso", True) or cfg.get("hoja_ruta", True) or cfg.get("recibo_ata", True) else 0
            res_gua = self._super_aplicar_guarda(carpetas) if cfg.get("aplicar_guarda", True) else 0
            res_pla = self._super_completar_planillas() if cfg.get("completar_planillas", True) else False
            self.after(0, lambda: self._super_popup_final(res_imp, res_gua, res_pla))

        t = threading.Thread(target=_cadena, daemon=True)
        t.start()

    def _super_imprimir(self, carpetas, cfg=None):
        """Imprime usando la misma lógica que el panel de Impresión Documental."""
        if cfg is None:
            cfg = {}
        hacer_sobre    = cfg.get("sobre", True)
        hacer_permiso  = cfg.get("permiso", True)
        hacer_hoja_ruta = cfg.get("hoja_ruta", True)
        hacer_recibo   = cfg.get("recibo_ata", True)
        anio = datetime.now().strftime("%y")
        prefijo_permiso = f"{anio}069EC"
        impresora = self._detectar_impresoras()[0] if self._detectar_impresoras() else "Default"
        total_ok = 0

        for ruta in carpetas:
            nombre = os.path.basename(ruta)
            self.log_queue.put(f"[...] 📁 {nombre}")
            try:
                archivos = sorted(os.listdir(ruta))
            except Exception:
                continue

            # Detectar archivos (misma lógica que _imp_worker)
            sobres = [a for a in archivos if "CONTENEDORES" in a.upper() and (a.upper().endswith(".XLSX") or a.upper().endswith(".XLS"))]
            permisos = [a for a in archivos if a.upper().startswith(prefijo_permiso) and a.upper().endswith(".PDF")]
            hojas_ruta = [a for a in archivos if "HOJA" in a.upper() and "RUTA" in a.upper() and (a.upper().endswith(".XLSX") or a.upper().endswith(".XLS"))]

            # 1. SOBRE: imprimir hoja "SOBRE" de Contenedores (1 copia)
            if hacer_sobre and sobres:
                for sobre in sobres:
                    ruta_excel = os.path.join(ruta, sobre)
                    hojas = self._imp_hojas_sobre(ruta_excel)
                    if hojas:
                        self.log_queue.put(f"[...]   📄 Sobre: {sobre} → {hojas}")
                        ok = self._imp_enviar(ruta_excel, impresora, f"  Sobre ({sobre[:40]})", hojas=hojas, copias=1)
                        if ok:
                            total_ok += 1
            else:
                self.log_queue.put(f"[...]   ⚠ Sin archivo Contenedores")

            # 2. Permisos (2 copias)
            if hacer_permiso and permisos:
                for a in permisos:
                    self.log_queue.put(f"[...]   📄 Permiso: {a} (2 copias)")
                    ok = self._imp_enviar(os.path.join(ruta, a), impresora, f"  Permiso ({a[:40]})", copias=2)
                    if ok:
                        total_ok += 1
            else:
                self.log_queue.put(f"[...]   ⚠ Sin Permiso (busca: {prefijo_permiso}*.PDF)")

            # 3. Hojas de Ruta (2 copias)
            if hacer_hoja_ruta and hojas_ruta:
                for a in hojas_ruta:
                    self.log_queue.put(f"[...]   📄 Hoja Ruta: {a} (2 copias)")
                    ok = self._imp_enviar(os.path.join(ruta, a), impresora, f"  Hoja Ruta ({a[:40]})", copias=2)
                    if ok:
                        total_ok += 1
            else:
                self.log_queue.put(f"[...]   ⚠ Sin Hoja de Ruta")

            # 4. Servicio ATA / Recibo ATA: hojas del Excel Contenedores
            if hacer_recibo and sobres:
                tipo = self._detectar_tipo_carpeta(nombre)
                if tipo in ("ISO", "FLEXI"):
                    self.log_queue.put(f"[...]   ⏭ Saltando Recibo ATA: carpeta marítima ({tipo})")
                else:
                    nombres_ata = ["RECIBO ATA"] + [f"RECIBO ATA {i}" for i in range(2, 9)]
                    hojas_ata = []
                    for sobre in sobres:
                        ruta_excel = os.path.join(ruta, sobre)
                        try:
                            if sobre.lower().endswith(".xlsx"):
                                wb_tmp = self._abrir_excel_seguro(ruta_excel)
                                for sn in wb_tmp.sheetnames:
                                    if sn.upper() in [n.upper() for n in nombres_ata]:
                                        hojas_ata.append((sobre, sn))
                                wb_tmp.close()
                            else:
                                book = xlrd.open_workbook(ruta_excel)
                                for sn in book.sheet_names():
                                    if sn.upper() in [n.upper() for n in nombres_ata]:
                                        hojas_ata.append((sobre, sn))
                        except Exception:
                            pass
                    if hojas_ata:
                        for archivo, hoja in hojas_ata:
                            self.log_queue.put(f"[...]   📄 Recibo ATA: {archivo} → {hoja}")
                            ok = self._imp_enviar(os.path.join(ruta, archivo), impresora, f"  Recibo ATA ({archivo[:30]})", hojas=[hoja])
                            if ok:
                                total_ok += 1
                    else:
                        self.log_queue.put(f"[...]   ⚠ Sin hoja Recibo ATA en Contenedores")
            else:
                self.log_queue.put(f"[...]   ⚠ Sin Contenedores para Recibo ATA")

        self.log_queue.put(f"[...] ✓ Impresión: {total_ok} documentos enviados de {len(carpetas)} carpetas")
        return total_ok

    def _super_aplicar_guarda(self, carpetas):
        """Escribe el guarda en los archivos Contenedores de las carpetas."""
        aplicados = 0
        for carpeta in carpetas:
            try:
                archivos = os.listdir(carpeta)
            except Exception:
                continue
            for archivo in archivos:
                up = archivo.upper()
                if "CONTENEDORES" not in up or not (up.endswith(".XLSX") or up.endswith(".XLS")):
                    continue
                ruta = os.path.join(carpeta, archivo)
                es_xls = ruta.lower().endswith(".xls") and not ruta.lower().endswith(".xlsx")
                # Reintentos automáticos: el archivo puede estar bloqueado brevemente
                # si COM lo usó para imprimir justo antes.
                max_reintentos = 6
                ultimo_error = None
                for _reintento in range(max_reintentos):
                    try:
                        _abrir_ok = True
                        if es_xls:
                            import xlrd as xlrd_local
                            xlrd_local.open_workbook(ruta, formatting_info=True).release_resources()
                        else:
                            openpyxl.load_workbook(ruta, read_only=True).close()
                        break
                    except PermissionError as _pe:
                        _abrir_ok = False
                        ultimo_error = _pe
                        if _reintento < max_reintentos - 1:
                            self.log_queue.put(f"[...]   ⏳ {archivo[:40]} en uso, esperando... ({_reintento+1}/{max_reintentos-1})")
                            time.sleep(1.5)
                else:
                    self.log_queue.put(f"[...]   ⚠ Error guarda en {archivo}: {ultimo_error}")
                    continue
                try:
                    if not self._escribir_guarda_en_archivo(ruta, self._super_guarda, self.log_queue.put):
                        self.log_queue.put(f"[...]   ⚠ 'Guarda' no hallada en {archivo[:50]}")
                    else:
                        aplicados += 1
                except Exception as e:
                    self.log_queue.put(f"[...]   ⚠ Error guarda en {archivo}: {e}")
                    import traceback
                    self.log_queue.put(f"[...]     {traceback.format_exc()}")
        self.log_queue.put(f"[...] ✓ Guarda '{self._super_guarda}': {aplicados} planillas actualizadas")
        return aplicados

    def _escribir_guarda_en_archivo(
        self, ruta: str, guarda_nombre: str, log_func=None
    ) -> bool:
        """Write guarda_nombre to column H in the CHOFER sheet of an Excel file.

        Handles both .xls (Excel COM via win32com) and .xlsx (openpyxl) formats.
        Returns True if written, False if no 'GUARDA' found or no CHOFER sheet.
        """
        es_xls = ruta.lower().endswith(".xls") and not ruta.lower().endswith(".xlsx")
        escrito = False

        if es_xls:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                try:
                    wb = excel.Workbooks.Open(ruta)
                    ws_chofer = None
                    for s in wb.Sheets:
                        if "CHOFER" in s.Name.upper():
                            ws_chofer = s
                            break
                    if not ws_chofer:
                        if log_func:
                            log_func("[...]   Hoja 'Choferes' no hallada")
                        return False
                    for row in range(2, 16):
                        val = ws_chofer.Cells(row, 7).Value  # col 7 = G
                        if val is not None and "GUARDA" in str(val).strip().upper():
                            ws_chofer.Cells(row, 8).Value = guarda_nombre  # col 8 = H
                            escrito = True
                            break
                    wb.Save()
                    wb.Close()
                finally:
                    excel.Quit()
            finally:
                pythoncom.CoUninitialize()
        else:
            wb = self._abrir_excel_seguro(ruta)
            try:
                ws_chofer = None
                for s in wb.sheetnames:
                    if "CHOFER" in s.upper():
                        ws_chofer = wb[s]
                        break
                if not ws_chofer:
                    if log_func:
                        log_func("[...]   Hoja 'Choferes' no hallada")
                    return False
                for row in range(2, 16):
                    cell = ws_chofer.cell(row=row, column=7)  # col 7 = G (1-indexed)
                    val = cell.value
                    if val is None:
                        for mr in ws_chofer.merged_cells.ranges:
                            if (mr.min_col <= 7 <= mr.max_col and mr.min_row <= row <= mr.max_row):
                                val = ws_chofer.cell(row=mr.min_row, column=mr.min_col).value
                                break
                    if val is not None and "GUARDA" in str(val).strip().upper():
                        ws_chofer.cell(row=row, column=8).value = guarda_nombre  # col 8 = H
                        escrito = True
                        break
                self._guardar_excel_seguro(wb, ruta)
            finally:
                wb.close()

        return escrito

    def _super_completar_planillas(self):
        """Ejecuta Completar Planillas con los datos extraídos."""
        try:
            self._planillas_core()
            self.log_queue.put(f"[...] ✓ Contenedores completadas (SOBRES, COBRO, PC)")
            return True
        except Exception as e:
            self.log_queue.put(f"[...] ⚠ Error completando planillas: {e}")
            return False

    def _super_popup_final(self, impresos, guardas, planillas_ok):
        """Popup resumen del Súper Auto."""
        popup = ctk.CTkToplevel(self)
        popup.title("Súper Auto Completado")
        popup.geometry("400x300")
        popup.configure(fg_color=Palette.BG_CARD)
        popup.transient(self)
        popup.lift()
        popup.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 400) // 2
        py = self.winfo_y() + (self.winfo_height() - 300) // 2
        popup.geometry(f"400x300+{px}+{py}")

        ctk.CTkLabel(popup, text="⚡ SÚPER AUTO COMPLETADO",
                     font=ctk.CTkFont(family=FONT_FAMILY, size=18, weight="bold"),
                     text_color=Palette.TEXT_PRIMARY).pack(pady=(24, 20))

        ctk.CTkLabel(popup, text=f"🖨  {impresos} archivos impresos",
                     font=ctk.CTkFont(family=FONT_FAMILY, size=13),
                     text_color=Palette.SUCCESS).pack(pady=3)
        txt_guarda = f"🛡  Guarda '{self._super_guarda}' aplicado a {guardas} planillas" if guardas > 0 else f"⚠  Guarda '{self._super_guarda}' — no se encontró etiqueta GUARDA"
        ctk.CTkLabel(popup, text=txt_guarda,
                     font=ctk.CTkFont(family=FONT_FAMILY, size=13),
                     text_color=Palette.SUCCESS if guardas > 0 else Palette.WARNING).pack(pady=3)
        ctk.CTkLabel(popup, text=f"{'✓' if planillas_ok else '✗'} Planillas completadas (SOBRES, COBRO, PC)",
                     font=ctk.CTkFont(family=FONT_FAMILY, size=13),
                     text_color=Palette.SUCCESS if planillas_ok else Palette.ERROR).pack(pady=3)

        ctk.CTkButton(popup, text="Cerrar", width=120, height=34,
                      font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
                      fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
                      text_color=Palette.WHITE, corner_radius=6,
                      command=popup.destroy).pack(pady=(20, 16))

    # ═══════════════════════════════════════════════════════════════════
    # PANEL: BACKUP
    # ═══════════════════════════════════════════════════════════════════
    def _panel_backup(self):
        if "backup" in self._panel_frames:
            return
        frame = ctk.CTkFrame(self.panel_container, fg_color="transparent")
        self._panel_frames["backup"] = frame
        # frame.pack(fill="both", expand=True) — packed by _cambiar_panel_forzado

        # ── Toolbar ──────────────────────────────────────────────────
        toolbar = ctk.CTkFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER, height=44
        )
        toolbar.pack(fill="x", pady=(0, 6))
        toolbar.pack_propagate(False)

        self.btn_backup_pendrive = ctk.CTkButton(
            toolbar,
            text="💿  Back Up",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.SECONDARY, hover_color=Palette.SECONDARY_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=34, width=170,
            command=self._backup_pendrive_iniciar,
        )
        self.btn_backup_pendrive.pack(side="left", padx=4, pady=4)

        self.lbl_estado_backup = ctk.CTkLabel(
            toolbar,
            text="Listo para respaldar",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.TEXT_PRIMARY,
        )
        self.lbl_estado_backup.pack(side="left", padx=(8, 0))

        self.progress_backup = ctk.CTkProgressBar(
            toolbar, width=140, height=8, corner_radius=4,
            fg_color=Palette.BG_INPUT, progress_color=Palette.ACCENT,
        )
        self.progress_backup.pack(side="right", padx=12)
        self.progress_backup.set(0)

        # ── Lista de carpetas con checkboxes ─────────────────────────
        self._backup_scroll = ctk.CTkScrollableFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER, height=180)
        self._backup_scroll.pack(fill="both", expand=True, pady=(0, 6))
        self._backup_checks = {}
        self._refrescar_lista_backup()

    def _refrescar_lista_backup(self):
        """Actualiza la lista de carpetas del escritorio con checkboxes."""
        for w in self._backup_scroll.winfo_children():
            w.destroy()
        self._backup_checks.clear()
        escritorio = self._resolver_ruta("planillas_carga", "Desktop")
        carpetas = []
        try:
            for item in sorted(os.listdir(escritorio)):
                ruta = os.path.join(escritorio, item)
                if not os.path.isdir(ruta) or item.startswith("."):
                    continue
                for archivo in os.listdir(ruta):
                    up = archivo.upper()
                    if "CONTENEDORES" in up or (up.startswith("PLANILLA DE CARGA") and (up.endswith(".XLSX") or up.endswith(".XLS"))):
                        carpetas.append((item, ruta)); break
        except Exception:
            pass
        if not carpetas:
            ctk.CTkLabel(self._backup_scroll,
                         text=f"No se encontraron carpetas de carga en:\n{escritorio}",
                         font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                         text_color=Palette.TEXT_MUTED, justify="center",
                         wraplength=400).pack(expand=True, pady=30)
        for nombre, _ in carpetas:
            var = ctk.BooleanVar(value=True)
            self._backup_checks[nombre] = var
            ctk.CTkCheckBox(self._backup_scroll, text=nombre, variable=var,
                            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
                            text_color=Palette.TEXT_PRIMARY).pack(anchor="w", padx=10, pady=3)

    # ── Backup Pendrive ──────────────────────────────────────────────
    def _backup_pendrive_iniciar(self):
        if self.tarea_activa:
            messagebox.showwarning("Agente ocupado", "Hay una tarea en ejecución.")
            return
        # Obtener carpetas chequeadas
        seleccionadas = []
        escritorio = self._resolver_ruta("planillas_carga", "Desktop")
        for nombre, var in self._backup_checks.items():
            if var.get():
                ruta = os.path.join(escritorio, nombre)
                if os.path.isdir(ruta):
                    seleccionadas.append(ruta)
        if not seleccionadas:
            messagebox.showwarning("Nada seleccionado", "Seleccioná al menos una carpeta.")
            return
        self._ejecutar_backup(seleccionadas)

    def _ejecutar_backup(self, seleccionadas):
        """Ejecuta el backup con las carpetas elegidas."""
        self.tarea_activa = True
        self._cancelar_tarea.clear()
        self.btn_backup_pendrive.configure(text="⏳  Moviendo...", state="disabled")
        self.progress_backup.configure(mode="indeterminate")
        self.progress_backup.start()
        self.lbl_estado_backup.configure(text="Detectando pendrive...", text_color=Palette.INFO)
        self._limpiar_log()
        self._log("💿 Iniciando Back Up...")
        t = threading.Thread(target=lambda: self._backup_pendrive_worker(seleccionadas), daemon=True)
        t.start()

    def _backup_pendrive_worker(self, carpetas=None):
        self._set_log_panel("backup")
        try:
            from ctypes import windll
            meses_es = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
                        7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
            ruta_base = self._cfg_obtener_rutas("backup_pendrive", "TRABAJO\\CARGAS")

            pendrive_encontrado = None
            # Primero buscar un pendrive que YA TENGA la carpeta base (evita Ventoy)
            for letra_ascii in range(ord('D'), ord('Z') + 1):
                unidad = f"{chr(letra_ascii)}:\\"
                try:
                    tipo = windll.kernel32.GetDriveTypeW(unidad)
                    if tipo == 2:  # DRIVE_REMOVABLE
                        test_path = os.path.join(unidad, ruta_base)
                        if os.path.exists(test_path):
                            pendrive_encontrado = unidad
                            self.log_queue.put(f"[...] ✓ Pendrive detectado con {ruta_base}: {pendrive_encontrado}")
                            break
                except Exception:
                    pass
            # Si ninguno tiene la carpeta, usar el primer removible
            if not pendrive_encontrado:
                for letra_ascii in range(ord('D'), ord('Z') + 1):
                    unidad = f"{chr(letra_ascii)}:\\"
                    try:
                        tipo = windll.kernel32.GetDriveTypeW(unidad)
                        if tipo == 2:
                            pendrive_encontrado = unidad
                            self.log_queue.put(f"[...] ✓ Pendrive (nuevo): {pendrive_encontrado}")
                            break
                    except Exception:
                        pass

            if not pendrive_encontrado:
                self.log_queue.put("[...] ⚠ No se detectó pendrive removible (D-Z)")
                self.after(0, self._backup_done)
                return

            self.log_queue.put(f"[...] ✓ Pendrive: {pendrive_encontrado}")

            self.after(0, lambda: self.progress_backup.configure(mode="determinate"))
            self.after(0, lambda: self.progress_backup.set(0))
            total = len(carpetas)
            nombres_movidos = []
            destinos_usados = set()
            for i, carpeta in enumerate(carpetas):
                if self._cancelar_tarea.is_set():
                    self.log_queue.put("[...] ⚠ Tarea cancelada."); break
                nombre = os.path.basename(carpeta)
                # Extraer año y mes de cada carpeta individualmente
                año = str(datetime.now().year)
                mes = datetime.now().month
                m = re.match(r"(\d{1,2})_(\d{1,2})_(\d{4})", nombre)
                if m:
                    mes = int(m.group(2))
                    año = m.group(3)
                mes_str = f"{mes:02d}_{meses_es.get(mes, meses_es[datetime.now().month])}"
                carpeta_destino = os.path.join(pendrive_encontrado, ruta_base, f"Cargas_{año}", mes_str)
                os.makedirs(carpeta_destino, exist_ok=True)
                destino = os.path.join(carpeta_destino, nombre)
                if os.path.exists(destino):
                    shutil.rmtree(destino)
                shutil.move(carpeta, destino)
                self.log_queue.put(f"[...]   ✓ {nombre} → {ruta_base}\\Cargas_{año}\\{mes_str}\\")
                nombres_movidos.append(nombre)
                destinos_usados.add(f"{ruta_base}\\Cargas_{año}\\{mes_str}")
                self.after(0, lambda p=i+1, t=total: self.progress_backup.set(p/t if t else 1))

            destinos_str = ", ".join(sorted(destinos_usados))
            self.log_queue.put(f"[...] ✓ COMPLETADO: Back Up — {len(carpetas)} carpetas movidas → {pendrive_encontrado}{destinos_str}")
            # Guardar para el popup
            self._backup_resultados = {"carpetas": nombres_movidos, "destino": f"{pendrive_encontrado}{destinos_str}"}
        except Exception as e:
            self.log_queue.put(f"[...] ⚠ Error: {e}")
            self._backup_resultados = None
        finally:
            self.after(0, self._backup_done)

    def _backup_done(self):
        self.tarea_activa = False
        try:
            self.btn_backup_pendrive.configure(text="💿  Back Up", state="normal", fg_color=Palette.SECONDARY)
            self.progress_backup.stop()
            self.progress_backup.set(1)
            self.lbl_estado_backup.configure(text="Backup completado", text_color=Palette.SUCCESS)
            self._refrescar_lista_backup()
        except (AttributeError, Exception):
            pass

        # Popup de resultados
        res = getattr(self, '_backup_resultados', None)
        if res and res.get("carpetas"):
            carpetas = res["carpetas"]
            destino = res["destino"]
            detalle = ""
            for c in carpetas:
                detalle += f"  ▸ {c}\n"
            self.after(300, lambda d=detalle, t=len(carpetas), dest=destino: messagebox.showinfo(
                "💿 Back Up Completado",
                f"{t} carpeta(s) movida(s) al pendrive:\n\n"
                f"{d}\n"
                f"📁 Destino: {dest}"))

    # ═══════════════════════════════════════════════════════════════════
    # ═══════════════════════════════════════════════════════════════════
    # AGREGAR GUARDA
    # ═══════════════════════════════════════════════════════════════════
    def _popup_agregar_guarda(self):
        """Popup para elegir un guarda y las carpetas donde aplicarlo."""
        if self.tarea_activa:
            messagebox.showwarning("Agente ocupado", "Hay una tarea en ejecución. Espere a que finalice.")
            return

        guardas = self._cfg_obtener("valores", "guardas", ["Gonzalez", "Rodriguez", "Martinez", "Perez"])
        escritorio = self._resolver_ruta("planillas_carga", "Desktop")

        # Buscar carpetas disponibles (igual que el análisis de planillas)
        carpetas_encontradas = []
        try:
            for item in sorted(os.listdir(escritorio)):
                ruta = os.path.join(escritorio, item)
                if not os.path.isdir(ruta):
                    continue
                if item.startswith("."):
                    continue
                # Buscar planillas de carga dentro
                try:
                    archivos = os.listdir(ruta)
                except Exception:
                    continue
                for a in archivos:
                    up = a.upper()
                    if "CONTENEDORES" in up and (up.endswith(".XLSX") or up.endswith(".XLS")):
                        carpetas_encontradas.append((item, ruta))
                        break
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el escritorio:\n{e}")
            return

        if not carpetas_encontradas:
            messagebox.showinfo("Sin planillas", "No se encontraron archivos Contenedores en el escritorio.")
            return

        # ── Popup ──────────────────────────────────────────────────
        popup = ctk.CTkToplevel(self)
        popup.title("Agregar Guarda")
        popup.geometry("550x520")
        popup.configure(fg_color=Palette.BG_CARD)
        popup.transient(self)
        popup.lift()
        popup.grab_set()
        popup.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 550) // 2
        py = self.winfo_y() + (self.winfo_height() - 520) // 2
        popup.geometry(f"550x520+{px}+{py}")

        ctk.CTkLabel(
            popup, text="Seleccioná un Guarda",
            font=ctk.CTkFont(family=FONT_FAMILY, size=15, weight="bold"),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(pady=(16, 8))

        # Dropdown de guardas
        guarda_var = ctk.StringVar(value=guardas[0] if guardas else "")
        ctk.CTkOptionMenu(
            popup, variable=guarda_var, values=guardas,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13),
            fg_color=Palette.BG_INPUT, button_color=Palette.ACCENT,
            button_hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=6,
            width=240, height=36,
        ).pack(pady=(0, 16))

        ctk.CTkLabel(
            popup, text="Carpetas donde se aplicará el guarda:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            text_color=Palette.TEXT_SECONDARY,
        ).pack(anchor="w", padx=16, pady=(0, 6))

        scroll = ctk.CTkScrollableFrame(popup, fg_color=Palette.BG_TABLE, corner_radius=8,
                                         border_width=1, border_color=Palette.BORDER)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        checks = {}
        for nombre, ruta_carp in carpetas_encontradas:
            var = ctk.BooleanVar(value=True)
            checks[nombre] = var
            ctk.CTkCheckBox(
                scroll, text=nombre, variable=var,
                font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
                text_color=Palette.TEXT_PRIMARY,
            ).pack(anchor="w", padx=12, pady=4)

        # Botones
        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(fill="x", padx=16, pady=(0, 12))

        ctk.CTkButton(
            btn_frame, text="Aplicar Guarda", width=180, height=34,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6,
            command=lambda: self._aplicar_guarda(
                popup, guarda_var.get(),
                [c[1] for c in carpetas_encontradas if checks[c[0]].get()]
            ),
        ).pack(side="right", padx=4)
        ctk.CTkButton(
            btn_frame, text="Cancelar", width=100, height=34,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_HOVER, hover_color=Palette.ERROR,
            text_color=Palette.TEXT_SECONDARY, corner_radius=6,
            command=popup.destroy,
        ).pack(side="right", padx=4)

    def _popup_editar_excels(self):
        """Popup con botones para abrir/editar archivos Excel."""
        popup = ctk.CTkToplevel(self)
        popup.title("Editar Excels")
        popup.geometry("360x360")
        popup.configure(fg_color=Palette.BG_CARD)
        popup.transient(self)
        popup.lift()
        popup.grab_set()
        popup.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 360) // 2
        py = self.winfo_y() + (self.winfo_height() - 360) // 2
        popup.geometry(f"360x360+{px}+{py}")

        ctk.CTkLabel(popup, text="EDITAR EXCELS",
                     font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
                     text_color=Palette.TEXT_PRIMARY).pack(pady=(20, 12))

        excels = [
            ("SOBRES_2026.xlsx", self._cfg_obtener_rutas("sobres", os.path.join("TRABAJO", "01_PLANILLAS"))),
            ("COBRO_2026.xlsx", self._cfg_obtener_rutas("cobro", os.path.join("TRABAJO", "01_PLANILLAS"))),
            ("PC.xlsx", self._cfg_obtener_rutas("pc", os.path.join("TRABAJO", "01_PLANILLAS"))),
            ("FECHA MIC Y SELLOS.xlsx", self._cfg_obtener_rutas("mic_sellos", os.path.join("TRABAJO", "01_PLANILLAS"))),
            ("FECHA CRT Y ORIGINAL.xlsx", self._cfg_obtener_rutas("crt_original", os.path.join("TRABAJO", "01_PLANILLAS"))),
            (self._cfg_obtener_rutas("carga_terrestre_nombre", "CARGA TERRESTRE.xlsx"),
             self._cfg_obtener_rutas("carga_terrestre_carpeta", os.path.join("TRABAJO", "01_PLANILLAS"))),
        ]

        def _abrir(nombre, carpeta):
            ruta = buscar_archivo_en_pendrive(nombre, carpeta)
            if ruta:
                os.startfile(ruta)
                popup.destroy()
            else:
                messagebox.showwarning("No encontrado", f"No se encontró:\n{nombre}\n\nEn: {carpeta}")

        for nombre, carpeta in excels:
            ctk.CTkButton(popup, text=f"📝 {nombre}",
                          font=ctk.CTkFont(family=FONT_FAMILY, size=13),
                          fg_color=Palette.BG_HOVER, hover_color=Palette.ACCENT_DIM,
                          text_color=Palette.TEXT_PRIMARY, corner_radius=6, height=34, width=280,
                          command=lambda n=nombre, c=carpeta: _abrir(n, c)).pack(pady=3)

        ctk.CTkButton(popup, text="Cerrar", width=100, height=32,
                      font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                      fg_color=Palette.BG_HOVER, hover_color=Palette.ERROR,
                      text_color=Palette.TEXT_PRIMARY, corner_radius=6,
                      command=popup.destroy).pack(pady=(8, 10))

    def _aplicar_guarda(self, popup, guarda_elegido, carpetas_seleccionadas):
        """Escribe el guarda elegido en la hoja Choferes de cada planilla de carga."""
        if not guarda_elegido or not carpetas_seleccionadas:
            messagebox.showwarning("Incompleto", "Elegí un guarda y al menos una carpeta.")
            return
        popup.destroy()

        if self.tarea_activa:
            return
        self.tarea_activa = True
        self._cancelar_tarea.clear()
        self.btn_agregar_guarda.configure(text="⏳  Procesando...", state="disabled")
        self.btn_ejecutar_planillas.configure(state="disabled")
        self._limpiar_log()
        self._log(f"🛡 Agregando Guarda: {guarda_elegido}")

        def worker():
            self._set_log_panel("planillas")
            for ruta_carp in carpetas_seleccionadas:
                if self._cancelar_tarea.is_set():
                    self.log_queue.put("[...] ⚠ Tarea cancelada.")
                    break
                nombre_carp = os.path.basename(ruta_carp)
                self.log_queue.put(f"[...]   📁 {nombre_carp}")

                ruta_contenedores = None
                try:
                    for archivo in os.listdir(ruta_carp):
                        up = archivo.upper()
                        if "CONTENEDORES" in up and (up.endswith(".XLSX") or up.endswith(".XLS")):
                            ruta_contenedores = os.path.join(ruta_carp, archivo)
                            break
                except Exception as e:
                    self.log_queue.put(f"[...]     ⚠ No se pudo leer: {e}")
                    continue

                if not ruta_contenedores:
                    self.log_queue.put(f"[...]     ⚠ Sin archivo Contenedores")
                    continue

                archivo_nombre = os.path.basename(ruta_contenedores)
                self.log_queue.put(f"[...]     📄 {archivo_nombre}")

                try:
                    if not self._escribir_guarda_en_archivo(ruta_contenedores, guarda_elegido, self.log_queue.put):
                        self.log_queue.put(f"[...]     ⚠ 'Guarda' no hallada en col G")
                    else:
                        self.log_queue.put(f"[...]     ✓ '{guarda_elegido}' escrito correctamente")
                except Exception as e:
                    self.log_queue.put(f"[...]     ⚠ Error: {e}")

            self.log_queue.put(f"[...] ✓ Guarda '{guarda_elegido}' completado.")
            self.after(0, self._guarda_done)

        t = threading.Thread(target=worker, daemon=True)
        t.start()

    def _guarda_done(self):
        self.tarea_activa = False
        try:
            self.btn_agregar_guarda.configure(
                text="🛡  Agregar Guarda", state="normal",
                fg_color=Palette.SECONDARY)
            self.btn_ejecutar_planillas.configure(state="normal")
        except (AttributeError, Exception):
            pass

    # ═══════════════════════════════════════════════════════════════════
    # POPUP: SELECCIONAR PLANILLAS A COMPLETAR
    # ═══════════════════════════════════════════════════════════════════
    def _popup_completar_planillas(self):
        """Popup para elegir qué planillas completar (SOBRES, COBRO, PC)."""
        if self.tarea_activa:
            messagebox.showwarning(
                "Agente ocupado",
                "Hay una tarea en ejecución. Espere a que finalice."
            )
            return

        popup = ctk.CTkToplevel(self)
        popup.title("Completar Planillas")
        popup.geometry("360x300")
        popup.configure(fg_color=Palette.BG_CARD)
        popup.transient(self)
        popup.lift()
        popup.resizable(False, False)

        popup.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 360) // 2
        py = self.winfo_y() + (self.winfo_height() - 300) // 2
        popup.geometry(f"360x300+{px}+{py}")

        ctk.CTkLabel(
            popup, text="Seleccionar Planillas",
            font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(pady=(20, 16))

        var_sobres = ctk.BooleanVar(value=True)
        var_cobro  = ctk.BooleanVar(value=True)
        var_pc     = ctk.BooleanVar(value=True)
        vars_tipos = {"sobres": var_sobres, "cobro": var_cobro, "pc": var_pc}

        def _actualizar_boton():
            alguna = any(v.get() for v in vars_tipos.values())
            btn_completar.configure(
                state="normal" if alguna else "disabled",
                fg_color=Palette.ACCENT if alguna else Palette.ACCENT_DIM,
            )

        def _seleccionar_todo():
            for v in vars_tipos.values():
                v.set(True)
            _actualizar_boton()

        def _ninguno():
            for v in vars_tipos.values():
                v.set(False)
            _actualizar_boton()

        # ── Checkboxes ────────────────────────────────────────────
        frame_cbs = ctk.CTkFrame(popup, fg_color="transparent")
        frame_cbs.pack(pady=(0, 8))

        for label, key in [("SOBRES", "sobres"), ("COBRO", "cobro"), ("PC (Precintos/Cables)", "pc")]:
            cb = ctk.CTkCheckBox(
                frame_cbs, text=label,
                font=ctk.CTkFont(family=FONT_FAMILY, size=13),
                variable=vars_tipos[key], command=_actualizar_boton,
                fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
                border_color=Palette.BORDER, checkmark_color=Palette.WHITE,
                text_color=Palette.TEXT_PRIMARY,
            )
            cb.pack(anchor="w", pady=4, padx=20)

        # ── Botones Seleccionar Todo / Ninguno ────────────────────
        frame_sel = ctk.CTkFrame(popup, fg_color="transparent")
        frame_sel.pack(pady=(4, 12))

        ctk.CTkButton(
            frame_sel, text="✓ Seleccionar Todo",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.SECONDARY, hover_color=Palette.SECONDARY_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=32, width=150,
            command=_seleccionar_todo,
        ).pack(side="left", padx=6)

        ctk.CTkButton(
            frame_sel, text="✗ Ninguno",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_HOVER, hover_color=Palette.ACCENT_DIM,
            text_color=Palette.TEXT_SECONDARY, corner_radius=6, height=32, width=110,
            command=_ninguno,
        ).pack(side="left", padx=6)

        # ── Botón Completar ───────────────────────────────────────
        btn_completar = ctk.CTkButton(
            popup, text="▶  Completar",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=38,
            command=lambda: self._confirmar_completar(popup, vars_tipos),
        )
        btn_completar.pack(fill="x", padx=30, pady=(6, 16))

    def _confirmar_completar(self, popup, vars_tipos):
        """Cierra el popup y lanza la ejecución con los tipos seleccionados."""
        tipos = [nombre for nombre, var in vars_tipos.items() if var.get()]
        if not tipos:
            return
        popup.destroy()
        self._ejecutar_agente_planillas(tipos)

    # ═══════════════════════════════════════════════════════════════════
    # AGENTE 2: EJECUCIÓN EN HILO
    # ═══════════════════════════════════════════════════════════════════
    def _ejecutar_agente_planillas(self, tipos=None):
        if self.tarea_activa:
            messagebox.showwarning(
                "Agente ocupado",
                "Hay una tarea en ejecución. Espere a que finalice."
            )
            return
        self.tarea_activa = True
        self._tipos_planillas = tipos
        self._cancelar_tarea.clear()
        self.btn_ejecutar_planillas.configure(
            text="⏳  Analizando...", state="disabled",
            fg_color=Palette.ACCENT_DIM
        )
        self.progress_planillas.configure(mode="indeterminate")
        self.progress_planillas.start()
        self.lbl_estado_planillas.configure(
            text="Escaneando Escritorio...", text_color=Palette.INFO
        )
        self._set_status("Analizando planillas del Escritorio...")

        # Limpiar pantalla antes de empezar
        self._limpiar_log()
        for row in self.tree_planillas.get_children():
            self.tree_planillas.delete(row)
        self.lbl_resumen_planillas.configure(text="Sin datos analizados")

        # Mostrar en log qué planillas se van a completar
        nombres = {"sobres": "SOBRES", "cobro": "COBRO", "pc": "PC"}
        seleccionadas = [nombres[t] for t in (tipos or ["sobres", "cobro", "pc"])]
        self.log_queue.put(f"[...] Planillas seleccionadas: {', '.join(seleccionadas)}")

        t = threading.Thread(target=self._agente_planillas_worker, daemon=True)
        t.start()

    def _agente_planillas_worker(self):
        """Ejecuta el agente de planillas en hilo de fondo."""
        self._set_log_panel("planillas")
        try:
            self._planillas_core(self._tipos_planillas)
        except Exception as e:
            self.after(0, lambda: self._planillas_error(str(e)))
        finally:
            self.after(0, self._planillas_done)

    def _buscar_buque_en_chofer(self, ruta_excel):
        """Busca 'BUQUE' en la hoja CHOFERES del Excel y devuelve el valor de la derecha.

        Soporta .xlsx (openpyxl) y .xls (xlrd). Case insensitive.
        Devuelve string vacío si no encuentra.
        """
        try:
            if str(ruta_excel).lower().endswith(".xlsx"):
                wb = openpyxl.load_workbook(ruta_excel, data_only=True)
                try:
                    pestana = next((s for s in wb.sheetnames if "CHOFER" in s.upper()), None)
                    if not pestana:
                        return ""
                    ws = wb[pestana]
                    for r in range(1, ws.max_row + 1):
                        for c in range(1, ws.max_column + 1):
                            val = ws.cell(row=r, column=c).value
                            if val and isinstance(val, str) and "BUQUE" in val.strip().upper():
                                nxt = ws.cell(row=r, column=c + 1).value
                                if nxt:
                                    return str(nxt).strip()
                finally:
                    wb.close()
            else:
                import xlrd
                book = xlrd.open_workbook(ruta_excel)
                pestana = next((s for s in book.sheet_names() if "CHOFER" in s.upper()), None)
                if not pestana:
                    return ""
                sheet = book.sheet_by_name(pestana)
                for r in range(sheet.nrows):
                    for c in range(sheet.ncols):
                        val = sheet.cell_value(r, c)
                        if val and isinstance(val, str) and "BUQUE" in val.strip().upper():
                            if c + 1 < sheet.ncols:
                                nxt = sheet.cell_value(r, c + 1)
                                if nxt:
                                    return str(nxt).strip()
        except Exception as e:
            self._log(f"  ⚠ Error buscando BUQUE en {os.path.basename(ruta_excel)}: {e}")
        return ""

    def _planillas_core(self, tipos=None):
        """Lógica real del agente de planillas (modificada para GUI)."""
        escritorio = self._resolver_ruta("planillas_carga", "Desktop")
        self._log(f"Escaneando: {escritorio}")

        # Tracking de resultados para el popup final
        resultados = {
            "sobres": {"ok": False, "detalle": "No procesado"},
            "cobro":  {"ok": False, "detalle": "No procesado"},
            "pc":     {"ok": False, "detalle": "No procesado"},
        }

        if not os.path.exists(escritorio):
            self._log("ERROR: No se pudo localizar la ruta del Escritorio.")
            self.after(0, lambda: self._mostrar_resumen(resultados))
            return

        # Buscar archivos en el pendrive (rutas configurables por tipo)
        ruta_sobres = buscar_archivo_en_pendrive(
            "SOBRES_2026.xlsx", self._cfg_obtener_rutas("sobres", os.path.join("TRABAJO", "01_PLANILLAS")))

        wb_sobres = None
        ws_sobres = None
        if ruta_sobres:
            try:
                wb_sobres = self._abrir_excel_seguro(ruta_sobres)
                meses_es = {
                    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
                    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
                    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
                }
                mes_actual = meses_es[datetime.now().month]
                for sheet in wb_sobres.sheetnames:
                    if mes_actual in sheet.upper():
                        ws_sobres = wb_sobres[sheet]
                        break
                if not ws_sobres:
                    self._log(f"ERROR: No se encontró hoja del mes '{mes_actual}' en SOBRES_2026.xlsx.")
                    self.after(0, lambda m=mes_actual: messagebox.showerror(
                        "Hoja del mes no encontrada",
                        f"No se encontró la hoja \"{m}\" en SOBRES_2026.xlsx.\n\n"
                        f"Verifique que el archivo contenga una hoja con el nombre del mes actual.\n\n"
                        f"Operación cancelada."
                    ))
                    wb_sobres.close()
                    return
                self._log(f"SOBRES_2026.xlsx detectada → {ruta_sobres} · hoja '{ws_sobres.title}'")
            except Exception as e:
                self._log(f"Error al abrir SOBRES_2026.xlsx: {e}")
        else:
            self._log("ERROR: No se encontró SOBRES_2026.xlsx.")
            pendrive_path = os.path.join("PENDRIVE:\\", "TRABAJO", "01_PLANILLAS", "SOBRES_2026.xlsx")
            self.after(0, lambda: messagebox.showwarning(
                "Planilla SOBRES no encontrada",
                f"No se encontró SOBRES_2026.xlsx.\n\n"
                f"Ruta esperada en el pendrive:\n"
                f"  {pendrive_path}\n\n"
                f"Conecte el pendrive y vuelva a ejecutar el análisis."
            ))
            # FRENAR: no seguir escaneando el Escritorio sin SOBRES
            return

        carpetas_encontradas = 0
        datos_extraidos = []

        for item in sorted(os.listdir(escritorio)):
            time.sleep(0.01) # Ceder el GIL
            ruta_carpeta = os.path.join(escritorio, item)
            if not os.path.isdir(ruta_carpeta):
                continue
            if item.startswith(".") or item.upper() in ("RECYCLED", "RECYCLER"):
                continue

            try:
                archivos_en_carpeta = os.listdir(ruta_carpeta)
            except Exception:
                continue

            # Filtrar excels de CONTENEDORES
            excels = []
            for archivo in archivos_en_carpeta:
                up = archivo.upper().strip()
                if archivo.startswith("~$"):
                    continue
                if (up.endswith(".XLSX") or up.endswith(".XLS")) and "CONTENEDORES" in up:
                    excels.append(archivo)

            if not excels:
                continue

            carpetas_encontradas += 1
            self._log(f"Carpeta localizada: \"{item}\"")

            match_frac = re.search(r"(F(?:RACCION)?\s*\d+)", item, re.IGNORECASE)
            frac_carpeta = match_frac.group(1).upper() if match_frac else "No detectada"

            match_pe_folder = re.search(r"(\d{3,4}[A-Z])", item, re.IGNORECASE)
            pe_carpeta = match_pe_folder.group(1).upper() if match_pe_folder else ""

            for archivo in excels:
                time.sleep(0.01)  # Ceder el GIL para evitar que la UI se congele
                ruta_excel = os.path.join(ruta_carpeta, archivo)
                self._log(f"  Leyendo: {archivo}")

                match_cant = re.match(r"^(\d+)", archivo.strip())
                cant_final = int(match_cant.group(1)) if match_cant else 1

                try:
                    if archivo.lower().endswith(".xlsx"):
                        datos = self._leer_xlsx_moderno(ruta_excel)
                    else:
                        datos = self._leer_xls_antiguo(ruta_excel)
                    f_celda, pe_celda, carp_celda, dest_celda, bl_celda, trans_final, primera_patente, precintos, guarda, puerto_salida, peso_flexi = datos

                    pe_origen = pe_celda if pe_celda else pe_carpeta
                    pe_recortado = pe_origen[-5:].lstrip("0") if pe_origen else "No detectado"
                    if not trans_final:
                        trans_final = "No detectado"

                    # Determinar tipo de transporte: primero desde Excel, fallback desde nombre de carpeta
                    tipo_transporte = self._clasificar_tipo_transporte(puerto_salida, peso_flexi)
                    if tipo_transporte == "TERRESTRE":
                        # Fallback: el nombre de carpeta ya tiene el tipo correcto (índice 4)
                        partes_carpeta = item.split("_")
                        if len(partes_carpeta) > 4 and partes_carpeta[4] in ("ISO", "FLEXI"):
                            tipo_transporte = partes_carpeta[4]
                    if tipo_transporte in ("ISO", "FLEXI"):
                        # Marítimo: buscar BUQUE en CHOFERES
                        bl_detectado = self._buscar_buque_en_chofer(ruta_excel)
                        bl_final = bl_detectado if bl_detectado else "-"
                    else:
                        bl_final = bl_celda if bl_celda else "-"

                    # Convertir carpeta a número si es posible (para evitar alerta de texto en Excel)
                    try:
                        carp_celda_num = int(carp_celda)
                    except (ValueError, TypeError):
                        carp_celda_num = carp_celda

                    registro = {
                        "f_celda": f_celda,
                        "cant_final": cant_final,
                        "pe_recortado": pe_recortado,
                        "pe_completo": pe_origen if pe_origen else "No detectado",
                        "carp_celda": carp_celda_num,
                        "tipo_transporte": "TERRES" if tipo_transporte == "TERRESTRE" else tipo_transporte,
                        "terminal": (puerto_salida[:9] if puerto_salida else "CHILE") if tipo_transporte in ("ISO", "FLEXI") else "CHILE",
                        "trans_final": trans_final,
                        "dest_celda": dest_celda,
                        "bl_final": bl_final,
                        "frac_carpeta": "-" if tipo_transporte in ("ISO", "FLEXI") else frac_carpeta,
                        "servicio": "-" if tipo_transporte in ("ISO", "FLEXI") else cant_final * int(self._cfg_obtener("valores", "ata_tares", 65000)),
                        "precintos": precintos,
                        "guarda": guarda if guarda else "No detectado",
                        "source_folder": item,
                        "source_file": archivo,
                    }
                    if "COMPARTIDO" in item.upper():
                        self._log(f"  🔀 Carpeta COMPARTIDO: {item}")
                    self.after(0, lambda r=registro: self._agregar_fila_planillas(r))

                    if ws_sobres is not None:
                        datos_extraidos.append(registro)

                except Exception as err:
                    self._log(f"  ERROR procesando {archivo}: {err}")

        # Guardar en SOBRES
        if datos_extraidos and ws_sobres is not None and (not tipos or "sobres" in tipos):
            self._log("Escribiendo datos en SOBRES_2026.xlsx...")
            try:

                def sort_date(item):
                    try:
                        p = item["f_celda"].split("/")
                        return datetime(int(p[2]), int(p[1]), int(p[0]))
                    except Exception:
                        return datetime.max

                datos_extraidos.sort(key=sort_date)

                def abreviar_dest(nombre):
                    n = (nombre or "").upper()
                    if "VITAPRO" in n: return "VITAPRO"
                    if "EWOS" in n: return "EWOS"
                    if "NUTRECO" in n: return "NUTRECO"
                    if "DICOAL" in n: return "DICOAL"
                    if "CARGILL" in n: return "CARGILL"
                    if "BIOMAR" in n: return "BIOMAR"
                    return nombre

                COLORES_EMPRESA = {
                    "VITAPRO": "7EB8E0",
                    "EWOS": "8CC56E", "DICOAL": "F4B183",
                    "NUTRECO": "FFEB9C", "CARGILL": "BDD7EE", "BIOMAR": "E2EFDA",
                }
                from openpyxl.styles import PatternFill, Border, Side, Alignment as Aln, Font

                thin = Side(style="thin")
                border = Border(left=thin, right=thin, top=thin, bottom=thin)
                center = Aln(horizontal="center", vertical="center", wrap_text=True)
                font_normal = Font(name="Calibri", size=11)
                font_mediano = Font(name="Calibri", size=9)   # columnas H, I
                font_chico = Font(name="Calibri", size=8)     # columna G

                # Agrupar pares comparte: detectar por carpeta (ambas con "COMPARTIDO")
                def _pares_comparte(datos):
                    usados = set()
                    items = []
                    for i, d1 in enumerate(datos):
                        if i in usados: continue
                        f1 = d1.get("source_folder", "")
                        if "COMPARTIDO" not in f1.upper():
                            items.append(d1); usados.add(i); continue
                        pareja = None
                        for j, d2 in enumerate(datos):
                            if j == i or j in usados: continue
                            f2 = d2.get("source_folder", "")
                            if ("COMPARTIDO" in f2.upper()
                                    and d1.get("dest_celda") == d2.get("dest_celda")
                                    and d1.get("f_celda") == d2.get("f_celda")):
                                pareja = j; break
                        if pareja is not None:
                            items.append((d1, datos[pareja]))
                            usados.add(i); usados.add(pareja)
                        else:
                            items.append(d1); usados.add(i)
                    return items

                def _escribir_celda(ws, row, col, val, font, fill=None):
                    cell = ws.cell(row=row, column=col)
                    cell.value = val
                    cell.number_format = 'General'
                    cell.border = border
                    cell.alignment = center
                    cell.font = font
                    if fill:
                        cell.fill = fill

                def _escribir_fecha(ws, row, f_celda_str):
                    nueva_fila = row
                    while True:
                        time.sleep(0.005)
                        v1 = ws.cell(row=nueva_fila, column=1).value
                        v5 = ws.cell(row=nueva_fila, column=5).value
                        if not v1 and not v5:
                            break
                        nueva_fila += 1
                    if nueva_fila != row:
                        self._log(f"  → Fila ajustada de {row} a {nueva_fila}")
                    r_prev = nueva_fila - 1
                    val_prev = None
                    while r_prev >= 3:
                        val_prev = ws.cell(row=r_prev, column=1).value
                        if val_prev is not None:
                            break
                        r_prev -= 1
                    if val_prev and formatear_fecha_excel(val_prev) == f_celda_str:
                        for m_range in list(ws.merged_cells.ranges):
                            if m_range.min_col == 1 and m_range.max_col == 1:
                                if (m_range.min_row <= nueva_fila <= m_range.max_row
                                        or m_range.min_row <= r_prev <= m_range.max_row):
                                    ws.merged_cells.ranges.remove(m_range)
                        ws.merge_cells(start_row=r_prev, start_column=1, end_row=nueva_fila, end_column=1)
                        fecha_merge = ws.cell(row=r_prev, column=1)
                        fecha_merge.number_format = 'General'
                        fecha_merge.alignment = Aln(horizontal="center", vertical="center")
                    else:
                        ws.cell(row=nueva_fila, column=1).value = f_celda_str
                    c1 = ws.cell(row=nueva_fila, column=1)
                    c1.number_format = 'General'
                    c1.border = border
                    c1.alignment = Aln(horizontal="center", vertical="center", wrap_text=True)
                    c1.font = font_normal
                    return nueva_fila

                items = _pares_comparte(datos_extraidos)
                for item in items:
                    es_par = isinstance(item, tuple) and len(item) == 2
                    if es_par:
                        d1, d2 = item
                        self._log(f"  🔀 Par COMPARTE: {d1.get('pe_recortado')} + {d2.get('pe_recortado')}")
                    else:
                        d1 = item
                        d2 = None

                    f_celda = d1["f_celda"]
                    dest_abrev = abreviar_dest(d1["dest_celda"])
                    color_hex = COLORES_EMPRESA.get(dest_abrev, "FFFFFF")
                    fill_color = PatternFill(fill_type="solid", fgColor=color_hex)

                    tipo = d1.get("tipo_transporte", "TERRESTRE")
                    es_maritimo = tipo in ("ISO", "FLEXI")
                    valores_check_1 = {
                        1: f_celda, 2: d1["cant_final"], 3: tipo, 4: d1["pe_recortado"],
                        5: d1["carp_celda"], 6: d1.get("terminal", "CHILE"), 7: d1["trans_final"], 8: dest_abrev,
                        9: d1["bl_final"], 10: "-" if es_maritimo else d1["frac_carpeta"],
                        11: "-" if es_maritimo else d1["cant_final"] * int(self._cfg_obtener("valores", "ata_tares", 65000)),
                    }
                    if ya_existe_en_hoja(ws_sobres, valores_check_1, excluir_columnas={9}):
                        self._log(f"  Omitido (ya existe): Carpeta {d1.get('carp_celda','?')} | P.E. {d1.get('pe_recortado','?')} | {f_celda}")
                        continue

                    r1 = _escribir_fecha(ws_sobres, 3, f_celda)
                    for col in (2, 3, 4, 5, 6, 7, 8, 9, 10, 11):
                        val = valores_check_1[col]
                        font = font_chico if col == 7 else (font_mediano if col in (6, 8, 9) else font_normal)
                        fl = fill_color if col == 8 else None
                        _escribir_celda(ws_sobres, r1, col, val, font, fl)

                    if es_par:
                        dest2 = abreviar_dest(d2["dest_celda"])
                        color2 = COLORES_EMPRESA.get(dest2, "FFFFFF")
                        fill2 = PatternFill(fill_type="solid", fgColor=color2)
                        tipo2 = d2.get("tipo_transporte", "TERRESTRE")
                        es_maritimo2 = tipo2 in ("ISO", "FLEXI")
                        valores_check_2 = {
                            4: d2["pe_recortado"], 5: d2["carp_celda"],
                            9: d2["bl_final"], 10: "-" if es_maritimo2 else d2["frac_carpeta"],
                        }
                        r2 = r1 + 1
                        # Fecha (col 1) merge con r1
                        for m_range in list(ws_sobres.merged_cells.ranges):
                            if m_range.min_col == 1 and m_range.max_col == 1:
                                if m_range.min_row <= r2 <= m_range.max_row or m_range.min_row <= r1 <= m_range.max_row:
                                    ws_sobres.merged_cells.ranges.remove(m_range)
                        ws_sobres.merge_cells(start_row=r1, start_column=1, end_row=r2, end_column=1)
                        ws_sobres.cell(row=r1, column=1).alignment = Aln(horizontal="center", vertical="center")

                        # Columnas individuales para fila 2
                        for col, val in valores_check_2.items():
                            font = font_chico if col == 7 else (font_mediano if col in (6, 8, 9) else font_normal)
                            fl = fill2 if col == 8 else None
                            _escribir_celda(ws_sobres, r2, col, val, font, fl)
                        # Cols compartidas (merge) para fila 2: B,C,F,G,H,K
                        for col in (2, 3, 6, 7, 8, 11):
                            val = valores_check_1[col]
                            font = font_chico if col == 7 else (font_mediano if col in (6, 8) else font_normal)
                            _escribir_celda(ws_sobres, r2, col, val, font)
                        # Merge vertical para B(2), C(3), F(6), G(7), H(8), K(11)
                        for col in (2, 3, 6, 7, 8, 11):
                            ws_sobres.merge_cells(start_row=r1, start_column=col, end_row=r2, end_column=col)
                            ws_sobres.cell(row=r1, column=col).alignment = center

                        self._log(f"Dato copiado en SOBRES_2026 (Filas {r1}-{r2}) → Par COMPARTE")
                    else:
                        self._log(f"Dato copiado en SOBRES_2026 (Fila {r1})")

                # Auto-ajustar alto de filas para que el texto entre
                for row_idx in range(1, ws_sobres.max_row + 1):
                    ws_sobres.row_dimensions[row_idx].height = 20

                self._guardar_excel_seguro(wb_sobres, ruta_sobres)

                # Post-procesado: completar B/L faltantes buscando por Carpeta
                bls_completados = self._completar_bl_por_carpeta(wb_sobres, ws_sobres)
                if bls_completados > 0:
                    self._guardar_excel_seguro(wb_sobres, ruta_sobres)
                    self._log(f"  → {bls_completados} B/L completados por coincidencia de Carpeta.")

                self._log("COMPLETADO: SOBRES_2026.xlsx guardada correctamente.")
                resultados["sobres"] = {"ok": True, "detalle": f"{len(datos_extraidos)} registros cargados"}
            except Exception as e:
                self._log(f"Error al guardar SOBRES_2026.xlsx: {e}")
                resultados["sobres"] = {"ok": False, "detalle": f"Error: {str(e)[:60]}"}

        if carpetas_encontradas == 0:
            self._log("No se detectaron carpetas con archivos de CONTENEDORES.")

        self.datos_planillas = datos_extraidos
        self._log(f"Escaneo finalizado. {len(datos_extraidos)} registros extraídos.")

        # ═══════════════════════════════════════════════════════════════
        # COMPLETAR PLANILLA DE COBRO
        # ═══════════════════════════════════════════════════════════════
        if datos_extraidos and (not tipos or "cobro" in tipos):
            resultados["cobro"] = self._completar_cobro(datos_extraidos)

        # ═══════════════════════════════════════════════════════════════
        # COMPLETAR PLANILLA PC (PRECINTOS/CABLES)
        # ═══════════════════════════════════════════════════════════════
        if datos_extraidos and (not tipos or "pc" in tipos):
            resultados["pc"] = self._completar_pc(datos_extraidos)

        # Guardar resultados para el popup (se muestra en el hilo principal)
        self._resultados_pendientes = resultados

    def _completar_cobro(self, datos_extraidos):
        """Llena la planilla COBRO_2026.xlsx con los datos extraídos."""
        ruta_cobro = buscar_archivo_en_pendrive(
            "COBRO_2026.xlsx", self._cfg_obtener_rutas("cobro", os.path.join("TRABAJO", "01_PLANILLAS")))

        if not ruta_cobro:
            self._log("ERROR: No se encontró COBRO_2026.xlsx.")
            pendrive_path = os.path.join("PENDRIVE:\\", "TRABAJO", "01_PLANILLAS", "COBRO_2026.xlsx")
            self.after(0, lambda: messagebox.showwarning(
                "Planilla COBRO no encontrada",
                f"No se encontró COBRO_2026.xlsx.\n\n"
                f"Ruta esperada en el pendrive:\n"
                f"  {pendrive_path}\n\n"
                f"Conecte el pendrive y vuelva a ejecutar el análisis."
            ))
            return {"ok": False, "detalle": "Archivo no encontrado"}

        wb_cobro = None
        ws_cobro = None
        try:
            wb_cobro = self._abrir_excel_seguro(ruta_cobro)
            meses_es = {
                1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
                5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
                9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
            }
            mes_actual = meses_es[datetime.now().month]
            for sheet in wb_cobro.sheetnames:
                if mes_actual in sheet.upper():
                    ws_cobro = wb_cobro[sheet]
                    break
            if not ws_cobro:
                self._log(f"ERROR: No se encontró hoja del mes '{mes_actual}' en COBRO_2026.xlsx.")
                self.after(0, lambda m=mes_actual: messagebox.showerror(
                    "Hoja del mes no encontrada",
                    f"No se encontró la hoja \"{m}\" en COBRO_2026.xlsx.\n\n"
                    f"Verifique que el archivo contenga una hoja con el nombre del mes actual.\n\n"
                    f"Operación cancelada."
                ))
                wb_cobro.close()
                return {"ok": False, "detalle": f"Hoja '{mes_actual}' no encontrada"}
            self._log(f"COBRO_2026.xlsx detectada → hoja '{ws_cobro.title}'")
        except Exception as e:
            self._log(f"ERROR al abrir COBRO_2026.xlsx: {e}")
            self.after(0, lambda: messagebox.showerror(
                "Error al abrir COBRO",
                f"No se pudo abrir COBRO_2026.xlsx:\n{e}\n\n"
                f"Verifique que el archivo no esté abierto por otro programa."
            ))
            if wb_cobro:
                try:
                    wb_cobro.close()
                except Exception:
                    pass
            return {"ok": False, "detalle": f"Error al abrir: {str(e)[:40]}"}

        try:
            from openpyxl.styles import Border, Side, Alignment as Aln
            thin = Side(style="thin")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            center = Aln(horizontal="center", vertical="center")

            # Ordenar por fecha
            def sort_date(item):
                try:
                    p = item["f_celda"].split("/")
                    return datetime(int(p[2]), int(p[1]), int(p[0]))
                except Exception:
                    return datetime.max

            datos_ordenados = sorted(datos_extraidos, key=sort_date)

            # Pre-escanear fechas que YA existen en COBRO
            fechas_existentes = set()
            for row in range(3, ws_cobro.max_row + 1):
                val_fecha = ws_cobro.cell(row=row, column=1).value
                if val_fecha:
                    fechas_existentes.add(formatear_fecha_excel(val_fecha))

            fecha_anterior = None

            # Agrupar pares comparte por carpeta (ambas con "COMPARTIDO")
            def _pares_cobro(datos):
                usados = set()
                items = []
                for i, d1 in enumerate(datos):
                    if i in usados: continue
                    f1 = d1.get("source_folder", "")
                    if "COMPARTIDO" not in f1.upper():
                        items.append(d1); usados.add(i); continue
                    pareja = None
                    for j, d2 in enumerate(datos):
                        if j == i or j in usados: continue
                        f2 = d2.get("source_folder", "")
                        if ("COMPARTIDO" in f2.upper()
                                and d1.get("dest_celda") == d2.get("dest_celda")
                                and d1.get("f_celda") == d2.get("f_celda")):
                            pareja = j; break
                    if pareja is not None:
                        items.append((d1, datos[pareja]))
                        usados.add(i); usados.add(pareja)
                    else:
                        items.append(d1); usados.add(i)
                return items

            items = _pares_cobro(datos_ordenados)
            for item in items:
                es_par = isinstance(item, tuple) and len(item) == 2
                if es_par:
                    d1, d2 = item
                    self._log(f"  🔀 Par COMPARTE COBRO: {d1.get('pe_recortado')} + {d2.get('pe_recortado')}")
                else:
                    d1 = item
                    d2 = None

                f_celda = d1["f_celda"]
                if f_celda != fecha_anterior:
                    fecha_anterior = f_celda
                    es_primero_del_dia = f_celda not in fechas_existentes
                else:
                    es_primero_del_dia = False

                precio_base = int(self._cfg_obtener("valores", "precio_carpeta", 49000))
                precio_carpeta = precio_base if es_primero_del_dia else precio_base // 2
                tipo_cobro = d1.get("tipo_transporte", "TERRESTRE")
                es_maritimo_cobro = tipo_cobro in ("ISO", "FLEXI")
                servicio_ata = "-" if es_maritimo_cobro else d1["cant_final"] * int(self._cfg_obtener("valores", "ata_tares", 65000))

                valores_fila = {
                    1: f_celda, 2: d1["cant_final"], 3: tipo_cobro, 4: d1["pe_recortado"],
                    5: d1["carp_celda"], 6: "-" if es_maritimo_cobro else d1["frac_carpeta"],
                    7: servicio_ata, 8: precio_carpeta,
                }
                valores_check = {c: v for c, v in valores_fila.items() if c != 8}
                if ya_existe_en_hoja(ws_cobro, valores_check):
                    self._log(f"  Omitido (ya existe en COBRO): Carpeta {d1.get('carp_celda','?')} | {f_celda}")
                    continue

                fechas_existentes.add(f_celda)
                r1 = primera_fila_libre(ws_cobro)

                for col, val in valores_fila.items():
                    cell = ws_cobro.cell(row=r1, column=col)
                    cell.value = val
                    cell.border = border
                    cell.alignment = center

                if es_par:
                    # Segunda fila con datos del otro Contenedores
                    precio_carpeta_2 = precio_base // 2
                    tipo2_cobro = d2.get("tipo_transporte", "TERRESTRE")
                    es_maritimo2_cobro = tipo2_cobro in ("ISO", "FLEXI")
                    servicio_ata_2 = "-" if es_maritimo2_cobro else d2["cant_final"] * int(self._cfg_obtener("valores", "ata_tares", 65000))
                    r2 = r1 + 1

                    vals2 = {1: f_celda, 2: d2["cant_final"], 3: tipo2_cobro, 4: d2["pe_recortado"],
                             5: d2["carp_celda"], 6: "-" if es_maritimo2_cobro else d2["frac_carpeta"],
                             7: servicio_ata_2, 8: precio_carpeta_2}
                    for col, val in vals2.items():
                        cell = ws_cobro.cell(row=r2, column=col)
                        cell.value = val
                        cell.border = border
                        cell.alignment = center

                    # Merge fecha (col 1)
                    ws_cobro.merge_cells(start_row=r1, start_column=1, end_row=r2, end_column=1)
                    ws_cobro.cell(row=r1, column=1).alignment = center
                    # Merge Cantidad (col 2), Tipo (col 3) y Servicio ATA (col 7)
                    for col in (2, 3, 7):
                        ws_cobro.merge_cells(start_row=r1, start_column=col, end_row=r2, end_column=col)
                        ws_cobro.cell(row=r1, column=col).alignment = center

                    self._log(f"Dato copiado en COBRO_2026 (Filas {r1}-{r2}) → Par COMPARTE")
                else:
                    # Merge de fecha normal
                    if not es_primero_del_dia:
                        r_prev = r1 - 1
                        val_prev = None
                        while r_prev >= 3:
                            val_prev = ws_cobro.cell(row=r_prev, column=1).value
                            if val_prev is not None:
                                break
                            r_prev -= 1
                        if val_prev and formatear_fecha_excel(val_prev) == f_celda:
                            for m_range in list(ws_cobro.merged_cells.ranges):
                                if m_range.min_col == 1 and m_range.max_col == 1:
                                    if (m_range.min_row <= r1 <= m_range.max_row
                                            or m_range.min_row <= r_prev <= m_range.max_row):
                                        ws_cobro.merged_cells.ranges.remove(m_range)
                            ws_cobro.merge_cells(start_row=r_prev, start_column=1, end_row=r1, end_column=1)
                            ws_cobro.cell(row=r_prev, column=1).alignment = center

                    self._log(f"Dato copiado en COBRO_2026 (Fila {r1}) → Carpetas: ${precio_carpeta:,}")

            self._guardar_excel_seguro(wb_cobro, ruta_cobro)
            self._log("COMPLETADO: COBRO_2026.xlsx guardada correctamente.")
            return {"ok": True, "detalle": f"{len(datos_extraidos)} registros cargados"}
        except Exception as e:
            self._log(f"ERROR al guardar COBRO_2026.xlsx: {e}")
            self.after(0, lambda: messagebox.showerror(
                "Error al guardar COBRO",
                f"No se pudo guardar COBRO_2026.xlsx:\n{e}\n\n"
                f"Verifique que el archivo no esté abierto por otro programa."
            ))
            return {"ok": False, "detalle": f"Error: {str(e)[:60]}"}
        finally:
            if wb_cobro:
                try:
                    wb_cobro.close()
                except Exception:
                    pass

    # ── Precintos disponibles ──────────────────────────────────────────
    def _contar_precintos_disponibles(self):
        """Cuenta precintos SIN ASIGNAR en PC.xlsx (col 1 llena, col 2 vacía).
        Retorna el número o None si no encuentra la planilla."""
        ruta_pc = buscar_archivo_en_pendrive(
            "PC.xlsx", self._cfg_obtener_rutas("pc", os.path.join("TRABAJO", "01_PLANILLAS")))
        if not ruta_pc:
            ruta_pc = buscar_archivo_en_pendrive(
                "PC_2026.xlsx", self._cfg_obtener_rutas("pc", os.path.join("TRABAJO", "01_PLANILLAS")))
        if not ruta_pc:
            return None
        try:
            wb = openpyxl.load_workbook(ruta_pc)
            try:
                ws = wb["SIN ASIGNACION"]
            except KeyError:
                ws = wb.active
            disponibles = 0
            for fila in range(2, ws.max_row + 1):
                nro = ws.cell(row=fila, column=1).value
                if nro and not ws.cell(row=fila, column=2).value:
                    disponibles += 1
            wb.close()
            return disponibles
        except Exception:
            return None

    def _actualizar_titulo_precintos(self):
        """Actualiza lbl_titulo_panel con el conteo de precintos disponibles."""
        disponibles = self._contar_precintos_disponibles()
        if disponibles is None:
            self.lbl_titulo_panel.configure(
                text="Completar Planillas — No se encontró planilla PC")
        else:
            self.lbl_titulo_panel.configure(
                text=f"Completar Planillas — Precintos disponibles: {disponibles}")

    def _completar_pc(self, datos_extraidos):
        """Llena la planilla PC (Precintos/Cables) en la hoja SIN ASIGNACION.
        Busca cada precinto de abajo hacia arriba y completa: Fecha, P.E., Guarda, Carpeta."""
        ruta_pc = buscar_archivo_en_pendrive(
            "PC.xlsx", self._cfg_obtener_rutas("pc", os.path.join("TRABAJO", "01_PLANILLAS")))
        if not ruta_pc:
            ruta_pc = buscar_archivo_en_pendrive(
                "PC_2026.xlsx", self._cfg_obtener_rutas("pc", os.path.join("TRABAJO", "01_PLANILLAS")))
        if not ruta_pc:
            self._log("ERROR: No se encontró PC.xlsx.")
            pendrive_path = os.path.join("PENDRIVE:\\", "TRABAJO", "01_PLANILLAS", "PC.xlsx")
            self.after(0, lambda: messagebox.showwarning(
                "Planilla PC no encontrada",
                f"No se encontró PC.xlsx.\n\n"
                f"Ruta esperada en el pendrive:\n"
                f"  {pendrive_path}\n\n"
                f"Conecte el pendrive y vuelva a ejecutar el análisis."
            ))
            return {"ok": False, "detalle": "Archivo no encontrado"}

        wb_pc = None
        ws_pc = None
        try:
            wb_pc = self._abrir_excel_seguro(ruta_pc)
            # Buscar hoja SIN ASIGNACION
            for sheet in wb_pc.sheetnames:
                if "SIN ASIGNACION" in sheet.upper() or "SINASIGNACION" in sheet.upper():
                    ws_pc = wb_pc[sheet]
                    break
            if not ws_pc:
                # Intentar la primera hoja como fallback
                ws_pc = wb_pc.active
                self._log(f"PC.xlsx: hoja 'SIN ASIGNACION' no encontrada, usando '{ws_pc.title}'")
            else:
                self._log(f"PC.xlsx detectada → hoja '{ws_pc.title}'")
        except Exception as e:
            self._log(f"ERROR al abrir PC.xlsx: {e}")
            self.after(0, lambda: messagebox.showerror(
                "Error al abrir PC",
                f"No se pudo abrir PC.xlsx:\n{e}"
            ))
            if wb_pc:
                try:
                    wb_pc.close()
                except Exception:
                    pass
            return {"ok": False, "detalle": f"Error al abrir: {str(e)[:40]}"}

        try:
            from openpyxl.styles import PatternFill, Border, Side, Alignment as Aln

            # Color de resalte para precintos encontrados (dorado claro)
            fill_encontrado = PatternFill(fill_type="solid", fgColor="FFD966")
            thin = Side(style="thin")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            center = Aln(horizontal="center", vertical="center")

            # Agrupar pares comparte por carpeta (misma lógica que SOBRES/COBRO)
            def _pares_pc(datos):
                usados = set()
                items = []
                for i, d1 in enumerate(datos):
                    if i in usados: continue
                    f1 = d1.get("source_folder", "")
                    if "COMPARTIDO" not in f1.upper():
                        items.append(d1); usados.add(i); continue
                    pareja = None
                    for j, d2 in enumerate(datos):
                        if j == i or j in usados: continue
                        f2 = d2.get("source_folder", "")
                        if ("COMPARTIDO" in f2.upper()
                                and d1.get("dest_celda") == d2.get("dest_celda")
                                and d1.get("f_celda") == d2.get("f_celda")):
                            pareja = j; break
                    if pareja is not None:
                        items.append((d1, datos[pareja]))
                        usados.add(i); usados.add(pareja)
                    else:
                        items.append(d1); usados.add(i)
                return items

            # Recolectar precintos, manejando pares comparte
            todos_precintos = []
            items_pc = _pares_pc(datos_extraidos)
            for item in items_pc:
                es_par = isinstance(item, tuple) and len(item) == 2
                if es_par:
                    d1, d2 = item
                    # d1 = comparte, d2 = cierra (o viceversa). Usar precintos del que NO tiene CERRADO? No, del que cierra.
                    # Identificar cuál es cierra: el que tiene "CERRADO" en source_folder
                    d_cierra = d1 if "CERRADO" in d1.get("source_folder", "").upper() else d2
                    d_comparte = d1 if d_cierra is d2 else d2

                    pe_cierra = d_cierra.get("pe_completo", d_cierra.get("pe_recortado", ""))
                    pe_comparte = d_comparte.get("pe_completo", d_comparte.get("pe_recortado", ""))
                    pe_comparte_5 = pe_comparte.strip()[-5:] if pe_comparte else ""
                    pe_combinado = f"{pe_cierra}/{pe_comparte_5}" if pe_cierra and pe_comparte_5 else pe_cierra
                    carp_combinado = f"{d_cierra.get('carp_celda','')}/{d_comparte.get('carp_celda','')}"

                    # Crear un dato combinado con los precintos del cierra
                    dato_combinado = dict(d_cierra)
                    dato_combinado["pe_completo"] = pe_combinado
                    dato_combinado["carp_celda"] = carp_combinado
                    dato_combinado["_es_par_comparte"] = True

                    for p in d_cierra.get("precintos", []):
                        if p:
                            todos_precintos.append((p, dato_combinado))
                    self._log(f"  PC Par COMPARTE: PE={pe_combinado} | Carpeta={carp_combinado}")
                else:
                    for p in item.get("precintos", []):
                        if p:
                            todos_precintos.append((p, item))

            if not todos_precintos:
                self._log("PC: No se encontraron precintos en los archivos de carga.")
                return {"ok": False, "detalle": "Sin precintos en archivos"}

            encontrados = 0

            for precinto, dato in todos_precintos:
                # Buscar de abajo hacia arriba (pocos sin asignar al final)
                encontrado = False
                for row in range(ws_pc.max_row, 0, -1):
                    val_celda = ws_pc.cell(row=row, column=1).value
                    if val_celda and str(val_celda).strip() == precinto:
                        # Completar datos: Fecha, P.E., Guarda, Carpeta (columnas 2-5)
                        ws_pc.cell(row=row, column=2).value = dato["f_celda"]
                        ws_pc.cell(row=row, column=3).value = dato.get("pe_completo", dato["pe_recortado"])
                        ws_pc.cell(row=row, column=4).value = dato.get("guarda", "")
                        ws_pc.cell(row=row, column=5).value = dato["carp_celda"]

                        # Aplicar formato y color a toda la fila
                        for c in range(1, 6):
                            cell = ws_pc.cell(row=row, column=c)
                            cell.fill = fill_encontrado
                            cell.border = border
                            cell.alignment = center

                        encontrados += 1
                        encontrado = True
                        self._log(f"  PC: Precinto {precinto} → Fila {row} | P.E. {dato.get('pe_completo', dato['pe_recortado'])} | Carpeta {dato['carp_celda']} | Guarda {dato.get('guarda','')}")
                        break

                # Si no se encontró, simplemente seguimos sin loguear (ya fue asignado antes)

            # Auto-ajustar alto de filas
            for row_idx in range(1, ws_pc.max_row + 1):
                ws_pc.row_dimensions[row_idx].height = 20
            self._guardar_excel_seguro(wb_pc, ruta_pc)
            self._log(f"COMPLETADO: PC.xlsx — {encontrados} precintos asignados.")
            return {"ok": True, "detalle": f"{encontrados} precintos asignados"}
        except Exception as e:
            self._log(f"ERROR al guardar PC.xlsx: {e}")
            self.after(0, lambda: messagebox.showerror(
                "Error al guardar PC",
                f"No se pudo guardar PC.xlsx:\n{e}\n\n"
                f"Verifique que el archivo no esté abierto por otro programa."
            ))
            return {"ok": False, "detalle": f"Error: {str(e)[:60]}"}
        finally:
            if wb_pc:
                try:
                    wb_pc.close()
                except Exception:
                    pass

    def _leer_xls_antiguo(self, ruta_excel):
        book = xlrd.open_workbook(ruta_excel)
        sheet_name = next(
            (s for s in book.sheet_names() if "CHOFER" in s.upper()), None
        )
        if not sheet_name:
            raise ValueError(f"Falta pestaña 'CHOFER'. Hojas: {book.sheet_names()}")
        sheet = book.sheet_by_name(sheet_name)

        f_celda = pe_celda = carp_celda = dest_celda = bl_celda = ""
        transporte_encontrado = primera_patente = guarda = ""
        puerto_salida = peso_flexi = ""
        col_patente = None

        for r in range(sheet.nrows):
            time.sleep(0.01) # Ceder GIL
            for c in range(sheet.ncols):
                val = sheet.cell_value(r, c)
                if val and isinstance(val, str):
                    val_up = val.strip().upper()
                    if "FECHA CARGA" in val_up or "FECHA DE CARGA" in val_up or val_up == "FECHA":
                        nxt = sheet.cell_value(r, c + 1) if c + 1 < sheet.ncols else None
                        if nxt and not f_celda:
                            if sheet.cell_type(r, c + 1) == xlrd.XL_CELL_DATE:
                                nxt = xlrd.xldate_as_datetime(nxt, book.datemode)
                            f_celda = formatear_fecha_excel(nxt)
                    elif val_up in ("P.E", "P.E.", "PE", "PERMISO"):
                        nxt = sheet.cell_value(r, c + 1) if c + 1 < sheet.ncols else None
                        if nxt:
                            pe_celda = str(nxt).strip()
                    elif "CARPETA" in val_up or val_up == "CARP.":
                        nxt = sheet.cell_value(r, c + 1) if c + 1 < sheet.ncols else None
                        if nxt and not carp_celda:
                            carp_celda = str(int(nxt)) if isinstance(nxt, float) else str(nxt).strip()
                    elif "DESTINATARIO" in val_up or "CLIENTE" in val_up:
                        nxt = sheet.cell_value(r, c + 1) if c + 1 < sheet.ncols else None
                        if nxt and not dest_celda:
                            dest_celda = str(nxt).strip().upper()
                    elif val_up in ("BL", "B/L", "CONOCIMIENTO"):
                        nxt = sheet.cell_value(r, c + 1) if c + 1 < sheet.ncols else None
                        if nxt and not bl_celda:
                            bl_celda = str(nxt).strip()
                    elif "PUERTO SALIDA" in val_up or "P. SALIDA" in val_up:
                        nxt = sheet.cell_value(r, c + 1) if c + 1 < sheet.ncols else None
                        if nxt and not puerto_salida:
                            puerto_salida = str(nxt).strip()
                    elif "PESO FLEXI" in val_up:
                        nxt = sheet.cell_value(r, c + 1) if c + 1 < sheet.ncols else None
                        if nxt != "" and nxt is not None and not peso_flexi:
                            # xlrd: float → int si es entero, sino string
                            if isinstance(nxt, float):
                                peso_flexi = str(int(nxt)) if nxt.is_integer() else str(nxt)
                            else:
                                peso_flexi = str(nxt).strip()
                    if "PATENTE" in val_up:
                        col_patente = c
                    # La columna de precintos no se busca más por etiqueta, se toma directo de col 1
                    if "TRANSPORT" in val_up or "EMPRESA" in val_up:
                        # Puede haber hasta 4 transportes en grilla 2x2 a la derecha
                        transportes = []
                        for dr in (0, 1):
                            for dc in (1, 2):
                                if r + dr < sheet.nrows and c + dc < sheet.ncols:
                                    val_t = sheet.cell_value(r + dr, c + dc)
                                    if val_t:
                                        transportes.append(str(val_t).strip())
                        # Eliminar duplicados preservando orden
                        unicos = list(dict.fromkeys(transportes))
                        transporte_encontrado = " - ".join(unicos) if unicos else ""
                    if "GUARDA" in val_up or "CHOFER" in val_up:
                        if c + 1 < sheet.ncols:
                            val_g = sheet.cell_value(r, c + 1)
                            if val_g:
                                guarda = str(val_g).strip()
                    if val_up == "GUARDA" and not guarda:
                        nxt = sheet.cell_value(r, c + 1) if c + 1 < sheet.ncols else None
                        if nxt:
                            guarda = str(nxt).strip()

        if col_patente is not None:
            for row_p in range(sheet.nrows):
                val_p = sheet.cell_value(row_p, col_patente)
                if val_p and not any(
                    k in str(val_p).upper() for k in ("PATENTE", "TOTAL", "CHOFER")
                ):
                    primera_patente = str(val_p).strip()
                    break

        # Precintos: SIEMPRE columna 1 (A), desde fila 2 (xlrd: row>=1)
        precintos = []
        for row_pr in range(1, sheet.nrows):
            val_pr = sheet.cell_value(row_pr, 0)  # columna 0 = A
            if val_pr:
                val_str = str(val_pr).strip()
                palabras = val_str.split()
                es_nombre_largo = len(palabras) >= 3 or len(val_str) > 30
                if len(val_str) >= 2 and val_str != "-" and not es_nombre_largo:
                    # Blacklist de palabras que son etiquetas/headers, no precintos reales
                    labels_prohibidas = (
                        "PRECINTO", "SELLO", "TOTAL", "CHOFER", "GUARDA", "ADUANA",
                        "PUERTO", "CARPETA", "BOOKING", "TRANSPORTE", "PESO", "FLEXI",
                        "FECHA", "DESTINATARIO", "TERMINAL", "PATENTE", "PERMISO",
                        "EMPRESA", "SALIDA", "CARGA", "EMBARQUE", "CONTENEDOR",
                    )
                    if not any(k in val_str.upper() for k in labels_prohibidas):
                        # Separar precintos múltiples unidos por guion (ej: "12345-67890")
                        for parte in val_str.split("-"):
                            parte = parte.strip()
                            if parte:
                                precintos.append(parte)
        # Si no se encontró B/L, buscar por número de carpeta en la misma hoja
        if not bl_celda and carp_celda:
            bl_celda = buscar_bl_por_carpeta_xls(sheet, carp_celda)
        return f_celda, pe_celda, carp_celda, dest_celda, bl_celda, transporte_encontrado, primera_patente, precintos, guarda, puerto_salida, peso_flexi

    def _completar_bl_por_carpeta(self, wb, ws_actual):
        """Recorre las filas de la hoja SOBRES. Si B/L (col 9) es '-', busca
           en filas de arriba (misma hoja) u hoja anterior otra fila con la
           misma Carpeta (col 5) que tenga B/L válido, y lo copia.
           Omite filas con Fracción F1 (col 10), que no llevan B/L."""
        def normalizar(val):
            if val is None:
                return ""
            if isinstance(val, float):
                if val == int(val):
                    return str(int(val))
                return str(val).strip()
            s = str(val).strip().lstrip("'")
            return s

        completados = 0
        nombre_actual = ws_actual.title
        idx_actual = wb.sheetnames.index(nombre_actual)

        def recolectar_datos(ws):
            filas = []  # [(r, carpeta, bl, fraccion)]
            for r in range(3, ws.max_row + 1):
                carp = ws.cell(row=r, column=5).value
                bl = ws.cell(row=r, column=9).value
                if carp is not None:
                    frac = normalizar(ws.cell(row=r, column=10).value)
                    filas.append((r, normalizar(carp), normalizar(bl), frac))
            return filas

        filas_actual = recolectar_datos(ws_actual)

        filas_anterior = []
        if idx_actual > 0:
            ws_ant = wb[wb.sheetnames[idx_actual - 1]]
            filas_anterior = recolectar_datos(ws_ant)

        for r, carpeta, bl, fraccion in filas_actual:
            if bl and bl not in ("-", "—"):
                continue
            if not carpeta or carpeta == "0":
                continue
            # Fracción F1 no lleva B/L
            if fraccion.upper() == "F1":
                continue

            bl_encontrado = ""
            for r2, carp2, bl2, _ in filas_actual:
                if r2 >= r:
                    continue
                if carp2 == carpeta and bl2 and bl2 not in ("-", "—"):
                    bl_encontrado = bl2
                    break

            if not bl_encontrado and filas_anterior:
                for r2, carp2, bl2, _ in filas_anterior:
                    if carp2 == carpeta and bl2 and bl2 not in ("-", "—"):
                        bl_encontrado = bl2
                        break

            if bl_encontrado:
                from openpyxl.styles import Border, Side, Alignment as Aln, Font
                thin = Side(style="thin")
                border = Border(left=thin, right=thin, top=thin, bottom=thin)
                center = Aln(horizontal="center", vertical="center", wrap_text=True)
                font_bl = Font(name="Calibri", size=9)
                cell_bl = ws_actual.cell(row=r, column=9)
                cell_bl.value = bl_encontrado
                cell_bl.number_format = 'General'
                cell_bl.border = border
                cell_bl.alignment = center
                cell_bl.font = font_bl
                completados += 1

        if completados > 0:
            self._log(f"  → {completados} B/L completados por coincidencia de Carpeta.")
        return completados

    def _leer_xlsx_moderno(self, ruta_excel):
        wb = openpyxl.load_workbook(ruta_excel, data_only=True)
        pestana_chofer = next(
            (s for s in wb.sheetnames if "CHOFER" in s.upper()), None
        )
        if not pestana_chofer:
            wb.close()
            raise ValueError(f"Falta pestaña 'CHOFER'. Hojas: {wb.sheetnames}")
        ws = wb[pestana_chofer]

        f_celda = pe_celda = carp_celda = dest_celda = bl_celda = ""
        transporte_encontrado = primera_patente = guarda = ""
        puerto_salida = peso_flexi = ""
        col_patente = None

        for r in range(1, ws.max_row + 1):
            time.sleep(0.01) # Ceder GIL
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=r, column=c).value
                if val and isinstance(val, str):
                    val_up = val.strip().upper()
                    if "FECHA CARGA" in val_up or "FECHA DE CARGA" in val_up or val_up == "FECHA":
                        nxt = ws.cell(row=r, column=c + 1).value
                        if nxt and not f_celda:
                            f_celda = formatear_fecha_excel(nxt)
                    elif val_up in ("P.E", "P.E.", "PE", "PERMISO"):
                        nxt = ws.cell(row=r, column=c + 1).value
                        if nxt:
                            pe_celda = str(nxt).strip()
                    elif "CARPETA" in val_up or val_up == "CARP.":
                        nxt = ws.cell(row=r, column=c + 1).value
                        if nxt and not carp_celda:
                            carp_celda = str(nxt).strip()
                    elif "DESTINATARIO" in val_up or "CLIENTE" in val_up:
                        nxt = ws.cell(row=r, column=c + 1).value
                        if nxt and not dest_celda:
                            dest_celda = str(nxt).strip().upper()
                    elif val_up in ("BL", "B/L", "CONOCIMIENTO"):
                        nxt = ws.cell(row=r, column=c + 1).value
                        if nxt and not bl_celda:
                            bl_celda = str(nxt).strip()
                    elif "PUERTO SALIDA" in val_up or "P. SALIDA" in val_up:
                        nxt = ws.cell(row=r, column=c + 1).value
                        if nxt and not puerto_salida:
                            puerto_salida = str(nxt).strip()
                    elif "PESO FLEXI" in val_up:
                        nxt = ws.cell(row=r, column=c + 1).value
                        if nxt != "" and nxt is not None and not peso_flexi:
                            # openpyxl: float → int si es entero, sino string
                            if isinstance(nxt, float):
                                peso_flexi = str(int(nxt)) if nxt.is_integer() else str(nxt)
                            else:
                                peso_flexi = str(nxt).strip()
                    if "PATENTE" in val_up:
                        col_patente = c
                    # La columna de precintos no se busca más por etiqueta, se toma directo de col 1
                    if "TRANSPORT" in val_up or "EMPRESA" in val_up:
                        # Puede haber hasta 4 transportes en grilla 2x2 a la derecha
                        transportes = []
                        for dr in (0, 1):
                            for dc in (1, 2):
                                val_t = ws.cell(row=r + dr, column=c + dc).value
                                if val_t:
                                    transportes.append(str(val_t).strip())
                        # Eliminar duplicados preservando orden
                        unicos = list(dict.fromkeys(transportes))
                        transporte_encontrado = " - ".join(unicos) if unicos else ""
                    if "GUARDA" in val_up or "CHOFER" in val_up:
                        val_g = ws.cell(row=r, column=c + 1).value
                        if val_g:
                            guarda = str(val_g).strip()

        if col_patente is not None:
            for row_p in range(1, ws.max_row + 1):
                val_p = ws.cell(row=row_p, column=col_patente).value
                if val_p and not any(
                    k in str(val_p).upper() for k in ("PATENTE", "TOTAL", "CHOFER")
                ):
                    primera_patente = str(val_p).strip()
                    break

        # Precintos: SIEMPRE columna 1 (A), desde fila 2 (openpyxl: row>=2)
        precintos = []
        for row_pr in range(2, ws.max_row + 1):
            val_pr = ws.cell(row=row_pr, column=1).value  # columna 1 = A
            if val_pr:
                val_str = str(val_pr).strip()
                palabras = val_str.split()
                es_nombre_largo = len(palabras) >= 3 or len(val_str) > 30
                if len(val_str) >= 2 and val_str != "-" and not es_nombre_largo:
                    # Blacklist de palabras que son etiquetas/headers, no precintos reales
                    labels_prohibidas = (
                        "PRECINTO", "SELLO", "TOTAL", "CHOFER", "GUARDA", "ADUANA",
                        "PUERTO", "CARPETA", "BOOKING", "TRANSPORTE", "PESO", "FLEXI",
                        "FECHA", "DESTINATARIO", "TERMINAL", "PATENTE", "PERMISO",
                        "EMPRESA", "SALIDA", "CARGA", "EMBARQUE", "CONTENEDOR",
                    )
                    if not any(k in val_str.upper() for k in labels_prohibidas):
                        # Separar precintos múltiples unidos por guion (ej: "12345-67890")
                        for parte in val_str.split("-"):
                            parte = parte.strip()
                            if parte:
                                precintos.append(parte)
        # Si no se encontró B/L, buscar por número de carpeta en la hoja actual y la anterior
        if not bl_celda and carp_celda:
            bl_celda = buscar_bl_por_carpeta_xlsx(wb, pestana_chofer, carp_celda)

        wb.close()
        return f_celda, pe_celda, carp_celda, dest_celda, bl_celda, transporte_encontrado, primera_patente, precintos, guarda, puerto_salida, peso_flexi

    def _agregar_fila_planillas(self, registro):
        """Agrega una fila a la tabla de planillas (thread-safe, en hilo principal)."""
        dest = (registro.get("dest_celda") or "").upper()

        if "VITAPRO" in dest:
            tag = "vitapro"
        elif "EWOS" in dest:
            tag = "ewos"
        elif "DICOAL" in dest:
            tag = "dicoal"
        elif "NUTRECO" in dest:
            tag = "nutreco"
        elif "CARGILL" in dest:
            tag = "cargill"
        elif "BIOMAR" in dest:
            tag = "biomar"
        else:
            tag = ""

        valores = (
            registro["f_celda"],
            registro["cant_final"],
            registro["pe_recortado"],
            registro["carp_celda"],
            registro["terminal"],
            registro["trans_final"],
            registro["dest_celda"],
            registro["bl_final"],
            registro["frac_carpeta"],
            "-" if registro["servicio"] == "-" else f'${registro["servicio"]:,}',
        )
        self.tree_planillas.insert("", "end", values=valores, tags=(tag,) if tag else ())

        # Actualizar resumen
        count = len(self.tree_planillas.get_children())
        total_cant = sum(
            int(self.tree_planillas.item(child, "values")[1])
            for child in self.tree_planillas.get_children()
        )
        self.lbl_resumen_planillas.configure(
            text=f"{count} registros  ·  {total_cant} contenedores totales"
        )

    def _planillas_error(self, error_msg):
        self._log(f"ERROR CRÍTICO: {error_msg}")
        messagebox.showerror(
            "Error en Agente de Contenedores",
            f"Ocurrió un error durante el análisis:\n\n{error_msg}",
        )

    def _planillas_done(self):
        self.tarea_activa = False
        try:
            self.btn_ejecutar_planillas.configure(
                text="▶  Completar Planillas",
                state="normal",
                fg_color=Palette.ACCENT,
            )
            self.progress_planillas.stop()
            self.progress_planillas.set(1)
            self.lbl_estado_planillas.configure(
                text="Análisis completado", text_color=Palette.SUCCESS
            )
        except (AttributeError, Exception):
            pass
        self._set_status("Análisis de planillas finalizado")

        # Actualizar conteo de precintos disponibles
        self._actualizar_titulo_precintos()

        # Mostrar popup de resumen si hay resultados pendientes
        if hasattr(self, "_resultados_pendientes") and self._resultados_pendientes:
            try:
                self._mostrar_resumen(self._resultados_pendientes)
            except Exception:
                pass
            self._resultados_pendientes = None

    # ═══════════════════════════════════════════════════════════════════
    # AGENTE 3: CORREOS EN HILO
    # ═══════════════════════════════════════════════════════════════════
    def _elegir_carpetas_correos(self):
        """Popup para elegir qué carpetas del escritorio enviar."""
        if self.tarea_activa:
            messagebox.showwarning("Agente ocupado", "Hay una tarea en ejecución. Espere a que finalice.")
            return

        escritorio = self._resolver_ruta("planillas_carga", "Desktop")
        try:
            items = sorted(os.listdir(escritorio))
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el escritorio:\n{e}")
            return

        carpetas_encontradas = []
        for item in items:
            ruta = os.path.join(escritorio, item)
            if not os.path.isdir(ruta):
                continue
            if item.startswith(".") or item.upper() in ("RECYCLED", "RECYCLER"):
                continue
            # Buscar archivos adjuntos válidos dentro
            try:
                archivos = os.listdir(ruta)
            except Exception:
                continue
            adjuntos = []
            año_actual = datetime.now().strftime("%y")
            for a in archivos:
                up = a.upper()
                if (up.startswith("PLT") and up.endswith(".PDF")) or \
                   (re.search(f"^{año_actual}AR", up) and up.endswith(".PDF")) or \
                   (re.search(r"^\d", a) and (up.endswith(".XLSX") or up.endswith(".XLS"))):
                    adjuntos.append(a)
            if adjuntos:
                carpetas_encontradas.append((item, ruta, adjuntos))

        if not carpetas_encontradas:
            messagebox.showinfo("Sin carpetas", "No se encontraron carpetas con archivos para enviar en el escritorio.")
            return

        # ── Popup ──────────────────────────────────────────────────
        popup = ctk.CTkToplevel(self)
        popup.title("Elegir Correos a Enviar")
        popup.geometry("600x480")
        popup.configure(fg_color=Palette.BG_CARD)
        popup.transient(self)
        popup.lift()
        popup.grab_set()
        popup.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 600) // 2
        py = self.winfo_y() + (self.winfo_height() - 480) // 2
        popup.geometry(f"600x480+{px}+{py}")

        ctk.CTkLabel(
            popup, text="Seleccioná las carpetas a enviar",
            font=ctk.CTkFont(family=FONT_FAMILY, size=15, weight="bold"),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(pady=(16, 8))

        # Frame scrolleable con checkboxes
        scroll = ctk.CTkScrollableFrame(popup, fg_color=Palette.BG_TABLE, corner_radius=8,
                                         border_width=1, border_color=Palette.BORDER)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        checks = {}
        for nombre, ruta_carp, adjuntos in carpetas_encontradas:
            var = ctk.BooleanVar(value=True)
            checks[nombre] = var
            frame_item = ctk.CTkFrame(scroll, fg_color=Palette.BG_CARD, corner_radius=6)
            frame_item.pack(fill="x", padx=4, pady=3)
            ctk.CTkCheckBox(
                frame_item, text="", variable=var, width=24,
                fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            ).pack(side="left", padx=(8, 4), pady=8)
            ctk.CTkLabel(
                frame_item, text=nombre,
                font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
                text_color=Palette.TEXT_PRIMARY, anchor="w",
            ).pack(side="left", padx=4, pady=4)
            ctk.CTkLabel(
                frame_item, text=f"{len(adjuntos)} archivos",
                font=ctk.CTkFont(family=FONT_FAMILY, size=11),
                text_color=Palette.TEXT_MUTED, anchor="w",
            ).pack(side="right", padx=12, pady=4)

        # Agregar planilla de carga terrestre al listado scrolleable (destilado por defecto)
        var_plt = ctk.BooleanVar(value=False)
        checks["_PLANILLA_DE_CARGA_TERRESTRE_"] = var_plt
        frame_item_plt = ctk.CTkFrame(scroll, fg_color=Palette.BG_CARD, corner_radius=6)
        frame_item_plt.pack(fill="x", padx=4, pady=3)
        ctk.CTkCheckBox(
            frame_item_plt, text="", variable=var_plt, width=24,
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
        ).pack(side="left", padx=(8, 4), pady=8)
        ctk.CTkLabel(
            frame_item_plt, text="Planilla de Carga (Carga Terrestre)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            text_color=Palette.TEXT_PRIMARY, anchor="w",
        ).pack(side="left", padx=4, pady=4)
        ctk.CTkLabel(
            frame_item_plt, text="Correo Grupal",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_MUTED, anchor="w",
        ).pack(side="right", padx=12, pady=4)

        # Agregar planilla de carga marítima (destildado por defecto)
        var_plt_mar = ctk.BooleanVar(value=False)
        checks["_PLANILLA_DE_CARGA_MARITIMA_"] = var_plt_mar
        frame_item_plt_mar = ctk.CTkFrame(scroll, fg_color=Palette.BG_CARD, corner_radius=6)
        frame_item_plt_mar.pack(fill="x", padx=4, pady=3)
        ctk.CTkCheckBox(
            frame_item_plt_mar, text="", variable=var_plt_mar, width=24,
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
        ).pack(side="left", padx=(8, 4), pady=8)
        ctk.CTkLabel(
            frame_item_plt_mar, text="Planillas de Carga Marítimas",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            text_color=Palette.TEXT_PRIMARY, anchor="w",
        ).pack(side="left", padx=4, pady=4)
        ctk.CTkLabel(
            frame_item_plt_mar, text="Correo Grupal",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_MUTED, anchor="w",
        ).pack(side="right", padx=12, pady=4)

        # Botones de acción
        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(fill="x", padx=16, pady=(0, 12))

        def seleccionar_todas():
            for var in checks.values():
                var.set(True)

        def deseleccionar_todas():
            for var in checks.values():
                var.set(False)

        ctk.CTkButton(
            btn_frame, text="Todas", width=70, height=30,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_HOVER, hover_color=Palette.BG_SIDEBAR,
            text_color=Palette.TEXT_SECONDARY, corner_radius=4,
            command=seleccionar_todas,
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            btn_frame, text="Ninguna", width=70, height=30,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_HOVER, hover_color=Palette.BG_SIDEBAR,
            text_color=Palette.TEXT_SECONDARY, corner_radius=4,
            command=deseleccionar_todas,
        ).pack(side="left", padx=4)

        def enviar_seleccionados():
            seleccionadas = [ruta for nombre, ruta, _ in carpetas_encontradas if checks[nombre].get()]
            incluir_plt = checks["_PLANILLA_DE_CARGA_TERRESTRE_"].get()
            incluir_plt_mar = checks["_PLANILLA_DE_CARGA_MARITIMA_"].get()
            if not seleccionadas and not incluir_plt and not incluir_plt_mar:
                messagebox.showwarning("Nada seleccionado", "Seleccioná al menos una carpeta o Planilla de Carga.")
                return
            popup.destroy()
            self._despachar_carpetas(seleccionadas, incluir_plt, incluir_plt_mar)

        ctk.CTkButton(
            btn_frame, text="📤  Enviar Seleccionados", width=200, height=34,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6,
            command=enviar_seleccionados,
        ).pack(side="right", padx=4)

    def _despachar_carpetas(self, rutas_seleccionadas, incluir_planilla=False, incluir_planilla_mar=False):
        """Procesa solo las carpetas elegidas por el usuario."""
        if self.tarea_activa:
            return
        self.tarea_activa = True
        self.btn_elegir_correos.configure(text="⏳  Procesando...", state="disabled")
        self.btn_ejecutar_correos.configure(state="disabled")
        self.progress_correos.configure(mode="indeterminate")
        self.progress_correos.start()
        self.lbl_estado_correos.configure(text="Procesando carpetas seleccionadas...", text_color=Palette.INFO)
        self._set_status("Despachando correos seleccionados...")

        self._limpiar_log()
        for row in self.tree_correos.get_children():
            self.tree_correos.delete(row)

        t = threading.Thread(
            target=lambda: self._agente_correos_worker(
                carpetas_filtro=rutas_seleccionadas,
                incluir_planilla=incluir_planilla,
                incluir_planilla_mar=incluir_planilla_mar),
            daemon=True)
        t.start()

    def _ejecutar_agente_correos(self):
        if self.tarea_activa:
            messagebox.showwarning(
                "Agente ocupado",
                "Hay una tarea en ejecución. Espere a que finalice."
            )
            return
        self.tarea_activa = True
        self.btn_ejecutar_correos.configure(
            text="⏳  Procesando...", state="disabled",
            fg_color=Palette.ACCENT_DIM
        )
        self.progress_correos.configure(mode="indeterminate")
        self.progress_correos.start()
        self.lbl_estado_correos.configure(
            text="Escaneando carpetas y preparando correos...", text_color=Palette.INFO
        )
        self._set_status("Despachando borradores de correo...")

        # Limpiar pantalla antes de empezar
        self._limpiar_log()
        for row in self.tree_correos.get_children():
            self.tree_correos.delete(row)

        t = threading.Thread(target=self._agente_correos_worker, daemon=True)
        t.start()

    def _agente_correos_worker(self, carpetas_filtro=None, incluir_planilla=False, incluir_planilla_mar=False):
        self._set_log_panel("correos")
        try:
            self._correos_core(carpetas_filtro=carpetas_filtro, incluir_planilla=incluir_planilla, incluir_planilla_mar=incluir_planilla_mar)
        except Exception as e:
            self.after(0, lambda: self._correos_error(str(e)))
        finally:
            self.after(0, self._correos_done)

    def _correos_core(self, carpetas_filtro=None, incluir_planilla=False, incluir_planilla_mar=False):
        escritorio = self._resolver_ruta("planillas_carga", "Desktop")
        año_actual = datetime.now().strftime("%y")

        terr_planillas = []
        mar_planillas = []
        correos_a_subir = []

        self._log(f"Escaneando Escritorio: {escritorio}")

        # Pre-escanear carpetas para detectar pares COMPARTIDO
        # Agrupar por destinatario (ej: DICOAL) cuando ambas tienen "COMPARTIDO" en el nombre
        carpetas_compartido = []
        for item in sorted(os.listdir(escritorio)):
            ruta = os.path.join(escritorio, item)
            if os.path.isdir(ruta) and "COMPARTIDO" in item.upper():
                carpetas_compartido.append((item, ruta))

        pares_compartido = {}  # {dest: (carpeta1, carpeta2)}
        dests_usados = set()
        for i, (n1, r1) in enumerate(carpetas_compartido):
            dest1 = None
            for d in ("VITAPRO","EWOS","NUTRECO","DICOAL","CARGILL","BIOMAR"):
                if d in n1.upper():
                    dest1 = d; break
            if not dest1:
                match = re.search(r"(?:TERRESTRES?|ISO|FLEXI)_[^_]+_[^_]+_([A-Z]+)", n1, re.IGNORECASE)
                dest1 = match.group(1) if match else n1
            if dest1 in dests_usados:
                continue
            for j, (n2, r2) in enumerate(carpetas_compartido):
                if j <= i: continue
                if dest1 in n2.upper() or any(d in n2.upper() for d in ("VITAPRO","EWOS","NUTRECO","DICOAL","CARGILL","BIOMAR") if d == dest1):
                    pares_compartido[dest1] = ((n1, r1), (n2, r2))
                    dests_usados.add(dest1)
                    self._log(f"  🔀 Par COMPARTIDO detectado: {dest1} → {n1} + {n2}")
                    break

        for item in sorted(os.listdir(escritorio)):
            ruta_carpeta = os.path.join(escritorio, item)
            if not os.path.isdir(ruta_carpeta):
                continue
            if item.startswith(".") or item.upper() in ("RECYCLED", "RECYCLER"):
                continue

            try:
                archivos = os.listdir(ruta_carpeta)
            except Exception:
                continue

            es_maritimo = bool(re.search(r"_(ISO|FLEXI)_", item))

            # Buscar todas las planillas de carga en el escritorio para el correo grupal
            for archivo in archivos:
                up = archivo.upper()
                if up.startswith("PLANILLA DE CARGA") and (
                    up.endswith(".XLSX") or up.endswith(".XLS")
                ):
                    if "_TERRESTRE_" in item.upper():
                        terr_planillas.append(os.path.join(ruta_carpeta, archivo))
                    elif "_ISO_" in item.upper() or "_FLEXI_" in item.upper():
                        mar_planillas.append(os.path.join(ruta_carpeta, archivo))

            # Si hay filtro de carpetas, solo procesar los individuales seleccionados
            if carpetas_filtro is not None and ruta_carpeta not in carpetas_filtro:
                continue

            # Filtrar archivos para correo
            adjuntos_validos = []
            if es_maritimo:
                for archivo in archivos:
                    up_name = archivo.upper()
                    es_contenedores = "CONTENEDORES" in up_name and (up_name.endswith(".XLSX") or up_name.endswith(".XLS"))
                    es_get_pdf = up_name.startswith("GET") and up_name.endswith(".PDF")
                    if es_contenedores or es_get_pdf:
                        adjuntos_validos.append(archivo)
            else:
                for archivo in archivos:
                    up_name = archivo.upper()
                    es_plt = up_name.startswith("PLT") and up_name.endswith(".PDF")
                    es_mic = re.search(f"^{año_actual}AR", up_name) and up_name.endswith(".PDF")
                    es_excel_num = re.search(r"^\d", archivo) and (
                        up_name.endswith(".XLSX") or up_name.endswith(".XLS")
                    )
                    if es_plt or es_mic or es_excel_num:
                        adjuntos_validos.append(archivo)

            # Detectar si esta carpeta es parte de un par COMPARTIDO
            dest_folder = None
            for d in ("VITAPRO","EWOS","NUTRECO","DICOAL","CARGILL","BIOMAR"):
                if d in item.upper():
                    dest_folder = d; break
            par_activo = pares_compartido.get(dest_folder) if dest_folder else None
            es_par_lider = par_activo and ruta_carpeta == par_activo[0][1]  # r1 es la líder
            es_par_segundo = par_activo and ruta_carpeta == par_activo[1][1]  # r2 se saltea

            if es_par_segundo:
                continue  # ya se procesó junto con el líder

            if adjuntos_validos:
                if es_par_lider:
                    # Combinar adjuntos de ambas carpetas del par
                    (n1, r1), (n2, r2) = par_activo
                    ruta_par2 = r2
                    archivos_2 = []
                    try:
                        archivos_2 = os.listdir(ruta_par2)
                    except Exception:
                        pass
                    adjuntos_2 = []
                    for a2 in archivos_2:
                        up2 = a2.upper()
                        if ((up2.startswith("PLT") and up2.endswith(".PDF")) or
                            (re.search(f"^{año_actual}AR", up2) and up2.endswith(".PDF")) or
                            (re.search(r"^\d", a2) and (up2.endswith(".XLSX") or up2.endswith(".XLS")))):
                            adjuntos_2.append(a2)
                    todos_adjuntos = adjuntos_validos + adjuntos_2

                    # Asunto
                    asunto = f"SALIDA, MIC, PLANILLA COMPLETA DE EXPORTACIÓN_{dest_folder}_COMPARTIDO"

                    # Cuerpo: nombre de carpetas desde el PE
                    def _nombre_desde_pe(nombre_carpeta):
                        m = re.search(r"(?:TERRESTRES?|ISO|FLEXI)_(.+)", nombre_carpeta, re.IGNORECASE)
                        return m.group(1) if m else nombre_carpeta
                    cuerpo = f"{asunto}\n\n{_nombre_desde_pe(n1)}\n{_nombre_desde_pe(n2)}\n"

                    msg_ind = MIMEMultipart()
                    msg_ind["Subject"] = asunto
                    msg_ind["From"] = self._cfg_obtener_correo("usuario", "")
                    msg_ind["To"] = ", ".join(self._cfg_obtener_correo("destinatarios_individual", DESTINATARIOS_INDIVIDUAL))
                    msg_ind.attach(MIMEText(cuerpo, "plain"))

                    for archivo_nombre in adjuntos_validos:
                        adjuntar_archivo(msg_ind, os.path.join(ruta_carpeta, archivo_nombre))
                    for archivo_nombre in adjuntos_2:
                        adjuntar_archivo(msg_ind, os.path.join(ruta_par2, archivo_nombre))
                    correos_a_subir.append(("Individual", asunto, msg_ind, len(todos_adjuntos)))

                    self.after(0, lambda a=asunto, n=len(todos_adjuntos): self._agregar_fila_correos(
                        "Individual", a, "3 destinatarios", str(n)))
                    self._log(f"Correo COMPARTIDO preparado: {asunto} ({len(todos_adjuntos)} adjuntos)")
                elif es_maritimo and adjuntos_validos:
                    match_sufijo = re.search(r"(?:TERRESTRES?|ISO|FLEXI)_(.*)", item, re.IGNORECASE)
                    sufijo = match_sufijo.group(1) if match_sufijo else item

                    c_get = sum(1 for a in adjuntos_validos if a.upper().startswith("GET") and a.upper().endswith(".PDF"))
                    t_salida = "SALIDAS" if c_get > 1 else "SALIDA"
                    asunto = f"{t_salida}, PLANILLA COMPLETA DE EXPORTACIÓN_{sufijo}"

                    msg_ind = MIMEMultipart()
                    msg_ind["Subject"] = asunto
                    msg_ind["From"] = self._cfg_obtener_correo("usuario", "")
                    msg_ind["To"] = ", ".join(self._cfg_obtener_correo("destinatarios_individual", DESTINATARIOS_INDIVIDUAL))
                    msg_ind.attach(MIMEText(f"{asunto}\n", "plain"))

                    for archivo_nombre in adjuntos_validos:
                        adjuntar_archivo(msg_ind, os.path.join(ruta_carpeta, archivo_nombre))
                    correos_a_subir.append(("Individual", asunto, msg_ind, len(adjuntos_validos)))

                    self.after(0, lambda a=asunto, n=len(adjuntos_validos): self._agregar_fila_correos(
                        "Individual", a, "3 destinatarios", str(n)))
                    self._log(f"Correo individual marítimo preparado: {asunto}")
                else:
                    match_sufijo = re.search(r"(?:TERRESTRES?|ISO|FLEXI)_(.*)", item, re.IGNORECASE)
                    sufijo = match_sufijo.group(1) if match_sufijo else item

                    c_plt = sum(1 for a in adjuntos_validos if a.upper().startswith("PLT"))
                    c_mics = sum(1 for a in adjuntos_validos if re.search(f"^{año_actual}AR", a.upper()))
                    t_salida = "SALIDAS" if c_plt > 1 else "SALIDA"
                    t_mic = "MICS" if c_mics > 1 else "MIC"
                    asunto = f"{t_salida}, {t_mic}, PLANILLA COMPLETA DE EXPORTACIÓN_{sufijo}"

                    msg_ind = MIMEMultipart()
                    msg_ind["Subject"] = asunto
                    msg_ind["From"] = self._cfg_obtener_correo("usuario", "")
                    msg_ind["To"] = ", ".join(self._cfg_obtener_correo("destinatarios_individual", DESTINATARIOS_INDIVIDUAL))
                    msg_ind.attach(MIMEText(f"{asunto}\n", "plain"))

                    for archivo_nombre in adjuntos_validos:
                        adjuntar_archivo(msg_ind, os.path.join(ruta_carpeta, archivo_nombre))
                    correos_a_subir.append(("Individual", asunto, msg_ind, len(adjuntos_validos)))

                    self.after(0, lambda a=asunto, n=len(adjuntos_validos): self._agregar_fila_correos(
                        "Individual", a, "3 destinatarios", str(n)))
                    self._log(f"Correo individual preparado: {asunto}")

        # Correo grupal (solo si no estamos en modo Elegir con planilla desactivada)
        incluir_plt = incluir_planilla or carpetas_filtro is None
        if terr_planillas and incluir_plt:
            msg_grupal = MIMEMultipart()
            msg_grupal["Subject"] = "CARGA TERRESTRE"
            msg_grupal["From"] = self._cfg_obtener_correo("usuario", "")
            msg_grupal["To"] = ", ".join(self._cfg_obtener_correo("destinatarios_grupal", DESTINATARIOS_GRUPAL))

            # Buscar CARGA TERRESTRE antes de armar el cuerpo
            nombre_ct = self._cfg_obtener_rutas("carga_terrestre_nombre", "CARGA TERRESTRE.xlsx")
            carpeta_ct = self._cfg_obtener_rutas("carga_terrestre_carpeta", os.path.join("TRABAJO", "01_PLANILLAS"))
            ruta_maestra = buscar_archivo_en_pendrive(nombre_ct, carpeta_ct)
            if not ruta_maestra and nombre_ct.lower().endswith(".xlsx"):
                ruta_maestra = buscar_archivo_en_pendrive(nombre_ct[:-5] + ".xls", carpeta_ct)

            # Ordenar planillas por nombre
            planillas_ordenadas = sorted(terr_planillas, key=lambda p: os.path.basename(p).upper())

            cuerpo = "Estimados,\n\nSe adjuntan las planillas de carga correspondientes:\n\n"
            for p in planillas_ordenadas:
                nombre_sin_ext = os.path.splitext(os.path.basename(p))[0]
                cuerpo += f"  • {nombre_sin_ext}\n"
            if ruta_maestra:
                cuerpo += f"  • {os.path.splitext(os.path.basename(ruta_maestra))[0]}\n"
            cuerpo += "\nSaludos cordiales."
            msg_grupal.attach(MIMEText(cuerpo, "plain"))

            n_adj_grupal = len(terr_planillas)
            for p in terr_planillas:
                adjuntar_archivo(msg_grupal, p)
            if ruta_maestra:
                adjuntar_archivo(msg_grupal, ruta_maestra)
                n_adj_grupal += 1
            else:
                self._log("Advertencia: No se detectó 'CARGA TERRESTRE' en el pendrive.")

            correos_a_subir.append(("Grupal", "CARGA TERRESTRE", msg_grupal, n_adj_grupal))
            self.after(
                0,
                lambda n=n_adj_grupal: self._agregar_fila_correos(
                    "Grupal", "CARGA TERRESTRE", "14 destinatarios", str(n)
                ),
            )
            self._log("Correo grupal preparado: CARGA TERRESTRE")

        # Correo grupal marítimo (solo si no estamos en modo Elegir con planilla desactivada)
        incluir_mar_plt = incluir_planilla_mar or carpetas_filtro is None
        if mar_planillas and incluir_mar_plt:
            msg_grupal = MIMEMultipart()
            asunto_grupal = "PLANILLA DE CARGA" if len(mar_planillas) == 1 else "PLANILLAS DE CARGA"
            msg_grupal["Subject"] = asunto_grupal
            msg_grupal["From"] = self._cfg_obtener_correo("usuario", "")
            msg_grupal["To"] = ", ".join(self._cfg_obtener_correo("destinatarios_grupal", DESTINATARIOS_GRUPAL))

            mar_ordenadas = sorted(mar_planillas, key=lambda p: os.path.basename(p).upper())

            if len(mar_planillas) == 1:
                cuerpo = "Estimados,\n\nSe adjunta la planilla de carga correspondiente:\n\n"
            else:
                cuerpo = "Estimados,\n\nSe adjuntan las planillas de carga correspondientes:\n\n"
            for p in mar_ordenadas:
                nombre_sin_ext = os.path.splitext(os.path.basename(p))[0]
                cuerpo += f"  • {nombre_sin_ext}\n"
            cuerpo += "\nSaludos cordiales."
            msg_grupal.attach(MIMEText(cuerpo, "plain"))

            for p in mar_planillas:
                adjuntar_archivo(msg_grupal, p)

            correos_a_subir.append(("Grupal", asunto_grupal, msg_grupal, len(mar_planillas)))
            self.after(
                0,
                lambda a=asunto_grupal, n=len(mar_planillas): self._agregar_fila_correos(
                    "Grupal", a, "14 destinatarios", str(n)
                ),
            )
            self._log(f"Correo grupal marítimo preparado: {asunto_grupal}")

        # Subir borradores en paralelo
        if correos_a_subir:
            srv = self._cfg_obtener_correo("imap_server", IMAP_SERVER)
            prt = self._cfg_obtener_correo("imap_puerto", PUERTO_IMAP)
            usr = self._cfg_obtener_correo("usuario", "")
            pwd = self._cfg_obtener_correo("password", "")
            self._log(f"Conectando con {srv}... Subiendo {len(correos_a_subir)} correos en paralelo...")
            mensajes = [m for _, _, m, _ in correos_a_subir]

            def subir_borrador(msg):
                try:
                    mail_conn = imaplib.IMAP4(srv, prt)
                    mail_conn.login(usr, pwd)
                    fecha_con_tz = datetime.now().astimezone()
                    for carpeta in ("Drafts", "Borradores", "INBOX.Drafts"):
                        try:
                            mail_conn.append(
                                carpeta, "",
                                imaplib.Time2Internaldate(fecha_con_tz),
                                msg.as_bytes(),
                            )
                            self._log(f"  Sincronizado: {msg['Subject'][:55]}...")
                            break
                        except Exception:
                            continue
                    mail_conn.logout()
                    return True
                except Exception as e:
                    self._log(f"  Error de conexión: {e}")
                    return False

            with ThreadPoolExecutor(max_workers=4) as executor:
                resultados = list(executor.map(subir_borrador, mensajes))

            exitos = sum(1 for r in resultados if r)
            self._log(f"Despacho finalizado: {exitos}/{len(mensajes)} correos subidos.")
        else:
            self._log("No se encontraron archivos para adjuntar en ninguna carpeta.")

    def _agregar_fila_correos(self, tipo, asunto, dests, adjuntos):
        self.tree_correos.insert("", "end", values=(tipo, asunto, dests, adjuntos))

    def _correos_error(self, error_msg):
        self._log(f"ERROR CRÍTICO: {error_msg}")
        messagebox.showerror(
            "Error en Agente de Correos",
            f"Ocurrió un error durante el despacho:\n\n{error_msg}",
        )

    def _correos_done(self):
        self.tarea_activa = False
        try:
            self.btn_ejecutar_correos.configure(
                text="▶  Procesar y Despachar Correos",
                state="normal",
                fg_color=Palette.ACCENT,
            )
            self.btn_elegir_correos.configure(
                text="📂  Elegir Correos",
                state="normal",
                fg_color=Palette.SECONDARY,
            )
            self.progress_correos.stop()
            self.progress_correos.set(1)
            self.lbl_estado_correos.configure(
                text="Despacho completado", text_color=Palette.SUCCESS
            )
        except (AttributeError, Exception):
            pass
        self._set_status("Despacho de correos finalizado")

        # Mostrar popup con resumen
        total = len(self.tree_correos.get_children())
        if total > 0:
            items_tree = []
            for item in self.tree_correos.get_children():
                vals = self.tree_correos.item(item, "values")
                if vals:
                    items_tree.append(vals)
            self.after(300, lambda: self._correos_popup_confirmacion(items_tree))

    def _correos_popup_confirmacion(self, items_tree):
        """Popup lindo mostrando los correos despachados."""
        popup = ctk.CTkToplevel(self)
        popup.title("Correos Despachados")
        n = len(items_tree)
        h = min(220 + n * 40, 640)
        popup.geometry(f"530x{h}")
        popup.configure(fg_color=Palette.BG_CARD)
        popup.transient(self)
        popup.lift()
        popup.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 520) // 2
        py = self.winfo_y() + (self.winfo_height() - h) // 2
        popup.geometry(f"520x{h}+{px}+{py}")

        ctk.CTkLabel(
            popup, text="✉  CORREOS DESPACHADOS",
            font=ctk.CTkFont(family=FONT_FAMILY, size=18, weight="bold"),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(pady=(24, 4))

        ctk.CTkLabel(
            popup, text=f"{n} borrador(es) subidos a Drafts",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.TEXT_MUTED,
        ).pack(pady=(0, 16))

        for vals in items_tree:
            tipo, asunto, dests, adj = vals[0], vals[1], vals[2], vals[3]
            row = ctk.CTkFrame(popup, fg_color="transparent")
            row.pack(fill="x", padx=24, pady=3)

            color_tipo = Palette.ACCENT if "Grupal" in tipo else Palette.SECONDARY
            ctk.CTkLabel(
                row, text=tipo,
                font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
                text_color=color_tipo, width=70, anchor="w",
            ).pack(side="left")

            ctk.CTkLabel(
                row, text=asunto[:55],
                font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                text_color=Palette.TEXT_PRIMARY, anchor="w",
            ).pack(side="left", padx=(8, 0))

            ctk.CTkLabel(
                row, text=f"{adj} adj.",
                font=ctk.CTkFont(family=FONT_FAMILY, size=11),
                text_color=Palette.TEXT_MUTED, width=50, anchor="e",
            ).pack(side="right")

        ctk.CTkFrame(popup, fg_color=Palette.BORDER, height=1).pack(fill="x", padx=32, pady=(12, 0))

        ctk.CTkLabel(
            popup, text="Revisá tu bandeja de borradores y envialos manualmente.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_SECONDARY,
        ).pack(pady=(8, 12))

        ctk.CTkButton(
            popup, text="Cerrar", width=120, height=34,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6,
            command=popup.destroy,
        ).pack(pady=(0, 16))

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
        
        Procesa max 100 msg/tick para no congelar el main thread.
        """
        try:
            _processed = 0
            while _processed < 100:
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

    def _procesar_resultado_ocr(self, data):
        """Inserta una fila por ticket con color verde/rojo según match.
        Almacena datos de comparación para abrir detalle al hacer click."""
        ticket = data.get("ticket", {})
        cont_data = data.get("contenedor")
        ruta_match = data.get("match")

        # Valores OCR
        archivo   = ticket.get("archivo", "")
        patente   = ticket.get("patente", "")
        semi      = ticket.get("semi", "")
        conductor = ticket.get("conductor", "")
        dni       = ticket.get("dni", "")
        neto_ocr  = ticket.get("neto", 0)
        tara_ocr  = ticket.get("tara", 0)
        contenedor_str = ticket.get("contenedor", "")
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
            valores = (archivo, patente, semi, conductor, dni, neto_ocr, tara_ocr, contenedor_str, permiso, estado)
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
        ok_neto      = _comparar_num(neto_ocr, peso_carga)
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

        # Insertar fila única en TreeView
        iid = f"ocr_{self._cargar_datos_idx}"
        self._cargar_datos_idx += 1
        nombre_contenedor = os.path.basename(ruta_match)
        valores = (
            f"📄 {archivo}",
            patente, semi, conductor, dni,
            f"{neto_ocr:.0f}", f"{tara_ocr:.0f}",
            contenedor_str,
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
                "Contenedor": contenedor_str,
                "Permiso": permiso,
            },
            "contenedor": {
                "Patente": cam_pat,
                "Semirremolque": cam_semi,
                "Conductor": cam_cond,
                "DNI": cam_dni,
                "Neto (kg)": f"{peso_carga:.0f}",
                "Tara (kg)": f"{tara_cont:.0f}",
                "Contenedor": camion_match.get("contenedor", ""),
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

            dlg = ctk.CTkToplevel(self)
            dlg.title(f"Comparación — {datos['archivo']}")
            dlg.transient(self)
            dlg.grab_set()
            dlg.resizable(False, False)

            # ── Top bar: header + font level override ──────────────────
            top_bar = ctk.CTkFrame(dlg, fg_color=Palette.BG_SIDEBAR, corner_radius=6, height=36)
            top_bar.pack(fill="x", padx=12, pady=(12, 0))
            top_bar.pack_propagate(False)

            for col_i, txt in enumerate(["Campo", "Ticket (OCR)", "Contenedor (Excel)"]):
                lbl = ctk.CTkLabel(
                    top_bar, text=txt, width=120 if col_i == 0 else 160,
                    font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["header"], weight="bold"),
                    text_color=Palette.TEXT_SECONDARY,
                )
                kwargs = {"side": "left", "padx": 4}
                if col_i > 0:
                    kwargs["fill"] = "x"
                    kwargs["expand"] = True
                lbl.pack(**kwargs)

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

            # Cuerpo: frame normal sin scroll
            body = ctk.CTkFrame(dlg, fg_color="transparent")
            body.pack(fill="x", padx=12, pady=8)

            campos = [("Camion", "Patente"), ("Semirremolque", "Semirremolque"), ("Conductor", "Conductor"),
                       ("DNI", "DNI"), ("Neto (kg)", "Neto (kg)"), ("Tara (kg)", "Tara (kg)"),
                       ("Contenedor", "Contenedor"), ("Permiso", "Permiso")]

            for label, key in campos:
                val_ticket = datos["ticket"].get(key, "")
                val_cont   = datos["contenedor"].get(key, "")
                ok         = datos["ok"].get(key, False)

                bg = "#C8FFC8" if ok else "#FFC8C8"
                fg = "#006400" if ok else "#8B0000"

                row = ctk.CTkFrame(body, fg_color="transparent")
                row.pack(fill="x", pady=1)

                ctk.CTkLabel(
                    row, text=label, width=120, anchor="w",
                    font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["data"], weight="bold"),
                    text_color=Palette.TEXT_PRIMARY,
                ).pack(side="left", padx=(4, 0))

                for val in [val_ticket, val_cont]:
                    lbl = ctk.CTkLabel(
                        row, text=val, width=160, anchor="center",
                        font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["data"]),
                        fg_color=bg, text_color=fg, corner_radius=4,
                    )
                    lbl.pack(side="left", padx=4, pady=2, fill="x", expand=True)

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

    # ═══════════════════════════════════════════════════════════════════
    # PANEL: CARGA DE DATOS (T3+)
    # ═══════════════════════════════════════════════════════════════════
    def _panel_cargar_datos(self):
        if "cargar-datos" in self._panel_frames and hasattr(self, 'tree_coordinacion'):
            return
        frame = ctk.CTkFrame(self.panel_container, fg_color="transparent")
        # NOTA: _panel_frames se setea al FINAL del método para evitar
        # que una inicialización incompleta deje el panel huérfano.
        # frame.pack(fill="both", expand=True) — packed by _cambiar_panel_forzado

        # ── T11: Verificar dependencias (Tesseract + Poppler) ───────────
        dependencias_ok = True
        errores_dep = []

        # ── Configurar rutas de binarios ──────────────────────────
        _app_base = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        _res_base = getattr(sys, '_MEIPASS', _app_base)
        _POPPLER_BIN = os.path.join(_res_base, "poppler", "Library", "bin")
        _TESSERACT_EXE = os.path.join(_res_base, "engines", "tesseract", "tesseract.exe")
        if not os.path.isfile(_TESSERACT_EXE):
            _TESSERACT_EXE = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

        # Verificar Tesseract (solo archivo, sin ejecutarlo — evita ventana DOS)
        import pytesseract
        tesseract_sidecar = os.path.join(
            getattr(sys, '_MEIPASS', os.path.dirname(__file__)),
            "engines", "tesseract", "tesseract.exe")
        tesseract_paths = [
            tesseract_sidecar,
            _TESSERACT_EXE,
        ]
        if not any(os.path.isfile(p) for p in tesseract_paths):
            dependencias_ok = False
            errores_dep.append(
                "⚠ Tesseract-OCR no está instalado. "
                "Descargalo de https://github.com/UB-Mannheim/tesseract/wiki"
            )
        else:
            pytesseract.pytesseract.tesseract_cmd = next(p for p in tesseract_paths if os.path.isfile(p))

        # Verificar Poppler (pdf2image)
        poppler_binarios = [
            os.path.join(_POPPLER_BIN, "pdftoppm.exe"),
            os.path.join(_POPPLER_BIN, "pdfinfo.exe"),
        ]
        if not any(os.path.isfile(b) for b in poppler_binarios):
            dependencias_ok = False
            errores_dep.append(
                "⚠ Poppler no está instalado. "
                "Descargalo de https://github.com/oschwartz10612/poppler-windows/releases/"
            )

        if not dependencias_ok:
            error_frame = ctk.CTkFrame(
                frame, fg_color="#FFE0E0", corner_radius=8,
                border_width=1, border_color="#FFB0B0",
            )
            error_frame.pack(fill="x", pady=(0, 6))

            for msg in errores_dep:
                ctk.CTkLabel(
                    error_frame, text=msg,
                    font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                    text_color="red", wraplength=600, justify="left",
                ).pack(padx=12, pady=6, anchor="w")

            # ── Toolbar con botones deshabilitados ──────────────────────
            toolbar = ctk.CTkFrame(
                frame, fg_color=Palette.BG_CARD, corner_radius=8,
                border_width=1, border_color=Palette.BORDER, height=44
            )
            toolbar.pack(fill="x", pady=(0, 6))
            toolbar.pack_propagate(False)

            self.btn_controlar_tickets = ctk.CTkButton(
                toolbar,
                text="🎫  Controlar Tickets",
                font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
                fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
                text_color=Palette.WHITE, corner_radius=6, height=34, width=170,
                state="disabled",
            )
            self.btn_controlar_tickets.pack(side="left", padx=4, pady=4)

            # Espacio para mantener layout
            ctk.CTkLabel(
                toolbar,
                text="Dependencias faltantes — resolver antes de continuar",
                font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                text_color=Palette.ERROR,
            ).pack(side="left", padx=(8, 0))

            # ── Resultados vacío ──────────────────────────────────────
            scroll = ctk.CTkScrollableFrame(
                frame, fg_color=Palette.BG_CARD, corner_radius=8,
                border_width=1, border_color=Palette.BORDER
            )
            scroll.pack(fill="both", expand=True, pady=(0, 6))

            self._panel_frames["cargar-datos"] = frame
            return  # No seguir con el panel normal

        # ── Toolbar ──────────────────────────────────────────────────
        toolbar = ctk.CTkFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER, height=44
        )
        toolbar.pack(fill="x", pady=(0, 6))
        toolbar.pack_propagate(False)

        # ── Grupo Control Tickets ──────────────────────────────────
        self.btn_controlar_tickets = ctk.CTkButton(
            toolbar,
            text="🎫  Controlar Tickets",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=34, width=150,
            command=self._cargar_datos_seleccionar_pdfs,
        )
        self.btn_controlar_tickets.pack(side="left", padx=(10, 2), pady=5)

        # ── Separador vertical ──────────────────────────────────────
        ctk.CTkFrame(
            toolbar, width=2, height=34, fg_color=Palette.BORDER
        ).pack(side="left", padx=6, pady=5)

        self.btn_controlar_coordinacion = ctk.CTkButton(
            toolbar,
            text="📋  Controlar Coordinación",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=34, width=150,
            command=self._controlar_coordinacion,
        )
        self.btn_controlar_coordinacion.pack(side="left", padx=2, pady=5)

        # ── Separador vertical ──────────────────────────────────────
        ctk.CTkFrame(
            toolbar, width=2, height=34, fg_color=Palette.BORDER
        ).pack(side="left", padx=6, pady=5)

        self.btn_control_final = ctk.CTkButton(
            toolbar,
            text="Control Final",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.SECONDARY, hover_color=Palette.SECONDARY_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=34, width=150,
            state="normal",
            command=self._control_final_seleccionar,
        )
        self.btn_control_final.pack(side="left", padx=2, pady=5)

        self.progress_carga = ctk.CTkProgressBar(
            toolbar, width=140, height=8, corner_radius=4,
            mode="indeterminate",
            fg_color=Palette.BG_INPUT, progress_color=Palette.ACCENT,
        )
        # Oculto por defecto — se muestra al iniciar procesamiento

        self.btn_limpiar_carga = ctk.CTkButton(
            toolbar,
            text="Limpiar",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color="transparent",
            hover_color=Palette.BG_HOVER,
            text_color=Palette.TEXT_PRIMARY,
            border_width=0,
            corner_radius=4,
            height=30,
            width=70,
            command=self._limpiar_carga,
        )
        self.btn_limpiar_carga.pack(side="right", padx=4, pady=4)

        # ── Resultados (Frame + TreeView con scrollbars nativas) ──────
        scroll = ctk.CTkFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER
        )
        scroll.pack(fill="both", expand=True, pady=(0, 6))

        # Treeview style oscuro
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "CargaDatos.Treeview",
            background=Palette.BG_TABLE,
            foreground=Palette.TEXT_PRIMARY,
            fieldbackground=Palette.BG_TABLE,
            borderwidth=0,
            font=(FONT_FAMILY, 10),
            rowheight=28,
        )
        style.map(
            "CargaDatos.Treeview",
            background=[("selected", Palette.ACCENT_DIM)],
            foreground=[("selected", Palette.WHITE)],
        )
        style.configure(
            "CargaDatos.Treeview.Heading",
            background=Palette.BG_SIDEBAR,
            foreground=Palette.TEXT_SECONDARY,
            font=(FONT_FAMILY, 10, "bold"),
            borderwidth=0,
            padding=(8, 6),
        )

        # ── Sección: Tickets ──────────────────────────────────────────
        self._section_tickets = ctk.CTkFrame(
            scroll, fg_color="transparent",
        )

        ctk.CTkLabel(
            self._section_tickets,
            text="  Control de Tickets",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=Palette.TEXT_SECONDARY, anchor="w",
        ).pack(fill="x", padx=8, pady=(6, 2))

        columns = ("archivo", "patente", "semi", "conductor", "dni",
                   "neto", "tara", "contenedor", "permiso", "estado")
        headers = ("Archivo", "Patente", "Semirremolque", "Conductor", "DNI",
                   "Neto", "Tara", "Contenedor", "Permiso", "Estado")
        anchos = (180, 80, 80, 100, 80, 70, 70, 100, 100, 100)

        self.tree_carga = ttk.Treeview(
            self._section_tickets, columns=columns, show="headings",
            style="CargaDatos.Treeview", selectmode="browse",
        )
        for col, hdr, w in zip(columns, headers, anchos):
            self.tree_carga.heading(col, text=hdr)
            self.tree_carga.column(col, width=w, anchor="center", minwidth=w)

        # Scrollbar vertical
        scroll_y = ctk.CTkScrollbar(
            self._section_tickets, orientation="vertical",
            fg_color=Palette.BG_CARD, button_color=Palette.TEXT_MUTED,
            button_hover_color=Palette.TEXT_SECONDARY,
        )
        scroll_y.pack(side="right", fill="y", padx=(0, 2), pady=2)

        self.tree_carga.configure(yscrollcommand=scroll_y.set)
        scroll_y.configure(command=self.tree_carga.yview)

        # Scrollbar horizontal
        self.scroll_carga_x = ctk.CTkScrollbar(
            self._section_tickets, orientation="horizontal",
            fg_color=Palette.BG_CARD, button_color=Palette.TEXT_MUTED,
            button_hover_color=Palette.TEXT_SECONDARY,
        )
        self.tree_carga.configure(xscrollcommand=self.scroll_carga_x.set)
        self.scroll_carga_x.configure(command=self.tree_carga.xview)
        self.scroll_carga_x.pack(fill="x", padx=2, pady=(2, 0))

        self.tree_carga.pack(fill="both", expand=True, padx=2, pady=(0, 0))

        # Tags
        self.tree_carga.tag_configure("tag_ok", background="#C8FFC8", foreground="#006400")
        self.tree_carga.tag_configure("tag_mismatch", background="#FFC8C8", foreground="#8B0000")
        self.tree_carga.tag_configure("tag_no_match", background="#FFF3E0", foreground="#E65100")

        self.tree_carga.bind("<Double-1>", self._abrir_comparacion)

        # ── Sección: Coordinación ISO/FLEXI ──────────────────────────
        self._section_coord = ctk.CTkFrame(
            scroll, fg_color="transparent",
        )

        ctk.CTkLabel(
            self._section_coord,
            text="  Coordinación ISO/FLEXI",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=Palette.TEXT_SECONDARY, anchor="w",
        ).pack(fill="x", padx=8, pady=(6, 2))

        cols_c = ("carpeta", "giro", "cliente", "destino",
                  "buque", "viaje", "booking", "pto_descarga", "pto_final",
                  "fecha_of_pe", "fecha_carga", "peso_flexi", "estado")
        hdrs_c = ("Carpeta", "Giro", "Cliente", "Destino",
                  "Buque", "Viaje", "Booking", "Pto Descarga", "Pto Final",
                  "Fec Of PE", "Fec Carga", "Peso Flexi", "Estado")
        ancho_c = (100, 70, 220, 120, 180, 70, 160, 120, 120, 110, 110, 80, 80)

        self.tree_coordinacion = ttk.Treeview(
            self._section_coord, columns=cols_c, show="headings",
            style="CargaDatos.Treeview", selectmode="browse",
        )
        for col, hdr, w in zip(cols_c, hdrs_c, ancho_c):
            self.tree_coordinacion.heading(col, text=hdr)
            self.tree_coordinacion.column(col, width=w, anchor="center", minwidth=w)

        self.tree_coordinacion.tag_configure("tag_ok", background="#C8FFC8", foreground="#006400")
        self.tree_coordinacion.tag_configure("tag_mismatch", background="#FFC8C8", foreground="#8B0000")
        self.tree_coordinacion.tag_configure("tag_error", background="#FFE0B0", foreground="#8B4500")

        self.tree_coordinacion.bind("<Double-1>", self._abrir_comparacion_coordinacion)

        # Scrollbar horizontal
        self.scroll_coord_x = ctk.CTkScrollbar(
            self._section_coord, orientation="horizontal",
            fg_color=Palette.BG_CARD, button_color=Palette.TEXT_MUTED,
            button_hover_color=Palette.TEXT_SECONDARY,
        )
        self.tree_coordinacion.configure(xscrollcommand=self.scroll_coord_x.set)
        self.scroll_coord_x.configure(command=self.tree_coordinacion.xview)
        self.scroll_coord_x.pack(fill="x", padx=2, pady=(2, 0))

        self.tree_coordinacion.pack(fill="both", expand=True, padx=2, pady=(0, 8))

        # ── Sección: Control Final ──────────────────────────────────
        self._section_final = ctk.CTkFrame(
            scroll, fg_color="transparent",
        )

        ctk.CTkLabel(
            self._section_final,
            text="  Control Final",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=Palette.TEXT_SECONDARY, anchor="w",
        ).pack(fill="x", padx=8, pady=(6, 2))

        cols_f = ("archivo", "patente", "semi", "conductor", "dni",
                  "neto", "tara", "contenedor", "permiso", "estado", "salida_aduana")
        hdrs_f = ("Archivo", "Patente", "Semirremolque", "Conductor", "DNI",
                  "Neto", "Tara", "Contenedor", "Permiso", "Estado", "Salida Aduana")
        anchos_f = (180, 80, 80, 100, 80, 70, 70, 100, 100, 100, 100)

        self.tree_control_final = ttk.Treeview(
            self._section_final, columns=cols_f, show="headings",
            style="CargaDatos.Treeview", selectmode="browse",
        )
        for col, hdr, w in zip(cols_f, hdrs_f, anchos_f):
            self.tree_control_final.heading(col, text=hdr)
            self.tree_control_final.column(col, width=w, anchor="center", minwidth=w)

        self.tree_control_final.tag_configure("tag_ok", background="#C8FFC8", foreground="#006400")
        self.tree_control_final.tag_configure("tag_mismatch", background="#FFC8C8", foreground="#8B0000")
        self.tree_control_final.tag_configure("tag_no_match", background="#FFF3E0", foreground="#E65100")
        self.tree_control_final.bind("<Double-1>", self._abrir_comparacion_final)

        # Scrollbar horizontal
        self.scroll_control_final_x = ctk.CTkScrollbar(
            self._section_final, orientation="horizontal",
            fg_color=Palette.BG_CARD, button_color=Palette.TEXT_MUTED,
            button_hover_color=Palette.TEXT_SECONDARY,
        )
        self.tree_control_final.configure(xscrollcommand=self.scroll_control_final_x.set)
        self.scroll_control_final_x.configure(command=self.tree_control_final.xview)
        self.scroll_control_final_x.pack(fill="x", padx=2, pady=(2, 0))

        self.tree_control_final.pack(fill="both", expand=True, padx=2, pady=(0, 8))

        # Pack todas las secciones; ocultar coord y final por defecto
        for section in (self._section_tickets, self._section_coord, self._section_final):
            section.pack(fill="both", expand=True, pady=(0, 6))
        self._section_coord.pack_forget()
        self._section_final.pack_forget()

        # Todo listo — registrar el frame como completo en _panel_frames
        self._panel_frames["cargar-datos"] = frame

    # ── Limpiar vista — planillas ──────────────────────────────────────
    def _limpiar_planillas(self):
        """Elimina filas y resetea estado del panel Planillas."""
        for row in self.tree_planillas.get_children():
            self.tree_planillas.delete(row)
        self.lbl_resumen_planillas.configure(text="Sin datos analizados")
        self.lbl_estado_planillas.configure(text="Listo para analizar el Escritorio")
        self.progress_planillas.set(0)

    # ── Limpiar vista — descargar ──────────────────────────────────────
    def _limpiar_mail(self):
        """Elimina filas y resetea estado del panel Descargar Mails."""
        for row in self._mail_tree.get_children():
            self._mail_tree.delete(row)
        self._mail_data.clear()
        self._mail_lbl_estado.configure(
            text="Listo — Buscar mails de 'papeles' o Últimos 20"
        )
        self._mail_progress.set(0)

    # ── Limpiar vista — correos ────────────────────────────────────────
    def _limpiar_correos(self):
        """Elimina filas y resetea estado del panel Enviar Correos."""
        for row in self.tree_correos.get_children():
            self.tree_correos.delete(row)
        self.lbl_estado_correos.configure(text="Listo para despachar borradores")
        self.progress_correos.set(0)

    # ── Limpiar vista — cargar datos ───────────────────────────────────
    def _limpiar_carga(self):
        """Elimina todas las filas y resetea a vista tickets."""
        self.tree_carga.delete(*self.tree_carga.get_children())
        if hasattr(self, 'tree_control_final'):
            self.tree_control_final.delete(*self.tree_control_final.get_children())
            self.tree_control_final.heading("salida_aduana", text="Salida Aduana")
        if hasattr(self, 'tree_coordinacion'):
            self.tree_coordinacion.delete(*self.tree_coordinacion.get_children())
        # Resetear a vista tickets
        self._section_coord.pack_forget()
        self._section_final.pack_forget()
        self._section_tickets.pack(fill="both", expand=True, pady=(0, 6))

    # ── Coordinación ISO/FLEXI ────────────────────────────────────────
    def _escanear_carpetas_coordinacion(self):
        """Escanea Desktop + 1 nivel buscando carpetas con CONTENEDORES Excel y PDF coordinación."""
        base = self._resolver_ruta("planillas_carga", "Desktop")
        carpetas = []

        def _tiene_coordinacion(nombre):
            """Detecta 'coordinacion' en el nombre sin importar acentos."""
            return "coordinaci" in nombre.lower()

        # 1 nivel de subcarpetas en Desktop
        for item in os.listdir(base):
            ruta_dir = os.path.join(base, item)
            if not os.path.isdir(ruta_dir):
                continue

            # Buscar CONTENEDORES Excel, PDF Coordinación, y Permiso dentro
            xls_path = None
            pdf_path = None
            permiso_path = None
            for f in os.listdir(ruta_dir):
                f_lower = f.lower()
                if "contenedor" in f_lower and (f_lower.endswith(".xlsx") or f_lower.endswith(".xls")):
                    xls_path = os.path.join(ruta_dir, f)
                elif f_lower.endswith(".pdf"):
                    if _tiene_coordinacion(f):
                        pdf_path = os.path.join(ruta_dir, f)
                    elif permiso_path is None:
                        permiso_path = os.path.join(ruta_dir, f)

            if xls_path:  # Si tiene Excel, lo consideramos
                carpetas.append({
                    "nombre": item,
                    "ruta": ruta_dir,
                    "pdf_path": pdf_path,
                    "xls_path": xls_path,
                    "permiso_path": permiso_path,
                    "pdf_ok": pdf_path is not None,
                    "xls_ok": True,
                    "permiso_ok": permiso_path is not None,
                })

        carpetas.sort(key=lambda c: c["nombre"])
        return carpetas

    def _controlar_coordinacion(self):
        """Escanea Desktop, muestra popup con checkboxes, procesa seleccionadas."""
        from procesar_tickets import (
            extraer_coordinacion, leer_choferes_coordinacion, comparar_coordinacion,
            extraer_fecha_permiso,
        )
        import unicodedata
        import tkinter as tk

        # 1. Escanear carpetas
        carpetas = self._escanear_carpetas_coordinacion()
        if not carpetas:
            messagebox.showinfo(
                "Sin resultados",
                "No se encontraron carpetas con archivos CONTENEDORES en el Desktop."
            )
            return

        # 2. Popup de selección con checkboxes
        seleccion = []

        top = ctk.CTkToplevel(self)
        top.title("Seleccionar envíos ISO/FLEXI para controlar")
        ancho, alto = 960, 520
        top.geometry(f"{ancho}x{alto}")
        top.transient(self)
        top.grab_set()
        top.update_idletasks()
        x = (top.winfo_screenwidth() - ancho) // 2
        y = (top.winfo_screenheight() - alto) // 2
        top.geometry(f"{ancho}x{alto}+{x}+{y}")

        # Encabezado
        ctk.CTkLabel(
            top,
            text="Carpetas con CONTENEDORES encontradas en el Desktop",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(pady=(12, 2))

        ctk.CTkLabel(
            top,
            text="Solo se pueden seleccionar carpetas con PDF + Excel presentes",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_MUTED,
        ).pack(pady=(0, 8))

        # Scroll frame para la lista
        scroll = ctk.CTkScrollableFrame(
            top, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER,
        )
        scroll.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        # Columnas fijas — una sola grilla compartida en el scroll
        cols_spec = [
            ("", 24, (6, 2)),           # checkbox
            ("Carpeta / Envío", 480, (4, 4)),
            ("PDF", 50, (0, 0)),
            ("Excel", 50, (0, 0)),
            ("Estado", 200, (4, 4)),
        ]

        # Configurar columnas UNA SOLA VEZ en el scroll
        for ci, (txt, w, pad) in enumerate(cols_spec):
            scroll.grid_columnconfigure(ci, minsize=w, weight=0)

        # Header de columnas — widgets directo en scroll, row 0
        for ci, (txt, w, pad) in enumerate(cols_spec):
            ctk.CTkLabel(
                scroll, text=txt, width=w,
                font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
                text_color=Palette.TEXT_SECONDARY, anchor="w",
            ).grid(row=0, column=ci, padx=pad, pady=(8, 4), sticky="w")

        # Línea separadora — row 1
        sep = ctk.CTkFrame(scroll, height=1, fg_color=Palette.BORDER)
        sep.grid(row=1, column=0, columnspan=5, sticky="ew", padx=4, pady=2)

        vars_check = []

        for i, carp in enumerate(carpetas):
            row = i + 2
            puede = carp["pdf_ok"] and carp["xls_ok"]
            bg_fila = Palette.BG_CARD

            var = tk.BooleanVar(value=puede)
            vars_check.append(var)

            chk = ctk.CTkCheckBox(
                scroll, text="", variable=var, width=24,
                checkbox_width=20, checkbox_height=20,
                state="normal" if puede else "disabled",
                fg_color=Palette.ACCENT if puede else Palette.TEXT_MUTED,
            )
            chk.grid(row=row, column=0, padx=cols_spec[0][2], pady=3)
            if not puede:
                chk.configure(text=" " * 2)

            # Carpeta — tk.Label trunca al ancho de columna (no empuja ✅/❌)
            nombre_mostrar = carp["nombre"] if len(carp["nombre"]) <= 55 else carp["nombre"][:52] + "..."
            tk.Label(
                scroll, text=nombre_mostrar, width=58, anchor="w",
                font=(FONT_FAMILY, 10), fg=Palette.TEXT_PRIMARY,
                bg=bg_fila, bd=0, highlightthickness=0,
            ).grid(row=row, column=1, padx=cols_spec[1][2], sticky="w", pady=3)

            # PDF
            pdf_txt = "✅" if carp["pdf_ok"] else "❌"
            pdf_color = Palette.SUCCESS if carp["pdf_ok"] else Palette.ERROR
            ctk.CTkLabel(
                scroll, text=pdf_txt, width=50,
                font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                text_color=pdf_color, anchor="center",
            ).grid(row=row, column=2, padx=cols_spec[2][2], pady=3)

            # Excel
            ctk.CTkLabel(
                scroll, text="✅", width=50,
                font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                text_color=Palette.SUCCESS, anchor="center",
            ).grid(row=row, column=3, padx=cols_spec[3][2], pady=3)

            # Estado
            estado_txt = "✅ Listo para controlar" if puede else "❌ No se encontró PDF de Coordinación"
            ctk.CTkLabel(
                scroll, text=estado_txt, width=cols_spec[4][1],
                font=ctk.CTkFont(family=FONT_FAMILY, size=10),
                text_color=Palette.TEXT_MUTED, anchor="w",
            ).grid(row=row, column=4, padx=cols_spec[4][2], sticky="w", pady=3)

        # Botones
        btn_frame = ctk.CTkFrame(top, fg_color="transparent")
        btn_frame.pack(fill="x", padx=12, pady=(0, 12))

        def confirmar():
            nonlocal seleccion
            for i, (var, carp) in enumerate(zip(vars_check, carpetas)):
                if var.get() and carp["pdf_ok"] and carp["xls_ok"]:
                    seleccion.append(carp)
            top.destroy()

        def cancelar():
            top.destroy()

        ctk.CTkButton(
            btn_frame, text="Procesar seleccionados",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=34, width=190,
            command=confirmar,
        ).pack(side="right", padx=4)

        ctk.CTkButton(
            btn_frame, text="Cancelar",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_INPUT, hover_color=Palette.BG_HOVER,
            text_color=Palette.TEXT_SECONDARY, corner_radius=6, height=34, width=120,
            command=cancelar,
        ).pack(side="right", padx=4)

        top.wait_window()

        if not seleccion:
            return

        # 3. Procesar seleccionadas
        if not hasattr(self, 'tree_coordinacion'):
            self._panel_cargar_datos()
            if not hasattr(self, 'tree_coordinacion'):
                self.log("ERROR: Panel de carga no disponible. Reinicie la aplicación.")
                return
        self.tree_coordinacion.delete(*self.tree_coordinacion.get_children())
        self._coord_comparaciones = {}

        # Mostrar sección coordinación, ocultar otras
        self._section_tickets.pack_forget()
        self._section_final.pack_forget()
        self._section_coord.pack(fill="both", expand=True, pady=(0, 6))

        ok_count = 0
        diff_count = 0
        error_count = 0

        for carp in seleccion:
            try:
                datos_pdf = extraer_coordinacion(carp["pdf_path"])
                if "_error" in datos_pdf:
                    self._insertar_fila_coordinacion(
                        carp["nombre"], "⚠ Error en PDF: " + datos_pdf["_error"],
                        "tag_error", None
                    )
                    error_count += 1
                    continue

                datos_xls = leer_choferes_coordinacion(carp["xls_path"])
                if "_error" in datos_xls:
                    self._insertar_fila_coordinacion(
                        carp["nombre"], "⚠ " + datos_xls["_error"],
                        "tag_error", None
                    )
                    error_count += 1
                    continue

                # OCR del permiso de exportación para fecha de oficialización
                if carp.get("permiso_path"):
                    fecha_of = extraer_fecha_permiso(carp["permiso_path"])
                    datos_pdf["fecha_of_pe"] = fecha_of
                else:
                    datos_pdf["fecha_of_pe"] = ""

                comparaciones, estado = comparar_coordinacion(datos_pdf, datos_xls)

                tag = "tag_ok" if estado == "ok" else "tag_mismatch"
                estado_label = "✅ OK" if estado == "ok" else "❌ Diferencias"

                valores = (
                    datos_pdf.get("carpeta", ""),
                    datos_pdf.get("giro", ""),
                    datos_pdf.get("cliente", ""),
                    datos_pdf.get("destino", ""),
                    datos_pdf.get("buque", ""),
                    datos_pdf.get("viaje", ""),
                    datos_pdf.get("booking", ""),
                    datos_pdf.get("pto_descarga", ""),
                    datos_pdf.get("pto_final", ""),
                    comparaciones.get("fecha_of_pe", {}).get("pdf", ""),
                    datos_xls.get("fecha_carga", ""),
                    datos_pdf.get("peso_flexi", ""),
                    estado_label,
                )

                iid = self.tree_coordinacion.insert("", "end", values=valores, tags=(tag,))
                self._coord_comparaciones[iid] = comparaciones

                if estado == "ok":
                    ok_count += 1
                else:
                    diff_count += 1

            except Exception as e:
                self._insertar_fila_coordinacion(
                    carp["nombre"], "⚠ Error: " + str(e),
                    "tag_error", None
                )
                error_count += 1

        total = ok_count + diff_count + error_count

    def _insertar_fila_coordinacion(self, carpeta, estado_texto, tag, comparaciones):
        """Inserta una fila en el treeview de coordinación."""
        valores = (carpeta, "", "", "", "", "", "", "", "", "", "", "", estado_texto)
        iid = self.tree_coordinacion.insert("", "end", values=valores, tags=(tag,))
        if comparaciones is not None:
            self._coord_comparaciones[iid] = comparaciones

    def _abrir_comparacion_coordinacion(self, event):
        """Popup con tabla verde/rojo campo por campo — mismo estilo que el de tickets."""
        sel = self.tree_coordinacion.selection()
        if not sel or not hasattr(self, "_coord_comparaciones"):
            return

        comps = self._coord_comparaciones.get(sel[0])
        if not comps:
            return

        etiquetas = {
            "giro": "Puerto Salida",
            "carpeta": "Carpeta", "cliente": "Cliente",
            "destino": "Destino", "buque": "Buque", "viaje": "Viaje",
            "booking": "Booking", "pto_descarga": "Pto Descarga",
            "pto_final": "Pto Final",
            "fecha_of_pe": "Fecha Ofic.",
            "fecha_carga": "Fecha Carga",
            "peso_flexi": "Peso Flexi (kg)",
        }

        def _build_popup(level):
            fsizes = self._get_font_sizes(level)
            geom = self._get_popup_geometry(760, 480, level)
            ancho = int(geom.split("x")[0])

            dlg = ctk.CTkToplevel(self)
            dlg.title("Comparación — Coordinación vs Excel")
            dlg.transient(self)
            dlg.resizable(False, False)

            # ── Top bar: header + font level override ──────────────────
            top_bar = ctk.CTkFrame(dlg, fg_color=Palette.BG_SIDEBAR, corner_radius=6, height=36)
            top_bar.pack(fill="x", padx=12, pady=(12, 0))
            top_bar.pack_propagate(False)

            for col_i, txt in enumerate(["Campo", "PDF (Coordinación)", "Excel (Choferes)"]):
                lbl = ctk.CTkLabel(
                    top_bar, text=txt, width=120 if col_i == 0 else 280,
                    font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["header"], weight="bold"),
                    text_color=Palette.TEXT_SECONDARY,
                )
                kwargs = {"side": "left", "padx": 4}
                if col_i > 0:
                    kwargs["fill"] = "x"
                    kwargs["expand"] = True
                lbl.pack(**kwargs)

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

            # ── Cuerpo: frame normal sin scroll ─────────────────────────
            body = ctk.CTkFrame(dlg, fg_color="transparent")
            body.pack(fill="x", padx=12, pady=8)

            for campo, c in comps.items():
                etq = etiquetas.get(campo, campo)
                v_pdf = c["pdf"] or "—"
                v_xls = c["excel"] or "—"
                ok = c["match"]

                bg = "#C8FFC8" if ok else "#FFC8C8"
                fg = "#006400" if ok else "#8B0000"

                row = ctk.CTkFrame(body, fg_color="transparent")
                row.pack(fill="x", pady=1)

                ctk.CTkLabel(
                    row, text=etq, width=120, anchor="w",
                    font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["data"], weight="bold"),
                    text_color=Palette.TEXT_PRIMARY,
                ).pack(side="left", padx=(4, 0))

                for val in [v_pdf, v_xls]:
                    lbl = ctk.CTkLabel(
                        row, text=val, width=280, anchor="center",
                        font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["data"]),
                        fg_color=bg, text_color=fg, corner_radius=4,
                    )
                    lbl.pack(side="left", padx=4, pady=2, fill="x", expand=True)

            # ── Leyenda al pie (centrada) ─────────────────────────────
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


    def _control_final_seleccionar(self):
        """Seleccionar PDFs (tickets + aduanas) + Excel para Control Final."""
        from tkinter import filedialog as tk_filedialog
        rutas = tk_filedialog.askopenfilenames(
            title="Seleccionar PDFs (tickets + salidas aduana) + Excel",
            filetypes=[("Archivos soportados", "*.pdf *.xls *.xlsx")],
        )
        if not rutas:
            return
        # Separar Excel de PDFs
        exceles = [r for r in rutas if r.lower().endswith(('.xls', '.xlsx'))]
        pdfs = [r for r in rutas if r.lower().endswith('.pdf')]
        if not exceles:
            self._log("ERROR: Debe seleccionar al menos un Excel de Contenedores")
            return
        if not pdfs:
            self._log("ERROR: Debe seleccionar al menos un PDF")
            return

        self._log(f"[Control Final] {len(pdfs)} PDFs, {len(exceles)} Excel(s)")

        if self.tarea_activa:
            messagebox.showwarning("Agente ocupado", "Hay una tarea en ejecución.")
            return
        self.tarea_activa = True

        for btn_name in ('btn_control_final', 'btn_controlar_tickets', 'btn_controlar_coordinacion'):
            btn = getattr(self, btn_name, None)
            if btn and btn.winfo_exists():
                btn.configure(state="disabled")

        if hasattr(self, 'progress_carga') and self.progress_carga.winfo_exists():
            self.progress_carga.pack(side="right", padx=12)
            self.progress_carga.configure(mode="indeterminate")
            self.progress_carga.start()

        # Mostrar sección control final, ocultar otras
        self._section_tickets.pack_forget()
        self._section_coord.pack_forget()
        self._section_final.pack(fill="both", expand=True, pady=(0, 6))

        # Limpiar tree viejo
        for item in self.tree_control_final.get_children():
            self.tree_control_final.delete(item)

        t = threading.Thread(
            target=self._control_final_worker,
            args=(list(pdfs), list(exceles)),
            daemon=True,
        )
        t.start()

    def _control_final_worker(self, pdf_paths, excel_paths):
        """Worker thread para Control Final."""
        import procesar_tickets
        from datetime import date
        import os

        try:
            result_count = 0
            # 1. Clasificar PDFs
            tickets_pdf = []   # scanned, need OCR
            aduanas_data = []  # PLT / Salida Aduana (texto)
            mic_data = []      # MIC/DTA (texto)
            modo = "flexi"     # por defecto

            for ruta in pdf_paths:
                import fitz
                doc = fitz.open(ruta)
                text = ""
                for p in doc:
                    text += p.get_text()
                doc.close()

                if not text.strip():
                    tickets_pdf.append(ruta)
                    continue

                if "MIC/DTA" in text.upper() or "MANIFIESTO INTERNACIONAL DE CARGA" in text.upper():
                    # MIC/DTA → activa modo terrestre
                    md = procesar_tickets.extraer_mic_dta(ruta)
                    if "_error" not in md:
                        md["_archivo"] = os.path.basename(ruta)
                        mic_data.append(md)
                    modo = "terrestre"
                elif "SALIDA DE ZONA PRIMARIA" in text.upper():
                    # PLT Aduana
                    data = procesar_tickets.extraer_salida_aduana(ruta, modo=modo)
                    data["_archivo"] = os.path.basename(ruta)
                    aduanas_data.append(data)
                else:
                    tickets_pdf.append(ruta)

            self._log(f"[Control Final] modo={modo}, {len(tickets_pdf)} tickets, "
                      f"{len(aduanas_data)} aduanas, {len(mic_data)} mic/dta")

            # 2. OCR tickets
            self._contenedores_cache = {}
            self._cargar_cache_contenedores(excel_paths)

            ocr_method = self.config.get("ocr_method", "api_vision")
            textos_por_pdf = {}
            api_datos_raw = {}

            if ocr_method == "api_vision":
                api_vision_conf = self.config.get("api_vision", {})
                api_key_raw = api_vision_conf.get("api_key", "")
                api_key = self._decrypt_api_key(api_key_raw)
                model = api_vision_conf.get("model", procesar_tickets.MODELO_VISION_DEFAULT)
                parallel_enabled = api_vision_conf.get("parallel_enabled", False)

                if parallel_enabled and len(tickets_pdf) > 1:
                    # Parallel mode: distribute across enabled models
                    model_states = api_vision_conf.get("parallel_model_states", {})
                    enabled_models = [m for m, on in model_states.items() if on]
                    if not enabled_models:
                        enabled_models = [model]
                    # Ensure selected model is first
                    if model in enabled_models:
                        enabled_models.remove(model)
                    enabled_models.insert(0, model)
                    temperature = api_vision_conf.get("temperature", 0.1)
                    max_tokens = api_vision_conf.get("max_tokens", 4000)
                    timeout = api_vision_conf.get("timeout", 60)
                    result = procesar_tickets.api_vision_con_fallback(
                        tickets_pdf, api_key, enabled_models,
                        temperature=temperature, max_tokens=max_tokens,
                        timeout=timeout, log_callback=self._log,
                    )
                    api_datos_raw = result["datos"]
                    textos_por_pdf = result["textos"]
                else:
                    # Sequential fallback: try models one by one per PDF
                    all_models = api_vision_conf.get("custom_models", [])
                    if not all_models:
                        all_models = list(procesar_tickets.MODELOS_VISION)
                    # Ensure selected model is first
                    if model in all_models:
                        all_models.remove(model)
                    all_models.insert(0, model)

                    self._log(f"[Control Final] API Visión secuencial: {len(tickets_pdf)} PDF(s), {len(all_models)} modelo(s)")
                    for ruta in tickets_pdf:
                        stem = os.path.splitext(os.path.basename(ruta))[0]
                        self._log(f"  API Visión: {stem}")
                        success = False
                        for m in all_models:
                            datos = procesar_tickets.api_vision_extraer_datos(
                                ruta, api_key, model=m
                            )
                            if "error" in datos:
                                err = datos["error"]
                                if "is not a valid model ID" in err:
                                    err = "modelo no disponible"
                                elif "429" in err or "rate" in err.lower():
                                    err = "rate limit"
                                elif "timeout" in err.lower():
                                    err = "timeout"
                                self._log(f"  ERROR {m}: {err}")
                                continue
                            # Success
                            api_datos_raw[stem] = datos
                            self._log(f"  ✓ {stem} procesado con {m}")
                            success = True
                            break
                        if not success:
                            self._log(f"  ⚠ Todos los modelos fallaron para {stem}, usando PaddleOCR")
                            texto = procesar_tickets.pdf_a_texto(ruta, engine="paddleocr")
                            textos_por_pdf[stem] = texto
            else:
                textos_por_pdf = procesar_tickets.pdfs_a_texto_batch(tickets_pdf, engine="paddleocr")

            # 3. Build lookup structures
            self._log(f"[Control Final] Procesando tickets...")

            if modo == "terrestre":
                # ── Terrestre: match por patente → precinto → DNI ──
                import re as _re

                def _norm_pat_match(s):
                    """Normalizar patente: sacar espacios, uppercase."""
                    if not s: return ""
                    return _re.sub(r'\s+', '', s).upper()

                # ── Build lookup structures desde MIC/DTA ──
                aduana_por_patente = {}
                aduana_por_precinto = {}
                aduana_por_dni = {}
                for md in mic_data:
                    pat = _norm_pat_match(md.get("patente_camion", ""))
                    if pat: aduana_por_patente[pat] = md
                    if md.get("precinto"):
                        for p in md["precinto"].upper().split():
                            p = p.strip()
                            if p: aduana_por_precinto[p] = md
                    dni = _re.sub(r'\D', '', md.get("cuil", ""))
                    if dni: aduana_por_dni[dni] = md

                # Excel lookups
                excel_por_patente = {}
                excel_por_precinto = {}
                excel_por_dni = {}
                for k, v in self._contenedores_cache.items():
                    camiones = v.get("camiones", [])
                    for ci, cam in enumerate(camiones):
                        pat = _norm_pat_match(cam.get("patente_camion", ""))
                        if pat: excel_por_patente[pat] = (v, ci)
                        prec = cam.get("precinto", "").upper().strip()
                        if prec: excel_por_precinto[prec] = (v, ci)
                        dni = _re.sub(r'\D', '', str(cam.get("dni", "")))
                        if dni: excel_por_dni[dni] = (v, ci)

                todos_pdfs = set(tickets_pdf)
                for ruta in todos_pdfs:
                    stem = os.path.splitext(os.path.basename(ruta))[0]
                    try:
                        if stem in api_datos_raw:
                            raw = api_datos_raw[stem]
                            ticket_data = {
                                "archivo":   stem,
                                "Patente":   raw.get("Patente", ""),
                                "Semirremolque": raw.get("Semirremolque", ""),
                                "Conductor": raw.get("Conductor", ""),
                                "DNI":       raw.get("DNI", ""),
                                "Neto":      raw.get("Neto", ""),
                                "Tara Contenedor": raw.get("Tara Contenedor", ""),
                                "Contenedor": raw.get("Contenedor", ""),
                                "Permiso":   raw.get("Permiso", ""),
                            }
                        else:
                            texto = textos_por_pdf.get(stem, "")
                            if not texto:
                                continue
                            extraido = procesar_tickets.extraer_datos(texto)
                            ticket_data = {
                                "archivo":   stem,
                                "Patente":   extraido.get("patente", ""),
                                "Semirremolque": extraido.get("semi", ""),
                                "Conductor": extraido.get("conductor", ""),
                                "DNI":       extraido.get("dni", ""),
                                "Neto":      extraido.get("neto", ""),
                                "Tara Contenedor": extraido.get("tara", ""),
                                "Contenedor": extraido.get("contenedor", ""),
                                "Permiso":   extraido.get("oferta", ""),
                            }

                        patente_ticket = _norm_pat_match(ticket_data.get("Patente", ""))
                        dni_ticket = _re.sub(r'\D', '', ticket_data.get("DNI", ""))

                        # ── Paso 1: Match por patente ──
                        cont_data, cont_idx = excel_por_patente.get(patente_ticket, (None, -1))
                        aduana = aduana_por_patente.get(patente_ticket, {})

                        # Si no hay match directo, probar variantes O/0 en patente
                        if not cont_data and not aduana:
                            # Generar variantes: O⇄0 en nuevas patentes Mercosur (AB123CD)
                            for i, ch in enumerate(patente_ticket):
                                variants = set()
                                if ch == 'O': variants.add(patente_ticket[:i] + '0' + patente_ticket[i+1:])
                                if ch == '0': variants.add(patente_ticket[:i] + 'O' + patente_ticket[i+1:])
                                if ch == 'I': variants.add(patente_ticket[:i] + '1' + patente_ticket[i+1:])
                                if ch == '1': variants.add(patente_ticket[:i] + 'I' + patente_ticket[i+1:])
                                if ch == 'S': variants.add(patente_ticket[:i] + '5' + patente_ticket[i+1:])
                                if ch == '5': variants.add(patente_ticket[:i] + 'S' + patente_ticket[i+1:])
                                for v in variants:
                                    if not cont_data and v in excel_por_patente:
                                        cont_data, cont_idx = excel_por_patente[v]
                                        self._log(f"  ⚠ Patente OCR corregida: {patente_ticket}→{v} (Excel)")
                                    if not aduana and v in aduana_por_patente:
                                        aduana = aduana_por_patente[v]
                                        self._log(f"  ⚠ Patente OCR corregida: {patente_ticket}→{v} (Aduana)")
                                    if cont_data and aduana: break
                                if cont_data and aduana: break

                        # ── Paso 2: Fallback por precinto ──
                        if not cont_data or not aduana:
                            # 2a: Tenemos Excel pero falta Aduana
                            if cont_data and cont_idx >= 0 and not aduana:
                                cam = cont_data["camiones"][cont_idx]
                                prec = cam.get("precinto", "").upper().strip()
                                if prec and prec in aduana_por_precinto:
                                    aduana = aduana_por_precinto[prec]
                                    self._log(f"  ⚠ Aduana x precinto Excel: {stem} — precinto={prec}")

                            # 2b: Tenemos Aduana pero falta Excel
                            if aduana and (not cont_data or cont_idx < 0):
                                prec = aduana.get("precinto", "").upper().strip()
                                for p in prec.split():
                                    if p in excel_por_precinto:
                                        cont_data, cont_idx = excel_por_precinto[p]
                                        self._log(f"  ⚠ Excel x precinto Aduana: {stem} — precinto={p}")
                                        break

                            # 2c: Sin patente match: DNI primero (único, no cruza)
                            if not cont_data and not aduana and dni_ticket:
                                if dni_ticket in excel_por_dni:
                                    cont_data, cont_idx = excel_por_dni[dni_ticket]
                                    self._log(f"  ⚠ Excel x DNI: {stem} — dni={dni_ticket}")
                                if dni_ticket in aduana_por_dni:
                                    aduana = aduana_por_dni[dni_ticket]
                                    self._log(f"  ⚠ Aduana x DNI: {stem} — dni={dni_ticket}")

                            # 2d: Después de DNI, puentear lo que falte
                            if not cont_data or not aduana:
                                if cont_data and cont_idx >= 0 and not aduana:
                                    cam = cont_data["camiones"][cont_idx]
                                    prec = cam.get("precinto", "").upper().strip()
                                    if prec and prec in aduana_por_precinto:
                                        aduana = aduana_por_precinto[prec]
                                        self._log(f"  ⚠ Aduana x precinto post-DNI: {stem} — precinto={prec}")
                                if aduana and (not cont_data or cont_idx < 0):
                                    prec = aduana.get("precinto", "").upper().strip()
                                    for p in prec.split():
                                        if p in excel_por_precinto:
                                            cont_data, cont_idx = excel_por_precinto[p]
                                            self._log(f"  ⚠ Excel x precinto post-DNI: {stem} — precinto={p}")
                                            break

                            # 2e: Sin ningún match, escanear todo por precinto
                            if not cont_data and not aduana:
                                for md in mic_data:
                                    prec = md.get("precinto", "").upper().strip()
                                    for p in prec.split():
                                        if p in excel_por_precinto:
                                            cont_data, cont_idx = excel_por_precinto[p]
                                            aduana = md
                                            self._log(f"  ⚠ Sin match x precinto (último): {stem} — precinto={p}")
                                            break
                                    if cont_data: break

                        result = self._build_fila_control_final(
                            ticket_data, cont_data, cont_idx, aduana, stem, modo=modo
                        )
                        if result:
                            self.log_queue.put(("_CONTROL_FINAL_RESULT_", result))
                            result_count += 1

                    except Exception as e:
                        self._log(f"  ✗ Error {stem}: {e}")
                        import traceback
                        self._log(traceback.format_exc())

            else:
                # ── Flexi/ISO: match por contenedor (comportamiento actual) ──
                aduana_por_contenedor = {}
                aduana_por_precinto = {}
                for ad in aduanas_data:
                    if ad.get("contenedor"):
                        aduana_por_contenedor[ad["contenedor"].upper()] = ad
                    if ad.get("precinto"):
                        for p in ad["precinto"].upper().split():
                            p = p.strip()
                            if p:
                                aduana_por_precinto[p] = ad

                todos_pdfs = set(tickets_pdf)
                for ruta in todos_pdfs:
                    stem = os.path.splitext(os.path.basename(ruta))[0]
                    try:
                        if stem in api_datos_raw:
                            raw = api_datos_raw[stem]
                            ticket_data = {
                                "archivo":   stem,
                                "Patente":   raw.get("Patente", ""),
                                "Semirremolque": raw.get("Semirremolque", ""),
                                "Conductor": raw.get("Conductor", ""),
                                "DNI":       raw.get("DNI", ""),
                                "Neto":      raw.get("Neto", ""),
                                "Tara Contenedor": raw.get("Tara Contenedor", ""),
                                "Contenedor": raw.get("Contenedor", ""),
                                "Permiso":   raw.get("Permiso", ""),
                            }
                        else:
                            texto = textos_por_pdf.get(stem, "")
                            if not texto:
                                continue
                            extraido = procesar_tickets.extraer_datos(texto)
                            ticket_data = {
                                "archivo":   stem,
                                "Patente":   extraido.get("patente", ""),
                                "Semirremolque": extraido.get("semi", ""),
                                "Conductor": extraido.get("conductor", ""),
                                "DNI":       extraido.get("dni", ""),
                                "Neto":      extraido.get("neto", ""),
                                "Tara Contenedor": extraido.get("tara", ""),
                                "Contenedor": extraido.get("contenedor", ""),
                                "Permiso":   extraido.get("oferta", ""),
                            }

                        # Match contenedor (normalizar espacios/guiones)
                        contenedor_num = ticket_data.get("Contenedor", "").upper().strip()
                        import re as _re
                        contenedor_num_norm = _re.sub(r'[\s\-]+', '', contenedor_num)
                        self._log(f"  {stem}: contenedor='{contenedor_num}' (norm='{contenedor_num_norm}')")

                        # Match Excel contenedores
                        cont_data = None
                        cont_idx = -1
                        for k, v in self._contenedores_cache.items():
                            camiones = v.get("camiones", [])
                            for ci, cam in enumerate(camiones):
                                c_raw = cam.get("contenedor", "").upper().strip()
                                c_norm = _re.sub(r'[\s\-]+', '', c_raw)
                                if c_norm and c_norm == contenedor_num_norm:
                                    cont_data = v
                                    cont_idx = ci
                                    break
                            if cont_data:
                                break

                        # Match Aduana by contenedor
                        aduana = aduana_por_contenedor.get(contenedor_num_norm, {})

                        # Fallback por precinto
                        if not cont_data or not aduana:
                            precinto_excel = ""
                            if cont_data and cont_idx >= 0:
                                cam = cont_data.get("camiones", [])
                                if cont_idx < len(cam):
                                    precinto_excel = cam[cont_idx].get("precinto", "").upper().strip()
                            else:
                                for k, v in self._contenedores_cache.items():
                                    for ci, cam in enumerate(v.get("camiones", [])):
                                        p = cam.get("precinto", "").upper().strip()
                                        if p and p in aduana_por_precinto:
                                            cont_data = v
                                            cont_idx = ci
                                            precinto_excel = p
                                            self._log(f"  ⚠ Fallback precinto: {stem} — Excel camión {ci} precinto={p}")
                                            break
                                    if cont_data:
                                        break
                            if precinto_excel and precinto_excel in aduana_por_precinto:
                                if not aduana:
                                    aduana = aduana_por_precinto[precinto_excel]
                                    self._log(f"  ⚠ Aduana encontrada por precinto: {precinto_excel}")

                        result = self._build_fila_control_final(
                            ticket_data, cont_data, cont_idx, aduana, stem, modo=modo
                        )
                        if result:
                            self.log_queue.put(("_CONTROL_FINAL_RESULT_", result))
                            result_count += 1

                    except Exception as e:
                        self._log(f"  ✗ Error {stem}: {e}")
                        import traceback
                        self._log(traceback.format_exc())

        except Exception as e:
            self._log(f"[Control Final] Error general: {e}")
            import traceback
            self._log(traceback.format_exc())
        finally:
            try: rc = result_count
            except NameError: rc = 0
            self._log(f"[Control Final] Finalizado — {rc} ticket(s) procesado(s)")
            self.log_queue.put(("_TAREA_COMPLETA_", None))

    def _build_fila_control_final(self, ticket_data, cont_data, cont_idx, aduana, stem,
                                   modo="flexi"):
        """Build row values and comparison data for Control Final.

        Args:
            modo: "flexi" (ISO/Flexi containers) or "terrestre" (bulk/granel).

        Returns dict with keys:
          valores    → tuple for tree row (11 cols)
          tag        → tag_ok / tag_mismatch
          ticket     → dict {Patente, Semirremolque, ..., Permiso} (normalized ticket)
          contenedor → dict {Patente, Semirremolque, ..., Permiso} (normalized Excel)
          aduana     → dict {Patente Camión, ..., CUIL, Contenedor, ...}
          ok         → dict per field: True if ticket matches reference
          modo       → the mode used
        """
        import procesar_tickets
        import re as _re

        # ── Normalization helpers ──────────────────────────────
        def _norm_pat(s):
            return str(s).replace(" ", "") if s else s

        def _norm_num(s):
            """Remove trailing .0 / .00 + thousands-separator dots from numeric values."""
            if not s:
                return s
            v = str(s).strip()
            v = _re.sub(r'\.0+$', '', v)      # "23440.0" → "23440"
            v = v.replace(".", "")             # "23.440" → "23440"
            return v

        def _fmt_neto(s):
            """Normalize number then format with dot as thousands separator: '27890' → '27.890'."""
            if not s:
                return s
            v = str(s).strip()
            # Quitar comas y puntos existentes
            v = v.replace(",", "")
            v = _re.sub(r'\.0+$', '', v)
            v = v.replace(".", "")
            if v.isdigit():
                return f"{int(v):,}".replace(",", ".")
            return v

        def _norm_ctn(s):
            """Remove spaces and dashes from contenedor numbers."""
            if not s:
                return s
            return _re.sub(r'[\s\-]+', '', str(s).strip())

        def _dni_solo(v):
            """Extraer solo DNI, limpiando CUIL/CUIT y no-dígitos."""
            if not v:
                return ""
            d = _re.sub(r'\D', '', v)
            if len(d) == 11 and d[:2] in ("20", "23", "27", "30"):
                return d[2:-1]
            return d

        # ── Ticket data (ORIGINAL, never overwrite) ──
        t_patente    = ticket_data.get("Patente", "")
        t_semi       = ticket_data.get("Semirremolque", "")
        t_conductor  = ticket_data.get("Conductor", "")
        t_dni        = ticket_data.get("DNI", "")
        t_neto       = ticket_data.get("Neto", "")
        t_tara       = ticket_data.get("Tara Contenedor", "")
        t_contenedor = ticket_data.get("Contenedor", "")
        t_permiso    = ticket_data.get("Permiso", "")
        t_dni        = _dni_solo(t_dni)

        # ── Aduana data ──
        a_plt            = aduana.get("plt", "")
        a_peso_bruto     = aduana.get("peso_bruto", "")
        a_cuil           = aduana.get("cuil", "")
        a_patente_camion = aduana.get("patente_camion", "")
        a_patente_semi   = aduana.get("patente_semi", "")
        a_conductor      = aduana.get("conductor", "")
        a_id_destinacion = aduana.get("id_destinacion", "")
        a_exportador     = aduana.get("exportador", "")
        a_contenedor     = aduana.get("contenedor", "")
        a_precinto       = aduana.get("precinto", "")

        a_dni = _dni_solo(a_cuil)

        # ── Excel / Contenedor data ──
        e_patente = ""
        e_semi = ""
        e_conductor = ""
        e_dni = ""
        e_neto = ""
        e_tara = ""
        e_contenedor = ""
        e_permiso = ""
        e_peso_flexi = "0"
        if cont_data and cont_idx >= 0:
            camiones = cont_data.get("camiones", [])
            if cont_idx < len(camiones):
                cam = camiones[cont_idx]
                e_patente    = cam.get("patente_camion", "")
                e_semi       = cam.get("patente_semi", "")
                e_conductor  = cam.get("conductor", "")
                e_dni        = cam.get("dni", "")
                e_neto       = str(cam.get("peso_carga", ""))
                e_tara       = str(cam.get("tara_cont", ""))
                e_contenedor = cam.get("contenedor", "")
                e_precinto   = cam.get("precinto", "")
                e_permiso    = cont_data.get("pe", "")
                e_peso_flexi = str(cam.get("peso_flexi", "0"))
        e_dni = _dni_solo(e_dni)

        # ── Normalize all values ──
        t_patente    = _norm_pat(t_patente)
        t_semi       = _norm_pat(t_semi)
        t_neto       = _norm_num(t_neto)
        t_contenedor = _norm_ctn(t_contenedor)

        e_patente    = _norm_pat(e_patente)
        e_semi       = _norm_pat(e_semi)
        e_neto       = _norm_num(e_neto)
        e_contenedor = _norm_ctn(e_contenedor)

        a_patente_camion = _norm_pat(a_patente_camion)
        a_patente_semi   = _norm_pat(a_patente_semi)
        a_peso_bruto     = _norm_num(a_peso_bruto)
        a_contenedor     = _norm_ctn(a_contenedor)

        # ── Calculate aduana neto (solo flexi: peso_bruto - flexi) ──
        if modo == "terrestre":
            a_neto_calc = a_peso_bruto
        else:
            a_neto_calc = a_peso_bruto
            if a_peso_bruto:
                try:
                    pb = float(a_peso_bruto.replace(",", ""))
                    pf_val = (cont_data.get("peso_flexi_global") if cont_data else None) or e_peso_flexi
                    pf = float(pf_val) if pf_val and pf_val != "0" else 0
                    a_neto_calc = str(int(pb - pf)) if pf > 0 else str(int(pb))
                except (ValueError, TypeError):
                    pass

        # ── Apply terrestre field overrides ──
        if modo == "terrestre":
            t_tara = "—"
            t_contenedor = "—"
            e_tara = "—"
            e_contenedor = "—"
            if not a_contenedor:
                a_contenedor = "—"

        # ── Format neto/tara with dot thousands separator ──
        t_neto       = _fmt_neto(t_neto)
        e_neto       = _fmt_neto(e_neto)
        a_neto_calc  = _fmt_neto(a_neto_calc)
        t_tara       = _fmt_neto(t_tara)
        e_tara       = _fmt_neto(e_tara)
        # No format a_peso_bruto display — a_neto_calc is used for comparison

        # ── Build display dicts (all normalized) ──
        ticket_d = {
            "Patente": t_patente, "Semirremolque": t_semi,
            "Conductor": t_conductor, "DNI": t_dni,
            "Neto (kg)": t_neto, "Tara (kg)": t_tara,
            "Contenedor": t_contenedor, "Permiso": t_permiso,
            "Precinto": "-",
        }
        cont_d = {
            "Patente": e_patente, "Semirremolque": e_semi,
            "Conductor": e_conductor, "DNI": e_dni,
            "Neto (kg)": e_neto, "Tara (kg)": e_tara,
            "Contenedor": e_contenedor, "Permiso": e_permiso,
            "Precinto": e_precinto,
        }
        dni_key = "DNI" if modo == "terrestre" else "CUIL"
        aduana_d = {
            "PLT": a_plt,
            "Peso Bruto Aduana": a_neto_calc,
            "Id Destinación": a_id_destinacion,
            "Exportador": a_exportador,
            "Patente Camión": a_patente_camion,
            "Patente Semi": a_patente_semi,
            "Conductor Aduana": a_conductor,
            dni_key: a_dni,
            "Contenedor": a_contenedor,
            "Precinto": a_precinto,
        }

        # ── 3-way comparison ──
        ok_map = {}
        all_ok = True
        compare_fields = [
            "Patente", "Semirremolque", "Conductor",
            "DNI", "Neto (kg)", "Tara (kg)",
            "Contenedor", "Permiso", "Precinto",
        ]

        for campo in compare_fields:
            v_ticket = ticket_d.get(campo, "").strip().upper()
            v_cont   = cont_d.get(campo, "").strip().upper()

            # Map field → aduana variable
            aduana_val = ""
            if   campo == "Patente":      aduana_val = a_patente_camion.strip().upper()
            elif campo == "Semirremolque": aduana_val = a_patente_semi.strip().upper()
            elif campo == "Conductor":    aduana_val = a_conductor.strip().upper()
            elif campo == "DNI":          aduana_val = a_dni.strip().upper()
            elif campo == "Neto (kg)":    aduana_val = a_neto_calc.strip().upper()
            elif campo == "Contenedor":   aduana_val = a_contenedor.strip().upper()
            elif campo == "Permiso":      aduana_val = a_id_destinacion.strip().upper()
            elif campo == "Precinto":     aduana_val = a_precinto.strip().upper()
            # Tara → no aduana equivalent, stays ""

            # Normalize conductor accents before compare
            if campo == "Conductor":
                import unicodedata
                _norm = lambda s: unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().strip()
                if v_ticket:   v_ticket   = _norm(v_ticket)
                if v_cont:     v_cont     = _norm(v_cont)
                if aduana_val: aduana_val = _norm(aduana_val)

                # Word-level matching: al menos 2 palabras de cada fuente en Aduana
                t_words = set(v_ticket.split()) if v_ticket else set()
                c_words = set(v_cont.split()) if v_cont else set()
                a_words = set(aduana_val.split()) if aduana_val else set()

                ticket_ok = len(t_words & a_words) >= 2 if len(t_words) >= 2 else t_words.issubset(a_words) if t_words else False
                excel_ok  = len(c_words & a_words) >= 2 if len(c_words) >= 2 else c_words.issubset(a_words) if c_words else False
                ok_map[campo] = ticket_ok and excel_ok
                if not ok_map[campo]:
                    all_ok = False
                continue  # skip the generic comparison below

            # Normalize DNI: strip non-digits for comparison
            if campo == "DNI":
                v_ticket   = _re.sub(r'\D', '', v_ticket)
                v_cont     = _re.sub(r'\D', '', v_cont)
                aduana_val = _re.sub(r'\D', '', aduana_val)

            # Collect non-empty values from all 3 sources
            vals = set()
            # Precinto: ticket siempre "-", no comparar
            if campo == "Precinto":
                if v_cont:     vals.add(v_cont)
                if aduana_val: vals.add(aduana_val)
            elif campo == "Tara (kg)":
                # Tara: solo Ticket vs Excel (Aduana no tiene tara)
                if v_ticket:   vals.add(v_ticket)
                if v_cont:     vals.add(v_cont)
            else:
                if v_ticket:   vals.add(v_ticket)
                if v_cont:     vals.add(v_cont)
                if aduana_val: vals.add(aduana_val)

            # 0 or 1 unique value → match; 2+ unique → mismatch
            if len(vals) <= 1:
                ok_map[campo] = True
            else:
                # Para Precinto: si hay multi-precinto, verificar solapamiento
                if campo == "Precinto" and len(vals) > 1:
                    cont_codes = set(v_cont.split()) if v_cont else set()
                    adu_codes = set(aduana_val.split()) if aduana_val else set()
                    if cont_codes & adu_codes:
                        ok_map[campo] = True
                    else:
                        ok_map[campo] = False
                        all_ok = False
                else:
                    ok_map[campo] = False
                    all_ok = False

        estado = "ok" if all_ok else "mismatch"
        estado_label = "✅ OK" if all_ok else "❌ Diferencia"
        tag = "tag_ok" if all_ok else "tag_mismatch"

        # Tree row: all values normalized
        valores = (
            f"📄 {stem}",
            t_patente,
            t_semi,
            t_conductor,
            t_dni,
            t_neto,
            t_tara,
            t_contenedor,
            t_permiso,
            estado_label,
            a_plt,
        )

        return {
            "valores": valores,
            "tag": tag,
            "ticket": ticket_d,
            "contenedor": cont_d,
            "aduana": aduana_d,
            "ok": ok_map,
            "modo": modo,
        }

    def _procesar_resultado_control_final(self, data):
        """Inserta fila en tree_control_final."""
        try:
            stem = data.get("valores", ())[0] if data.get("valores") else "?"
            iid = self.tree_control_final.insert("", "end", values=data["valores"], tags=(data["tag"],))
            # Actualizar header según modo (terrestre → MIC/DTA, flexi → Salida Aduana)
            if len(self.tree_control_final.get_children()) == 1:
                modo = data.get("modo", "flexi")
                self.tree_control_final.heading(
                    "salida_aduana",
                    text="MIC/DTA" if modo == "terrestre" else "Salida Aduana"
                )
            self._log(f"  + Fila insertada: {stem}")
            # Store comparison data
            if not hasattr(self, '_control_final_comparacion'):
                self._control_final_comparacion = {}
            self._control_final_comparacion[iid] = {
                "ticket": data["ticket"],
                "contenedor": data["contenedor"],
                "ok": data["ok"],
                "aduana": data["aduana"],
                "modo": data.get("modo", "flexi"),
            }
        except Exception as e:
            self._log(f"Error insertando fila: {e}")

    def _abrir_comparacion_final(self, event):
        """Popup con tabla comparativa: Ticket vs Aduana vs Excel."""
        sel = self.tree_control_final.selection()
        if not sel or not hasattr(self, '_control_final_comparacion'):
            return

        datos = self._control_final_comparacion.get(sel[0])
        if not datos:
            return

        modo = datos.get("modo", "flexi")
        aduana_header = "MIC/DTA" if modo == "terrestre" else "Salida Aduana"

        def _build_popup(level):
            fsizes = self._get_font_sizes(level)
            geom = self._get_popup_geometry(700, 420, level)
            ancho = int(geom.split("x")[0])

            dlg = ctk.CTkToplevel(self)
            dlg.title("Comparación — Control Final")
            dlg.transient(self)
            dlg.grab_set()
            dlg.resizable(False, False)

            # ── Top bar: header + font level override ──────────────────
            top_bar = ctk.CTkFrame(dlg, fg_color=Palette.BG_SIDEBAR, corner_radius=6, height=36)
            top_bar.pack(fill="x", padx=12, pady=(12, 0))
            top_bar.pack_propagate(False)

            for col_i, txt in enumerate(["Campo", "Ticket (OCR)", aduana_header, "Contenedor (Excel)"]):
                lbl = ctk.CTkLabel(
                    top_bar, text=txt, width=100 if col_i == 0 else 160,
                    font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["header"], weight="bold"),
                    text_color=Palette.TEXT_SECONDARY,
                )
                kwargs = {"side": "left", "padx": 4}
                if col_i > 0:
                    kwargs["fill"] = "x"
                    kwargs["expand"] = True
                lbl.pack(**kwargs)

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

            # Cuerpo: frame normal sin scroll
            body = ctk.CTkFrame(dlg, fg_color="transparent")
            body.pack(fill="x", padx=12, pady=8)

            if modo == "terrestre":
                campos = [
                    ("Camion", "Patente", "Patente Camión"),
                    ("Semi", "Semirremolque", "Patente Semi"),
                    ("Conductor", "Conductor", "Conductor Aduana"),
                    ("DNI", "DNI", "DNI"),
                    ("Neto (kg)", "Neto (kg)", "Peso Bruto Aduana"),
                    ("Tara (kg)", "Tara (kg)", None),
                    ("Contenedor", "Contenedor", "Contenedor"),
                    ("Permiso", "Permiso", "Id Destinación"),
                    ("Precinto", "Precinto", "Precinto"),
                ]
            else:
                campos = [
                    ("Camion", "Patente", "Patente Camión"),
                    ("Semi", "Semirremolque", "Patente Semi"),
                    ("Conductor", "Conductor", "Conductor Aduana"),
                    ("DNI", "DNI", "CUIL"),
                    ("Neto (kg)", "Neto (kg)", "Peso Bruto Aduana"),
                    ("Tara (kg)", "Tara (kg)", None),
                    ("Contenedor", "Contenedor", "Contenedor"),
                    ("Permiso", "Permiso", "Id Destinación"),
                    ("Precinto", "Precinto", "Precinto"),
                ]

            # Campos con coloreado individual por celda (mayoría decide)
            campos_mayoria = {"Camion", "Semi", "DNI",
                              "Neto (kg)", "Contenedor", "Permiso"}

            for campo, key_ticket, key_aduana in campos:
                val_ticket = datos["ticket"].get(key_ticket, "")
                val_cont = datos["contenedor"].get(key_ticket, "")
                val_aduana = datos["aduana"].get(key_aduana, "—") if key_aduana else "—"
                ok = datos["ok"].get(key_ticket, False)

                vals_list = [str(val_ticket), str(val_aduana), str(val_cont)]

                if campo in campos_mayoria:
                    colors = []
                    for v in vals_list:
                        count = sum(1 for x in vals_list if x == v)
                        if count >= 2:
                            colors.append(("#C8FFC8", "#006400"))
                        else:
                            colors.append(("#FFC8C8", "#8B0000"))
                else:
                    bg = "#C8FFC8" if ok else "#FFC8C8"
                    fg = "#006400" if ok else "#8B0000"
                    colors = [(bg, fg)] * 3

                row = ctk.CTkFrame(body, fg_color="transparent")
                row.pack(fill="x", pady=1)

                ctk.CTkLabel(
                    row, text=campo, width=100, anchor="w",
                    font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["data"], weight="bold"),
                    text_color=Palette.TEXT_PRIMARY,
                ).pack(side="left", padx=(4, 0))

                for val, (bg, fg) in zip(vals_list, colors):
                    lbl = ctk.CTkLabel(
                        row, text=val, width=160, anchor="center",
                        font=ctk.CTkFont(family=FONT_FAMILY, size=fsizes["data"]),
                        fg_color=bg, text_color=fg, corner_radius=4,
                    )
                    lbl.pack(side="left", padx=4, pady=2, fill="x", expand=True)

            # Legend
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
                inner, text="Diferencia",
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

    def _cargar_datos_seleccionar_pdfs(self):
        from tkinter import filedialog as tk_filedialog
        rutas = tk_filedialog.askopenfilenames(
            title="Seleccionar PDFs y Excel CONTENEDORES",
            filetypes=[("Archivos soportados", "*.pdf *.xlsx *.xls")],
        )
        if not rutas:
            return
        rutas_pdf = [r for r in rutas if r.lower().endswith('.pdf')]
        rutas_excel = [r for r in rutas if r.lower().endswith(('.xlsx', '.xls'))]
        if not rutas_pdf:
            self._log("[Control Datos] ⚠ No seleccionaste ningún PDF.")
            return
        if not rutas_excel:
            self._log("[Control Datos] ⚠ No seleccionaste ningún archivo Excel CONTENEDORES.")
            return
        # Mostrar sección tickets, ocultar otras
        self._section_coord.pack_forget()
        self._section_final.pack_forget()
        self._section_tickets.pack(fill="both", expand=True, pady=(0, 6))
        self._log(f"[Control Datos] Procesando {len(rutas_pdf)} PDFs contra {len(rutas_excel)} Excel(s)...")
        if self.tarea_activa:
            messagebox.showwarning("Agente ocupado", "Hay una tarea en ejecución.")
            return
        self.tarea_activa = True
        for btn_name in ('btn_controlar_tickets',):
            btn = getattr(self, btn_name, None)
            if btn and btn.winfo_exists():
                btn.configure(state="disabled")
        if hasattr(self, 'progress_carga') and self.progress_carga.winfo_exists():
            self.progress_carga.pack(side="right", padx=12)
            self.progress_carga.configure(mode="indeterminate")
            self.progress_carga.start()
        t = threading.Thread(
            target=self._cargar_datos_worker,
            args=("local", list(rutas_pdf), list(rutas_excel)),
            daemon=True,
        )
        t.start()

    def _cargar_datos_escribir(self):
        """T8: Escribe valores OCR validados en las celdas DATOS del CONTENEDOR."""
        ok_count = 0
        fail_count = 0
        items = list(self.tree_carga.get_children())

        if not items:
            self._log("[Control Datos] No hay filas para escribir.")
            return

        for item in items:
            vals = self.tree_carga.item(item, "values")
            if not vals or len(vals) < 9:
                continue
            estado = vals[-1]
            if estado != "✅ Ok":
                continue

            iid = item
            if iid not in self._cargar_datos_rutas:
                self._log(f"[Control Datos] ⚠ Sin ruta almacenada para {vals[0]}")
                fail_count += 1
                continue

            ruta, k_idx = self._cargar_datos_rutas[iid]
            neto_val = vals[5]   # columna neto
            tara_val = vals[6]   # columna tara
            nombre_archivo = vals[0]

            try:
                wb = openpyxl.load_workbook(ruta)
                if "DATOS" not in wb.sheetnames:
                    self._log(f"[Control Datos] ⚠ {nombre_archivo}: no tiene hoja DATOS")
                    wb.close()
                    fail_count += 1
                    continue
                ws = wb["DATOS"]

                # Calcular coordenadas según bloque k
                row_base = 20 + (k_idx // 2) * 13
                label_col = 1 + (k_idx % 2) * 12
                value_col = label_col + 6

                # PESO CARGA = row_base + 8, value_col
                ws.cell(row=row_base + 8, column=value_col, value=float(neto_val))
                # TARA CONT = row_base + 9, value_col
                ws.cell(row=row_base + 9, column=value_col, value=float(tara_val))

                wb.save(ruta)
                wb.close()
                ok_count += 1
                self._log(f"[Control Datos] ✓ {nombre_archivo}: Neto={neto_val}, Tara={tara_val}")

            except PermissionError:
                self._log(
                    f"[Control Datos] ✗ El archivo {os.path.basename(ruta)} "
                    "está bloqueado. Cerralo e intentá de nuevo."
                )
                fail_count += 1
            except Exception as e:
                self._log(f"[Control Datos] ✗ Error escribiendo {nombre_archivo}: {e}")
                fail_count += 1

        resumen = f"Escritura completada: {ok_count} ok, {fail_count} errores"
        self._log(f"[Control Datos] {resumen}")

    # ── T5: Matching CONTENEDOR ──────────────────────────────────────
    def _cargar_cache_contenedores(self, rutas_excel=None):
        """Escanea Desktop UNA VEZ y llena self._contenedores_cache abriendo
        cada workbook UNA SOLA VEZ (PE + DATOS + Choferes).
        Si se pasan rutas_excel explícitas, las usa en lugar de escanear."""
        if rutas_excel:
            self._contenedores_cache = {}
            for ruta in rutas_excel:
                data = self._leer_wb_completo(ruta)
                if data:
                    self._contenedores_cache[ruta] = data
            return
        self._contenedores_cache = {}
        base = self._resolver_ruta("planillas_carga", "Desktop")
        if not os.path.isdir(base):
            self._log("[CACHE] Desktop no encontrado")
            return

        archivos = []
        try:
            # Archivos sueltos en Desktop
            for f in os.listdir(base):
                ruta = os.path.join(base, f)
                if os.path.isfile(ruta) and 'CONTENEDORES' in f.upper() and \
                   (f.lower().endswith('.xlsx') or f.lower().endswith('.xls')):
                    archivos.append(ruta)
            # 1 nivel de subcarpetas
            for item in os.listdir(base):
                ruta_item = os.path.join(base, item)
                if os.path.isdir(ruta_item):
                    for f in os.listdir(ruta_item):
                        if 'CONTENEDORES' in f.upper() and \
                           (f.lower().endswith('.xlsx') or f.lower().endswith('.xls')):
                            archivos.append(os.path.join(ruta_item, f))
        except Exception:
            pass

        self._log(f"[CACHE] Escaneando {len(archivos)} archivos CONTENEDORES...")
        for ruta in archivos:
            data = self._leer_wb_completo(ruta)
            if data:
                self._contenedores_cache[ruta] = data
        n = len(self._contenedores_cache)
        self._log(f"[CACHE] Caché completado: {n} archivos")

    def _match_contenedor(self, permiso_ticket):
        """Busca archivos CONTENEDORES asociados al permiso del ticket.

        Usa self._contenedores_cache (cargado por _cargar_cache_contenedores).
        NO abre workbooks — lee PE desde caché.

        Retorna lista de rutas de archivos que coinciden (ordenadas por
        estrategia 1 primero, luego estrategia 2), o lista vacía."""
        import re as _re

        if not permiso_ticket:
            self.log_queue.put((self._log_warning, "[MATCH] permiso_ticket vacío, abortando"))
            return []

        # ── Preparar patrones desde el permiso del ticket ────────────
        ticket_alfanum = _re.sub(r'[^a-zA-Z0-9]', '', str(permiso_ticket))
        self.log_queue.put((self._log_warning, f"[MATCH] Permiso limpio: '{ticket_alfanum}'"))

        if not ticket_alfanum:
            return []

        sufijo_ticket = ticket_alfanum[-5:].upper()
        self.log_queue.put((self._log_warning, f"[MATCH] Sufijo (últimos 5): '{sufijo_ticket}'"))

        codigo_corto = ""
        m = _re.search(r'0+([1-9A-Z][A-Z0-9]+)$', ticket_alfanum)
        if m:
            codigo_corto = m.group(1).upper()
            self.log_queue.put((self._log_warning, f"[MATCH] Código corto (tras ceros): '{codigo_corto}'"))

        # ── Buscar en caché en vez de abrir workbooks ──────────────
        if not self._contenedores_cache:
            self.log_queue.put((self._log_warning,
                f"[MATCH] Caché vacío, sin archivos CONTENEDORES"))
            return []

        self.log_queue.put((self._log_warning,
            f"[MATCH] Evaluando {len(self._contenedores_cache)} archivos desde caché..."))

        # ── Evaluar TODOS los archivos desde caché ─────────────────
        estrategia1 = []  # match exacto por sufijo PE
        estrategia2 = []  # código corto en nombre de carpeta

        for ruta, data in self._contenedores_cache.items():
            carpeta = os.path.basename(os.path.dirname(ruta))
            nombre = os.path.basename(ruta)
            pe_val = data.get("pe")
            try:
                if pe_val is not None:
                    pe_alfanum = _re.sub(r'[^a-zA-Z0-9]', '', str(pe_val))
                    pe_sufijo = pe_alfanum[-5:].upper() if len(pe_alfanum) >= 5 else pe_alfanum.upper()
                    if pe_sufijo == sufijo_ticket:
                        estrategia1.append(ruta)
                        self.log_queue.put((self._log_warning,
                            f"[MATCH]   ✅ E1: '{pe_sufijo}' == '{sufijo_ticket}' -> {ruta}"))
                        continue  # no evaluar E2 si ya es E1
                # Estrategia 2: código corto en carpeta
                if codigo_corto and codigo_corto in carpeta.upper():
                    estrategia2.append(ruta)
                    self.log_queue.put((self._log_warning,
                        f"[MATCH]   📁 E2: '{codigo_corto}' en carpeta -> {ruta}"))
            except Exception:
                pass

        # Combinar: E1 primero (match exacto), luego E2
        todas = estrategia1 + [r for r in estrategia2 if r not in estrategia1]

        if todas:
            self.log_queue.put((self._log_warning,
                f"[MATCH] ✅ Total candidatos: {len(todas)}"))
            return todas

        self.log_queue.put((self._log_warning,
            f"[MATCH] ❌ Sin candidatos para permiso '{permiso_ticket}' "
            f"(sufijo='{sufijo_ticket}')"))
        return []

    # ── Helper: leer PE de hoja Choferes ──────────────────────────
    def _leer_pe_choferes(self, ruta):
        """Busca etiqueta 'PE' en hoja Choferes y retorna el valor
        de la celda de la derecha. Retorna None si no encuentra."""
        # Cache check
        if ruta in getattr(self, '_contenedores_cache', {}):
            return self._contenedores_cache[ruta].get("pe")
        try:
            if ruta.lower().endswith('.xlsx'):
                wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
                if 'Choferes' not in wb.sheetnames:
                    self.log_queue.put((self._log_warning,
                        f"[MATCH]   {os.path.basename(ruta)}: NO tiene hoja Choferes"))
                    wb.close()
                    return None
                ws = wb['Choferes']
                for r in range(1, ws.max_row + 1):
                    for c in range(1, ws.max_column + 1):
                        v = ws.cell(r, c).value
                        if v is not None and str(v).strip().upper() == 'PE':
                            pe = ws.cell(r, c + 1).value
                            wb.close()
                            if pe:
                                return str(pe).strip()
                            return None
                wb.close()
                self.log_queue.put((self._log_warning,
                    f"[MATCH]   {os.path.basename(ruta)}: No se encontró etiqueta 'PE' en Choferes"))
                return None
            else:
                book = xlrd.open_workbook(ruta)
                if 'Choferes' not in book.sheet_names():
                    self.log_queue.put((self._log_warning,
                        f"[MATCH]   {os.path.basename(ruta)}: NO tiene hoja Choferes"))
                    return None
                ws = book.sheet_by_name('Choferes')
                for r in range(ws.nrows):
                    for c in range(ws.ncols):
                        v = ws.cell_value(r, c)
                        if v and str(v).strip().upper() == 'PE':
                            pe = ws.cell_value(r, c + 1)
                            if pe:
                                return str(pe).strip()
                            return None
                self.log_queue.put((self._log_warning,
                    f"[MATCH]   {os.path.basename(ruta)}: No se encontró etiqueta 'PE' en Choferes"))
                return None
        except Exception as ex:
            self.log_queue.put((self._log_warning,
                f"[MATCH]   Error al leer {os.path.basename(ruta)}: {ex}"))
            return None

    def _leer_contenedores_desde_choferes(self, ruta_wb):
        """Lee la columna 'NUMERO DE CONTENEDORES' (col D) de hoja Choferes.

        Retorna lista de strings (container numbers vacíos o no),
        en el mismo orden que aparecen en la hoja (skip header row).
        """
        # Cache check
        if ruta_wb in getattr(self, '_contenedores_cache', {}):
            return self._contenedores_cache[ruta_wb].get("contenedores", [])
        import re as _re
        contenedores = []
        try:
            if ruta_wb.lower().endswith('.xlsx'):
                wb = openpyxl.load_workbook(ruta_wb, read_only=True, data_only=True)
                if 'Choferes' not in wb.sheetnames:
                    wb.close()
                    return contenedores
                ws = wb['Choferes']
                # Buscar header "NUMERO DE CONTENEDORES" en fila 1, col D
                col_idx = None
                for c in range(1, ws.max_column + 1):
                    v = ws.cell(1, c).value
                    if v and 'NUMERO' in str(v).upper() and 'CONTENEDOR' in str(v).upper():
                        col_idx = c
                        break
                if col_idx is None:
                    col_idx = 4  # fallback a col D
                # Leer desde fila 2 hasta el final
                for r in range(2, ws.max_row + 1):
                    val = ws.cell(r, col_idx).value
                    contenedores.append(str(val).strip() if val else "")
                wb.close()
            else:
                book = xlrd.open_workbook(ruta_wb)
                if 'Choferes' not in book.sheet_names():
                    return contenedores
                ws = book.sheet_by_name('Choferes')
                col_idx = None
                for c in range(ws.ncols):
                    v = ws.cell_value(0, c)
                    if v and 'NUMERO' in str(v).upper() and 'CONTENEDOR' in str(v).upper():
                        col_idx = c
                        break
                if col_idx is None:
                    col_idx = 3  # fallback a col D (0-indexed)
                for r in range(1, ws.nrows):
                    val = ws.cell_value(r, col_idx)
                    contenedores.append(str(val).strip() if val else "")

            n_reales = sum(1 for c in contenedores if c)
            self.log_queue.put((self._log_warning,
                f"[Choferes] {n_reales} contenedores encontrados en col {col_idx+1}"))
            return contenedores
        except Exception as ex:
            self.log_queue.put((self._log_warning,
                f"[Choferes] Error leyendo contenedores: {ex}"))
            return contenedores

    # ── T5b: Lector combinado (una sola apertura) ──────────────────────
    def _leer_wb_completo(self, ruta_wb):
        """Abre el workbook UNA VEZ y extrae PE + contenedores + camiones.

        Retorna dict {pe, camiones[], contenedores[]} o None si error.
        Evita 3 opens separados para PE / DATOS / Choferes.
        """
        import re as _re
        pe = None
        contenedores = []
        precintos = []
        camiones = []
        col_idx = None
        peso_flexi_global = None

        try:
            if ruta_wb.lower().endswith('.xlsx'):
                wb = openpyxl.load_workbook(ruta_wb, read_only=True, data_only=True)

                # ── Hoja Choferes: PE ──
                if 'Choferes' in wb.sheetnames:
                    ws = wb['Choferes']
                    for r in range(1, ws.max_row + 1):
                        for c in range(1, ws.max_column + 1):
                            v = ws.cell(r, c).value
                            if v is not None and str(v).strip().upper() == 'PE':
                                pe_v = ws.cell(r, c + 1).value
                                if pe_v:
                                    pe = str(pe_v).strip()
                                break
                        if pe is not None:
                            break

                    # ── Hoja Choferes: contenedores + precintos ──
                    col_idx = None
                    precinto_col = None
                    for c in range(1, ws.max_column + 1):
                        v = ws.cell(1, c).value
                        if v:
                            h = str(v).upper()
                            if 'NUMERO' in h and 'CONTENEDOR' in h:
                                col_idx = c
                            if 'PRECINTO' in h and 'ADUANA' in h:
                                precinto_col = c
                    if col_idx is None:
                        col_idx = 4
                    precintos = []
                    for r in range(2, ws.max_row + 1):
                        val = ws.cell(r, col_idx).value
                        contenedores.append(str(val).strip() if val else "")
                        if precinto_col:
                            pv = ws.cell(r, precinto_col).value
                            precintos.append(str(pv).strip() if pv else "")
                        else:
                            precintos.append("")

                    # ── Buscar PESO FLEXI como etiqueta (no columna) ──
                    peso_flexi_global = None
                    for r in range(1, ws.max_row + 1):
                        for c in range(1, ws.max_column + 1):
                            v = ws.cell(r, c).value
                            if v is not None and 'PESO FLEXI' in str(v).upper():
                                pfv = ws.cell(r, c + 1).value
                                if pfv is not None:
                                    try:
                                        peso_flexi_global = float(str(pfv).strip())
                                    except ValueError:
                                        pass
                                break
                        if peso_flexi_global is not None:
                            break

                # ── Hoja DATOS: camiones ──
                if 'DATOS' in wb.sheetnames:
                    ws = wb['DATOS']
                    max_row = ws.max_row or 0
                    max_col = ws.max_column or 0

                    def _celda(row, col):
                        if row > max_row or col > max_col:
                            return None
                        return ws.cell(row=row, column=col).value

                    def _val_num(raw_val):
                        if raw_val is None:
                            return 0
                        if isinstance(raw_val, (int, float)):
                            return float(raw_val)
                        return 0

                    # ── Scan dinámico: busca PATENTE CAMION en DATOS ──
                    bloques = []
                    for r in range(1, max_row + 1):
                        for c in range(1, max_col + 1):
                            raw = _celda(r, c)
                            if raw is not None:
                                etiqueta = str(raw).strip().upper()
                                if 'PATENTE' in etiqueta and 'CAMION' in etiqueta:
                                    if c + 6 <= max_col:
                                        bloques.append((r, c))
                    bloques.sort(key=lambda b: (b[0], b[1]))

                    for k, (patente_row, label_col) in enumerate(bloques):
                        row_base = patente_row - 1
                        value_col = label_col + 6
                        patente_camion = str(_celda(patente_row, value_col) or "").strip()
                        patente_semi   = str(_celda(patente_row + 1, value_col) or "").strip()
                        conductor_val  = str(_celda(patente_row + 2, value_col) or "").strip()
                        dni_raw = _celda(patente_row + 3, value_col)
                        if isinstance(dni_raw, (int, float)):
                            dni_celda = str(int(dni_raw))
                        else:
                            dni_celda = str(dni_raw or "").strip()
                        peso_carga_num = _val_num(_celda(patente_row + 7, value_col))
                        tara_cont_num  = _val_num(_celda(patente_row + 8, value_col))
                        camion = {
                            "k": k,
                            "patente_camion": patente_camion,
                            "patente_semi": patente_semi,
                            "conductor": conductor_val,
                            "dni": dni_celda,
                            "contenedor": "",
                            "peso_carga": peso_carga_num,
                            "tara_cont": tara_cont_num,
                        }
                        camiones.append(camion)

                    # Merge contenedores
                    if contenedores:
                        for i, camion in enumerate(camiones):
                            if i < len(contenedores) and contenedores[i]:
                                camion["contenedor"] = contenedores[i]
                    if precintos:
                        for i, camion in enumerate(camiones):
                            if i < len(precintos) and precintos[i]:
                                camion["precinto"] = precintos[i]

                wb.close()

            else:
                # .xls con xlrd
                book = xlrd.open_workbook(ruta_wb)

                if 'Choferes' in book.sheet_names():
                    ws = book.sheet_by_name('Choferes')
                    for r in range(ws.nrows):
                        for c in range(ws.ncols):
                            v = ws.cell_value(r, c)
                            if v and str(v).strip().upper() == 'PE':
                                pe_v = ws.cell_value(r, c + 1)
                                if pe_v:
                                    pe = str(pe_v).strip()
                                break
                        if pe is not None:
                            break

                    col_idx = None
                    precinto_col = None
                    for c in range(ws.ncols):
                        v = ws.cell_value(0, c)
                        if v:
                            h = str(v).upper()
                            if 'NUMERO' in h and 'CONTENEDOR' in h:
                                col_idx = c
                            if 'PRECINTO' in h and 'ADUANA' in h:
                                precinto_col = c
                    if col_idx is None:
                        col_idx = 3
                    precintos = []
                    for r in range(1, ws.nrows):
                        val = ws.cell_value(r, col_idx)
                        contenedores.append(str(val).strip() if val else "")
                        if precinto_col is not None:
                            pv = ws.cell_value(r, precinto_col)
                            precintos.append(str(pv).strip() if pv else "")
                        else:
                            precintos.append("")

                    # ── Buscar PESO FLEXI como etiqueta (no columna) ──
                    peso_flexi_global = None
                    for r in range(ws.nrows):
                        for c in range(ws.ncols):
                            v = ws.cell_value(r, c)
                            if v and 'PESO FLEXI' in str(v).upper():
                                pfv = ws.cell_value(r, c + 1)
                                if pfv:
                                    try:
                                        peso_flexi_global = float(str(pfv).strip())
                                    except ValueError:
                                        pass
                                break
                        if peso_flexi_global is not None:
                            break

                if 'DATOS' in book.sheet_names():
                    ws = book.sheet_by_name('DATOS')
                    max_row = ws.nrows
                    max_col = ws.ncols

                    def _celda(row, col):
                        if row > max_row or col > max_col:
                            return None
                        return ws.cell_value(row - 1, col - 1)

                    def _val_num(raw_val):
                        if raw_val is None:
                            return 0
                        if isinstance(raw_val, (int, float)):
                            return float(raw_val)
                        return 0

                    # ── Scan dinámico: busca PATENTE CAMION en DATOS ──
                    bloques = []
                    for r in range(1, max_row + 1):
                        for c in range(1, max_col + 1):
                            raw = _celda(r, c)
                            if raw is not None:
                                etiqueta = str(raw).strip().upper()
                                if 'PATENTE' in etiqueta and 'CAMION' in etiqueta:
                                    if c + 6 <= max_col:
                                        bloques.append((r, c))
                    bloques.sort(key=lambda b: (b[0], b[1]))

                    for k, (patente_row, label_col) in enumerate(bloques):
                        row_base = patente_row - 1
                        value_col = label_col + 6
                        patente_camion = str(_celda(patente_row, value_col) or "").strip()
                        patente_semi   = str(_celda(patente_row + 1, value_col) or "").strip()
                        conductor_val  = str(_celda(patente_row + 2, value_col) or "").strip()
                        dni_raw = _celda(patente_row + 3, value_col)
                        if isinstance(dni_raw, (int, float)):
                            dni_celda = str(int(dni_raw))
                        else:
                            dni_celda = str(dni_raw or "").strip()
                        peso_carga_num = _val_num(_celda(patente_row + 7, value_col))
                        tara_cont_num  = _val_num(_celda(patente_row + 8, value_col))
                        camion = {
                            "k": k,
                            "patente_camion": patente_camion,
                            "patente_semi": patente_semi,
                            "conductor": conductor_val,
                            "dni": dni_celda,
                            "contenedor": "",
                            "peso_carga": peso_carga_num,
                            "tara_cont": tara_cont_num,
                        }
                        camiones.append(camion)

                    if contenedores:
                        for i, camion in enumerate(camiones):
                            if i < len(contenedores) and contenedores[i]:
                                camion["contenedor"] = contenedores[i]
                    if precintos:
                        for i, camion in enumerate(camiones):
                            if i < len(precintos) and precintos[i]:
                                camion["precinto"] = precintos[i]

        except Exception as ex:
            self.log_queue.put((self._log_warning,
                f"[CACHE] Error al leer {os.path.basename(ruta_wb)}: {ex}"))
            return None

        n_ctn = sum(1 for c in contenedores if c)
        self.log_queue.put((self._log_warning,
            f"[CACHE] Cacheado: {os.path.basename(ruta_wb)} — "
            f"PE={pe} camiones={len(camiones)} contenedores={n_ctn}"))
        return {"pe": pe, "camiones": camiones, "contenedores": contenedores,
                "peso_flexi_global": peso_flexi_global,
                "precintos": precintos}

    # ── T6: Lectura de DATOS y comparación ────────────────────────────
    def _leer_datos_contenedor(self, ruta_wb):
        """Lee la hoja DATOS del CONTENEDOR y devuelve dict con camiones.

        Retorna:
            {"camiones": [{patente_camion, patente_semi, conductor,
                           dni, peso_carga, tara_cont}, ...]}
            o None si no hay hoja DATOS.
        """
        # Cache check: si ya fue cargado por _leer_wb_completo, evitar otro open
        if ruta_wb in getattr(self, '_contenedores_cache', {}):
            camiones = self._contenedores_cache[ruta_wb].get("camiones", [])
            return {"camiones": camiones} if camiones else None

        if ruta_wb.lower().endswith('.xlsx'):
            try:
                wb = openpyxl.load_workbook(ruta_wb, read_only=True, data_only=True)
            except Exception as e:
                self.log_queue.put((self._log_warning, f"No se pudo abrir {os.path.basename(ruta_wb)}: {e}"))
                return None
            if 'DATOS' not in wb.sheetnames:
                wb.close()
                self.log_queue.put((self._log_warning, f"El archivo {os.path.basename(ruta_wb)} no tiene hoja DATOS"))
                return None
            ws = wb['DATOS']
            usar_xlrd = False
        else:
            # .xls con xlrd
            try:
                book = xlrd.open_workbook(ruta_wb)
            except Exception as e:
                self.log_queue.put((self._log_warning, f"No se pudo abrir {os.path.basename(ruta_wb)}: {e}"))
                return None
            if 'DATOS' not in book.sheet_names():
                self.log_queue.put((self._log_warning, f"El archivo {os.path.basename(ruta_wb)} no tiene hoja DATOS"))
                return None
            ws = book.sheet_by_name('DATOS')
            usar_xlrd = True

        # Obtener límites del sheet para evitar IndexError
        if usar_xlrd:
            max_row = ws.nrows
            max_col = ws.ncols
        else:
            max_row = ws.max_row or 0
            max_col = ws.max_column or 0

        def _celda(row, col):
            """Lee valor de celda (1-indexed row/col)."""
            if row > max_row or col > max_col:
                return None
            if usar_xlrd:
                return ws.cell_value(row - 1, col - 1)
            else:
                return ws.cell(row=row, column=col).value

        def _val_num(raw_val):
            """Retorna float si raw_val es numérico, 0 en otro caso."""
            if raw_val is None:
                return 0
            if isinstance(raw_val, (int, float)):
                return float(raw_val)
            return 0

        camiones = []

        # ── Scan dinámico: busca todas las etiquetas PATENTE CAMION en DATOS ──
        # Soporta layout horizontal (C1..C3 en cols 1,13,25 misma fila)
        # y el layout antiguo de 2 por bloque de filas.
        bloques = []
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                raw = _celda(r, c)
                if raw is not None:
                    etiqueta = str(raw).strip().upper()
                    if 'PATENTE' in etiqueta and 'CAMION' in etiqueta:
                        if c + 6 <= max_col:
                            bloques.append((r, c))

        # Ordenar por fila, luego columna (izquierda→derecha, arriba→abajo)
        bloques.sort(key=lambda b: (b[0], b[1]))

        for k, (patente_row, label_col) in enumerate(bloques):
            row_base = patente_row - 1  # PATENTE CAMION está en row_base + 1
            value_col = label_col + 6

            patente_camion = str(_celda(patente_row, value_col) or "").strip()
            patente_semi   = str(_celda(patente_row + 1, value_col) or "").strip()
            conductor      = str(_celda(patente_row + 2, value_col) or "").strip()
            dni_raw        = _celda(patente_row + 3, value_col)
            if isinstance(dni_raw, (int, float)):
                dni_celda = str(int(dni_raw))
            else:
                dni_celda = str(dni_raw or "").strip()
            peso_carga_num = _val_num(_celda(patente_row + 7, value_col))
            tara_cont_num  = _val_num(_celda(patente_row + 8, value_col))

            camion = {
                "k": k,
                "patente_camion": patente_camion,
                "patente_semi": patente_semi,
                "conductor": conductor,
                "dni": dni_celda,
                "contenedor": "",  # se mergea desde Choferes después
                "peso_carga": peso_carga_num,
                "tara_cont": tara_cont_num,
            }
            camiones.append(camion)

        if not usar_xlrd:
            wb.close()

        # ── Merge container numbers from Choferes sheet (col D: NUMERO DE CONTENEDORES) ──
        try:
            contenedores = self._leer_contenedores_desde_choferes(ruta_wb)
            if contenedores and camiones:
                for i, camion in enumerate(camiones):
                    if i < len(contenedores) and contenedores[i]:
                        camion["contenedor"] = contenedores[i]
                        self.log_queue.put((self._log_warning,
                            f"[DATOS]   Contenedor[{i}]: '{contenedores[i]}' -> Patente: {camion['patente_camion']}"))
        except Exception as e:
            self.log_queue.put((self._log_warning, f"[DATOS] Error merge contenedores: {e}"))

        resultado = {"camiones": camiones} if camiones else None
        self.log_queue.put((self._log_warning,
            f"[DATOS] _leer_datos_contenedor: {len(camiones)} camiones encontrados "
            f"en {os.path.basename(ruta_wb)}"
            + (f" -> {[c['patente_camion'] for c in camiones]}" if camiones else "")))
        return resultado

    # ── T4: Worker OCR ────────────────────────────────────────────────
    def _cargar_datos_worker(self, fuente, rutas, rutas_excel=None):
        """Worker que corre en threading.Thread daemon.
        
        Procesa cada PDF: OCR → extracción → matching CONTENEDOR → lectura DATOS.
        Pushea resultados a log_queue como ("_OCR_RESULT_", data).
        Al finalizar pushea ("_OCR_DONE_", None).
        
        Args:
            rutas_excel: lista de rutas a archivos CONTENEDORES seleccionados
                         por el usuario. Si es None, escanea Desktop (viejo).
        """
        self._set_log_panel("cargar-datos")
        import procesar_tickets
        import re as _re
        from concurrent.futures import ThreadPoolExecutor, as_completed
        import threading

        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put(f"[{timestamp}] [Control Datos] Fuente: {fuente} — {len(rutas)} PDF(s)")

        total = len(rutas)
        # ── 0. Cargar caché de CONTENEDORES (una vez, zero opens repetidos) ──
        self._contenedores_cache = {}
        if rutas_excel:
            self._log(f"[Control Datos] Usando {len(rutas_excel)} Excel(s) seleccionados por el usuario")
            self._cargar_cache_contenedores(rutas_excel)
        else:
            self._cargar_cache_contenedores()  # escanea Desktop (compatibilidad)
        # ── 1. OCR: según método seleccionado ──
        ocr_method = self.config.get("ocr_method", "api_vision")
        textos_por_pdf = {}
        api_datos_raw = {}  # dict crudo de API Visión para mapeo directo

        if ocr_method == "api_vision":
            api_vision_conf = self.config.get("api_vision", {})
            api_key_raw = api_vision_conf.get("api_key", "")
            api_key = self._decrypt_api_key(api_key_raw)

            if not api_key:
                self.log_queue.put(f"[{timestamp}] [Control Datos] ⚠ No hay API Key configurada, usando OCR local")
                textos_por_pdf = procesar_tickets.pdfs_a_texto_batch(rutas, engine="paddleocr")
            else:
                model = api_vision_conf.get("model", procesar_tickets.MODELO_VISION_DEFAULT)
                temperature = api_vision_conf.get("temperature", 0.1)
                max_tokens = api_vision_conf.get("max_tokens", 4000)
                timeout = api_vision_conf.get("timeout", 60)
                parallel_enabled = api_vision_conf.get("parallel_enabled", False)

                if parallel_enabled and total > 1:
                    # Parallel mode: distribute across enabled models
                    model_states = api_vision_conf.get("parallel_model_states", {})
                    enabled_models = [m for m, on in model_states.items() if on]
                    if not enabled_models:
                        enabled_models = [model]
                    # Ensure selected model is first
                    if model in enabled_models:
                        enabled_models.remove(model)
                    enabled_models.insert(0, model)

                    def _log_fn(msg):
                        self.log_queue.put(f"[{timestamp}] {msg}")

                    self.log_queue.put(f"[{timestamp}] [Control Datos] API Visión paralelo: {total} PDF(s), {len(enabled_models)} modelo(s)...")
                    result = procesar_tickets.api_vision_con_fallback(
                        rutas, api_key, enabled_models,
                        temperature=temperature, max_tokens=max_tokens,
                        timeout=timeout, log_callback=_log_fn,
                    )
                    api_datos_raw = result["datos"]
                    textos_por_pdf = result["textos"]
                else:
                    # Sequential fallback: try models one by one per PDF
                    all_models = api_vision_conf.get("custom_models", [])
                    if not all_models:
                        all_models = list(procesar_tickets.MODELOS_VISION)
                    # Ensure selected model is first
                    if model in all_models:
                        all_models.remove(model)
                    all_models.insert(0, model)

                    self.log_queue.put(f"[{timestamp}] [Control Datos] API Visión secuencial: {total} PDF(s), {len(all_models)} modelo(s) disponible(s)")
                    for i, ruta in enumerate(rutas):
                        stem = os.path.splitext(os.path.basename(ruta))[0]
                        self.log_queue.put(f"[{timestamp}] [Control Datos]   [{i+1}/{total}] {stem}")
                        success = False
                        for m in all_models:
                            try:
                                datos = procesar_tickets.api_vision_extraer_datos(
                                    ruta, api_key, model=m,
                                    temperature=temperature, max_tokens=max_tokens,
                                    timeout=timeout,
                                )
                                if "error" in datos:
                                    # Clean error message
                                    err = datos["error"]
                                    if "is not a valid model ID" in err:
                                        err = "modelo no disponible"
                                    elif "429" in err or "rate" in err.lower():
                                        err = "rate limit"
                                    elif "timeout" in err.lower():
                                        err = "timeout"
                                    self.log_queue.put(f"[{timestamp}] [Control Datos]   ERROR {m}: {err}")
                                    continue
                                # Success
                                api_datos_raw[stem] = datos
                                texto_simulado = "\n".join(f"{k}: {v}" for k, v in datos.items() if v)
                                textos_por_pdf[stem] = texto_simulado
                                self.log_queue.put(f"[{timestamp}] [Control Datos]   ✓ {stem} procesado con {m}")
                                success = True
                                break
                            except Exception as e:
                                self.log_queue.put(f"[{timestamp}] [Control Datos]   ERROR {m}: {e}")
                                continue
                        if not success:
                            self.log_queue.put(f"[{timestamp}] [Control Datos]   ⚠ Todos los modelos fallaron para {stem}, usando PaddleOCR")
                            textos_local = procesar_tickets.pdfs_a_texto_batch([ruta], engine="paddleocr")
                            textos_por_pdf[stem] = textos_local.get(stem, "")

                vision_ok = len(api_datos_raw)
                self.log_queue.put(f"[{timestamp}] [Control Datos] API Visión completado: {vision_ok}/{total} PDFs OK")
        else:
            self.log_queue.put(f"[{timestamp}] [Control Datos] OCR batch: {total} PDF(s)...")
            textos_por_pdf = procesar_tickets.pdfs_a_texto_batch(rutas, engine="paddleocr")
            self.log_queue.put(f"[{timestamp}] [Control Datos] OCR batch completado ({len(textos_por_pdf)} PDFs)")

        # ── 2. Parsear cada resultado + match CONTENEDOR ──
        def _normalizar_simple(s):
            return _re.sub(r'[^A-Z0-9]', '', str(s).upper().strip())

        def _procesar_texto(stem, texto):
            """Parsea el texto OCR y busca match CONTENEDOR.
            Si hay datos crudos de API Visión, mapea directamente."""
            
            if stem in api_datos_raw:
                # Mapeo directo desde JSON de API Visión
                raw = api_datos_raw[stem]
                patente  = raw.get("Patente", "")
                semi     = raw.get("Semirremolque", "")
                conductor = raw.get("Conductor", "")
                dni      = raw.get("DNI", "")
                permiso  = raw.get("Permiso", "")
                neto_str = raw.get("Neto", "")
                tara_str = raw.get("Tara Contenedor", "")
                contenedor_str = raw.get("Contenedor", "")
            else:
                # Parseo tradicional desde texto OCR
                datos = procesar_tickets.extraer_datos(texto)
                patente  = datos.get("Patente Camion", "")
                semi     = datos.get("Patente Acoplado", "")
                conductor_full = datos.get("Conductor", "")
                permiso  = datos.get("Merc./Permiso", "")
                neto_str = datos.get("Peso Neto (kg)", "")
                tara_str = datos.get("Tara Contenedor", "")
                contenedor_str = datos.get("Contenedor", "")
                dni = datos.get("DNI Conductor", "") or ""
                conductor = datos.get("Conductor", conductor_full) or conductor_full

            try:
                neto = float(neto_str.replace('.', '').replace(',', '.')) if neto_str else 0
            except (ValueError, AttributeError):
                neto = 0
            try:
                tara = float(tara_str.replace('.', '').replace(',', '.')) if tara_str else 0
            except (ValueError, AttributeError):
                tara = 0

            ticket_data = {
                "archivo": stem,
                "patente": patente, "semi": semi,
                "conductor": conductor, "dni": dni,
                "neto": neto, "tara": tara,
                "contenedor": contenedor_str,
                "permiso": permiso,
            }

            # Match CONTENEDOR
            rutas_match = self._match_contenedor(permiso)
            ruta_match = None
            cont_data = None
            pe_val = None

            for ruta_cand in rutas_match:
                cd = self._leer_datos_contenedor(ruta_cand)
                cv = self._leer_pe_choferes(ruta_cand)
                if cd and cd.get("camiones"):
                    pn = _normalizar_simple(patente)
                    sn = _normalizar_simple(semi)
                    dn = _re.sub(r'\D', '', str(dni).strip())
                    for camion in cd["camiones"]:
                        cp = _normalizar_simple(camion.get("patente_camion", ""))
                        cs = _normalizar_simple(camion.get("patente_semi", ""))
                        cdn = _re.sub(r'\D', '', str(camion.get("dni", "")).strip())
                        if (pn and cp and pn == cp) or \
                           (sn and cs and sn == cs) or \
                           (dn and cdn and dn == cdn):
                            ruta_match = ruta_cand
                            cont_data = cd
                            pe_val = cv
                            break
                if ruta_match:
                    break

            if not ruta_match and rutas_match:
                ruta_match = rutas_match[0]
                cont_data = self._leer_datos_contenedor(ruta_match)
                pe_val = self._leer_pe_choferes(ruta_match)

            return ticket_data, cont_data, ruta_match, pe_val, permiso

        # ── 3. Procesar resultados y enviarlos a la UI ──
        for hechos, ruta in enumerate(rutas, start=1):
            stem = os.path.splitext(os.path.basename(ruta))[0]
            has_data = stem in api_datos_raw or stem in textos_por_pdf

            if has_data:
                texto = textos_por_pdf.get(stem, "")
                ticket_data, cont_data, ruta_match, pe_val, permiso = \
                    _procesar_texto(stem, texto)
                nombre = stem

                self.log_queue.put(
                    f"[{timestamp}] [{hechos}/{total}] ✓ {nombre} — permiso {permiso}")
                if ruta_match:
                    self.log_queue.put(
                        f"     ✓ CONTENEDOR: {os.path.basename(ruta_match)}")
                else:
                    self.log_queue.put(
                        f"     ⚠ Sin CONTENEDOR match para permiso '{permiso}'")
                self.log_queue.put(("_OCR_RESULT_", {
                    "ticket": ticket_data,
                    "contenedor": cont_data,
                    "match": ruta_match,
                    "pe": pe_val,
                }))
            else:
                nombre = os.path.basename(ruta)
                self.log_queue.put(
                    f"[{timestamp}] [{hechos}/{total}] ✗ {nombre}: OCR falló")

        # Finalizar
        self.log_queue.put(("_OCR_DONE_", None))
        self._contenedores_cache = {}  # liberar memoria

    def _cargar_datos_done(self):
        """Callback al recibir _OCR_DONE_. Desbloquea UI y habilita escritura."""
        self.tarea_activa = False
        try:
            for btn_name in ('btn_controlar_tickets',):
                btn = getattr(self, btn_name, None)
                if btn and btn.winfo_exists():
                    btn.configure(state="normal")
            if hasattr(self, 'progress_carga') and self.progress_carga.winfo_exists():
                self.progress_carga.stop()
                self.progress_carga.pack_forget()
        except Exception:
            pass
        self._log("[Control Datos] OCR finalizado. Revisar resultados en la tabla.")

    def _finalizar_tarea(self):
        """Callback al recibir _TAREA_COMPLETA_. Rehabilita botones."""
        self.tarea_activa = False
        try:
            for btn_name in ('btn_control_final', 'btn_controlar_tickets', 'btn_controlar_coordinacion'):
                btn = getattr(self, btn_name, None)
                if btn and btn.winfo_exists():
                    btn.configure(state="normal")
            if hasattr(self, 'progress_carga') and self.progress_carga.winfo_exists():
                self.progress_carga.stop()
                self.progress_carga.pack_forget()
        except Exception:
            pass
        self._log("[Control] Tarea completada.")

    def _on_ocr_method_change(self, choice):
        """Persiste el método OCR elegido."""
        val = "api_vision" if choice == "API Visión" else "local"
        self.config["ocr_method"] = val
        # Habilitar/deshabilitar selector de modelo según método
        es_api = val == "api_vision"
        self._modelo_vision_menu.configure(state="normal" if es_api else "disabled")

    def _on_modelo_vision_change(self, choice):
        """Persiste el modelo de API Visión elegido."""
        self.config.setdefault("api_vision", {})["model"] = choice
        # Sincronizar también el dropdown de Ajustes si existe
        if hasattr(self, '_ent_vision_model') and self._ent_vision_model.winfo_exists():
            self._ent_vision_model.set(choice)

    def _actualizar_modelos_vision(self):
        """Sincroniza el dropdown de modelo con la config actual."""
        import procesar_tickets
        modelos_guardados = self.config.get("api_vision", {}).get("custom_models", [])
        modelos = modelos_guardados if modelos_guardados else list(procesar_tickets.MODELOS_VISION)
        modelo_default = self.config.get("api_vision", {}).get("model", procesar_tickets.MODELO_VISION_DEFAULT)
        self._modelo_vision_menu.configure(values=modelos)
        if modelo_default in modelos:
            self._modelo_vision_menu.set(modelo_default)
        elif modelos:
            self._modelo_vision_menu.set(modelos[0])

    # ═══════════════════════════════════════════════════════════════════
    # PANEL: AJUSTES
    # ═══════════════════════════════════════════════════════════════════
    def _panel_ajustes(self):
        if "ajustes" in self._panel_frames:
            return
        frame = ctk.CTkFrame(self.panel_container, fg_color="transparent")
        self._panel_frames["ajustes"] = frame
        # frame.pack(fill="both", expand=True) — packed by _cambiar_panel_forzado

        # ── Toolbar de tabs ──────────────────────────────────────────
        tabs_frame = ctk.CTkFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER,
        )
        tabs_frame.pack(fill="x", pady=(0, 4))

        self._ajustes_tabs = {}
        self._ajustes_frames = {}
        tab_names = [
            ("correo", "📧  Correo"),
            ("documentos", "📄  Documentos"),
            ("descarga", "📩  Descarga Mails"),
            ("rutas", "📁  Rutas"),
            ("valores", "💰  Valores"),
            ("seguridad", "🔒  Seguridad"),
            ("ocr", "🤖  OCR"),
            ("apariencia", "🎨  Apariencia"),
        ]

        def cambiar_tab(nombre):
            for key, fr in self._ajustes_frames.items():
                fr.pack_forget() if fr.winfo_ismapped() else None
            for key, btn in self._ajustes_tabs.items():
                if key == nombre:
                    btn.configure(fg_color=Palette.ACCENT, text_color=Palette.WHITE)
                else:
                    btn.configure(fg_color="transparent", text_color=Palette.TEXT_PRIMARY)
            # Construir contenido del tab si es primera vez que se clickea
            fr = self._ajustes_frames[nombre]
            if not fr.winfo_children() and nombre in self._ajustes_builders:
                self._ajustes_builders[nombre](fr)
            fr.pack(fill="both", expand=True)

        tab_keys = [k for k, _ in tab_names]

        def _reflow_tabs(event=None):
            fw = tabs_frame.winfo_width()
            if fw < 50:
                return
            pitch = 114  # 110 ancho + 4 padx
            cols = max(1, fw // pitch)
            for i, key in enumerate(tab_keys):
                self._ajustes_tabs[key].grid(row=i // cols, column=i % cols, padx=2, pady=2, sticky="w")

        for key, label in tab_names:
            btn = ctk.CTkButton(
                tabs_frame, text=label,
                font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
                fg_color="transparent", hover_color=Palette.BG_HOVER,
                text_color=Palette.TEXT_SECONDARY,
                border_width=1, border_color=Palette.BORDER,
                corner_radius=6, height=32, width=110,
                command=lambda k=key: cambiar_tab(k),
            )
            btn.grid(row=0, column=0)  # placeholder, _reflow_tabs reposiciona
            self._ajustes_tabs[key] = btn

        tabs_frame.bind("<Configure>", _reflow_tabs)
        tabs_frame.after_idle(_reflow_tabs)

        self._ajustes_tabs["correo"].configure(fg_color=Palette.ACCENT, text_color=Palette.WHITE)

        # ── Contenedor de contenido ───────────────────────────────────
        content_frame = ctk.CTkFrame(frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True)

        for key, _ in tab_names:
            sub = ctk.CTkScrollableFrame(
                content_frame, fg_color=Palette.BG_CARD, corner_radius=8,
                border_width=1, border_color=Palette.BORDER,
            )
            self._ajustes_frames[key] = sub

        # Builders diferidos — el contenido de cada tab se crea al primer clic
        self._ajustes_builders = {
            "correo": self._ajustes_tab_correo,
            "documentos": self._ajustes_tab_documentos,
            "descarga": self._ajustes_tab_descarga,
            "rutas": self._ajustes_tab_rutas,
            "valores": self._ajustes_tab_valores,
            "seguridad": self._ajustes_tab_seguridad,
            "ocr": self._ajustes_tab_ocr,
            "apariencia": self._ajustes_tab_apariencia,
        }
        # Solo construir el tab Correo (visible por defecto)
        self._ajustes_tab_correo(self._ajustes_frames["correo"])
        self._ajustes_frames["correo"].pack(fill="both", expand=True)

        # ── Botón Guardar ────────────────────────────────────────────
        save_frame = ctk.CTkFrame(frame, fg_color=Palette.BG_CARD, corner_radius=8,
                                  border_width=1, border_color=Palette.BORDER, height=44)
        save_frame.pack(fill="x", pady=(4, 0))
        save_frame.pack_propagate(False)

        ctk.CTkButton(
            save_frame,
            text="💾  Guardar Cambios",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            fg_color=Palette.SUCCESS, hover_color="#00c853",
            text_color=Palette.WHITE, corner_radius=6, height=34, width=190,
            command=self._guardar_ajustes,
        ).pack(side="left", padx=10, pady=4)

        self._ajustes_lbl_status = ctk.CTkLabel(
            save_frame, text="",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.SUCCESS,
        )
        self._ajustes_lbl_status.pack(side="left", padx=8)

    def _ajustes_tab_apariencia(self, parent):
        """Tab Apariencia en Ajustes: configuración visual (tamaño de fuentes en popups)."""
        self._ajustes_seccion(parent, "Tamaño de Tablas Comparativas")

        ctk.CTkLabel(
            parent,
            text="Ajusta el tamaño de fuente de las ventanas de comparación (Tickets, Coordinación, Final). "
                 "El nivel se aplica al reiniciar o abrir un popup.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_MUTED,
            wraplength=480,
            justify="left",
        ).pack(anchor="w", padx=14, pady=(0, 8))

        # Font level selector
        font_level_row = ctk.CTkFrame(parent, fg_color="transparent")
        font_level_row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(
            font_level_row, text="Tamaño de tablas (1-3)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.TEXT_SECONDARY,
        ).pack(anchor="w")
        self._ent_font_level = ctk.CTkOptionMenu(
            font_level_row,
            values=["1", "2", "3"],
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_INPUT, button_color=Palette.ACCENT,
            button_hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.TEXT_PRIMARY,
            dropdown_fg_color=Palette.BG_CARD,
            dropdown_hover_color=Palette.BG_HOVER,
            dropdown_text_color=Palette.TEXT_PRIMARY,
            width=80, height=30,
        )
        self._ent_font_level.set(str(self.config.get("font_level", 1)))
        self._ent_font_level.pack(anchor="w", pady=(2, 0))

    # ── Helpers de layout ────────────────────────────────────────────
    def _ajustes_seccion(self, parent, titulo):
        ctk.CTkLabel(
            parent, text=titulo,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(anchor="w", padx=14, pady=(12, 4))

    def _ajustes_row(self, parent, label, default="", show="", extra=None, width=260, toggle_pw=False):
        """Crea fila compacta: label arriba, entry abajo (ancho fijo). Retorna el CTkEntry."""
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(
            row, text=label,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.TEXT_SECONDARY, anchor="w",
        ).pack(anchor="w")
        if toggle_pw:
            # Entry + botón lado a lado
            erow = ctk.CTkFrame(row, fg_color="transparent")
            erow.pack(anchor="w", pady=(2, 0))
            e = ctk.CTkEntry(
                erow, width=width - 34, height=30,
                font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
                text_color=Palette.TEXT_PRIMARY, corner_radius=4,
                show=show,
            )
            e.pack(side="left")
            self._pw_visible = False
            def _toggle():
                self._pw_visible = not self._pw_visible
                e.configure(show="" if self._pw_visible else "*")
            ctk.CTkButton(erow, text="👁", width=30, height=30,
                          font=ctk.CTkFont(size=14),
                          fg_color=Palette.BG_HOVER, hover_color=Palette.ACCENT_DIM,
                          text_color=Palette.TEXT_MUTED, corner_radius=4,
                          command=_toggle).pack(side="left", padx=(4, 0))
        else:
            e = ctk.CTkEntry(
                row, width=width, height=30,
                font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
                text_color=Palette.TEXT_PRIMARY, corner_radius=4,
                show=show,
            )
            e.pack(anchor="w", pady=(2, 0))
        e.insert(0, str(default) if default else "")
        if extra:
            ctk.CTkLabel(
                row, text=extra,
                font=ctk.CTkFont(family=FONT_FAMILY, size=10),
                text_color=Palette.TEXT_MUTED, anchor="w",
            ).pack(anchor="w", pady=(1, 0))
        return e

    # ── TAB: CORREO ──────────────────────────────────────────────────
    def _ajustes_tab_correo(self, parent):
        self._ajustes_seccion(parent, "Credenciales IMAP")
        self._ent_correo_usuario = self._ajustes_row(
            parent, "Usuario (email):", self._cfg_obtener_correo("usuario", ""), width=300)
        self._ent_correo_password = self._ajustes_row(
            parent, "Contraseña:", self._cfg_obtener_correo("password", ""),
            show="*", width=240, toggle_pw=True)
        self._ent_correo_imap = self._ajustes_row(
            parent, "Servidor IMAP:", self._cfg_obtener_correo("imap_server", IMAP_SERVER), width=280)
        self._ent_correo_puerto = self._ajustes_row(
            parent, "Puerto IMAP:", str(self._cfg_obtener_correo("imap_puerto", PUERTO_IMAP)), width=80)

        self._ajustes_seccion(parent, "Destinatarios — Planillas de Carga")
        default_grupal = "\n".join(self._cfg_obtener_correo("destinatarios_grupal", DESTINATARIOS_GRUPAL))
        self._ajustes_texto_grupal = ctk.CTkTextbox(
            parent, height=100, width=400,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=4, border_width=1,
        )
        self._ajustes_texto_grupal.insert("1.0", default_grupal)
        self._ajustes_texto_grupal.pack(anchor="w", padx=14, pady=(2, 8))

        self._ajustes_seccion(parent, "Destinatarios — Correo Individual")
        default_ind = "\n".join(self._cfg_obtener_correo("destinatarios_individual", DESTINATARIOS_INDIVIDUAL))
        self._ajustes_texto_ind = ctk.CTkTextbox(
            parent, height=70, width=400,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=4, border_width=1,
        )
        self._ajustes_texto_ind.insert("1.0", default_ind)
        self._ajustes_texto_ind.pack(anchor="w", padx=14, pady=(2, 8))

        self._ajustes_seccion(parent, "Remitente Balanza")
        self._ent_remitente_balanza = self._ajustes_row(
            parent, "Remitente Balanza:", self.config.get("correo", {}).get("remitente_balanza", ""), width=300)

    # ── TAB: DOCUMENTOS ──────────────────────────────────────────────
    def _ajustes_tab_documentos(self, parent):
        # Contenedor horizontal: izquierda (dorsos) | derecha (copias)
        split = ctk.CTkFrame(parent, fg_color="transparent")
        split.pack(fill="x", padx=10, pady=(4, 0))

        # ── Columna Izquierda: Dorsos ────────────────────────────────
        col_izq = ctk.CTkFrame(split, fg_color="transparent")
        col_izq.pack(side="left", fill="y", padx=(4, 8))

        self._ajustes_seccion(col_izq, "Hojas — Dorsos")
        self._ent_dorso_mic = self._ajustes_row(
            col_izq, "Dorso MIC:", str(self._cfg_obtener_docs("dorso_mic", 15)), width=70)
        self._ent_dorso_crt = self._ajustes_row(
            col_izq, "Dorso CRT:", str(self._cfg_obtener_docs("dorso_crt", 4)), width=70)
        self._ent_dorso_pe = self._ajustes_row(
            col_izq, "Dorso PE:", str(self._cfg_obtener_docs("dorso_pe", 2)), width=70)

        # ── Columna Derecha: Copias ──────────────────────────────────
        col_der = ctk.CTkFrame(split, fg_color="transparent")
        col_der.pack(side="left", fill="y", padx=(8, 4))

        self._ajustes_seccion(col_der, "Copias — Impresión")
        self._ent_permiso_exp = self._ajustes_row(
            col_der, "Permiso de Exportación:", str(self._cfg_obtener_docs("permiso_exportacion", 2)), width=70)
        self._ent_hoja_ruta = self._ajustes_row(
            col_der, "Hoja de Ruta:", str(self._cfg_obtener_docs("hoja_ruta", 2)), width=70)
        self._ent_sobre = self._ajustes_row(
            col_der, "Sobre (Planilla de Carga):", str(self._cfg_obtener_docs("sobre", 1)), width=70)

    def _ajustes_row_browse(self, parent, label, default="", extra=None, width=290):
        """Como _ajustes_row pero con botón 📂 Examinar.
           Detecta automáticamente si es pendrive (ruta relativa) o disco fijo (absoluta)."""
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(
            row, text=label,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.TEXT_SECONDARY, anchor="w",
        ).pack(anchor="w")
        entry_row = ctk.CTkFrame(row, fg_color="transparent")
        entry_row.pack(fill="x", pady=(2, 0))
        e = ctk.CTkEntry(
            entry_row, width=width, height=30,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=4,
        )
        e.insert(0, str(default) if default else "")
        e.pack(side="left")
        btn = ctk.CTkButton(
            entry_row, text="📂 Examinar", width=100, height=30,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=Palette.BG_HOVER, hover_color=Palette.ACCENT_DIM,
            text_color=Palette.TEXT_SECONDARY, corner_radius=4,
            command=lambda: self._ajustes_examinar_carpeta(e),
        )
        btn.pack(side="left", padx=(6, 0))
        if extra:
            ctk.CTkLabel(
                row, text=extra,
                font=ctk.CTkFont(family=FONT_FAMILY, size=10),
                text_color=Palette.TEXT_MUTED, anchor="w",
            ).pack(anchor="w", pady=(1, 0))
        return e

    def _ajustes_examinar_carpeta(self, entry_widget):
        r"""Abre diálogo para seleccionar carpeta.
           Si la unidad es removible (pendrive), recorta la letra (F:\TRABAJO → TRABAJO).
           Si es disco fijo/red, deja la ruta absoluta completa."""
        from tkinter import filedialog as tk_filedialog
        ruta = tk_filedialog.askdirectory(title="Seleccionar carpeta")
        if ruta:
            ruta = ruta.replace("/", "\\")
            match = re.match(r"^([A-Za-z]):\\(.*)", ruta)
            if match:
                letra = match.group(1)
                resto = match.group(2)
                # Detectar si es pendrive (DRIVE_REMOVABLE = 2) o disco fijo
                try:
                    from ctypes import windll
                    tipo = windll.kernel32.GetDriveTypeW(f"{letra}:\\")
                    if tipo == 2:  # DRIVE_REMOVABLE → pendrive
                        ruta = resto
                except Exception:
                    pass  # si falla, deja la ruta como está
            entry_widget.delete(0, "end")
            entry_widget.delete(0, "end")
            entry_widget.insert(0, ruta)

    # ── TAB: RUTAS ───────────────────────────────────────────────────
    def _ajustes_tab_rutas(self, parent):
        self._ajustes_seccion(parent, "Archivos en Pendrive / PC — Carpetas de búsqueda")
        ctk.CTkLabel(
            parent,
            text="Usá Examinar para elegir la carpeta. Si es un pendrive, la letra\n"
                 "se recorta sola. Si es un disco fijo, se guarda la ruta absoluta.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_MUTED,
        ).pack(anchor="w", padx=14, pady=(0, 4))

        self._ent_ruta_sobres = self._ajustes_row_browse(
            parent, "SOBRES (carpeta):",
            self._cfg_obtener_rutas("sobres", "TRABAJO\\01_PLANILLAS"),
            extra="Busca SOBRES_2026.xlsx")
        self._ent_ruta_cobro = self._ajustes_row_browse(
            parent, "COBRO (carpeta):",
            self._cfg_obtener_rutas("cobro", "TRABAJO\\01_PLANILLAS"),
            extra="Busca COBRO_2026.xlsx")
        self._ent_ruta_pc = self._ajustes_row_browse(
            parent, "PC (carpeta):",
            self._cfg_obtener_rutas("pc", "TRABAJO\\01_PLANILLAS"),
            extra="Busca PC.xlsx / PC_2026.xlsx")

        # Separador
        ctk.CTkFrame(parent, fg_color=Palette.DIVIDER, height=1).pack(fill="x", padx=14, pady=(10, 8))

        self._ajustes_seccion(parent, "Planilla Maestra — CARGA TERRESTRE")
        self._ent_ruta_ct_carpeta = self._ajustes_row_browse(
            parent, "Carpeta:",
            self._cfg_obtener_rutas("carga_terrestre_carpeta", "TRABAJO\\01_PLANILLAS"))
        self._ent_ruta_ct_nombre = self._ajustes_row(
            parent, "Nombre del archivo:",
            self._cfg_obtener_rutas("carga_terrestre_nombre", "CARGA TERRESTRE.xlsx"), width=280)

        # Separador
        ctk.CTkFrame(parent, fg_color=Palette.DIVIDER, height=1).pack(fill="x", padx=14, pady=(10, 8))

        self._ajustes_seccion(parent, "Contenedores de Carga y Descargas")
        self._ent_ruta_planillas = self._ajustes_row_browse(
            parent, "Carpeta de Contenedores de Carga:",
            self._cfg_obtener_rutas("planillas_carga", "Desktop"),
            extra="Donde se buscan las carpetas con 'PLANILLA DE CARGA'")
        self._ent_ruta_descarga = self._ajustes_row_browse(
            parent, "Carpeta de descarga de mails:",
            self._cfg_obtener_rutas("descarga_mails", "Desktop"),
            extra="Donde se guardan los adjuntos descargados de los mails")

        # Escritorio
        ctk.CTkFrame(parent, fg_color=Palette.DIVIDER, height=1).pack(fill="x", padx=14, pady=(10, 8))
        self._ajustes_seccion(parent, "Nombre del Escritorio")
        self._ent_ruta_escritorio = self._ajustes_row(
            parent, "Nombre de la carpeta Escritorio:",
            self._cfg_obtener_rutas("escritorio_nombre", "Desktop"),
            extra="Solo cambiar si tu SO usa otro nombre (ej: 'Desktop' en inglés)", width=200)

        # Backup y Sellos
        ctk.CTkFrame(parent, fg_color=Palette.DIVIDER, height=1).pack(fill="x", padx=14, pady=(10, 8))
        self._ajustes_seccion(parent, "Backup Pendrive y Sellos")
        self._ent_ruta_backup = self._ajustes_row_browse(
            parent, "Backup Pendrive (carpeta):",
            self._cfg_obtener_rutas("backup_pendrive", "TRABAJO\\CARGAS"))
        self._ent_ruta_mic_sellos = self._ajustes_row_browse(
            parent, "FECHA MIC Y SELLOS (carpeta):",
            self._cfg_obtener_rutas("mic_sellos", "TRABAJO\\01_PLANILLAS"),
            extra="Busca FECHA MIC Y SELLOS.xlsx")
        self._ent_ruta_crt_original = self._ajustes_row_browse(
            parent, "FECHA CRT Y ORIGINAL (carpeta):",
            self._cfg_obtener_rutas("crt_original", "TRABAJO\\01_PLANILLAS"),
            extra="Busca FECHA CRT Y ORIGINAL.xlsx")

    # ── TAB: VALORES ──────────────────────────────────────────────────
    def _ajustes_tab_descarga(self, parent):
        self._ajustes_seccion(parent, "Cantidad de Mails a Buscar")
        ctk.CTkLabel(
            parent,
            text="Estos valores se usan en el panel Correos para limitar la búsqueda.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_MUTED,
        ).pack(anchor="w", padx=14, pady=(0, 4))

        dm = self.config.get("descarga_mails", {})
        papeles_default = dm.get("papeles", "2")
        reglas_default = dm.get("reglas", "4")
        sin_filtro_default = dm.get("sin_filtro", "20")

        self._ent_descarga_papeles = self._ajustes_row(
            parent, "📥  Buscar y Descargar (papeles):", papeles_default,
            extra="Mails más nuevos con 'papeles' en el asunto.", width=80)
        self._ent_descarga_reglas = self._ajustes_row(
            parent, "🔍  Buscar con reglas:", reglas_default,
            extra="Mails que coinciden con las reglas configuradas.", width=80)
        self._ent_descarga_sin_filtro = self._ajustes_row(
            parent, "📋  Mail sin filtros:", sin_filtro_default,
            extra="Últimos mails sin ningún filtro.", width=80)

    def _ajustes_tab_valores(self, parent):
        self._ajustes_seccion(parent, "Tarifas de Planilla COBRO")
        self._ent_precio_carpeta = self._ajustes_row(
            parent, "Valor Carpeta ($):",
            str(self._cfg_obtener("valores", "precio_carpeta", 49000)),
            extra="Primer ítem de cada fecha. Los siguientes llevan la mitad.",
            width=120,
        )
        self._ent_ata_tares = self._ajustes_row(
            parent, "ATA y TARES ($):",
            str(self._cfg_obtener("valores", "ata_tares", 65000)),
            extra="Valor del servicio ATA por contenedor.",
            width=120,
        )

        self._ajustes_seccion(parent, "Guardas Disponibles")
        ctk.CTkLabel(
            parent,
            text="Un guarda por línea. Se usan en 'Agregar Guarda' del panel Contenedores.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_MUTED,
        ).pack(anchor="w", padx=14, pady=(0, 4))

        guardas_actuales = self._cfg_obtener("valores", "guardas", ["Gonzalez", "Rodriguez", "Martinez", "Perez"])
        self._ajustes_texto_guardas = ctk.CTkTextbox(
            parent, height=100, width=320,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=4,
        )
        self._ajustes_texto_guardas.pack(anchor="w", padx=14, pady=(0, 8))
        self._ajustes_texto_guardas.insert("1.0", "\n".join(guardas_actuales))

        # ── Config Súper Auto ──────────────────────────────────────
        self._ajustes_seccion(parent, "⚡ Súper Auto — Funciones Automáticas")
        sa = self._cfg_obtener("super_auto", "pasos", {})
        self._super_check_sobre    = ctk.BooleanVar(value=sa.get("sobre", True))
        self._super_check_permiso  = ctk.BooleanVar(value=sa.get("permiso", True))
        self._super_check_hoja_ruta = ctk.BooleanVar(value=sa.get("hoja_ruta", True))
        self._super_check_recibo   = ctk.BooleanVar(value=sa.get("recibo_ata", True))
        self._super_check_guarda   = ctk.BooleanVar(value=sa.get("aplicar_guarda", True))
        self._super_check_planillas = ctk.BooleanVar(value=sa.get("completar_planillas", True))

        for var, txt in [(self._super_check_sobre, "☑ Imprimir Sobre (1 copia)"),
                          (self._super_check_permiso, "☑ Imprimir Permiso (2 copias)"),
                          (self._super_check_hoja_ruta, "☑ Imprimir Hoja de Ruta (2 copias)"),
                          (self._super_check_recibo, "☑ Imprimir Recibo ATA (1 copia)"),
                          (self._super_check_guarda, "☑ Aplicar Guarda"),
                          (self._super_check_planillas, "☑ Completar Planillas")]:
            ctk.CTkCheckBox(parent, text=txt, variable=var,
                            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
                            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
                            text_color=Palette.TEXT_PRIMARY,
                            border_color=Palette.BORDER).pack(anchor="w", padx=20, pady=3)

    # ── TAB: SEGURIDAD ───────────────────────────────────────────────
    def _ajustes_tab_seguridad(self, parent):
        self._ajustes_seccion(parent, "Contraseña Maestra")
        ctk.CTkLabel(
            parent,
            text="Protege el acceso a la aplicación y los Ajustes.\n"
                 "Dejá ambos campos vacíos para desactivar la protección.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_MUTED,
        ).pack(anchor="w", padx=14, pady=(0, 8))

        pw_actual = self._cfg_obtener("seguridad", "password", "")
        self._ent_master_pw = self._ajustes_row(
            parent, "Contraseña nueva:", pw_actual, show="*", width=220)
        self._ent_master_pw_confirm = self._ajustes_row(
            parent, "Confirmar contraseña:", pw_actual, show="*", width=220,
            extra="Debe coincidir con la anterior. Vacío = sin protección.")

        self._ajustes_seccion(parent, "Recuperación")
        ctk.CTkLabel(
            parent,
            text="Si olvidás la contraseña, se enviará al correo configurado en la pestaña Correo.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_MUTED,
        ).pack(anchor="w", padx=14, pady=(0, 4))

    # ── TAB: OCR ─────────────────────────────────────────────────────
    def _verificar_motores_ocr(self):
        """Verifica Tesseract y PaddleOCR después de que el panel se renderice (evita demora al abrir)."""
        import procesar_tickets
        # Tesseract — priorizar sidecar portátil sobre instalación del sistema
        import pytesseract
        tess_sidecar = os.path.join(procesar_tickets.TESSERACT_SIDECAR, "tesseract.exe")
        tess_ok = False
        tess_path = ""
        if os.path.isfile(tess_sidecar):
            try:
                pytesseract.pytesseract.tesseract_cmd = tess_sidecar
                pytesseract.get_tesseract_version()
                tess_ok = True
                tess_path = tess_sidecar
            except Exception:
                pass
        if not tess_ok:
            try:
                pytesseract.pytesseract.tesseract_cmd = procesar_tickets.TESSERACT_CMD
                pytesseract.get_tesseract_version()
                tess_ok = True
                tess_path = procesar_tickets.TESSERACT_CMD
            except Exception:
                tess_ok = False
                tess_path = ""
        if hasattr(self, '_ocr_lbl_tesseract') and self._ocr_lbl_tesseract.winfo_exists():
            self._ocr_lbl_tesseract.configure(
                text="✓ Disponible" if tess_ok else "✗ No encontrado",
                text_color=Palette.SUCCESS if tess_ok else Palette.ERROR,
            )
        if hasattr(self, '_ocr_lbl_tess_path') and self._ocr_lbl_tess_path.winfo_exists():
            self._ocr_lbl_tess_path.configure(text=f"({tess_path})" if tess_path else "")

        # PaddleOCR — background thread (no congelar la UI)
        if hasattr(self, '_ocr_lbl_paddle') and self._ocr_lbl_paddle.winfo_exists():
            self._ocr_lbl_paddle.configure(
                text="⏳ Verificando...",
                text_color=Palette.TEXT_MUTED,
            )
        threading.Thread(target=self._verificar_paddle_bg, daemon=True).start()

    def _verificar_paddle_bg(self):
        """Corre _detectar_paddle() en background y actualiza UI via after()."""
        import procesar_tickets
        paddle_ok = procesar_tickets._detectar_paddle() is not None
        paddle_path = procesar_tickets.PADDLE_PORTABLE_PYTHON if paddle_ok else procesar_tickets.PADDLE_SIDECAR
        self.after(0, lambda: self._actualizar_estado_paddle(paddle_ok, paddle_path))

    def _actualizar_estado_paddle(self, ok: bool, path: str):
        """Callback desde el thread: actualiza labels de PaddleOCR en la UI."""
        if hasattr(self, '_ocr_lbl_paddle') and self._ocr_lbl_paddle.winfo_exists():
            self._ocr_lbl_paddle.configure(
                text="✓ Disponible" if ok else "✗ No encontrado",
                text_color=Palette.SUCCESS if ok else Palette.ERROR,
            )
        if hasattr(self, '_ocr_lbl_paddle_path') and self._ocr_lbl_paddle_path.winfo_exists():
            self._ocr_lbl_paddle_path.configure(text=f"({path})" if path else "")

    def _ajustes_tab_ocr(self, parent):
        """Tab OCR en Ajustes: muestra estado de disponibilidad de motores."""
        import procesar_tickets

        self._ajustes_seccion(parent, "Motores OCR")

        ctk.CTkLabel(
            parent,
            text="Tesseract → permiso de exportación (rápido).  PaddleOCR → tickets escaneados (preciso).",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_MUTED,
            wraplength=480,
            justify="left",
        ).pack(anchor="w", padx=14, pady=(0, 8))

        self._ajustes_seccion(parent, "Estado de Motores")

        # Crear labels (placeholder), verificar motores después
        row1 = ctk.CTkFrame(parent, fg_color="transparent")
        row1.pack(fill="x", padx=14, pady=2)
        ctk.CTkLabel(
            row1, text="Tesseract:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.TEXT_SECONDARY,
        ).pack(side="left")
        self._ocr_lbl_tesseract = ctk.CTkLabel(
            row1, text="⌛ Verificando...",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=Palette.TEXT_MUTED,
        )
        self._ocr_lbl_tesseract.pack(side="left", padx=(8, 0))
        self._ocr_lbl_tess_path = ctk.CTkLabel(
            row1, text="",
            font=ctk.CTkFont(family=FONT_MONO, size=10),
            text_color=Palette.TEXT_MUTED,
        )
        self._ocr_lbl_tess_path.pack(side="left", padx=(4, 0))

        row2 = ctk.CTkFrame(parent, fg_color="transparent")
        row2.pack(fill="x", padx=14, pady=2)
        ctk.CTkLabel(
            row2, text="PaddleOCR:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.TEXT_SECONDARY,
        ).pack(side="left")
        self._ocr_lbl_paddle = ctk.CTkLabel(
            row2, text="⌛ Verificando...",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            text_color=Palette.TEXT_MUTED,
        )
        self._ocr_lbl_paddle.pack(side="left", padx=(8, 0))
        self._ocr_lbl_paddle_path = ctk.CTkLabel(
            row2, text="",
            font=ctk.CTkFont(family=FONT_MONO, size=10),
            text_color=Palette.TEXT_MUTED,
        )
        self._ocr_lbl_paddle_path.pack(side="left", padx=(4, 0))

        # Verificar motores después de que el panel se haya renderizado
        parent.after(100, self._verificar_motores_ocr)

        # ── API Visión (OpenRouter) ───────────────────────────────────
        self._ajustes_seccion(parent, "API Visión (OpenRouter)")

        ctk.CTkLabel(
            parent,
            text="API para extraer datos de tickets escaneados mediante modelos con visión. Usa OpenRouter como gateway.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_MUTED, wraplength=480, justify="left",
        ).pack(anchor="w", padx=14, pady=(0, 8))

        # API Key (encriptada)
        self._ent_vision_api_key = self._ajustes_row(
            parent, "API Key",
            default=self.config.get("api_vision", {}).get("api_key", ""),
            show="*", width=300, toggle_pw=True,
        )

        # Modelo (dropdown) — mismo layout que _ajustes_row: label arriba, control abajo
        modelos_guardados = self.config.get("api_vision", {}).get("custom_models", [])
        modelos = modelos_guardados if modelos_guardados else list(procesar_tickets.MODELOS_VISION)
        modelo_default = self.config.get("api_vision", {}).get("model", procesar_tickets.MODELO_VISION_DEFAULT)
        vis_row = ctk.CTkFrame(parent, fg_color="transparent")
        vis_row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(
            vis_row, text="Modelo",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.TEXT_SECONDARY, anchor="w",
        ).pack(anchor="w")
        self._ent_vision_model = ctk.CTkOptionMenu(
            vis_row,
            values=modelos,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_INPUT, button_color=Palette.ACCENT,
            button_hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.TEXT_PRIMARY,
            dropdown_fg_color=Palette.BG_CARD,
            dropdown_hover_color=Palette.BG_HOVER,
            dropdown_text_color=Palette.TEXT_PRIMARY,
            corner_radius=4, width=320, height=30,
        )
        self._ent_vision_model.set(modelo_default if modelo_default in modelos else modelos[0] if modelos else "")
        self._ent_vision_model.pack(anchor="w", pady=(2, 0))

        # Modelos disponibles (text box, editar/quitar modelos)
        ctk.CTkLabel(
            parent, text="Modelos disponibles (uno por línea)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_MUTED, anchor="w",
        ).pack(anchor="w", padx=14, pady=(6, 0))
        self._ent_vision_custom_models = ctk.CTkTextbox(
            parent, height=70, width=400,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=4, border_width=1,
        )
        self._ent_vision_custom_models.insert("1.0", "\n".join(modelos))
        self._ent_vision_custom_models.pack(anchor="w", padx=14, pady=(2, 8))
        self._ent_vision_custom_models.bind("<KeyRelease>", lambda e: self._sync_modelos_desde_textbox())
        self._sync_modelos_desde_textbox()

        # ── API Vision en Paralelo ───────────────────────────────────
        parallel_cfg = self.config.get("api_vision", {})
        self._chk_parallel_enabled = ctk.CTkCheckBox(
            parent, text="Habilitar API Vision en Paralelo",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            border_color=Palette.BORDER, checkmark_color=Palette.WHITE,
            text_color=Palette.TEXT_PRIMARY,
            command=self._toggle_parallel_panel,
        )
        self._chk_parallel_enabled.pack(anchor="w", padx=14, pady=(4, 2))
        if parallel_cfg.get("parallel_enabled", False):
            self._chk_parallel_enabled.select()

        # Model panel frame (hidden initially)
        self._parallel_panel = ctk.CTkFrame(parent, fg_color="transparent")
        ctk.CTkLabel(
            self._parallel_panel, text="Modelos para procesamiento paralelo",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_MUTED, anchor="w",
        ).pack(anchor="w", padx=14, pady=(2, 0))
        self._parallel_checks_frame = ctk.CTkFrame(self._parallel_panel, fg_color="transparent")
        self._parallel_checks_frame.pack(anchor="w", padx=14, pady=(0, 4))
        self._parallel_model_checks = {}

        # Build checkboxes for existing models
        saved_states = parallel_cfg.get("parallel_model_states", {})
        for m in modelos:
            is_enabled = saved_states.get(m, True)
            chk = ctk.CTkCheckBox(
                self._parallel_checks_frame, text=m,
                font=ctk.CTkFont(family=FONT_FAMILY, size=11),
                fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
                border_color=Palette.BORDER, checkmark_color=Palette.WHITE,
                text_color=Palette.TEXT_PRIMARY,
            )
            chk.pack(anchor="w", padx=2, pady=1)
            if is_enabled:
                chk.select()
            self._parallel_model_checks[m] = chk

        self._toggle_parallel_panel()

        # Temperature
        self._ent_vision_temperature = self._ajustes_row(
            parent, "Temperature (0-1)",
            default=str(self.config.get("api_vision", {}).get("temperature", 0.1)),
            width=100,
        )

        # Max Tokens
        self._ent_vision_max_tokens = self._ajustes_row(
            parent, "Max Tokens",
            default=str(self.config.get("api_vision", {}).get("max_tokens", 4000)),
            width=100,
        )

        # Timeout
        self._ent_vision_timeout = self._ajustes_row(
            parent, "Timeout (segundos)",
            default=str(self.config.get("api_vision", {}).get("timeout", 60)),
            width=100,
        )

    # ── Sincronizar dropdown de modelo desde el textbox ─────────────
    def _sync_modelos_desde_textbox(self):
        raw = self._ent_vision_custom_models.get("1.0", "end-1c")
        lines = [l.strip() for l in raw.split("\n") if l.strip()]
        if lines:
            current = self._ent_vision_model.get()
            self._ent_vision_model.configure(values=lines)
            if current in lines:
                self._ent_vision_model.set(current)
            else:
                self._ent_vision_model.set(lines[0])
            # También sincronizar el dropdown de la top_bar si existe
            if hasattr(self, '_modelo_vision_menu') and self._modelo_vision_menu.winfo_exists():
                top_current = self._modelo_vision_menu.get()
                self._modelo_vision_menu.configure(values=lines)
                if top_current in lines:
                    self._modelo_vision_menu.set(top_current)
                elif lines:
                    self._modelo_vision_menu.set(lines[0])
        # Rebuild parallel model checkboxes when model list changes
        if hasattr(self, '_parallel_model_checks'):
            self._rebuild_parallel_model_checks()

    # ── Toggle panel de modelos paralelos ───────────────────────────
    def _toggle_parallel_panel(self):
        if self._chk_parallel_enabled.get() == 1:
            self._parallel_panel.pack(anchor="w", padx=0, pady=(0, 4), after=self._chk_parallel_enabled)
        else:
            self._parallel_panel.pack_forget()

    # ── Reconstruir checkboxes de modelos paralelos ─────────────────
    def _rebuild_parallel_model_checks(self):
        """Rebuild parallel model checkboxes from the textbox content."""
        raw = self._ent_vision_custom_models.get("1.0", "end-1c")
        modelos = [l.strip() for l in raw.split("\n") if l.strip()]
        # Preserve existing states
        old_states = {}
        for m, chk in self._parallel_model_checks.items():
            try:
                old_states[m] = chk.get() == 1
            except Exception:
                old_states[m] = True
        # Clear frame
        for w in self._parallel_checks_frame.winfo_children():
            w.destroy()
        self._parallel_model_checks = {}
        for m in modelos:
            is_enabled = old_states.get(m, True)
            chk = ctk.CTkCheckBox(
                self._parallel_checks_frame, text=m,
                font=ctk.CTkFont(family=FONT_FAMILY, size=11),
                fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
                border_color=Palette.BORDER, checkmark_color=Palette.WHITE,
                text_color=Palette.TEXT_PRIMARY,
            )
            chk.pack(anchor="w", padx=2, pady=1)
            if is_enabled:
                chk.select()
            self._parallel_model_checks[m] = chk

    # ── GUARDAR AJUSTES ──────────────────────────────────────────────
    def _guardar_ajustes(self):
        try:
            _g = lambda a: getattr(self, a, None)

            # Seguridad — validar que coincidan y actualizar cache
            ent_pw1 = _g('_ent_master_pw')
            ent_pw2 = _g('_ent_master_pw_confirm')
            pw1 = ent_pw1.get().strip() if ent_pw1 else self.config.get("seguridad", {}).get("password", "")
            pw2 = ent_pw2.get().strip() if ent_pw2 else ""
            if ent_pw1 and ent_pw2 and pw1 != pw2:
                self._ajustes_lbl_status.configure(
                    text="✗ Las contraseñas no coinciden. Corregí y volvé a guardar.",
                    text_color=Palette.ERROR)
                return

            self._master_pw_cache = pw1

            # Correo (siempre encriptado con la key del .env, no con la master)
            ent = _g('_ent_correo_password')
            mail_pw = ent.get().strip() if ent else ""
            key = os.environ["MULTIAGENTE_SECRET_KEY"]
            correo_cfg = {}
            for attr, cfg_key in [('_ent_correo_usuario', 'usuario'), ('_ent_correo_password', 'password'),
                                  ('_ent_correo_imap', 'imap_server'), ('_ent_correo_puerto', 'imap_puerto'),
                                  ('_ent_remitente_balanza', 'remitente_balanza')]:
                w = _g(attr)
                if w is not None:
                    raw = w.get().strip()
                    correo_cfg[cfg_key] = int(raw) if cfg_key == 'imap_puerto' else raw
            if mail_pw:
                correo_cfg["password"] = self._encrypt_val(mail_pw, key)
            w = _g('_ajustes_texto_grupal')
            if w is not None:
                correo_cfg["destinatarios_grupal"] = [
                    l.strip() for l in w.get("1.0", "end-1c").split("\n") if l.strip()
                ]
            w = _g('_ajustes_texto_ind')
            if w is not None:
                correo_cfg["destinatarios_individual"] = [
                    l.strip() for l in w.get("1.0", "end-1c").split("\n") if l.strip()
                ]
            if correo_cfg:
                self.config["correo"] = {**self.config.get("correo", {}), **correo_cfg}

            # Documentos
            docs_cfg = {}
            for key_doc, attr in [("dorso_mic", "_ent_dorso_mic"), ("dorso_crt", "_ent_dorso_crt"),
                                  ("dorso_pe", "_ent_dorso_pe"), ("permiso_exportacion", "_ent_permiso_exp"),
                                  ("hoja_ruta", "_ent_hoja_ruta"), ("sobre", "_ent_sobre")]:
                w = _g(attr)
                if w is not None:
                    try:
                        docs_cfg[key_doc] = int(w.get().strip())
                    except ValueError:
                        pass
            if docs_cfg:
                self.config["documentos"] = {**self.config.get("documentos", {}), **docs_cfg}

            # Rutas
            rutas_cfg = {}
            for attr, cfg_key in [("_ent_ruta_sobres", "sobres"), ("_ent_ruta_cobro", "cobro"),
                                  ("_ent_ruta_pc", "pc"), ("_ent_ruta_ct_carpeta", "carga_terrestre_carpeta"),
                                  ("_ent_ruta_ct_nombre", "carga_terrestre_nombre"),
                                  ("_ent_ruta_planillas", "planillas_carga"),
                                  ("_ent_ruta_descarga", "descarga_mails"),
                                  ("_ent_ruta_escritorio", "escritorio_nombre"),
                                  ("_ent_ruta_backup", "backup_pendrive"),
                                  ("_ent_ruta_mic_sellos", "mic_sellos"),
                                  ("_ent_ruta_crt_original", "crt_original")]:
                w = _g(attr)
                if w is not None:
                    rutas_cfg[cfg_key] = w.get().strip()
            if rutas_cfg:
                self.config["rutas"] = {**self.config.get("rutas", {}), **rutas_cfg}

            # Descarga Mails — guardar y sincronizar entries del panel Correos
            descarga_cfg = {}
            for attr, cfg_key in [("_ent_descarga_papeles", "papeles"),
                                  ("_ent_descarga_reglas", "reglas"),
                                  ("_ent_descarga_sin_filtro", "sin_filtro")]:
                w = _g(attr)
                if w is not None:
                    descarga_cfg[cfg_key] = w.get().strip()
            if descarga_cfg:
                self.config["descarga_mails"] = {**self.config.get("descarga_mails", {}), **descarga_cfg}
            # Sincronizar los entries del panel Correos (si aún existen)
            for attr, src in [('_mail_entry_cantidad', '_ent_descarga_papeles'),
                              ('_mail_entry_cantidad_reglas', '_ent_descarga_reglas'),
                              ('_mail_entry_sin_filtro', '_ent_descarga_sin_filtro')]:
                w = _g(attr)
                if w is not None and w.winfo_exists():
                    src_w = _g(src)
                    if src_w is not None:
                        w.delete(0, "end")
                        w.insert(0, src_w.get().strip())

            # Valores
            valores_cfg = {}
            w = _g('_ent_precio_carpeta')
            if w is not None:
                try:
                    valores_cfg["precio_carpeta"] = int(w.get().strip())
                except ValueError:
                    valores_cfg["precio_carpeta"] = 49000
            w = _g('_ent_ata_tares')
            if w is not None:
                try:
                    valores_cfg["ata_tares"] = int(w.get().strip())
                except ValueError:
                    valores_cfg["ata_tares"] = 65000
            w = _g('_ajustes_texto_guardas')
            if w is not None:
                guardas = [l.strip() for l in w.get("1.0", "end-1c").split("\n") if l.strip()]
                if guardas:
                    valores_cfg["guardas"] = guardas
            if valores_cfg:
                self.config["valores"] = {**self.config.get("valores", {}), **valores_cfg}

            # Súper Auto configuración
            super_cfg = {}
            for attr, cfg_key in [("_super_check_sobre", "sobre"), ("_super_check_permiso", "permiso"),
                                  ("_super_check_hoja_ruta", "hoja_ruta"), ("_super_check_recibo", "recibo_ata"),
                                  ("_super_check_guarda", "aplicar_guarda"),
                                  ("_super_check_planillas", "completar_planillas")]:
                w = _g(attr)
                if w is not None:
                    super_cfg[cfg_key] = w.get()
            if super_cfg:
                pasos = dict(self.config.get("super_auto", {}).get("pasos", {}))
                pasos.update(super_cfg)
                self.config["super_auto"] = {"pasos": pasos}

            # Seguridad — guardar contraseña encriptada
            if ent_pw1 is not None:
                self.config["seguridad"] = {
                    "password": self._encrypt_val(pw1, os.environ["MULTIAGENTE_SECRET_KEY"]),
                }

            # API Visión
            api_cfg = {}
            w = _g('_ent_vision_custom_models')
            if w is not None:
                raw = w.get("1.0", "end-1c")
                api_cfg["custom_models"] = [l.strip() for l in raw.split("\n") if l.strip()]
            for attr, cfg_key, conv in [
                ("_ent_vision_api_key", "api_key", lambda v: self._encrypt_val(v.strip(), self._clave_encriptacion())),
                ("_ent_vision_model", "model", lambda v: v.strip() or procesar_tickets.MODELO_VISION_DEFAULT),
                ("_ent_vision_temperature", "temperature", lambda v: float(v.strip() or "0.1")),
                ("_ent_vision_max_tokens", "max_tokens", lambda v: int(v.strip() or "4000")),
                ("_ent_vision_timeout", "timeout", lambda v: int(v.strip() or "60")),
            ]:
                w = _g(attr)
                if w is not None:
                    api_cfg[cfg_key] = conv(w.get())
            if api_cfg:
                self.config["api_vision"] = {**self.config.get("api_vision", {}), **api_cfg}

            # Parallel settings
            w = _g('_chk_parallel_enabled')
            if w is not None:
                self.config["api_vision"]["parallel_enabled"] = w.get() == 1
            model_states = {}
            for attr_name, widget in getattr(self, '_parallel_model_checks', {}).items():
                model_states[attr_name] = widget.get() == 1
            if model_states:
                self.config["api_vision"]["parallel_model_states"] = model_states

            # Apariencia — font level
            w = _g('_ent_font_level')
            if w is not None:
                try:
                    self.config["font_level"] = int(w.get())
                except (ValueError, TypeError):
                    self.config["font_level"] = 1

            self._guardar_config()
            self._ajustes_lbl_status.configure(text="✓ Configuración guardada correctamente.")
            self.after(3000, lambda: self._ajustes_lbl_status.configure(text="")
                       if self._ajustes_lbl_status.winfo_exists() else None)
        except Exception as e:
            self._ajustes_lbl_status.configure(text=f"✗ Error al guardar: {e}", text_color=Palette.ERROR)

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
