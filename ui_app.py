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

from constants import Palette, FONT_FAMILY, FONT_MONO
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
        self.datos_planillas = []  # Resultados del agente 2
        self._mail_data = {}       # {item_id: {mid, subject, date, checked, downloaded}}
        self.panel_actual = ""      # Panel activo actual
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
        }

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
                self._diagnostico_inicial()
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
            self.sidebar, fg_color="transparent", height=52
        )
        header.pack(fill="x", padx=14, pady=(18, 2))
        header.pack_propagate(False)

        ctk.CTkLabel(
            header,
            text="MULTIAGENTE",
            font=ctk.CTkFont(family=FONT_FAMILY, size=20, weight="bold"),
            text_color=Palette.ACCENT,
        ).pack(anchor="w")

        ctk.CTkFrame(
            self.sidebar, fg_color=Palette.DIVIDER, height=1
        ).pack(fill="x", padx=14, pady=(8, 12))

        # ── Navegación ───────────────────────────────────────────────
        nav_label = ctk.CTkLabel(
            self.sidebar,
            text="AGENTES",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            text_color=Palette.TEXT_MUTED,
        )
        nav_label.pack(anchor="w", padx=14, pady=(0, 4))

        self.btn_descargar = self._crear_btn_nav(
            "Descargar Mails", "📥", 0, lambda: self._cambiar_panel("descargar")
        )
        self.btn_impresion = self._crear_btn_nav(
            "Impresión Documental", "📄", 1, lambda: self._cambiar_panel("impresion")
        )
        self.btn_planillas = self._crear_btn_nav(
            "Completar Planillas", "📊", 2, lambda: self._cambiar_panel("planillas")
        )
        self.btn_correos = self._crear_btn_nav(
            "Enviar Correos", "📤", 3, lambda: self._cambiar_panel("correos")
        )
        self.btn_backup = self._crear_btn_nav(
            "Backup", "💾", 4, lambda: self._cambiar_panel("backup")
        )

        # ── Separador ────────────────────────────────────────────────
        ctk.CTkFrame(
            self.sidebar, fg_color=Palette.DIVIDER, height=1
        ).pack(fill="x", padx=14, pady=(16, 10))

        # ── Ajustes (abajo del todo) ─────────────────────────────────
        nav_ajustes_label = ctk.CTkLabel(
            self.sidebar,
            text="CONFIGURACIÓN",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            text_color=Palette.TEXT_MUTED,
        )
        nav_ajustes_label.pack(anchor="w", padx=14, pady=(6, 4))

        self.btn_ajustes = self._crear_btn_nav(
            "Ajustes", "⚙", 5, lambda: self._cambiar_panel("ajustes")
        )

        # ── Súper Auto Toggle ────────────────────────────────────────
        ctk.CTkFrame(self.sidebar, fg_color=Palette.DIVIDER, height=1).pack(fill="x", padx=14, pady=(14, 8))

        ctk.CTkLabel(
            self.sidebar, text="AUTOMATIZACIÓN",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            text_color=Palette.TEXT_MUTED,
        ).pack(anchor="w", padx=14, pady=(0, 4))

        self._super_switch = ctk.CTkSwitch(
            self.sidebar, text="⚡ Súper Auto",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13),
            progress_color=Palette.ACCENT,
            button_color=Palette.TEXT_SECONDARY,
            button_hover_color=Palette.TEXT_PRIMARY,
            command=self._super_toggle,
        )
        self._super_switch.pack(anchor="w", padx=12, pady=(2, 2))

        self._super_lbl_guarda = ctk.CTkLabel(
            self.sidebar, text="",
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
            text_color=Palette.TEXT_MUTED,
            border_color=Palette.BORDER,
            border_width=1,
            corner_radius=6,
            height=36,
            command=self._confirmar_salida,
        )
        self.btn_salir.pack(side="bottom", fill="x", padx=10, pady=10)

        # ── Versión ──────────────────────────────────────────────────
        ctk.CTkLabel(
            self.sidebar,
            text="v2.0 · 2026",
            font=ctk.CTkFont(family=FONT_FAMILY, size=9),
            text_color=Palette.TEXT_MUTED,
        ).pack(side="bottom", pady=(0, 12))

    def _crear_btn_nav(self, texto, icono, idx, comando):
        btn = ctk.CTkButton(
            self.sidebar,
            text=f"  {icono}  {texto}",
            font=ctk.CTkFont(family=FONT_FAMILY, size=14),
            fg_color="transparent",
            hover_color=Palette.BG_HOVER,
            text_color=Palette.TEXT_SECONDARY,
            anchor="w",
            corner_radius=8,
            height=46,
            command=comando,
        )
        btn.pack(fill="x", padx=8, pady=2)
        btn._nav_idx = idx
        return btn

    def _marcar_nav_activo(self, idx):
        botones = [self.btn_descargar, self.btn_impresion, self.btn_planillas, self.btn_correos, self.btn_backup, self.btn_ajustes]
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
                    text_color=Palette.TEXT_SECONDARY,
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
            text_color=Palette.ACCENT,
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
        self.log_text.tag_config("warning", foreground=Palette.WARNING)
        self.log_text.tag_config("info", foreground=Palette.INFO)

        self._log_line_count = 0
        self._log_tags = {}

    # ── Navegación de paneles ───────────────────────────────────────────
    def _cambiar_panel(self, nombre):
        # Si hay una tarea activa, preguntar si cancelar
        if self.tarea_activa:
            ok = messagebox.askyesno(
                "Tarea en ejecución",
                "Hay una operación en curso.\n\n"
                "¿Desea cancelar la tarea actual y cambiar de panel?"
            )
            if not ok:
                return  # seguir en el panel actual
            # Cancelar la tarea
            self._cancelar_tarea.set()
            self._log("⚠ Tarea cancelada por el usuario.")
            # Esperar un momento para que el hilo se detenga
            self.after(500, lambda n=nombre: self._cambiar_panel_forzado(n))
            return

        self._cambiar_panel_forzado(nombre)

    def _cambiar_panel_forzado(self, nombre):
        # Verificar contraseña maestra para Ajustes
        if nombre == "ajustes" and not self._verificar_password_maestra():
            return

        # Guardar consola del panel actual antes de salir
        if self.panel_actual and self.panel_actual in self.logs_por_panel:
            self.logs_por_panel[self.panel_actual] = self._capturar_lineas_log()

        # Limpiar contenedor
        for w in self.panel_container.winfo_children():
            w.destroy()

        idx_map = {"descargar": 0, "impresion": 1, "planillas": 2, "correos": 3, "backup": 4, "ajustes": 5}
        idx = idx_map.get(nombre, 1)
        self._marcar_nav_activo(idx)

        if nombre == "impresion":
            self._panel_impresion()
            self.lbl_titulo_panel.configure(text="Agente de Impresión Documental")
        elif nombre == "planillas":
            self._panel_planillas()
            self.lbl_titulo_panel.configure(text="Completar Planillas (Sobres, Cobro, PC)")
        elif nombre == "descargar":
            self._panel_descargar()
            self.lbl_titulo_panel.configure(text="Descargar Mails")
        elif nombre == "correos":
            self._panel_correos()
            self.lbl_titulo_panel.configure(text="Enviar Correos")
        elif nombre == "backup":
            self._panel_backup()
            self.lbl_titulo_panel.configure(text="Backup")
        elif nombre == "ajustes":
            self._panel_ajustes()
            self.lbl_titulo_panel.configure(text="Configuración del Sistema")

        # Mostrar/ocultar consola según panel
        if nombre == "ajustes":
            self.log_container.pack_forget()
        else:
            if not self.log_container.winfo_ismapped():
                self.paned_window.add(self.log_container, minsize=100, stretch="never")

        # Guardar sash del panel que vamos a dejar
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
        # Restaurar la consola guardada para este panel (si tiene historial)
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
        self._imp_carpetas_vars = {}
        self._imp_select_all_var = ctk.BooleanVar(value=True)

        frame = ctk.CTkFrame(self.panel_container, fg_color="transparent")
        frame.pack(fill="both", expand=True)

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
            text_color=Palette.TEXT_SECONDARY, corner_radius=6, height=34,
            command=self._imp_escanear_carpetas,
        )
        self._imp_btn_refresh.pack(side="left", padx=4, pady=4)

        self._imp_btn_dorsos = ctk.CTkButton(
            toolbar,
            text="📄 Dorsos",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_HOVER, hover_color=Palette.ACCENT_DIM,
            text_color=Palette.TEXT_SECONDARY, corner_radius=6, height=34,
            command=self._imp_popup_dorsos,
        )
        self._imp_btn_dorsos.pack(side="left", padx=4, pady=4)

        self._imp_btn_sellos_mic = ctk.CTkButton(
            toolbar, text="🔖 Sellos MIC",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_HOVER, hover_color=Palette.ACCENT_DIM,
            text_color=Palette.TEXT_SECONDARY, corner_radius=6, height=34,
            command=self._imp_popup_sellos_mic,
        )
        self._imp_btn_sellos_mic.pack(side="left", padx=4, pady=4)

        self._imp_btn_sellos_crt = ctk.CTkButton(
            toolbar, text="🔖 Sellos CRT",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_HOVER, hover_color=Palette.ACCENT_DIM,
            text_color=Palette.TEXT_SECONDARY, corner_radius=6, height=34,
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
            text_color=Palette.ACCENT,
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
            text_color=Palette.ACCENT,
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
                text_color=Palette.TEXT_SECONDARY,
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
                text_color=Palette.TEXT_SECONDARY,
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
        self._imp_btn_imprimir.configure(state="normal")
        self._imp_btn_refresh.configure(state="normal")

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
            text_color=Palette.ACCENT,
        ).pack(pady=(20, 16))

        var_mic = ctk.BooleanVar(value=True)
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
                     text_color=Palette.ACCENT).pack(pady=(20, 16))

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
                     text_color=Palette.ACCENT).pack(pady=(20, 16))

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
        self._imp_btn_imprimir.configure(
            text="🖨  IMPRIMIR", state="normal",
            fg_color=Palette.ACCENT,
        )
        self._imp_btn_refresh.configure(text="🔄  Refrescar", state="normal", command=self._imp_escanear_carpetas)

    # ═══════════════════════════════════════════════════════════════════
    # PANEL 2: PLANILLAS / SOBRES
    # ═══════════════════════════════════════════════════════════════════
    def _panel_planillas(self):
        frame = ctk.CTkFrame(self.panel_container, fg_color="transparent")
        frame.pack(fill="both", expand=True)

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
            command=self._ejecutar_agente_planillas,
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
            text_color=Palette.TEXT_SECONDARY, corner_radius=6, height=34,
            command=self._popup_editar_excels,
        )
        self.btn_editar_excels.pack(side="left", padx=4, pady=4)

        self.lbl_estado_planillas = ctk.CTkLabel(
            toolbar,
            text="Listo para analizar el Escritorio",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.TEXT_MUTED,
        )
        self.lbl_estado_planillas.pack(side="left", padx=(8, 0))

        self.progress_planillas = ctk.CTkProgressBar(
            toolbar, width=160, height=8, corner_radius=4,
            fg_color=Palette.BG_INPUT, progress_color=Palette.ACCENT,
        )
        self.progress_planillas.pack(side="right", padx=16)
        self.progress_planillas.set(0)

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

        # Scrollbar
        scroll_y = ctk.CTkScrollbar(
            tabla_frame, orientation="vertical",
            fg_color=Palette.BG_CARD, button_color=Palette.TEXT_MUTED,
            button_hover_color=Palette.TEXT_SECONDARY,
        )
        scroll_y.pack(side="right", fill="y", padx=(0, 2), pady=2)

        self.tree_planillas.configure(yscrollcommand=scroll_y.set)
        scroll_y.configure(command=self.tree_planillas.yview)

        self.tree_planillas.pack(fill="both", expand=True, padx=2, pady=2)

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
        frame = ctk.CTkFrame(self.panel_container, fg_color="transparent")
        frame.pack(fill="both", expand=True)

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
            text_color=Palette.TEXT_MUTED,
        )
        self.lbl_estado_correos.pack(side="left", padx=(8, 0))

        self.progress_correos = ctk.CTkProgressBar(
            toolbar, width=160, height=8, corner_radius=4,
            fg_color=Palette.BG_INPUT, progress_color=Palette.ACCENT,
        )
        self.progress_correos.pack(side="right", padx=16)
        self.progress_correos.set(0)

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

        scroll_y2 = ctk.CTkScrollbar(
            preview_frame, orientation="vertical",
            fg_color=Palette.BG_CARD, button_color=Palette.TEXT_MUTED,
            button_hover_color=Palette.TEXT_SECONDARY,
        )
        scroll_y2.pack(side="right", fill="y", padx=(0, 2), pady=2)
        self.tree_correos.configure(yscrollcommand=scroll_y2.set)
        scroll_y2.configure(command=self.tree_correos.yview)
        self.tree_correos.pack(fill="both", expand=True, padx=2, pady=2)

        # Placeholder
        self.tree_correos.insert(
            "", "end",
            values=("—", "Sin correos procesados", "—", "—")
        )

    # ═══════════════════════════════════════════════════════════════════
    # PANEL: DESCARGAR MAILS (placeholder)
    # ═══════════════════════════════════════════════════════════════════
    def _panel_descargar(self):
        frame = ctk.CTkFrame(self.panel_container, fg_color="transparent")
        frame.pack(fill="both", expand=True)

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
        self._mail_entry_cantidad.insert(0, "2")
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
        self._mail_entry_cantidad_reglas.insert(0, "10")
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
            fg_color=Palette.WARNING, hover_color=Palette.WARNING_HOVER,
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
        self._mail_entry_sin_filtro.insert(0, "20")
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
            fg_color="#6B8E23", hover_color="#5A7A1D",
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
        """Helper: busca headers y devuelve [(fecha_dt, fecha_str, mid, asunto), ...]."""
        from email.utils import parsedate_to_datetime
        resultados = []
        for mid in todos_ids:
            status, data = mail.fetch(mid, "(BODY.PEEK[HEADER.FIELDS (SUBJECT FROM DATE)])")
            if status != "OK":
                continue
            header = data[0][1].decode("utf-8", errors="replace") if data[0][1] else ""
            match_subj = re.search(r"Subject:\s*(.+)", header, re.IGNORECASE)
            match_from = re.search(r"From:\s*(.+)", header, re.IGNORECASE)
            if not match_subj:
                continue
            asunto = match_subj.group(1).strip()
            remitente = match_from.group(1).strip() if match_from else ""
            if solo_papeles:
                if not asunto.lower().startswith("papeles"):
                    continue
                if "correo" not in remitente.lower():
                    continue
            match_date = re.search(r"Date:\s*(.+)", header, re.IGNORECASE)
            fecha_str = match_date.group(1).strip() if match_date else ""
            try:
                fecha_dt = parsedate_to_datetime(fecha_str)
            except Exception:
                fecha_dt = datetime.min
            resultados.append((fecha_dt, fecha_str, mid, asunto))
        return resultados

    def _mail_worker(self):
        """Modo 0: Descarga automática de los N mails más nuevos que cumplen las reglas."""
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
            scope = max(cantidad, 20) if modo == 1 else cantidad
            todos_ids = todos_ids[-scope:]
            solo_papeles = (modo == 1)
            label_modo = "'papeles'" if solo_papeles else "sin filtro"
            self._log(f"Buscando entre los últimos {scope} mails ({label_modo})...")
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
        self._mail_btn_buscar.configure(text="📥  Buscar y Descargar", state="normal")
        self._mail_btn_buscar_reglas.configure(text="🔍  Buscar", state="normal")
        self._mail_btn_ultimos.configure(text="📋  Mail sin filtros", state="normal")
        self._mail_btn_descargar_sel.configure(text="⬇  Descargar seleccionados", state="normal")
        self._mail_progress.stop()
        self._mail_progress.set(1)
        if modo == 0:
            self._mail_lbl_estado.configure(text="Descarga completada")
        elif modo in (1, 2):
            pendientes = sum(1 for d in self._mail_data.values() if d.get("checked") and not d.get("downloaded"))
            self._mail_lbl_estado.configure(
                text=f"Listo — {pendientes} mail(s) marcados para descargar"
            )
        elif modo == 3:
            self._mail_lbl_estado.configure(text="Descarga de seleccionados completada")
        if resultados:
            if not self._super_auto:
                self._mail_popup_resumen(resultados)

        # Súper Auto: si está activo y hay resultados, disparar cadena
        if self._super_auto and resultados:
            self._log("⚡ Súper Auto: iniciando cadena automática...")
            self.after(500, lambda: self._super_ejecutar_cadena(resultados))

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
            f_celda, pe_celda, carp_celda, dest_celda, bl_celda, trans_final, patente, precintos, guarda = datos
        except Exception as e:
            self._log(f"     ⚠ No se pudo leer el Excel: {e}")
            return os.path.basename(carpeta_temp).replace("_tmp_", "Papeles_")

        # Abreviar destinatario
        def abreviar(nombre):
            n = (nombre or "").upper()
            if "VITAPRO" in n: return "VITAPRO"
            if "EWOS" in n: return "EWOS"
            if "NUTRECO" in n: return "NUTRECO"
            if "DICOAL" in n: return "DICOAL"
            if "CARGILL" in n: return "CARGILL"
            if "BIOMAR" in n: return "BIOMAR"
            return nombre[:15]

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
        fecha_clean = f_celda.replace("/", "_") if f_celda else "sin_fecha"
        # Asegurar formato DD_MM_YYYY con leading zeros
        m = re.match(r"(\d{1,2})_(\d{1,2})_(\d{4})", fecha_clean)
        if m:
            fecha_clean = f"{int(m.group(1)):02d}_{int(m.group(2)):02d}_{m.group(3)}"
        partes = [
            fecha_clean,
            cant,
            "TERRESTRE",
            pe_recortado or "PE",
            str(carp_celda)[:10] if carp_celda else "0",
            abreviar(dest_celda or ""),
        ]
        if frac:
            partes.append(frac)

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
                pe = (pe_celda or "").strip()[-5:].lstrip("0")
                return f_celda, pe, carp_celda, dest_celda
            except Exception as e:
                self._log(f"  ⚠ Error leyendo {os.path.basename(ruta)}: {e}")
                return "", "", "", ""

        def _abreviar_dest(dest):
            d = (dest or "").upper()
            for abr in ("VITAPRO", "EWOS", "NUTRECO", "DICOAL", "CARGILL", "BIOMAR"):
                if abr in d: return abr
            return (dest or "")[:15]

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

        f_a, pe_a, carp_a, dest_a = _extraer_datos(ruta_comparte)
        f_b, pe_b, carp_b, dest_b = _extraer_datos(ruta_cierra)

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

        # Carpeta A: COMPARTIDO
        pa = [fecha_str, cant, "TERRESTRE", pe_a, str(carp_a)[:10] if carp_a else "0", _abreviar_dest(dest_a)]
        if frac_a: pa.append(frac_a)
        pa.append("COMPARTIDO")
        archivos_a = [comparte_fn]
        if hoja_a: archivos_a.append(hoja_a)
        if pdf_a: archivos_a.append(pdf_a)
        ruta_a = _mover("_".join(pa), archivos_a)
        carpetas_creadas.append(("_".join(pa), ruta_a))
        self._log(f"  📁 {'/'.join(ruta_a.split(chr(92))[-2:])} ({len(archivos_a)} archivos)")

        # Carpeta B: COMPARTIDO_CERRA
        pb = [fecha_str, cant, "TERRESTRE", pe_b, str(carp_b)[:10] if carp_b else "0", _abreviar_dest(dest_b)]
        if frac_b: pb.append(frac_b)
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
                    if es_xls:
                        # Formato .xls: usar xlrd + xlutils
                        import xlrd as xlrd_local
                        from xlutils.copy import copy as xl_copy
                        rb = xlrd_local.open_workbook(ruta, formatting_info=True)
                        chofer_idx = None
                        for i, sname in enumerate(rb.sheet_names()):
                            if "CHOFER" in sname.upper():
                                chofer_idx = i; break
                        if chofer_idx is None:
                            self.log_queue.put(f"[...]   ⚠ Hoja 'Choferes' no hallada en {archivo[:50]}")
                            rb.release_resources(); continue
                        rs = rb.sheet_by_index(chofer_idx)
                        wb_w = xl_copy(rb)
                        ws_w = wb_w.get_sheet(chofer_idx)
                        escrito = False
                        valores_col_g = []
                        for row in range(1, 16):
                            val = rs.cell_value(row, 6) if row < rs.nrows else ""
                            if val:
                                valores_col_g.append(f"  fila {row}: '{str(val)[:60]}'")
                                if isinstance(val, str) and "GUARDA" in val.strip().upper():
                                    ws_w.write(row, 7, self._super_guarda)
                                    self.log_queue.put(f"[...]   🛡 Guarda '{self._super_guarda}' → {archivo[:50]} fila {row+1}")
                                    escrito = True; break
                        if not escrito:
                            self.log_queue.put(f"[...]   ⚠ 'Guarda' no hallada en col G de {archivo[:50]}")
                            self.log_queue.put(f"[...]     Hojas: {rb.sheet_names()}")
                            self.log_queue.put(f"[...]     Hoja '{rs.name}' col G (filas 1-15):")
                            if valores_col_g:
                                for v in valores_col_g:
                                    self.log_queue.put(f"[...]     {v}")
                            else:
                                self.log_queue.put(f"[...]     (todas vacías)")
                        wb_w.save(ruta)
                        rb.release_resources()
                        aplicados += 1
                    else:
                        # .xlsx: openpyxl
                        wb = self._abrir_excel_seguro(ruta)
                        ws_chofer = None
                        for s in wb.sheetnames:
                            if "CHOFER" in s.upper():
                                ws_chofer = wb[s]; break
                        if not ws_chofer:
                            self.log_queue.put(f"[...]   ⚠ Hoja 'Choferes' no hallada en {archivo[:50]}")
                            wb.close(); continue
                        escrito = False
                        valores_col_g = []
                        for row in range(1, 16):
                            cell = ws_chofer.cell(row=row, column=7)
                            val = cell.value
                            if val is None:
                                for mr in ws_chofer.merged_cells.ranges:
                                    if (mr.min_col <= 7 <= mr.max_col and mr.min_row <= row <= mr.max_row):
                                        val = ws_chofer.cell(row=mr.min_row, column=mr.min_col).value; break
                            if val is not None and val != "":
                                valores_col_g.append(f"  fila {row}: '{str(val)[:60]}'")
                                if "GUARDA" in str(val).strip().upper():
                                    ws_chofer.cell(row=row, column=8).value = self._super_guarda
                                    self.log_queue.put(f"[...]   🛡 Guarda '{self._super_guarda}' → {archivo[:50]} fila {row}")
                                    escrito = True; break
                        if not escrito:
                            self.log_queue.put(f"[...]   ⚠ 'Guarda' no hallada en col G de {archivo[:50]}")
                            self.log_queue.put(f"[...]     Hojas: {wb.sheetnames}")
                            self.log_queue.put(f"[...]     Hoja '{ws_chofer.title}' col G (filas 1-15):")
                            if valores_col_g:
                                for v in valores_col_g:
                                    self.log_queue.put(f"[...]     {v}")
                            else:
                                self.log_queue.put(f"[...]     (todas vacías)")
                        self._guardar_excel_seguro(wb, ruta)
                        aplicados += 1
                except Exception as e:
                    self.log_queue.put(f"[...]   ⚠ Error guarda en {archivo}: {e}")
                    import traceback
                    self.log_queue.put(f"[...]     {traceback.format_exc()}")
        self.log_queue.put(f"[...] ✓ Guarda '{self._super_guarda}': {aplicados} planillas actualizadas")
        return aplicados

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
                     text_color=Palette.ACCENT).pack(pady=(24, 20))

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
        frame = ctk.CTkFrame(self.panel_container, fg_color="transparent")
        frame.pack(fill="both", expand=True)

        # ── Toolbar ──────────────────────────────────────────────────
        toolbar = ctk.CTkFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER, height=44
        )
        toolbar.pack(fill="x", pady=(0, 6))
        toolbar.pack_propagate(False)

        self.btn_backup_drive = ctk.CTkButton(
            toolbar,
            text="📀  Backup",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.BG_HOVER, text_color=Palette.TEXT_MUTED,
            corner_radius=6, height=34, width=160, state="disabled",
        )
        self.btn_backup_drive.pack(side="left", padx=4, pady=4)

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
            text_color=Palette.TEXT_MUTED,
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
        self.btn_backup_drive.configure(state="disabled")
        self.progress_backup.configure(mode="indeterminate")
        self.progress_backup.start()
        self.lbl_estado_backup.configure(text="Detectando pendrive...", text_color=Palette.INFO)
        self._limpiar_log()
        self._log("💿 Iniciando Back Up...")
        t = threading.Thread(target=lambda: self._backup_pendrive_worker(seleccionadas), daemon=True)
        t.start()

    def _backup_pendrive_worker(self, carpetas=None):
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

    # ── Backup ─────────────────────────────────────────────────
    def _backup_drive_iniciar(self):
        if self.tarea_activa:
            messagebox.showwarning("Agente ocupado", "Hay una tarea en ejecución.")
            return
        if not self._cfg_obtener("google_drive", "client_id", ""):
            messagebox.showwarning("Drive no configurado",
                "Configure las credenciales de Google Drive en Ajustes > Drive primero.")
            return
        self.tarea_activa = True
        self._cancelar_tarea.clear()
        self.btn_backup_drive.configure(text="⏳  Subiendo...", state="disabled")
        self.btn_backup_pendrive.configure(state="disabled")
        self.progress_backup.configure(mode="indeterminate")
        self.progress_backup.start()
        self.lbl_estado_backup.configure(text="Conectando a Google Drive...", text_color=Palette.INFO)
        self._limpiar_log()
        self._log("📀 Iniciando Backup...")
        t = threading.Thread(target=self._backup_drive_worker, daemon=True)
        t.start()

    def _drive_autenticar(self):
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        from google.auth.credentials import Credentials
        from googleapiclient.discovery import build

        client_id = self._cfg_obtener("google_drive", "client_id", "")
        client_secret = self._cfg_obtener("google_drive", "client_secret", "")
        token_info = self._cfg_obtener("google_drive", "token", None)

        creds = None
        if token_info:
            try:
                creds = Credentials.from_authorized_user_info(token_info, ["https://www.googleapis.com/auth/drive.file"])
            except Exception:
                pass

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                self.log_queue.put("[...] 🔄 Renovando token OAuth...")
                creds.refresh(Request())
            else:
                self.log_queue.put("[...] 🔐 Abriendo navegador para autenticación...")
                flow = InstalledAppFlow.from_client_config(
                    {"installed": {
                        "client_id": client_id, "client_secret": client_secret,
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token"
                    }},
                    ["https://www.googleapis.com/auth/drive.file"])
                creds = flow.run_local_server(port=0)

            token_data = {
                "refresh_token": creds.refresh_token,
                "token": creds.token,
                "client_id": creds.client_id,
                "client_secret": creds.client_secret,
            }
            self.config.setdefault("google_drive", {})["token"] = token_data
            self._guardar_config()

        return build("drive", "v3", credentials=creds)

    def _drive_buscar_o_crear_carpeta(self, service, nombre, parent_id):
        query = f"name='{nombre}' and '{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        results = service.files().list(q=query, fields="files(id)").execute()
        if results.get("files"):
            fid = results["files"][0]["id"]
            self.log_queue.put(f"[...]   📁 {nombre} (existente)")
            return fid
        metadata = {"name": nombre, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_id]}
        fid = service.files().create(body=metadata, fields="id").execute()["id"]
        self.log_queue.put(f"[...]   + Creando carpeta: {nombre}")
        return fid

    def _drive_subir_sobrescrito(self, service, ruta_local, nombre, folder_id):
        from googleapiclient.http import MediaFileUpload
        query = f"name='{nombre}' and '{folder_id}' in parents and trashed=false"
        results = service.files().list(q=query, fields="files(id)").execute()
        media = MediaFileUpload(ruta_local, resumable=True)
        if results.get("files"):
            service.files().update(fileId=results["files"][0]["id"], media_body=media).execute()
            self.log_queue.put(f"[...]   ↑ {nombre} (sobreescrito)")
        else:
            metadata = {"name": nombre, "parents": [folder_id]}
            service.files().create(body=metadata, media_body=media).execute()
            self.log_queue.put(f"[...]   ↑ {nombre} (subido)")

    def _backup_drive_worker(self):
        try:
            service = self._drive_autenticar()
            folder_id_raiz = self._cfg_obtener("google_drive", "folder_id", "")
            if not folder_id_raiz:
                self.log_queue.put("[...] ⚠ Falta Folder ID en Ajustes > Drive")
                self.after(0, self._backup_done)
                return

            self.log_queue.put("[...] ✓ Conectado a Google Drive")

            año = str(datetime.now().year)
            mes = datetime.now().month
            meses_es = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
                        7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
            mes_str = f"{mes:02d}_{meses_es[mes]}"

            # Buscar/crear estructura de carpetas
            id_cargas     = self._drive_buscar_o_crear_carpeta(service, "CARGAS", folder_id_raiz)
            id_cargas_año = self._drive_buscar_o_crear_carpeta(service, f"Cargas_{año}", id_cargas)
            id_mes        = self._drive_buscar_o_crear_carpeta(service, mes_str, id_cargas_año)

            # Subir CARGA TERRESTRE.xlsx
            ruta_ct = buscar_archivo_en_pendrive(
                "CARGA TERRESTRE.xlsx",
                self._cfg_obtener_rutas("carga_terrestre_carpeta", os.path.join("TRABAJO", "01_PLANILLAS")))
            if not ruta_ct:
                ruta_ct = buscar_archivo_en_pendrive(
                    "CARGA TERRESTRE.xls",
                    self._cfg_obtener_rutas("carga_terrestre_carpeta", os.path.join("TRABAJO", "01_PLANILLAS")))
            if ruta_ct:
                self._drive_subir_sobrescrito(service, ruta_ct, "CARGA TERRESTRE.xlsx", id_cargas)

            # Escanear y subir carpetas del escritorio
            escritorio = self._resolver_ruta("planillas_carga", "Desktop")
            carpetas = []
            for item in sorted(os.listdir(escritorio)):
                ruta_carp = os.path.join(escritorio, item)
                if not os.path.isdir(ruta_carp):
                    continue
                if item.startswith("."):
                    continue
                for archivo in os.listdir(ruta_carp):
                    up = archivo.upper()
                    if "CONTENEDORES" in up or (up.startswith("PLANILLA DE CARGA") and (up.endswith(".XLSX") or up.endswith(".XLS"))):
                        carpetas.append(ruta_carp)
                        break

            self.after(0, lambda: self.progress_backup.configure(mode="determinate"))
            self.after(0, lambda: self.progress_backup.set(0))
            total = len(carpetas)
            for i, carpeta in enumerate(carpetas):
                if self._cancelar_tarea.is_set():
                    self.log_queue.put("[...] ⚠ Tarea cancelada.")
                    break
                nombre = os.path.basename(carpeta)
                id_sub = self._drive_buscar_o_crear_carpeta(service, nombre, id_mes)
                for archivo in os.listdir(carpeta):
                    if self._cancelar_tarea.is_set():
                        break
                    ruta_archivo = os.path.join(carpeta, archivo)
                    if os.path.isfile(ruta_archivo):
                        self._drive_subir_sobrescrito(service, ruta_archivo, archivo, id_sub)
                self.after(0, lambda p=i+1, t=total: self.progress_backup.set(p/t if t else 1))

            self.log_queue.put(f"[...] ✓ Backup: {len(carpetas)} carpetas subidas a Google Drive")
        except Exception as e:
            self.log_queue.put(f"[...] ⚠ Error: {e}")
        finally:
            self.after(0, self._backup_done)

    def _backup_done(self):
        self.tarea_activa = False
        self.btn_backup_drive.configure(text="📀  Backup", state="disabled", fg_color=Palette.BG_HOVER)
        self.btn_backup_pendrive.configure(text="💿  Back Up", state="normal", fg_color=Palette.SECONDARY)
        self.progress_backup.stop()
        self.progress_backup.set(1)
        self.lbl_estado_backup.configure(text="Backup completado", text_color=Palette.SUCCESS)
        self._refrescar_lista_backup()

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
        popup.geometry("360x320")
        popup.configure(fg_color=Palette.BG_CARD)
        popup.transient(self)
        popup.lift()
        popup.grab_set()
        popup.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 360) // 2
        py = self.winfo_y() + (self.winfo_height() - 320) // 2
        popup.geometry(f"360x320+{px}+{py}")

        ctk.CTkLabel(popup, text="EDITAR EXCELS",
                     font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
                     text_color=Palette.ACCENT).pack(pady=(20, 12))

        excels = [
            ("SOBRES_2026.xlsx", self._cfg_obtener_rutas("sobres", os.path.join("TRABAJO", "01_PLANILLAS"))),
            ("COBRO_2026.xlsx", self._cfg_obtener_rutas("cobro", os.path.join("TRABAJO", "01_PLANILLAS"))),
            ("PC.xlsx", self._cfg_obtener_rutas("pc", os.path.join("TRABAJO", "01_PLANILLAS"))),
            ("FECHA MIC Y SELLOS.xlsx", self._cfg_obtener_rutas("mic_sellos", os.path.join("TRABAJO", "01_PLANILLAS"))),
            ("FECHA CRT Y ORIGINAL.xlsx", self._cfg_obtener_rutas("crt_original", os.path.join("TRABAJO", "01_PLANILLAS"))),
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
                      text_color=Palette.TEXT_SECONDARY, corner_radius=6,
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
                es_xls = ruta_contenedores.lower().endswith(".xls") and not ruta_contenedores.lower().endswith(".xlsx")

                try:
                    if es_xls:
                        # Formato .xls antiguo: usar xlrd + xlutils
                        import xlrd
                        from xlutils.copy import copy as xl_copy
                        rb = xlrd.open_workbook(ruta_contenedores, formatting_info=True)
                        chofer_idx = None
                        for i, sname in enumerate(rb.sheet_names()):
                            if "CHOFER" in sname.upper():
                                chofer_idx = i
                                break
                        if chofer_idx is None:
                            self.log_queue.put(f"[...]     ⚠ Hoja 'Choferes' no hallada")
                            rb.release_resources()
                            continue

                        rs = rb.sheet_by_index(chofer_idx)
                        wb_w = xl_copy(rb)
                        ws_w = wb_w.get_sheet(chofer_idx)

                        escrito = False
                        for row in range(1, 16):
                            val = rs.cell_value(row, 6) if row < rs.nrows else ""  # col 6 = G
                            if val and isinstance(val, str) and "GUARDA" in val.strip().upper():
                                ws_w.write(row, 7, guarda_elegido)  # col 7 = H
                                escrito = True
                                self.log_queue.put(f"[...]     ✓ '{guarda_elegido}' → fila {row+1}, col H")
                                break
                        if not escrito:
                            self.log_queue.put(f"[...]     ⚠ 'Guarda' no hallada en col G")

                        self.log_queue.put(f"[...]     💾 Guardando...")
                        wb_w.save(ruta_contenedores)
                        rb.release_resources()
                        self.log_queue.put(f"[...]     ✅ Listo")
                    else:
                        # Formato .xlsx: usar openpyxl
                        wb = self._abrir_excel_seguro(ruta_contenedores)
                        ws_chofer = None
                        for s in wb.sheetnames:
                            if "CHOFER" in s.upper():
                                ws_chofer = wb[s]
                                break
                        if not ws_chofer:
                            self.log_queue.put(f"[...]     ⚠ Hoja 'Choferes' no hallada")
                            wb.close()
                            continue

                        escrito = False
                        for row in range(1, 16):
                            cell = ws_chofer.cell(row=row, column=7)
                            val = cell.value
                            if val is None:
                                for mr in ws_chofer.merged_cells.ranges:
                                    if (mr.min_col <= 7 <= mr.max_col
                                            and mr.min_row <= row <= mr.max_row):
                                        val = ws_chofer.cell(row=mr.min_row, column=mr.min_col).value
                                        break
                            if val is not None and val != "":
                                val_str = str(val).strip().upper()
                                if val_str == "GUARDA" or val_str.startswith("GUARDA"):
                                    ws_chofer.cell(row=row, column=8).value = guarda_elegido
                                    escrito = True
                                    self.log_queue.put(f"[...]     ✓ '{guarda_elegido}' → fila {row}, col H")
                                    break

                        if not escrito:
                            self.log_queue.put(f"[...]     ⚠ 'Guarda' no hallada en col G")

                        self.log_queue.put(f"[...]     💾 Guardando...")
                        self._guardar_excel_seguro(wb, ruta_contenedores)
                        self.log_queue.put(f"[...]     ✅ Listo")
                except Exception as e:
                    self.log_queue.put(f"[...]     ⚠ Error: {e}")

            self.log_queue.put(f"[...] ✓ Guarda '{guarda_elegido}' completado.")
            self.after(0, self._guarda_done)

        t = threading.Thread(target=worker, daemon=True)
        t.start()

    def _guarda_done(self):
        self.tarea_activa = False
        self.btn_agregar_guarda.configure(
            text="🛡  Agregar Guarda", state="normal",
            fg_color=Palette.SECONDARY)
        self.btn_ejecutar_planillas.configure(state="normal")

    # ═══════════════════════════════════════════════════════════════════
    # AGENTE 2: EJECUCIÓN EN HILO
    # ═══════════════════════════════════════════════════════════════════
    def _ejecutar_agente_planillas(self):
        if self.tarea_activa:
            messagebox.showwarning(
                "Agente ocupado",
                "Hay una tarea en ejecución. Espere a que finalice."
            )
            return
        self.tarea_activa = True
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

        t = threading.Thread(target=self._agente_planillas_worker, daemon=True)
        t.start()

    def _agente_planillas_worker(self):
        """Ejecuta el agente de planillas en hilo de fondo."""
        try:
            self._planillas_core()
        except Exception as e:
            self.after(0, lambda: self._planillas_error(str(e)))
        finally:
            self.after(0, self._planillas_done)

    def _planillas_core(self):
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
                    ws_sobres = wb_sobres.active
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
                    f_celda, pe_celda, carp_celda, dest_celda, bl_celda, trans_final, primera_patente, precintos, guarda = datos

                    pe_origen = pe_celda if pe_celda else pe_carpeta
                    pe_recortado = pe_origen[-5:].lstrip("0") if pe_origen else "No detectado"
                    if not trans_final:
                        trans_final = "No detectado"
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
                        "terminal": "CHILE",
                        "trans_final": trans_final,
                        "dest_celda": dest_celda,
                        "bl_final": bl_final,
                        "frac_carpeta": frac_carpeta,
                        "servicio": cant_final * 60000,
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
        if datos_extraidos and ws_sobres is not None:
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

                    valores_check_1 = {
                        1: f_celda, 2: d1["cant_final"], 3: "TERRES", 4: d1["pe_recortado"],
                        5: d1["carp_celda"], 6: "CHILE", 7: d1["trans_final"], 8: dest_abrev,
                        9: d1["bl_final"], 10: d1["frac_carpeta"], 11: d1["cant_final"] * 60000,
                    }
                    if ya_existe_en_hoja(ws_sobres, valores_check_1, excluir_columnas={9}):
                        self._log(f"  Omitido (ya existe): Carpeta {d1.get('carp_celda','?')} | P.E. {d1.get('pe_recortado','?')} | {f_celda}")
                        continue

                    r1 = _escribir_fecha(ws_sobres, 3, f_celda)
                    for col in (2, 3, 4, 5, 6, 7, 8, 9, 10, 11):
                        val = valores_check_1[col]
                        font = font_chico if col == 7 else (font_mediano if col in (8, 9) else font_normal)
                        fl = fill_color if col == 8 else None
                        _escribir_celda(ws_sobres, r1, col, val, font, fl)

                    if es_par:
                        dest2 = abreviar_dest(d2["dest_celda"])
                        color2 = COLORES_EMPRESA.get(dest2, "FFFFFF")
                        fill2 = PatternFill(fill_type="solid", fgColor=color2)
                        valores_check_2 = {
                            4: d2["pe_recortado"], 5: d2["carp_celda"],
                            9: d2["bl_final"], 10: d2["frac_carpeta"],
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
                            font = font_chico if col == 7 else (font_mediano if col in (8, 9) else font_normal)
                            fl = fill2 if col == 8 else None
                            _escribir_celda(ws_sobres, r2, col, val, font, fl)
                        # Cols compartidas (merge) para fila 2: B,C,F,G,H,K
                        for col in (2, 3, 6, 7, 8, 11):
                            val = valores_check_1[col]
                            font = font_chico if col == 7 else (font_mediano if col == 8 else font_normal)
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
        if datos_extraidos:
            resultados["cobro"] = self._completar_cobro(datos_extraidos)

        # ═══════════════════════════════════════════════════════════════
        # COMPLETAR PLANILLA PC (PRECINTOS/CABLES)
        # ═══════════════════════════════════════════════════════════════
        if datos_extraidos:
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
                ws_cobro = wb_cobro.active
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
                servicio_ata = d1["cant_final"] * 60000

                valores_fila = {
                    1: f_celda, 2: d1["cant_final"], 3: "TERRES", 4: d1["pe_recortado"],
                    5: d1["carp_celda"], 6: d1["frac_carpeta"], 7: servicio_ata, 8: precio_carpeta,
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
                    servicio_ata_2 = d2["cant_final"] * 60000
                    r2 = r1 + 1

                    vals2 = {1: f_celda, 2: d2["cant_final"], 3: "TERRES", 4: d2["pe_recortado"],
                             5: d2["carp_celda"], 6: d2["frac_carpeta"], 7: servicio_ata_2, 8: precio_carpeta_2}
                    for col, val in vals2.items():
                        cell = ws_cobro.cell(row=r2, column=col)
                        cell.value = val
                        cell.border = border
                        cell.alignment = center

                    # Merge fecha (col 1)
                    ws_cobro.merge_cells(start_row=r1, start_column=1, end_row=r2, end_column=1)
                    ws_cobro.cell(row=r1, column=1).alignment = center
                    # Merge Cantidad (col 2), TERRES (col 3) y Servicio ATA (col 7)
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
        return f_celda, pe_celda, carp_celda, dest_celda, bl_celda, transporte_encontrado, primera_patente, precintos, guarda

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
        return f_celda, pe_celda, carp_celda, dest_celda, bl_celda, transporte_encontrado, primera_patente, precintos, guarda

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
            f'${registro["servicio"]:,}',
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
        self._set_status("Análisis de planillas finalizado")

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
            if not seleccionadas and not incluir_plt:
                messagebox.showwarning("Nada seleccionado", "Seleccioná al menos una carpeta o la Planilla de Carga.")
                return
            popup.destroy()
            self._despachar_carpetas(seleccionadas, incluir_plt)

        ctk.CTkButton(
            btn_frame, text="📤  Enviar Seleccionados", width=200, height=34,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6,
            command=enviar_seleccionados,
        ).pack(side="right", padx=4)

    def _despachar_carpetas(self, rutas_seleccionadas, incluir_planilla=False):
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
                incluir_planilla=incluir_planilla),
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

    def _agente_correos_worker(self, carpetas_filtro=None, incluir_planilla=False):
        try:
            self._correos_core(carpetas_filtro=carpetas_filtro, incluir_planilla=incluir_planilla)
        except Exception as e:
            self.after(0, lambda: self._correos_error(str(e)))
        finally:
            self.after(0, self._correos_done)

    def _correos_core(self, carpetas_filtro=None, incluir_planilla=False):
        escritorio = self._resolver_ruta("planillas_carga", "Desktop")
        año_actual = datetime.now().strftime("%y")

        todas_las_planillas = []
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
                match = re.search(r"TERRESTRE_[^_]+_[^_]+_([A-Z]+)", n1, re.IGNORECASE)
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

            # Buscar todas las planillas de carga en el escritorio para el correo grupal
            for archivo in archivos:
                up = archivo.upper()
                if up.startswith("PLANILLA DE CARGA") and (
                    up.endswith(".XLSX") or up.endswith(".XLS")
                ):
                    todas_las_planillas.append(os.path.join(ruta_carpeta, archivo))

            # Si hay filtro de carpetas, solo procesar los individuales seleccionados
            if carpetas_filtro is not None and ruta_carpeta not in carpetas_filtro:
                continue

            # Filtrar archivos para correo
            adjuntos_validos = []
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
                        m = re.search(r"TERRESTRE_(.+)", nombre_carpeta, re.IGNORECASE)
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
                else:
                    match_sufijo = re.search(r"TERRESTRES?_(.*)", item, re.IGNORECASE)
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
        if todas_las_planillas and incluir_plt:
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
            planillas_ordenadas = sorted(todas_las_planillas, key=lambda p: os.path.basename(p).upper())

            cuerpo = "Estimados,\n\nSe adjuntan las planillas de carga correspondientes:\n\n"
            cuerpo += "PLANILLA DE CARGA:\n"
            for p in planillas_ordenadas:
                nombre_sin_ext = os.path.splitext(os.path.basename(p))[0]
                cuerpo += f"  • {nombre_sin_ext}\n"
            if ruta_maestra:
                cuerpo += f"  • {os.path.splitext(os.path.basename(ruta_maestra))[0]}\n"
            cuerpo += "\nSaludos cordiales."
            msg_grupal.attach(MIMEText(cuerpo, "plain"))

            n_adj_grupal = len(todas_las_planillas)
            for p in todas_las_planillas:
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
        h = min(180 + n * 36, 480)
        popup.geometry(f"520x{h}")
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
            text_color=Palette.ACCENT,
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
            text_color=Palette.ACCENT,
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
        """Agrega un mensaje al log desde cualquier hilo."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put(f"[{timestamp}] {mensaje}")

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
        elif "ADVERTENCIA" in txt_up or "FALTA" in txt_up or "NO SE ENCONTRÓ" in txt_up:
            tag = "warning"
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
        """Procesa la cola de mensajes de log periódicamente."""
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.emit_log(msg)
        except queue.Empty:
            pass
        self.after(100, self._poll_log_queue)

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
            key = self._master_pw_cache if self._master_pw_cache else os.environ["MULTIAGENTE_SECRET_KEY"]
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
    # PANEL: AJUSTES
    # ═══════════════════════════════════════════════════════════════════
    def _panel_ajustes(self):
        frame = ctk.CTkFrame(self.panel_container, fg_color="transparent")
        frame.pack(fill="both", expand=True)

        # ── Toolbar de tabs ──────────────────────────────────────────
        tabs_frame = ctk.CTkFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER, height=42
        )
        tabs_frame.pack(fill="x", pady=(0, 4))
        tabs_frame.pack_propagate(False)

        self._ajustes_tabs = {}
        self._ajustes_frames = {}
        tab_names = [
            ("correo", "📧  Correo"),
            ("documentos", "📄  Documentos"),
            ("rutas", "📁  Rutas"),
            ("valores", "💰  Valores"),
            ("drive", "💾  Drive"),
            ("seguridad", "🔒  Seguridad"),
        ]

        def cambiar_tab(nombre):
            for key, fr in self._ajustes_frames.items():
                fr.pack_forget() if fr.winfo_ismapped() else None
            for key, btn in self._ajustes_tabs.items():
                if key == nombre:
                    btn.configure(fg_color=Palette.ACCENT, text_color=Palette.WHITE)
                else:
                    btn.configure(fg_color="transparent", text_color=Palette.TEXT_SECONDARY)
            self._ajustes_frames[nombre].pack(fill="both", expand=True)

        for key, label in tab_names:
            btn = ctk.CTkButton(
                tabs_frame, text=label,
                font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
                fg_color="transparent", hover_color=Palette.BG_HOVER,
                text_color=Palette.TEXT_SECONDARY,
                corner_radius=6, height=32, width=110,
                command=lambda k=key: cambiar_tab(k),
            )
            btn.pack(side="left", padx=2, pady=4)
            self._ajustes_tabs[key] = btn

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

        self._ajustes_tab_correo(self._ajustes_frames["correo"])
        self._ajustes_tab_documentos(self._ajustes_frames["documentos"])
        self._ajustes_tab_rutas(self._ajustes_frames["rutas"])
        self._ajustes_tab_valores(self._ajustes_frames["valores"])
        self._ajustes_tab_drive(self._ajustes_frames["drive"])
        self._ajustes_tab_seguridad(self._ajustes_frames["seguridad"])

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

    # ── Helpers de layout ────────────────────────────────────────────
    def _ajustes_seccion(self, parent, titulo):
        ctk.CTkLabel(
            parent, text=titulo,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            text_color=Palette.ACCENT,
        ).pack(anchor="w", padx=14, pady=(12, 4))

    def _ajustes_row(self, parent, label, default="", show="", extra=None, width=260):
        """Crea fila compacta: label arriba, entry abajo (ancho fijo). Retorna el CTkEntry."""
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(
            row, text=label,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.TEXT_SECONDARY, anchor="w",
        ).pack(anchor="w")
        e = ctk.CTkEntry(
            row, width=width, height=30,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=4,
            show=show,
        )
        e.insert(0, str(default) if default else "")
        e.pack(anchor="w", pady=(2, 0))
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
            parent, "Contraseña:", self._cfg_obtener_correo("password", ""), show="*", width=220)
        self._ent_correo_imap = self._ajustes_row(
            parent, "Servidor IMAP:", self._cfg_obtener_correo("imap_server", IMAP_SERVER), width=280)
        self._ent_correo_puerto = self._ajustes_row(
            parent, "Puerto IMAP:", str(self._cfg_obtener_correo("imap_puerto", PUERTO_IMAP)), width=80)

        self._ajustes_seccion(parent, "Destinatarios — Correo Grupal (CARGA TERRESTRE)")
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
    def _ajustes_tab_valores(self, parent):
        self._ajustes_seccion(parent, "Tarifas de Planilla COBRO")
        self._ent_precio_carpeta = self._ajustes_row(
            parent, "Valor Carpeta ($):",
            str(self._cfg_obtener("valores", "precio_carpeta", 49000)),
            extra="Primer ítem de cada fecha. Los siguientes llevan la mitad.",
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

    # ── TAB: GOOGLE DRIVE ────────────────────────────────────────────
    def _ajustes_tab_drive(self, parent):
        self._ajustes_seccion(parent, "Backup en Google Drive")
        ctk.CTkLabel(
            parent,
            text="1. Ir a console.cloud.google.com → APIs y Servicios → Biblioteca\n"
                 "   Buscar 'Google Drive API' → Habilitar\n"
                 "2. Credenciales → Crear credenciales → ID de cliente OAuth\n"
                 "   Tipo: 'Aplicación de escritorio' → Crear\n"
                 "3. Copiar Client ID y Client Secret acá abajo\n"
                 "4. El Folder ID lo sacás de la URL de Drive:\n"
                 "   drive.google.com/drive/folders/XXXXXX ← eso es el ID",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_MUTED, justify="left",
        ).pack(anchor="w", padx=14, pady=(0, 8))

        self._ent_drive_client_id = self._ajustes_row(
            parent, "Client ID:", self._cfg_obtener("google_drive", "client_id", ""), width=400)
        self._ent_drive_client_secret = self._ajustes_row(
            parent, "Client Secret:", self._cfg_obtener("google_drive", "client_secret", ""), show="*", width=400)
        self._ent_drive_folder_id = self._ajustes_row(
            parent, "Folder ID (trabajo_2024):", self._cfg_obtener("google_drive", "folder_id", ""), width=400)

        # Botón para verificar conexión
        verify_frame = ctk.CTkFrame(parent, fg_color="transparent")
        verify_frame.pack(fill="x", padx=14, pady=(8, 4))
        self._drive_verify_btn = ctk.CTkButton(
            verify_frame, text="🔗  Verificar Conexión Drive",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=32, width=220,
            command=self._drive_verificar_conexion,
        )
        self._drive_verify_btn.pack(side="left")
        self._drive_verify_lbl = ctk.CTkLabel(
            verify_frame, text="",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_SECONDARY,
        )
        self._drive_verify_lbl.pack(side="left", padx=10)

    def _drive_verificar_conexion(self):
        """Prueba la autenticación OAuth con Google Drive en un hilo de fondo."""
        client_id = self._ent_drive_client_id.get().strip()
        client_secret = self._ent_drive_client_secret.get().strip()
        if not client_id or not client_secret:
            self._drive_verify_lbl.configure(text="✗ Completá Client ID y Client Secret", text_color=Palette.ERROR)
            return
        self._drive_verify_btn.configure(text="⏳  Verificando...", state="disabled")
        self._drive_verify_lbl.configure(text="Conectando...", text_color=Palette.INFO)

        def _verificar():
            try:
                from google_auth_oauthlib.flow import InstalledAppFlow
                from google.auth.transport.requests import Request
                from google.auth.credentials import Credentials
                from googleapiclient.discovery import build
                import json

                token_info = self._cfg_obtener("google_drive", "token", None)
                creds = None
                if token_info:
                    try:
                        creds = Credentials.from_authorized_user_info(token_info, ["https://www.googleapis.com/auth/drive.file"])
                    except Exception:
                        pass

                if not creds or not creds.valid:
                    if creds and creds.expired and creds.refresh_token:
                        creds.refresh(Request())
                    else:
                        flow = InstalledAppFlow.from_client_config(
                            {"installed": {"client_id": client_id, "client_secret": client_secret,
                                           "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                                           "token_uri": "https://oauth2.googleapis.com/token"}},
                            ["https://www.googleapis.com/auth/drive.file"])
                        creds = flow.run_local_server(port=0)

                    # Guardar token para futuro
                    token_data = {
                        "refresh_token": creds.refresh_token,
                        "token": creds.token,
                        "client_id": creds.client_id,
                        "client_secret": creds.client_secret,
                    }
                    self.config.setdefault("google_drive", {})["token"] = token_data
                    self._guardar_config()

                service = build("drive", "v3", credentials=creds)
                service.files().list(pageSize=1, fields="files(id)").execute()
                self.after(0, lambda: self._drive_verify_lbl.configure(
                    text="✓ Conexión exitosa", text_color=Palette.SUCCESS))
            except Exception as e:
                self.after(0, lambda: self._drive_verify_lbl.configure(
                    text=f"✗ Error: {str(e)[:60]}", text_color=Palette.ERROR))
            finally:
                self.after(0, lambda: self._drive_verify_btn.configure(
                    text="🔗  Verificar Conexión Drive", state="normal"))

        t = threading.Thread(target=_verificar, daemon=True)
        t.start()

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

    # ── GUARDAR AJUSTES ──────────────────────────────────────────────
    def _guardar_ajustes(self):
        try:
            # Seguridad — validar que coincidan y actualizar cache
            pw1 = self._ent_master_pw.get().strip()
            pw2 = self._ent_master_pw_confirm.get().strip()
            if pw1 != pw2:
                self._ajustes_lbl_status.configure(
                    text="✗ Las contraseñas no coinciden. Corregí y volvé a guardar.",
                    text_color=Palette.ERROR)
                return

            self._master_pw_cache = pw1

            # Correo
            mail_pw = self._ent_correo_password.get().strip()
            key = pw1 if pw1 else os.environ["MULTIAGENTE_SECRET_KEY"]
            correo_cfg = {
                "usuario": self._ent_correo_usuario.get().strip(),
                "password": self._encrypt_val(mail_pw, key),
                "imap_server": self._ent_correo_imap.get().strip(),
                "imap_puerto": int(self._ent_correo_puerto.get().strip() or "143"),
                "destinatarios_grupal": [
                    l.strip() for l in self._ajustes_texto_grupal.get("1.0", "end-1c").split("\n") if l.strip()
                ],
                "destinatarios_individual": [
                    l.strip() for l in self._ajustes_texto_ind.get("1.0", "end-1c").split("\n") if l.strip()
                ],
            }
            self.config["correo"] = correo_cfg

            # Documentos
            docs_cfg = {}
            for key_doc, ent in [("dorso_mic", self._ent_dorso_mic), ("dorso_crt", self._ent_dorso_crt),
                             ("dorso_pe", self._ent_dorso_pe), ("permiso_exportacion", self._ent_permiso_exp),
                             ("hoja_ruta", self._ent_hoja_ruta), ("sobre", self._ent_sobre)]:
                try:
                    docs_cfg[key_doc] = int(ent.get().strip())
                except ValueError:
                    pass
            self.config["documentos"] = docs_cfg

            # Rutas
            self.config["rutas"] = {
                "sobres": self._ent_ruta_sobres.get().strip(),
                "cobro": self._ent_ruta_cobro.get().strip(),
                "pc": self._ent_ruta_pc.get().strip(),
                "carga_terrestre_carpeta": self._ent_ruta_ct_carpeta.get().strip(),
                "carga_terrestre_nombre": self._ent_ruta_ct_nombre.get().strip(),
                "planillas_carga": self._ent_ruta_planillas.get().strip(),
                "descarga_mails": self._ent_ruta_descarga.get().strip(),
                "escritorio_nombre": self._ent_ruta_escritorio.get().strip(),
                "backup_pendrive": self._ent_ruta_backup.get().strip(),
                "mic_sellos": self._ent_ruta_mic_sellos.get().strip(),
                "crt_original": self._ent_ruta_crt_original.get().strip(),
            }

            # Valores
            try:
                precio_carpeta = int(self._ent_precio_carpeta.get().strip())
            except ValueError:
                precio_carpeta = 49000
            if hasattr(self, '_ajustes_texto_guardas'):
                guardas = [l.strip() for l in self._ajustes_texto_guardas.get("1.0", "end-1c").split("\n") if l.strip()]
            else:
                guardas = self._cfg_obtener("valores", "guardas", ["Gonzalez"])
            self.config["valores"] = {
                "precio_carpeta": precio_carpeta,
                "guardas": guardas if guardas else ["Gonzalez"],
            }

            # Súper Auto configuración
            self.config["super_auto"] = {
                "pasos": {
                    "sobre": self._super_check_sobre.get(),
                    "permiso": self._super_check_permiso.get(),
                    "hoja_ruta": self._super_check_hoja_ruta.get(),
                    "recibo_ata": self._super_check_recibo.get(),
                    "aplicar_guarda": self._super_check_guarda.get(),
                    "completar_planillas": self._super_check_planillas.get(),
                }
            }

            # Google Drive (mantener token si ya existe)
            token_existente = self._cfg_obtener("google_drive", "token", None)
            self.config["google_drive"] = {
                "client_id": self._ent_drive_client_id.get().strip(),
                "client_secret": self._ent_drive_client_secret.get().strip(),
                "folder_id": self._ent_drive_folder_id.get().strip(),
            }
            if token_existente:
                self.config["google_drive"]["token"] = token_existente

            self.config["seguridad"] = {
                "password": self._encrypt_val(pw1, os.environ["MULTIAGENTE_SECRET_KEY"]),
            }

            self._guardar_config()
            self._ajustes_lbl_status.configure(text="✓ Configuración guardada correctamente.")
            self.after(3000, lambda: self._ajustes_lbl_status.configure(text=""))
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

    # ═══════════════════════════════════════════════════════════════════
    # CONFIRMACIÓN DE SALIDA
    # ═══════════════════════════════════════════════════════════════════
    def _confirmar_salida(self):
        if self.tarea_activa:
            ok = messagebox.askyesno(
                "Tarea en ejecución",
                "Hay una tarea en curso. ¿Está seguro de que desea salir?"
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
