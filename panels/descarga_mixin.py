"""DescargaMixin — Panel de Descarga de Mails (IMAP).

Legacy block extracted from ui_app.py. Methods are mixed into the main
App class; they rely on attributes/methods living on the composed instance
(_log, _set_log_panel, _cfg_obtener, _cfg_obtener_rutas, _resolver_ruta,
_scan_desktop_folders, _detectar_tipo_carpeta, _clasificar_tipo_transporte,
_limpiar_mail, _set_status, _icons, _panel_frames, panel_container, tarea_activa).
"""

import os
import re
import string
import time
import threading
import shutil
import imaplib
import email
import xlrd
from datetime import datetime

import customtkinter as ctk
from tkinter import ttk, messagebox

from constants import Palette, FONT_FAMILY, IMAP_SERVER, PUERTO_IMAP


class DescargaMixin:
    # ═══════════════════════════════════════════════════════════════════
    # PANEL: DESCARGAR MAILS (placeholder)
    # ═══════════════════════════════════════════════════════════════════
    def _panel_descargar(self):
        if "descargar" in self._panel_frames:
            return
        frame = ctk.CTkFrame(self.panel_container, fg_color="transparent")
        self._panel_frames["descargar"] = frame
        # frame.pack(fill="both", expand=True) — packed by _cambiar_panel_forzado

        # ── Toolbar 1: botones de búsqueda (fila 1) ──────────────────
        self._mail_row1 = ctk.CTkFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER, height=44
        )
        self._mail_row1.pack(fill="x", pady=(0, 2))
        self._mail_row1.pack_propagate(False)

        # Botón 1: Buscar y Descargar (modo automático, comportamiento original)
        self._mail_btn_buscar = ctk.CTkButton(
            self._mail_row1,
            text="Buscar y Descargar",
            image=self._icons["mail-search"], compound="left",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=30, width=160,
            command=self._mail_ejecutar,
        )
        self._mail_btn_buscar.pack(side="left", padx=3, pady=3)

        ctk.CTkLabel(
            self._mail_row1, text="Últimos:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(side="left", padx=(1, 0))

        self._mail_entry_cantidad = ctk.CTkEntry(
            self._mail_row1, width=35, height=26,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=4,
        )
        self._mail_entry_cantidad.insert(0, self.config.get("descarga_mails", {}).get("papeles", "2"))
        self._mail_entry_cantidad.pack(side="left")

        ctk.CTkLabel(
            self._mail_row1, text="mails",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(side="left", padx=(1, 6))

        # Separador
        ctk.CTkLabel(
            self._mail_row1, text="│",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.BORDER,
        ).pack(side="left", padx=2)

        # Botón 2: Buscar con reglas (modo interactivo, el usuario elige qué descargar)
        self._mail_btn_buscar_reglas = ctk.CTkButton(
            self._mail_row1,
            text="Buscar",
            image=self._icons["search"], compound="left",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.SECONDARY, hover_color=Palette.SECONDARY_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=30, width=160,
            command=self._mail_ejecutar_buscar,
        )
        self._mail_btn_buscar_reglas.pack(side="left", padx=3, pady=3)

        ctk.CTkLabel(
            self._mail_row1, text="Últimos:",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(side="left", padx=(1, 0))

        self._mail_entry_cantidad_reglas = ctk.CTkEntry(
            self._mail_row1, width=35, height=26,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=4,
        )
        self._mail_entry_cantidad_reglas.insert(0, self.config.get("descarga_mails", {}).get("reglas", "4"))
        self._mail_entry_cantidad_reglas.pack(side="left")

        ctk.CTkLabel(
            self._mail_row1, text="mails",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(side="left", padx=(1, 6))

        # Separador + Botón 3: Mail sin filtros + entry (se baja a fila 2 si no entra)
        self._mail_ultimos_row1 = ctk.CTkFrame(self._mail_row1, fg_color="transparent")

        ctk.CTkLabel(
            self._mail_ultimos_row1, text="│",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.BORDER,
        ).pack(side="left", padx=2)

        self._mail_btn_ultimos = ctk.CTkButton(
            self._mail_ultimos_row1,
            text="Mail sin filtros",
            image=self._icons["mail"], compound="left",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.SECONDARY, hover_color=Palette.SECONDARY_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=30, width=160,
            command=self._mail_ejecutar_ultimos,
        )
        self._mail_btn_ultimos.pack(side="left", padx=3, pady=3)

        self._mail_entry_sin_filtro_r1 = ctk.CTkEntry(
            self._mail_ultimos_row1, width=40, height=28,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=4,
        )
        self._mail_entry_sin_filtro_r1.insert(0, self.config.get("descarga_mails", {}).get("sin_filtro", "20"))
        self._mail_entry_sin_filtro_r1.pack(side="left", padx=(4, 0))

        ctk.CTkLabel(
            self._mail_ultimos_row1, text="mails",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(side="left", padx=(1, 4))

        # Botón Limpiar (siempre en fila 1, a la derecha)
        self.btn_limpiar_mail = ctk.CTkButton(
            self._mail_row1,
            text="Limpiar",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color="transparent",
            hover_color=Palette.BG_HOVER,
            text_color=Palette.TEXT_PRIMARY,
            corner_radius=4, height=30, width=70,
            command=self._limpiar_mail,
        )
        self.btn_limpiar_mail.pack(side="right", padx=4)

        # ── Fila 2: Descargar seleccionados + Mail sin filtros (baja si no entra arriba) ──
        self._mail_row2 = ctk.CTkFrame(
            frame, fg_color=Palette.BG_CARD, corner_radius=8,
            border_width=1, border_color=Palette.BORDER, height=40
        )
        self._mail_row2.pack(fill="x", pady=(0, 2))
        self._mail_row2.pack_propagate(False)

        self._mail_btn_descargar_sel = ctk.CTkButton(
            self._mail_row2,
            text="Descargar Selección",
            image=self._icons["cloud-download"], compound="left",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=30, width=160,
            command=self._mail_descargar_seleccionados,
        )
        self._mail_btn_descargar_sel.pack(side="left", padx=4, pady=5)

        # Botón nuevo: descarga los adjuntos PDF de los mails marcados como
        # archivos sueltos Escaneos\Scan1.pdf, Scan2.pdf... (numeración continua)
        self._mail_btn_descargar_tickets = ctk.CTkButton(
            self._mail_row2,
            text="Descargar Tickets",
            image=self._icons["cloud-download"], compound="left",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.SECONDARY, hover_color=Palette.SECONDARY_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=30, width=160,
            command=self._mail_tickets_seleccionados,
        )
        self._mail_btn_descargar_tickets.pack(side="left", padx=4, pady=5)

        # Contenedor para "Mail sin filtros" en fila 2 (se muestra/oculta dinámicamente)
        self._mail_ultimos_container = ctk.CTkFrame(self._mail_row2, fg_color="transparent")
        # No se packea aquí — se gestiona por _mail_on_resize

        ctk.CTkLabel(
            self._mail_ultimos_container, text="│",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.BORDER,
        ).pack(side="left", padx=2)

        self._mail_btn_ultimos_2 = ctk.CTkButton(
            self._mail_ultimos_container,
            text="Mail sin filtros",
            image=self._icons["mail"], compound="left",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            fg_color=Palette.SECONDARY, hover_color=Palette.SECONDARY_HOVER,
            text_color=Palette.WHITE, corner_radius=6, height=30, width=160,
            command=self._mail_ejecutar_ultimos,
        )
        self._mail_btn_ultimos_2.pack(side="left", padx=3, pady=3)

        self._mail_entry_sin_filtro = ctk.CTkEntry(
            self._mail_ultimos_container, width=40, height=28,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=4,
        )
        self._mail_entry_sin_filtro.insert(0, self.config.get("descarga_mails", {}).get("sin_filtro", "20"))
        self._mail_entry_sin_filtro.pack(side="left", padx=(4, 0))

        ctk.CTkLabel(
            self._mail_ultimos_container, text="mails",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=Palette.TEXT_PRIMARY,
        ).pack(side="left", padx=(1, 4))

        self._mail_progress = ctk.CTkProgressBar(
            self._mail_row2, width=120, height=8, corner_radius=4,
            fg_color=Palette.BG_INPUT, progress_color=Palette.ACCENT,
        )
        self._mail_progress.pack(side="right", padx=10)
        self._mail_progress.set(0)

        self._mail_lbl_estado = ctk.CTkLabel(
            self._mail_row2,
            text="",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_MUTED,
        )
        self._mail_lbl_estado.pack(side="right", padx=6)

        # ── Responsive: si no entra "Mail sin filtros" en fila 1, bajarlo a fila 2 ──
        def _mail_sync_entries():
            """Sincronizar el valor del entry visible al oculto."""
            try:
                if self._mail_ultimos_row1.winfo_ismapped():
                    val = self._mail_entry_sin_filtro_r1.get().strip()
                else:
                    val = self._mail_entry_sin_filtro.get().strip()
                for entry in (self._mail_entry_sin_filtro_r1, self._mail_entry_sin_filtro):
                    entry.delete(0, "end")
                    entry.insert(0, val)
            except Exception:
                pass

        def _mail_on_resize(event):
            was_mapped = self._mail_ultimos_row1.winfo_ismapped()
            if event.width >= 850:
                if not was_mapped:
                    _mail_sync_entries()
                self._mail_ultimos_row1.pack(side="left")
                self._mail_ultimos_container.pack_forget()
            else:
                if was_mapped:
                    _mail_sync_entries()
                self._mail_ultimos_row1.pack_forget()
                self._mail_ultimos_container.pack(side="left", padx=(4, 0))

        frame.bind("<Configure>", _mail_on_resize)

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

        # Restaurar anchos de columna guardados (si existen)
        _saved_cols = self.config.get("mail_tree_columns", {})
        self._mail_tree.column("sel", width=_saved_cols.get("sel", 35), anchor="center", stretch=False)
        self._mail_tree.column("asunto", width=_saved_cols.get("asunto", 280), anchor="w")
        self._mail_tree.column("fecha", width=_saved_cols.get("fecha", 130), anchor="center")
        self._mail_tree.column("adjuntos", width=_saved_cols.get("adjuntos", 70), anchor="center")
        self._mail_tree.column("carpeta", width=_saved_cols.get("carpeta", 260), anchor="w")

        # Bind para toggle del checkbox al clickear en la columna "✓"
        self._mail_tree.bind("<ButtonRelease-1>", self._mail_toggle_check)

        # Scrollbar vertical
        scroll_v = ctk.CTkScrollbar(
            tabla_frame, orientation="vertical",
            fg_color=Palette.BG_CARD, button_color=Palette.TEXT_MUTED,
            button_hover_color=Palette.TEXT_SECONDARY,
        )
        scroll_v.pack(side="right", fill="y", padx=(0, 2), pady=2)
        self._mail_tree.configure(yscrollcommand=scroll_v.set)
        scroll_v.configure(command=self._mail_tree.yview)

        # Scrollbar horizontal
        scroll_h = ctk.CTkScrollbar(
            tabla_frame, orientation="horizontal",
            fg_color=Palette.BG_CARD, button_color=Palette.TEXT_MUTED,
            button_hover_color=Palette.TEXT_SECONDARY,
        )
        scroll_h.pack(side="bottom", fill="x", padx=2, pady=(0, 2))
        self._mail_tree.configure(xscrollcommand=scroll_h.set)
        scroll_h.configure(command=self._mail_tree.xview)

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
        try:
            cantidad = int(self._mail_entry_cantidad.get().strip())
        except (ValueError, AttributeError):
            cantidad = 2
        t = threading.Thread(target=self._mail_worker, args=(cantidad,), daemon=True)
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
            # Leer del entry visible (fila 1 o fila 2)
            if self._mail_ultimos_row1.winfo_ismapped():
                raw = self._mail_entry_sin_filtro_r1.get().strip()
            else:
                raw = self._mail_entry_sin_filtro.get().strip()
            cantidad = int(raw)
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
            remitente_papeles = self._cfg_obtener_correo("remitente_papeles", "")
            if remitente_papeles and remitente_papeles.lower() not in remitente.lower():
                return
        match_date = re.search(r"Date:\s*(.+)", header, re.IGNORECASE)
        fecha_str = match_date.group(1).strip() if match_date else ""
        try:
            fecha_dt = parsedate_to_datetime(fecha_str)
        except Exception:
            fecha_dt = datetime.min
        resultados.append((fecha_dt, fecha_str, mid, asunto))

    def _mail_worker(self, cantidad):
        """Modo 0: Descarga automática de los N mails más nuevos que cumplen las reglas."""
        self._set_log_panel("descargar")
        resultados = []
        try:
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
            self._log(f"{len(seleccionados)} mails encontrados. Marcá los que quieras descargar y presioná 'Descargar Selección'.")

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

    # ── Descargar Tickets: adjuntos sueltos en Escritorio\Escaneos\ScanN ──

    def _mail_siguiente_scan(self, carpeta_escaneos):
        """Devuelve el próximo número ScanN libre según los archivos existentes."""
        import re
        max_n = 0
        if os.path.isdir(carpeta_escaneos):
            for f in os.listdir(carpeta_escaneos):
                m = re.match(r"^Scan(\d+)", f)
                if m:
                    max_n = max(max_n, int(m.group(1)))
        return max_n + 1

    def _mail_tickets_seleccionados(self):
        """Descarga los adjuntos de los mails marcados como Escaneos/ScanN.ext."""
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
        self._mail_btn_descargar_sel.configure(state="disabled")
        self._mail_btn_descargar_tickets.configure(text="⏳  Descargando...", state="disabled")
        self._mail_progress.configure(mode="indeterminate")
        self._mail_progress.start()
        self._mail_lbl_estado.configure(text=f"Descargando tickets de {len(items)} mail(s)...")
        self._limpiar_log()
        t = threading.Thread(target=self._mail_tickets_worker, args=(items,), daemon=True)
        t.start()

    def _mail_tickets_worker(self, items_a_descargar):
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
            carpeta_escaneos = os.path.join(escritorio, "Escaneos")
            os.makedirs(carpeta_escaneos, exist_ok=True)
            n_scan = self._mail_siguiente_scan(carpeta_escaneos)
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
                adj_nombres = []
                for part in msg.walk():
                    if part.get_content_disposition() != "attachment":
                        continue
                    filename = part.get_filename()
                    if not filename:
                        continue
                    ext = os.path.splitext(filename)[1].lower()
                    if not ext:
                        ext = ".pdf"
                    destino = os.path.join(carpeta_escaneos, f"Scan{n_scan}{ext}")
                    with open(destino, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    self._log(f"     → Escaneos\\Scan{n_scan}{ext}  ({filename})")
                    adj_nombres.append(os.path.basename(destino))
                    n_scan += 1
                if adj_nombres:
                    resultados.append((asunto, adj_nombres, "Escaneos"))
                    self._mail_data[item]["downloaded"] = True
                    self._mail_data[item]["checked"] = False
                    total_adj = len(adj_nombres)
                    self.after(0, lambda i=item, a=asunto, adj=total_adj:
                        self._mail_tree.set(i, "sel", "✓") or
                        self._mail_tree.set(i, "adjuntos", str(adj)) or
                        self._mail_tree.set(i, "carpeta", "Escaneos"))
                else:
                    self._log(f"     ⚠ El mail no tiene adjuntos")
            mail.logout()
            self._log(f"COMPLETADO: {sum(len(r[1]) for r in resultados)} archivo(s) guardado(s) en {carpeta_escaneos}.")
        except Exception as e:
            self._log(f"ERROR: {e}")
            resultados = []
        finally:
            self.after(0, lambda: self._mail_done(resultados, modo=4))

    def _mail_done(self, resultados, modo=0):
        self.tarea_activa = False
        try:
            self._mail_btn_buscar.configure(text="Buscar y Descargar", state="normal")
            self._mail_btn_buscar_reglas.configure(text="Buscar", state="normal")
            self._mail_btn_ultimos.configure(text="Mail sin filtros", state="normal")
            self._mail_btn_descargar_sel.configure(text="Descargar Selección", state="normal")
            self._mail_btn_descargar_tickets.configure(text="Descargar Tickets", state="normal")
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
            elif modo == 4:
                total = sum(len(r[1]) for r in resultados)
                self._mail_lbl_estado.configure(text=f"Tickets descargados: {total} archivo(s) en Escaneos")
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

