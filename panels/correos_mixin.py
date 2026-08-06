"""CorreosMixin — Panel de Correos.

Legacy block extracted from ui_app.py. Methods are mixed into the main
App class; they rely on attributes/methods living on the composed instance
(_log, _set_log_panel, _cfg_obtener, _cfg_obtener_rutas, _resolver_ruta,
_imap_conectar, _mail_toggle_check, _limpiar_correos, _mostrar_resumen,
_set_status, _icons, _panel_frames, panel_container, tarea_activa).
"""

import os
import re
import threading
import imaplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import customtkinter as ctk
from tkinter import ttk, messagebox

from constants import Palette, FONT_FAMILY


class CorreosMixin:
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
            text="Procesar y Despachar Correos",
            image=self._icons["play"], compound="left",
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
            text="Elegir Correos",
            image=self._icons["folder-open"], compound="left",
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
            btn_frame, text="Enviar Seleccionados",
            image=self._icons["send"], compound="left",
            width=200, height=34,
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

            # Ordenar planillas: agrupar por destino, compartidos juntos (CERRADO primero)
            def _sort_key_terrestre(p):
                nombre = os.path.basename(p).upper()
                # Extraer destino (VITAPRO, EWOS, etc.)
                dest = ""
                for d in ("VITAPRO","EWOS","NUTRECO","DICOAL","CARGILL","BIOMAR"):
                    if f"_{d}_" in nombre:
                        dest = d; break
                if not dest:
                    m = re.search(r"(?:TERRESTRE|ISO|FLEXI)_[^_]+_[^_]+_([A-Z]+)", nombre)
                    dest = m.group(1) if m else "ZZZ"
                # Detectar si es compartido y si cierra
                es_compartido = "COMPARTIDO" in nombre
                cierra_primero = 0 if es_compartido and "CERRADO" in nombre else 1
                # Extraer fracción (F1, F2, F3, F4)
                m_frac = re.search(r"_(F\d+)_", nombre)
                fraccion = m_frac.group(1) if m_frac else ""
                # Destino → compartido → cierra → fracción
                return (dest, 0 if es_compartido else 1, cierra_primero, fraccion, nombre)

            planillas_ordenadas = sorted(terr_planillas, key=_sort_key_terrestre)

            cuerpo = "Estimados,\n\nSe adjuntan las planillas de carga correspondientes:\n\n"
            for p in planillas_ordenadas:
                nombre_sin_ext = os.path.splitext(os.path.basename(p))[0]
                cuerpo += f"  • {nombre_sin_ext}\n"
            if ruta_maestra:
                cuerpo += f"  • {os.path.splitext(os.path.basename(ruta_maestra))[0]}\n"
            cuerpo += "\nSaludos cordiales."
            msg_grupal.attach(MIMEText(cuerpo, "plain"))

            n_adj_grupal = len(planillas_ordenadas)
            for p in planillas_ordenadas:
                adjuntar_archivo(msg_grupal, p)
            if ruta_maestra:
                adjuntar_archivo(msg_grupal, ruta_maestra)
                n_adj_grupal += 1
            else:
                self._log("ERROR: No se detectó 'CARGA TERRESTRE' en el pendrive.")
                self.after(0, lambda: messagebox.showerror(
                    "Planilla no encontrada",
                    "No se encontró la planilla 'CARGA TERRESTRE' en el pendrive.\n\n"
                    "Verificá que el pendrive esté conectado y que el archivo exista."
                ))

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

            # Ordenar planillas marítimas: agrupar por destino, compartidos juntos (CERRADO primero)
            def _sort_key_maritimo(p):
                nombre = os.path.basename(p).upper()
                dest = ""
                for d in ("VITAPRO","EWOS","NUTRECO","DICOAL","CARGILL","BIOMAR"):
                    if f"_{d}_" in nombre:
                        dest = d; break
                if not dest:
                    m = re.search(r"(?:TERRESTRE|ISO|FLEXI)_[^_]+_[^_]+_([A-Z]+)", nombre)
                    dest = m.group(1) if m else "ZZZ"
                es_compartido = "COMPARTIDO" in nombre
                cierra_primero = 0 if es_compartido and "CERRADO" in nombre else 1
                m_frac = re.search(r"_(F\d+)_", nombre)
                fraccion = m_frac.group(1) if m_frac else ""
                return (dest, 0 if es_compartido else 1, cierra_primero, fraccion, nombre)

            mar_ordenadas = sorted(mar_planillas, key=_sort_key_maritimo)

            if len(mar_planillas) == 1:
                cuerpo = "Estimados,\n\nSe adjunta la planilla de carga correspondiente:\n\n"
            else:
                cuerpo = "Estimados,\n\nSe adjuntan las planillas de carga correspondientes:\n\n"
            for p in mar_ordenadas:
                nombre_sin_ext = os.path.splitext(os.path.basename(p))[0]
                cuerpo += f"  • {nombre_sin_ext}\n"
            cuerpo += "\nSaludos cordiales."
            msg_grupal.attach(MIMEText(cuerpo, "plain"))

            for p in mar_ordenadas:
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
                text="Procesar y Despachar Correos",
                state="normal",
                fg_color=Palette.ACCENT,
            )
            self.btn_elegir_correos.configure(
                text="Elegir Correos",
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

