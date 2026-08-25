"""AjustesMixin — Panel de Configuración.

Legacy block extracted from ui_app.py. Mixed into the main App class; relies
on attributes/methods living on the composed instance (_log, _set_log_panel,
_cfg_obtener, _cfg_obtener_correo, _cfg_obtener_docs, _cfg_obtener_rutas,
_resolver_ruta, _guardar_config, _encrypt_val, _decrypt_val, _clave_encriptacion,
_decrypt_api_key, _verificar_password_maestra, _mostrar_dialogo_password_ajustes,
_get_font_sizes, _get_popup_geometry, _calc_popup_height, _icons, _panel_frames).
"""

import os
import re
import threading

import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog

from constants import (Palette, FONT_FAMILY, FONT_MONO,
                       IMAP_SERVER, PUERTO_IMAP,
                       DESTINATARIOS_GRUPAL, DESTINATARIOS_INDIVIDUAL)


class AjustesMixin:
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
            ("correo", "Correo", "settings"),
            ("documentos", "Documentos", "file-text"),
            ("descarga", "Descarga Mails", "inbox"),
            ("rutas", "Rutas", "folder"),
            ("valores", "Valores", "banknote"),
            ("seguridad", "Seguridad", "lock"),
            ("ocr", "OCR", "bot"),
            ("apariencia", "Apariencia", "palette"),
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

        tab_keys = [t[0] for t in tab_names]

        def _reflow_tabs(event=None):
            fw = tabs_frame.winfo_width()
            if fw < 50:
                return
            pitch = 114  # 110 ancho + 4 padx
            cols = max(1, fw // pitch)
            for i, key in enumerate(tab_keys):
                self._ajustes_tabs[key].grid(row=i // cols, column=i % cols, padx=2, pady=2, sticky="w")

        for key, label, icon_key in tab_names:
            btn = ctk.CTkButton(
                tabs_frame, text=label,
                image=self._icons[icon_key] if icon_key else None,
                compound="left" if icon_key else "none",
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

        for key, _, _ in tab_names:
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
            text="Guardar",
            image=self._icons["save"], compound="left",
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
            pw_toggle_btn = ctk.CTkButton(
                erow, text="",
                image=self._icons["eye"],
                width=30, height=30,
                font=ctk.CTkFont(size=14),
                fg_color=Palette.BG_HOVER, hover_color=Palette.ACCENT_DIM,
                text_color=Palette.TEXT_MUTED, corner_radius=4,
                command=None,
            )
            pw_toggle_btn.pack(side="left", padx=(4, 0))
            def _toggle():
                self._pw_visible = not self._pw_visible
                e.configure(show="" if self._pw_visible else "*")
            pw_toggle_btn.configure(command=_toggle)
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

    def _ajustes_row_optionmenu(self, parent, label, values, default=""):
        """Crea fila con OptionMenu para seleccionar método."""
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=3)
        ctk.CTkLabel(
            row, text=label,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            text_color=Palette.TEXT_SECONDARY, anchor="w",
        ).pack(anchor="w")
        # Mapear display value a valor guardado
        value_map = {"PDF (rápido)": "pdf", "Excel (COM)": "excel"}
        reverse_map = {v: k for k, v in value_map.items()}
        display_default = reverse_map.get(default, default)
        var = ctk.StringVar(value=display_default)
        opt = ctk.CTkOptionMenu(
            row, variable=var, values=values,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_INPUT, button_color=Palette.ACCENT,
            button_hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.TEXT_PRIMARY,
            dropdown_fg_color=Palette.BG_CARD,
            dropdown_hover_color=Palette.BG_HOVER,
            dropdown_text_color=Palette.TEXT_PRIMARY,
            corner_radius=4, width=160, height=30,
        )
        opt.pack(anchor="w", pady=(2, 0))
        # Guardar reference para obtener valor real al guardar
        opt._value_map = value_map
        opt._var = var
        return opt

    def _ajustes_row_grid(self, parent, label, config_key, default, row, col, extra=None):
        """Crea Entry en grid para tab Escalas: label arriba, entry abajo con '85-100' al lado."""
        cell = ctk.CTkFrame(parent, fg_color="transparent")
        cell.grid(row=row, column=col, padx=4, pady=2, sticky="ew")
        ctk.CTkLabel(
            cell, text=label,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            text_color=Palette.TEXT_SECONDARY, anchor="w",
        ).pack(anchor="w")
        entry_row = ctk.CTkFrame(cell, fg_color="transparent")
        entry_row.pack(fill="x", pady=(2, 0))
        e = ctk.CTkEntry(
            entry_row, width=65, height=28,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=4,
        )
        e.insert(0, str(self._cfg_obtener_docs(config_key, default)))
        e.pack(side="left")
        if extra:
            ctk.CTkLabel(
                entry_row, text=extra,
                font=ctk.CTkFont(family=FONT_FAMILY, size=9),
                text_color=Palette.TEXT_MUTED, anchor="w",
            ).pack(side="left", padx=(2, 0))
        return e

    # ── TAB: CORREO ──────────────────────────────────────────────────
    def _ajustes_tab_correo(self, parent):
        # Contenedor horizontal: 2 columnas - Credenciales a la izquierda, Destinatarios a la derecha
        split = ctk.CTkFrame(parent, fg_color="transparent")
        split.pack(fill="x", padx=10, pady=(4, 0))

        # ── Columna Izquierda: Credenciales IMAP ─────────────────────────────
        col_izq = ctk.CTkFrame(split, fg_color="transparent")
        col_izq.pack(side="left", fill="y", padx=(4, 8))

        self._ajustes_seccion(col_izq, "Credenciales IMAP")
        self._ent_correo_usuario = self._ajustes_row(
            col_izq, "Usuario (email):", self._cfg_obtener_correo("usuario", ""), width=300)
        self._ent_correo_password = self._ajustes_row(
            col_izq, "Contraseña:", self._cfg_obtener_correo("password", ""),
            show="*", width=240, toggle_pw=True)
        self._ent_correo_imap = self._ajustes_row(
            col_izq, "Servidor IMAP:", self._cfg_obtener_correo("imap_server", IMAP_SERVER), width=280)
        self._ent_correo_puerto = self._ajustes_row(
            col_izq, "Puerto IMAP:", str(self._cfg_obtener_correo("imap_puerto", PUERTO_IMAP)), width=80)
        self._sent_correo_remitente_papeles = self._ajustes_row(
            col_izq, "Remitente Papeles (filtro):", self._cfg_obtener_correo("remitente_papeles", ""), width=280)

        # ── Columna Derecha: Destinatarios (apilados verticalmente) ─────────────
        col_der = ctk.CTkFrame(split, fg_color="transparent")
        col_der.pack(side="left", fill="both", expand=True, padx=(8, 4))

        self._ajustes_seccion(col_der, "Destinatarios — Planillas de Carga")
        default_grupal = "\n".join(self._cfg_obtener_correo("destinatarios_grupal", DESTINATARIOS_GRUPAL))
        self._ajustes_texto_grupal = ctk.CTkTextbox(
            col_der, height=100, width=400,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=4, border_width=1,
        )
        self._ajustes_texto_grupal.insert("1.0", default_grupal)
        self._ajustes_texto_grupal.pack(anchor="w", padx=14, pady=(2, 8))

        self._ajustes_seccion(col_der, "Destinatarios — Correo Individual")
        default_ind = "\n".join(self._cfg_obtener_correo("destinatarios_individual", DESTINATARIOS_INDIVIDUAL))
        self._ajustes_texto_ind = ctk.CTkTextbox(
            col_der, height=70, width=400,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color=Palette.BG_INPUT, border_color=Palette.BORDER,
            text_color=Palette.TEXT_PRIMARY, corner_radius=4, border_width=1,
        )
        self._ajustes_texto_ind.insert("1.0", default_ind)
        self._ajustes_texto_ind.pack(anchor="w", padx=14, pady=(2, 8))

        self._ajustes_seccion(col_der, "Remitente Balanza")
        self._ent_remitente_balanza = self._ajustes_row(
            col_der, "Remitente Balanza:", self.config.get("correo", {}).get("remitente_balanza", ""), width=300)

    # ── TAB: DOCUMENTOS ──────────────────────────────────────────────
    def _ajustes_tab_documentos(self, parent):
        # Tabview con 2 pestañas: Copias y Método | Escalas
        tabview = ctk.CTkTabview(
            parent,
            fg_color=Palette.BG_CARD,
            segmented_button_fg_color=Palette.BG_INPUT,
            segmented_button_selected_color=Palette.ACCENT,
            segmented_button_selected_hover_color=Palette.ACCENT_HOVER,
            text_color=Palette.TEXT_PRIMARY,
        )
        tabview.pack(fill="both", expand=True, padx=10, pady=(4, 0))

        # ── TAB 1: COPIAS Y MÉTODO ───────────────────────────────────
        tab_copias = tabview.add("Copias y Método")

        # Copias de Impresión - 1 fila x 6 columnas (grid)
        self._ajustes_seccion(tab_copias, "Copias de Impresión")
        grid_copias = ctk.CTkFrame(tab_copias, fg_color="transparent")
        grid_copias.pack(fill="x", padx=14, pady=4)
        for i in range(6):
            grid_copias.columnconfigure(i, weight=1)

        self._ent_permiso_exp = self._ajustes_row_grid(
            grid_copias, "Permiso", "permiso_exportacion", 2, row=0, col=0)
        self._ent_hoja_ruta = self._ajustes_row_grid(
            grid_copias, "Hoja Ruta", "hoja_ruta", 2, row=0, col=1)
        self._ent_sobre = self._ajustes_row_grid(
            grid_copias, "Sobre", "sobre", 1, row=0, col=2)
        self._ent_dorso_mic = self._ajustes_row_grid(
            grid_copias, "D. MIC", "dorso_mic", 15, row=0, col=3)
        self._ent_dorso_crt = self._ajustes_row_grid(
            grid_copias, "D. CRT", "dorso_crt", 4, row=0, col=4)
        self._ent_dorso_pe = self._ajustes_row_grid(
            grid_copias, "D. PE", "dorso_pe", 2, row=0, col=5)

        # Método de Impresión
        self._ajustes_seccion(tab_copias, "Método de Impresión")
        self._ent_metodo_sobre = self._ajustes_row_optionmenu(
            tab_copias, "Método Sobre:", ["PDF (rápido)", "Excel (COM)"],
            self._cfg_obtener_docs("metodo_sobre", "pdf"))
        self._ent_metodo_hoja_ruta = self._ajustes_row_optionmenu(
            tab_copias, "Método Hoja Ruta:", ["PDF (rápido)", "Excel (COM)"],
            self._cfg_obtener_docs("metodo_hoja_ruta", "pdf"))

        # Opciones
        self._ajustes_seccion(tab_copias, "Opciones")
        self._chk_intercalar = ctk.CTkCheckBox(
            tab_copias, text="Intercalar copias",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=Palette.ACCENT, hover_color=Palette.ACCENT_HOVER,
            border_color=Palette.BORDER, checkmark_color=Palette.WHITE,
            text_color=Palette.TEXT_PRIMARY,
        )
        if self._cfg_obtener_docs("intercalar", False):
            self._chk_intercalar.select()
        self._chk_intercalar.pack(anchor="w", padx=14, pady=(8, 0))
        ctk.CTkLabel(
            tab_copias,
            text="Marcado: 1,2,1,2 (juego completo por copia)\nDesmarcado: 1,1,2,2 (cada hoja junta)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_MUTED, anchor="w", justify="left",
        ).pack(anchor="w", padx=14, pady=(2, 0))

        # ── TAB 2: ESCALAS ───────────────────────────────────────────
        tab_escalas = tabview.add("Escalas")

        # Título ANTES del grid
        self._ajustes_seccion(tab_escalas, "Escalas Individuales (%)")

        # Grid 2 filas: Documentos (4 cols) + Dorsos (3 cols)
        grid_frame = ctk.CTkFrame(tab_escalas, fg_color="transparent")
        grid_frame.pack(fill="x", padx=14, pady=4)
        grid_frame.columnconfigure((0, 1, 2, 3), weight=1)

        # Fila 0: Documentos - 4 columnas: Sobre, Permiso, Hoja Ruta, Global
        self._ent_escala_sobre = self._ajustes_row_grid(
            grid_frame, "Sobre", "escala_sobre", 90, row=0, col=0)
        self._ent_escala_permiso = self._ajustes_row_grid(
            grid_frame, "Permiso", "escala_permiso", 90, row=0, col=1)
        self._ent_escala_hoja_ruta = self._ajustes_row_grid(
            grid_frame, "Hoja Ruta", "escala_hoja_ruta", 90, row=0, col=2)
        self._ent_escala_pdf = self._ajustes_row_grid(
            grid_frame, "Global", "escala_pdf", 90, row=0, col=3,
            extra="85-100. 92 = Adobe Ajustar. 100 = sin achicar")

        # Fila 1: Dorsos - 3 columnas: D. MIC, D. CRT, D. PE
        self._ent_escala_dorso_mic = self._ajustes_row_grid(
            grid_frame, "D. MIC", "escala_dorso_mic", 90, row=1, col=0)
        self._ent_escala_dorso_crt = self._ajustes_row_grid(
            grid_frame, "D. CRT", "escala_dorso_crt", 90, row=1, col=1)
        self._ent_escala_dorso_pe = self._ajustes_row_grid(
            grid_frame, "D. PE", "escala_dorso_pe", 90, row=1, col=2)

        # Nota
        ctk.CTkLabel(
            tab_escalas,
            text="✏️  Typear valor a mano  |  Validación 85-100 al guardar",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_MUTED,
        ).pack(anchor="w", padx=14, pady=(4, 0))

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
        ctk.CTkLabel(
            parent,
            text="Usá Examinar para elegir la carpeta. Si es un pendrive, la letra\n"
                 "se recorta sola. Si es un disco fijo, se guarda la ruta absoluta.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color=Palette.TEXT_MUTED,
        ).pack(anchor="w", padx=14, pady=(4, 8))

        # Contenedor horizontal: 3 columnas
        split = ctk.CTkFrame(parent, fg_color="transparent")
        split.pack(fill="x", padx=10, pady=(0, 4))

        # ── Columna Izquierda: Archivos Pendrive/PC + Escritorio + Backup/Sellos ─────
        col_izq = ctk.CTkFrame(split, fg_color="transparent")
        col_izq.pack(side="left", fill="y", padx=(4, 8))

        self._ajustes_seccion(col_izq, "Archivos en Pendrive / PC — Carpetas de búsqueda")
        self._ent_ruta_sobres = self._ajustes_row_browse(
            col_izq, "SOBRES (carpeta):",
            self._cfg_obtener_rutas("sobres", "TRABAJO\\01_PLANILLAS"),
            extra="Busca SOBRES_2026.xlsx")
        self._ent_ruta_cobro = self._ajustes_row_browse(
            col_izq, "COBRO (carpeta):",
            self._cfg_obtener_rutas("cobro", "TRABAJO\\01_PLANILLAS"),
            extra="Busca COBRO_2026.xlsx")
        self._ent_ruta_pc = self._ajustes_row_browse(
            col_izq, "PC (carpeta):",
            self._cfg_obtener_rutas("pc", "TRABAJO\\01_PLANILLAS"),
            extra="Busca PC.xlsx / PC_2026.xlsx")

        ctk.CTkFrame(col_izq, fg_color=Palette.DIVIDER, height=1).pack(fill="x", padx=14, pady=(10, 8))

        self._ajustes_seccion(col_izq, "Nombre del Escritorio")
        self._ent_ruta_escritorio = self._ajustes_row(
            col_izq, "Nombre de la carpeta Escritorio:",
            self._cfg_obtener_rutas("escritorio_nombre", "Desktop"),
            extra="Solo cambiar si tu SO usa otro nombre (ej: 'Desktop' en inglés)", width=200)

        ctk.CTkFrame(col_izq, fg_color=Palette.DIVIDER, height=1).pack(fill="x", padx=14, pady=(10, 8))

        self._ajustes_seccion(col_izq, "Backup Pendrive y Sellos")
        self._ent_ruta_backup = self._ajustes_row_browse(
            col_izq, "Backup Pendrive (carpeta):",
            self._cfg_obtener_rutas("backup_pendrive", "TRABAJO\\CARGAS"))
        self._ent_ruta_mic_sellos = self._ajustes_row_browse(
            col_izq, "FECHA MIC Y SELLOS (carpeta):",
            self._cfg_obtener_rutas("mic_sellos", "TRABAJO\\01_PLANILLAS"),
            extra="Busca FECHA MIC Y SELLOS.xlsx")
        self._ent_ruta_crt_original = self._ajustes_row_browse(
            col_izq, "FECHA CRT Y ORIGINAL (carpeta):",
            self._cfg_obtener_rutas("crt_original", "TRABAJO\\01_PLANILLAS"),
            extra="Busca FECHA CRT Y ORIGINAL.xlsx")

        # ── Columna Centro: Planilla Maestra + Contenedores ───────────────────────
        col_der = ctk.CTkFrame(split, fg_color="transparent")
        col_der.pack(side="left", fill="y", padx=(8, 4))

        self._ajustes_seccion(col_der, "Planilla Maestra — CARGA TERRESTRE")
        self._ent_ruta_ct_carpeta = self._ajustes_row_browse(
            col_der, "Carpeta:",
            self._cfg_obtener_rutas("carga_terrestre_carpeta", "TRABAJO\\01_PLANILLAS"))
        self._ent_ruta_ct_nombre = self._ajustes_row(
            col_der, "Nombre del archivo:",
            self._cfg_obtener_rutas("carga_terrestre_nombre", "CARGA TERRESTRE.xlsx"), width=280)

        ctk.CTkFrame(col_der, fg_color=Palette.DIVIDER, height=1).pack(fill="x", padx=14, pady=(10, 8))

        self._ajustes_seccion(col_der, "Contenedores de Carga y Descargas")
        self._ent_ruta_planillas = self._ajustes_row_browse(
            col_der, "Carpeta de Contenedores de Carga:",
            self._cfg_obtener_rutas("planillas_carga", "Desktop"),
            extra="Donde se buscan las carpetas con 'PLANILLA DE CARGA'")
        self._ent_ruta_descarga = self._ajustes_row_browse(
            col_der, "Carpeta de descarga de mails:",
            self._cfg_obtener_rutas("descarga_mails", "Desktop"),
            extra="Donde se guardan los adjuntos descargados de los mails")

        # ── Columna Derecha: (vacía para futuras opciones) ────────────────────────
        col_pdf = ctk.CTkFrame(split, fg_color="transparent")
        col_pdf.pack(side="left", fill="y", padx=(8, 4))
        # Reservada para futuras configuraciones

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

        self._ajustes_seccion(parent, "Planilla de Carga")
        self._ent_clave_pdf = self._ajustes_row(
            parent, "Clave de protección:",
            self._cfg_obtener("valores", "clave_pdf", "123"),
            extra="Contraseña que se aplica a la hoja 'Planilla de Carga' al extraerla.",
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
                                  ('_sent_correo_remitente_papeles', 'remitente_papeles'),
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
                                  ("hoja_ruta", "_ent_hoja_ruta"), ("sobre", "_ent_sobre"),
                                  ("escala_pdf", "_ent_escala_pdf"),
                                  ("escala_sobre", "_ent_escala_sobre"),
                                  ("escala_permiso", "_ent_escala_permiso"),
                                  ("escala_hoja_ruta", "_ent_escala_hoja_ruta"),
                                  ("escala_dorso_mic", "_ent_escala_dorso_mic"),
                                  ("escala_dorso_crt", "_ent_escala_dorso_crt"),
                                  ("escala_dorso_pe", "_ent_escala_dorso_pe")]:
                w = _g(attr)
                if w is not None:
                    try:
                        val = int(w.get().strip())
                        # Clamp 85-100 para todas las escalas
                        if key_doc.startswith("escala_"):
                            val = max(85, min(100, val))
                        docs_cfg[key_doc] = val
                    except ValueError:
                        pass
            # Métodos de impresión (pdf/excel)
            for key_doc, attr in [("metodo_sobre", "_ent_metodo_sobre"), ("metodo_hoja_ruta", "_ent_metodo_hoja_ruta")]:
                w = _g(attr)
                if w is not None:
                    # OptionMenu guarda "PDF (rápido)" / "Excel (COM)" → mapear a "pdf"/"excel"
                    value_map = {"PDF (rápido)": "pdf", "Excel (COM)": "excel"}
                    display_val = w._var.get()
                    docs_cfg[key_doc] = value_map.get(display_val, "pdf")
            w_chk = _g("_chk_intercalar")
            if w_chk is not None:
                docs_cfg["intercalar"] = bool(w_chk.get())
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
                              ('_mail_entry_sin_filtro', '_ent_descarga_sin_filtro'),
                              ('_mail_entry_sin_filtro_r1', '_ent_descarga_sin_filtro')]:
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
            w = _g('_ent_clave_pdf')
            if w is not None:
                valores_cfg["clave_pdf"] = w.get().strip() or "123"
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

