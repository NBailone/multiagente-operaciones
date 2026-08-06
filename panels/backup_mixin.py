"""BackupMixin.

Legacy block extracted from ui_app.py. Mixed into the main App class; relies
on attributes/methods living on the composed instance (_log, _set_log_panel,
_cfg_obtener, _cfg_obtener_rutas, _resolver_ruta, _abrir_excel_seguro,
_guardar_excel_seguro, _icons, _panel_frames, panel_container, tarea_activa).
"""

import os
import re
import threading
import shutil
from datetime import datetime

import customtkinter as ctk
from tkinter import messagebox

from constants import Palette, FONT_FAMILY


class BackupMixin:
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
            text="Back Up",
            image=self._icons["hard-drive"], compound="left",
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
            self.btn_backup_pendrive.configure(text="Back Up", state="normal", fg_color=Palette.SECONDARY)
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


