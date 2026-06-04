# Tasks: Usability Fixes

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~190 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | ask-on-risk |
| Chain strategy | size-exception |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Low

## Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | All 7 fixes | PR 1 (single commit) | ~190 lines, well under budget |

## Phase 1: Foundation (1 task)

- [x] 1.1 [T1] **Add xlutils to deps** — `requirements.txt` add `xlutils>=1.2.0`; `ui_app.py:33-37` add `("xlutils", "xlutils")` to `_instalar_deps_ui()` list

## Phase 2: Independent Fixes (4 tasks, any order)

- [x] 2.1 [T4] **Remove "PLANILLA DE CARGA:" line** — delete `cuerpo += "PLANILLA DE CARGA:\n"` at `ui_app.py:5402`
- [x] 2.2 [T5] **Add ATA y TARES config field** — label "ATA y TARES ($):" in `_ajustes_tab_valores()`, tkinter Entry with default `65000`; save in `_guardar_ajustes()` as `self.config["valores"]["ata_tares"]`; replace 4 `* 60000` at lines 4077, 4220, 4424, 4447 with `int(self._cfg_obtener("valores", "ata_tares", 65000))`
- [x] 2.3 [T6] **Increase dialog max height** — `ui_app.py:5512` change `480` to `560` in `_correos_popup_confirmacion()`
- [x] 2.4 [T8] **Remove dead btn_backup_drive** — delete button creation + pack (3248-3255), disable reference (3345), callback reference in `_backup_done()` (3595)

## Phase 3: Guarda Excel (2 tasks, sequential — 3.2 depends on 3.1)

- [x] 3.1 [T2] **Create `_escribir_guarda_en_archivo()` helper** — new App method: detect .xls (xlrd+xlutils) vs .xlsx (openpyxl), find "CHOFER" sheet, search col G rows 2-15 for "GUARDA", write `guarda_nombre` to col H same row, save workbook, return bool. Place near line 3189.
- [x] 3.2 [T3] **Replace 4 call sites with helper** — one-line call each: super-auto .xlsx (3158-3170), super-auto .xls (3126-3133), manual .xlsx (3868-3883), manual .xls (3840-3846). Each becomes `self._escribir_guarda_en_archivo(ruta, guarda_elegido)`.

## Phase 4: Drive Elimination (1 task, after T8/P2.4)

- [x] 4.1 [T7] **Eliminate Drive tab** — remove `("drive", ...)` from tabs list (5998), tab func call (6041); delete `_ajustes_tab_drive()` (6347-6386), `_drive_verificar_conexion()`, `_drive_autenticar()` (3454-3495), `_drive_buscar_o_crear_carpeta()` (3497-3507), `_drive_subir_sobrescrito()` (3509-3519), `_backup_drive_iniciar()` (3434-3452), `_backup_drive_worker()` (3522-3591); delete Drive save block in `_guardar_ajustes()` (6558-6566); remove `google_drive` from `ui_config.json` defaults
