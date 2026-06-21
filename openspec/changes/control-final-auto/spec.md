# Control Final — Auto Mode

## Purpose

Add an auto-mode toggle to the existing Control Final feature. When enabled,
replaces the manual file dialog with an automatic Desktop folder scan that
filters to folders containing Excel CONTENEDORES files, presents a selection
popup, and feeds selected folders to the existing `_control_final_worker`.

## Requirements

### Requirement: Auto mode toggle

The system MUST add a CTkSwitch labeled "Auto" next to the "Control Final"
button in the toolbar. The toggle state MUST persist in `ui_config.json`
under `config["control_final_auto"]`. Default: off.

#### Scenario: Toggle persists across sessions

- GIVEN the user enables the Auto toggle
- WHEN the app restarts
- THEN the Auto toggle MUST be ON
- AND clicking "Control Final" triggers auto-scan flow

#### Scenario: Toggle off restores manual flow

- GIVEN the Auto toggle is ON
- WHEN the user disables it
- THEN clicking "Control Final" MUST open the file dialog (current behavior)

### Requirement: Auto-scan Desktop folders

When Auto mode is ON, clicking "Control Final" MUST scan
`_resolver_ruta("planillas_carga", "Desktop")` + 1 level of subdirectories
for folders containing at least one Excel file with "CONTENEDORES" in the
filename (case-insensitive, `.xlsx` or `.xls`).

#### Scenario: Folders with CONTENEDORES found

- GIVEN Auto mode is ON
- WHEN the user clicks "Control Final"
- THEN the system scans Desktop + 1 level
- AND returns only folders containing *CONTENEDORES*.xlsx or *CONTENEDORES*.xls

#### Scenario: No folders found

- GIVEN Auto mode is ON
- AND no Desktop folders contain CONTENEDORES Excel
- WHEN the user clicks "Control Final"
- THEN a message MUST be shown: "No se encontraron carpetas con CONTENEDORES en el Desktop"
- AND no worker is started

### Requirement: Selection popup with checkboxes

The system MUST display a popup (CTkToplevel) listing qualifying folders.
Each folder shows: checkbox, folder name, PDF count, Excel filename, and
detected PDF types summary.

#### Scenario: Popup shows folder list

- GIVEN 3 Desktop folders contain CONTENEDORES Excel
- WHEN the auto-scan completes
- THEN the popup lists all 3 folders
- AND each has a checkbox (default: checked)
- AND the popup has "Todas" and "Ninguna" buttons

#### Scenario: Todas/Ninguna buttons

- GIVEN the popup is open with 5 folders
- WHEN the user clicks "Todas"
- THEN all checkboxes MUST be checked
- WHEN the user clicks "Ninguna"
- THEN all checkboxes MUST be unchecked

#### Scenario: User confirms selection

- GIVEN the popup is open
- WHEN the user clicks "Aceptar" with 2+ folders checked
- THEN the popup closes
- AND the worker processes ALL PDFs from ALL selected folders combined

#### Scenario: User cancels

- GIVEN the popup is open
- WHEN the user closes the popup or clicks "Cancelar"
- THEN no worker is started

### Requirement: PDF type detection hints

The system MUST classify PDFs by reading text content (same logic as current
`_control_final_worker`). For the popup display, the system MAY show
filename-based hints but MUST NOT use filename for final classification.

#### Scenario: Filename hints in popup

- GIVEN a folder with `scan_001.pdf` and `26AR12345.pdf`
- WHEN the popup renders
- THEN the folder row MAY show "2 scanned, 1 MIC" as a hint
- AND the worker still classifies by text content

### Requirement: Integration with existing worker

The auto-scan flow MUST call the existing `_control_final_worker(pdfs, excels)`
with the combined list of PDFs and Excel files from all selected folders.
No changes to worker logic.

#### Scenario: Worker receives combined file lists

- GIVEN the user selects 2 folders from the popup
- WHEN the worker starts
- THEN `pdf_paths` contains ALL PDFs from both folders
- AND `excel_paths` contains ALL CONTENEDORES Excel from both folders
- AND worker classification logic is unchanged

#### Scenario: Worker error handling unchanged

- GIVEN auto-scan passes files to the worker
- WHEN a PDF fails OCR or classification
- THEN the same error handling as manual mode applies
