# App UI Specification — Carga de Datos

## Purpose

Add a "Cargar Datos" button and panel to the app sidebar, providing two input modes (IMAP email and local file picker), a color-coded comparison table, and a write action for validated ticket data.

## Requirements

### Requirement: Sidebar button and panel navigation

The system MUST add a "Cargar Datos" button in the sidebar between "Completar Planillas" (index 2) and "Enviar Correos" (index 3), shifting downstream buttons by one position. Clicking the button SHALL switch the main content area to `_panel_cargar_datos`.

#### Scenario: Button renders in sidebar

- GIVEN the main app window is open
- WHEN the sidebar is rendered
- THEN a button labeled "Cargar Datos" appears between "Completar Planillas" and "Enviar Correos"
- AND the active nav highlight moves correctly when clicked

### Requirement: Panel with dual-mode input and results table

The `_panel_cargar_datos` MUST contain a toolbar with two source buttons ("Buscar en Correo", "Seleccionar PDFs"), a progress bar, a scrollable TreeView with 4 columns (campo, ticket, contenedor, estado), and a "Escribir en CONTENEDORES" action button.

#### Scenario: Panel opens with empty state

- GIVEN the user clicks "Cargar Datos"
- WHEN the panel is rendered
- THEN the toolbar buttons are visible
- AND the TreeView is empty
- AND the "Escribir en CONTENEDORES" button is disabled

#### Scenario: Table populated after OCR

- GIVEN OCR results are available for at least one ticket
- WHEN the results queue is processed
- THEN the TreeView shows one row per ticket
- AND each field cell is colored green/red/gray per validation result
- AND the "Escribir en CONTENEDORES" button is enabled if any row has "✅ Ok"

### Requirement: Remitente Balanza in Ajustes

The system MUST add a text input labeled "Remitente Balanza" in the Correo tab of the Ajustes panel. The value SHALL persist in `self.config["correo"]["remitente_balanza"]` and be saved via `_guardar_ajustes`.

#### Scenario: Remitente field saves and restores

- GIVEN the user opens Ajustes > Correo
- THEN a "Remitente Balanza" text field is present
- WHEN the user enters a value and saves
- THEN the value is persisted in `ui_config.json` under `correo.remitente_balanza`
- AND it is restored when the panel reopens
