# App UI — Cargar Datos Panel Container Fields

## Purpose

Extend the "Cargar Datos" panel in `ui_app.py` to display and compare the new `Contenedor` and `Tara Contenedor` fields extracted by the OCR pipeline.

## Requirements

### R1: Worker data propagation

The `_cargar_datos_worker()` function's inner `_procesar_texto()` MUST extract `Contenedor` and `Tara Contenedor` from `datos` (returned by `extraer_datos()`) and include them in the `ticket_data` dict.

#### Scenario: Worker includes container fields in ticket_data

- GIVEN a FLEXI/ISO ticket OCR result with container data
- WHEN `_procesar_texto()` builds `ticket_data`
- THEN `ticket_data["contenedor"]` MUST be the extracted container number
- AND `ticket_data["tara_contenedor"]` MUST be the extracted container tare weight

#### Scenario: Worker passes empty fields for terrestrial tickets

- GIVEN a terrestrial ticket OCR result without container data
- WHEN `_procesar_texto()` builds `ticket_data`
- THEN `ticket_data["contenedor"]` MUST be an empty string
- AND `ticket_data["tara_contenedor"]` MUST be 0 (zero float)

### R2: TreeView column expansion

The TreeView in `_panel_cargar_datos()` MUST add two new columns `"contenedor"` and `"tara_contenedor"` after the existing `"tara"` column and before `"permiso"`. Corresponding headers MUST be `"Contenedor"` and `"Tara Cont."` with widths 100 and 90 respectively.

#### Scenario: TreeView shows container columns

- GIVEN the Cargar Datos TreeView is initialized
- WHEN `_panel_cargar_datos()` configures columns
- THEN the TreeView MUST include `"contenedor"` and `"tara_contenedor"` columns
- AND their position MUST be between `"tara"` and `"permiso"`

### R3: Result processing includes new fields

The `_procesar_resultado_ocr()` function MUST read the new fields from `ticket` data and include them in both the match and no-match branch TreeView inserts, as well as in the comparison popup data dicts.

#### Scenario: Match branch inserts with container values

- GIVEN a match result with container fields populated
- WHEN `_procesar_resultado_ocr()` inserts the TreeView row
- THEN the values tuple MUST include `contenedor` and `tara_contenedor` between "tara" and "permiso" positions
- AND `_cargar_datos_comparacion[iid]["ticket"]` MUST contain `"Contenedor"` and `"Tara Contenedor"` keys
- AND `_cargar_datos_comparacion[iid]["contenedor"]` MUST contain the same keys from CONTENEDOR data
- AND `_cargar_datos_comparacion[iid]["ok"]` MUST contain comparison booleans for both new fields

#### Scenario: No-match branch inserts with empty container values

- GIVEN a ticket with no CONTENEDOR match
- WHEN `_procesar_resultado_ocr()` inserts the no-match row
- THEN the values tuple MUST include `contenedor` and `tara_contenedor` (as empty string and 0) at the correct position
- AND the insertion MUST NOT raise an exception due to column count mismatch

### R4: Comparison popup includes new fields

The `_abrir_comparacion()` popup MUST display the two new container fields in its comparison table, positioned after "Tara (kg)" and before "Permiso".

#### Scenario: Comparison popup shows container comparison

- GIVEN a ticket with container data and a CONTENEDOR match
- WHEN `_abrir_comparacion()` builds the popup
- THEN the `campos` list MUST include `"Contenedor"` and `"Tara Contenedor"` between `"Tara (kg)"` and `"Permiso"`
- AND each field MUST show green/red background based on the comparison result
