# Ticket Pesada Validacion Specification

## Purpose

Match OCR-extracted ticket data against existing CONTENEDOR records by permit digits, validate field-by-field, present a color-coded comparison table, and write approved values to the DATOS sheet.

## Requirements

### Requirement: Find matching CONTENEDOR file

The system MUST search the `planillas_carga` subdirectories for files matching `*CONTENEDORES*.*xls*`. It MUST match a ticket to a file by comparing the last 5 alphanumeric characters of `permiso_embarque` against the PE cell (Datos R5, column dependent on block offset). If no match exists, the user SHALL be prompted to select manually.

#### Scenario: CONTENEDOR file found by permit match

- GIVEN a ticket with `permiso_embarque` ending in "12345"
- AND a `*CONTENEDORES*.*xls*` file contains "12345" in its PE cell
- WHEN the matching algorithm runs
- THEN the file path is returned
- AND the DATOS sheet is opened for validation

#### Scenario: No CONTENEDOR matches permit

- GIVEN a ticket with `permiso_embarque` ending in "99999"
- AND no `*CONTENEDORES*.*xls*` file has that permit
- WHEN the matching algorithm runs
- THEN the ticket row status is set to "⚠ No encontrado"
- AND the user MAY manually assign a CONTENEDOR file via dropdown

### Requirement: Validate fields against CONTENEDOR data

The system MUST compare each extracted field against the corresponding value in the DATOS sheet. Fields SHALL be color-coded: green if matching, red if mismatching, gray if the ticket lacks the field.

#### Scenario: All fields match

- GIVEN a ticket where all 7 fields match CONTENEDOR DATOS values
- WHEN validation completes
- THEN each cell is colored green
- AND the row status is set to "✅ Ok"

#### Scenario: Some fields mismatch

- GIVEN a ticket where neto and tara differ from CONTENEDOR DATOS values
- WHEN validation completes
- THEN the matching cells are green
- AND the mismatching cells (neto, tara) are red
- AND the row status is set to "❌ Diferencia"

### Requirement: Write validated data to DATOS sheet

The system MUST write `Neto` → `PESO CARGA` cell and `Tara` → `TARA CONT` cell in the DATOS sheet using `openpyxl`. The "Escribir en CONTENEDORES" button MUST be enabled only when at least one row has "✅ Ok" status.

#### Scenario: Successful write to unlocked file

- GIVEN a row with "✅ Ok" status
- WHEN the user clicks "Escribir en CONTENEDORES"
- THEN Neto is written to the PESO CARGA cell (R{k*13+28}, C+6)
- AND Tara is written to the TARA CONT cell (R{k*13+29}, C+6)
- AND the file is saved
- AND a success message is logged

#### Scenario: CONTENEDOR file is locked

- GIVEN the CONTENEDOR file is open in Excel or otherwise locked
- WHEN the system attempts to write
- THEN an error is logged: "El archivo {filename} está bloqueado. Cerralo e intentá de nuevo."
- AND no cells are modified
