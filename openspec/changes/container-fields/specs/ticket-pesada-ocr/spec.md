# Ticket Pesada OCR — Container Fields

## Purpose

Extend the `extraer_datos()` field extraction and Excel export in `procesar_tickets.py` to recognize container number (`Sigla Contenedor`) and container tare weight (`Tara Contenedor`) from FLEXI/ISO AGD tickets.

## Requirements

### R1: CAMPOS list extension

The `CAMPOS` list MUST append two new entries: `"Contenedor"` and `"Tara Contenedor"`, in that order, at the end of the existing list.

#### Scenario: CAMPOS order

- GIVEN the `CAMPOS` list
- WHEN the module initializes
- THEN `"Contenedor"` MUST be the second-to-last entry AND `"Tara Contenedor"` MUST be the last entry

### R2: Container number extraction

The system MUST extract the container number from OCR text using the label `Sigla Contenedor:`. The value follows the label on the same line with format `LETTERS SPACE DIGITS-HYPHEN-DIGITS` (e.g., "MSMU 258531-2").

#### Scenario: FLEXI/ISO ticket with container number

- GIVEN OCR text containing `Sigla Contenedor: MSMU 258531-2`
- WHEN `extraer_datos()` runs
- THEN `datos["Contenedor"]` MUST equal `"MSMU 258531-2"`

#### Scenario: Terrestrial ticket without container number

- GIVEN OCR text that does NOT contain "Sigla Contenedor:"
- WHEN `extraer_datos()` runs
- THEN `datos["Contenedor"]` MUST be an empty string

### R3: Container tare extraction

The system MUST extract the tare weight that appears BEFORE the `Tara Contenedor:` label in the OCR text. The value is a 3-4 digit decimal number (e.g., `2.100`) followed by `Cert.Verif.INTI` text. The regex SHALL use a lookahead approach: capture the number preceding `Cert.Verif.INTI`.

#### Scenario: FLEXI/ISO ticket with tare weight

- GIVEN OCR text containing `...2.100 Cert.Verif.INTI - Balanza Egr. ... Tara Contenedor:...`
- WHEN `extraer_datos()` runs
- THEN `datos["Tara Contenedor"]` MUST equal `"2.100"`

#### Scenario: Terrestrial ticket without tare weight

- GIVEN OCR text that does NOT contain "Cert.Verif.INTI" or "Tara Contenedor:"
- WHEN `extraer_datos()` runs
- THEN `datos["Tara Contenedor"]` MUST be an empty string

### R4: Universal extraction (no transport type branching)

The two new fields SHALL be extracted unconditionally for every ticket, without checking transport type. Terrestrial tickets naturally yield empty strings because their OCR text lacks the container labels.

#### Scenario: No regression for terrestrial tickets

- GIVEN OCR text from a terrestrial ticket with no "Sigla Contenedor:" or "Cert.Verif.INTI"
- WHEN `extraer_datos()` runs
- THEN all previously extracted fields MUST remain unchanged
- AND `datos["Contenedor"]` AND `datos["Tara Contenedor"]` MUST be empty strings

### R5: Excel column widths

The `anchos` dict in `crear_excel()` MUST include widths for the two new fields: `"Contenedor": 20` and `"Tara Contenedor": 16`.

#### Scenario: Excel output includes new columns

- GIVEN a list of ticket data with container fields
- WHEN `crear_excel()` generates the workbook
- THEN the worksheet MUST contain columns for "Contenedor" and "Tara Contenedor"
- AND those columns MUST have the specified widths
