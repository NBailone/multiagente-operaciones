# Spec: mail-folder-naming

## Purpose

Derive the transport type of a mail-downloaded shipment from the
attached CONTENEDORES workbook and build the destination folder name
with a type-matching suffix. Type is data-derived, not user-selected.
Scope: CONTENEDORES reading (`_leer_xls_antiguo`,
`_leer_xlsx_moderno`) and folder-naming callers
(`_mail_nombre_carpeta`, `_mail_procesar_comparte`).

## Requirements

### Requirement: Detect transport type from CONTENEDORES

The system MUST classify each shipment from `Puerto Salida` and
`Peso Flexi` on the CHOFER sheet. `None`, `""`, and `-` (after `.strip()`) count as empty.

| `Puerto Salida` | `Peso Flexi` | Type |
|-----------------|--------------|------|
| empty / `-` | any | TERRESTRE |
| set | `0` or empty | MARITIMO ISO |
| set | positive number | MARITIMO FLEXI |

#### Scenario: Terrestrial — Puerto Salida empty or dash

- GIVEN `Puerto Salida` is empty or `-`
- WHEN the type is detected
- THEN the shipment is classified as TERRESTRE

#### Scenario: Maritime ISO — Puerto set, Peso Flexi zero or empty

- GIVEN `Puerto Salida = "TRP"` and `Peso Flexi` is `0` or empty
- WHEN the type is detected
- THEN the shipment is classified as MARITIMO ISO

#### Scenario: Maritime FLEXI — Puerto set, Peso Flexi has a number

- GIVEN `Puerto Salida = "EXOLGAN"` and `Peso Flexi = 1200`
- WHEN the type is detected
- THEN the shipment is classified as MARITIMO FLEXI

### Requirement: Build folder name with type-appropriate suffix

The folder name MUST consist of six common segments (`DD_MM_YYYY`, count, type, PE, carpeta, destinatario) plus one type-specific trailing segment: fraction for TERRESTRE, port code for MARITIMO.

| Type | Trailing | Example (tail) |
|------|----------|----------------|
| TERRESTRE | fraction | `...VITAPRO_F6` |
| MARITIMO ISO | port | `...ALLIED CHEMICALS_TRP` |
| MARITIMO FLEXI | port | `...PT SURI TANI PEMUKA_EXOLGAN` |

#### Scenario: Terrestrial folder ends with fraction

- GIVEN a TERRESTRE shipment with fraction `F6`
- WHEN the folder name is built
- THEN it ends with `_F6` and `TERRESTRE` is the third segment

#### Scenario: Maritime ISO folder ends with port

- GIVEN a MARITIMO ISO shipment with port `TRP`
- WHEN the folder name is built
- THEN it ends with `_TRP` and `ISO` is the third segment

#### Scenario: Maritime FLEXI folder ends with port

- GIVEN a MARITIMO FLEXI shipment with port `EXOLGAN`
- WHEN the folder name is built
- THEN it ends with `_EXOLGAN` and `FLEXI` is the third segment

### Requirement: Share case classifies each CONTENEDORES independently

When a mail has two CONTENEDORES files, the system MUST detect each type independently and create two folders, each per the suffix rules.

#### Scenario: Share with one terrestrial and one maritime

- GIVEN file A (empty `Puerto Salida`) and file B (`Puerto Salida = "TRP"`, `Peso Flexi = 0`) in one mail
- WHEN the share is processed
- THEN two folders are created: A ends with A's fraction, B ends with `_TRP`

### Requirement: Fallback when CONTENEDORES cannot be read

When CONTENEDORES is missing, unreadable, or has no CHOFER sheet, the system MUST return a generic name from the temp-folder basename and log a warning.

#### Scenario: Excel cannot be read

- GIVEN a CONTENEDORES file that raises an exception when opened
- WHEN the folder name is built
- THEN the system returns the temp-folder fallback and logs a warning

#### Scenario: CHOFER sheet missing

- GIVEN a CONTENEDORES file with no `CHOFER` sheet
- WHEN the folder name is built
- THEN the system returns the temp-folder fallback and logs a warning

### Requirement: Tuple read by downstream workers tolerates extra fields

The CONTENEDORES read result appends `puerto_salida` and `peso_flexi` as trailing elements. Workers that don't read the new fields MUST tolerate them on unpack.

#### Scenario: Planillas worker unpacks an 11-element tuple

- GIVEN a CONTENEDORES read result with 11 elements (9 original + `puerto_salida` + `peso_flexi`)
- WHEN `_planillas_worker` unpacks the tuple
- THEN the worker does not raise `ValueError: too many values to unpack`


