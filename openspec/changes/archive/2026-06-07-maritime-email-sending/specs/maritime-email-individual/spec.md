# Maritime Email Individual Specification

## Purpose

Send individual email notifications for ISO/FLEXI (maritime) folders — one
email per folder — with maritime-specific attachment selection (Contenedores.xlsx
+ get*.pdf), subject/body construction that omits MIC references, and suffix
extraction from ISO/FLEXI folder names.

## Requirements

### Requirement: Identify ISO/FLEXI folders for maritime individual email

The system MUST treat folders classified as MARITIMO ISO or MARITIMO FLEXI
(per folder-naming rules, not TERRESTRE and not COMPARTIDO) as candidates for
maritime individual email processing.

#### Scenario: ISO folder is processed individually

- GIVEN a MARITIMO ISO folder on the desktop
- WHEN `_correos_core` iterates folders for individual email
- THEN the folder enters the maritime individual branch
- AND it is NOT processed by the terrestrial individual branch

#### Scenario: FLEXI folder is processed individually

- GIVEN a MARITIMO FLEXI folder on the desktop
- WHEN `_correos_core` iterates folders for individual email
- THEN the folder enters the maritime individual branch

#### Scenario: TERRESTRE folder is excluded from maritime

- GIVEN a TERRESTRE folder on the desktop
- WHEN the folder is classified
- THEN it MUST NOT enter the maritime individual branch

### Requirement: Select attachments for ISO/FLEXI individual email

The system MUST attach `Contenedores.xlsx` and all `get*.pdf` files from the
folder. It MUST exclude `PLT*.pdf` and `MIC*.pdf` (current-year AR pattern)
from the attachment list.

#### Scenario: Maritime folder with mixed files

- GIVEN an ISO folder containing `Contenedores.xlsx`, `get001.pdf`, `get002.pdf`,
  `PLT1234.pdf`, and `MIC2456.pdf`
- WHEN the maritime individual email is assembled
- THEN the attachment list contains `Contenedores.xlsx`, `get001.pdf`, `get002.pdf`
- AND it does NOT contain `PLT1234.pdf` or `MIC2456.pdf`

#### Scenario: Maritime folder with only Contenedores

- GIVEN a FLEXI folder containing only `Contenedores.xlsx`
- WHEN the maritime individual email is assembled
- THEN the attachment list contains `Contenedores.xlsx`
- AND no PDF files are attached

### Requirement: Build subject with SALIDA/SALIDAS and no MIC

The subject MUST be `SALIDA, PLANILLA COMPLETA DE EXPORTACIÓN_{sufijo}` when
there is exactly one `get*.pdf` attachment, and `SALIDAS, PLANILLA COMPLETA DE
EXPORTACIÓN_{sufijo}` when there are two or more. The subject MUST NOT include
any MIC reference.

#### Scenario: Single get*.pdf — singular SALIDA

- GIVEN a maritime folder with exactly one `get*.pdf` file
- WHEN the subject is built
- THEN the subject starts with `SALIDA, PLANILLA COMPLETA DE EXPORTACIÓN_`

#### Scenario: Multiple get*.pdf — plural SALIDAS

- GIVEN a maritime folder with two `get*.pdf` files
- WHEN the subject is built
- THEN the subject starts with `SALIDAS, PLANILLA COMPLETA DE EXPORTACIÓN_`

#### Scenario: No get*.pdf — no subject generated

- GIVEN a maritime folder with zero `get*.pdf` files
- WHEN the subject is built
- THEN the system falls back to a reasonable default without SALIDA/SALIDAS

### Requirement: Build email body identical to subject

The body MUST be plain text identical to the subject line.

#### Scenario: Body equals subject

- GIVEN a maritime individual email with subject `SALIDA, PLANILLA COMPLETA DE EXPORTACIÓN_TRP`
- WHEN the email body is assembled
- THEN the body text equals `SALIDA, PLANILLA COMPLETA DE EXPORTACIÓN_TRP`

### Requirement: Extract suffix from ISO/FLEXI folder name

The system MUST derive the `{sufijo}` from the ISO/FLEXI folder name using a
meaningful portion of the name (e.g., the portion after the transport-type
segment, analogous to how TERRESTRE captures via `TERRESTRES?_(.*)`).

#### Scenario: Suffix from ISO folder

- GIVEN a folder named `10_06_26_ISO_ALLIED CHEMICALS_TRP`
- WHEN the suffix is extracted
- THEN the resulting suffix identifies the meaningful portion (e.g., `ALLIED CHEMICALS_TRP`)

#### Scenario: Suffix from FLEXI folder

- GIVEN a folder named `10_06_26_FLEXI_PT SURI TANI PEMUKA_EXOLGAN`
- WHEN the suffix is extracted
- THEN the resulting suffix identifies the meaningful portion

### Requirement: Send to individual recipients

The system MUST send maritime individual emails to the same three configured
recipients as terrestrial individual emails (`destinatarios_individual`).

#### Scenario: Recipient list matches terrestrial

- GIVEN a maritime individual email ready to send
- WHEN the `To` field is populated
- THEN it contains the same three recipients as a terrestrial individual email
