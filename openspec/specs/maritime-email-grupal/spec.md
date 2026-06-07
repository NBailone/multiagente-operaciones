# Maritime Email Grupal Specification

## Purpose

Send a single grupal email aggregating all `PLANILLA DE CARGA*.xlsx` files
found across ISO/FLEXI (maritime) folders, with a subject that reflects the
file count (singular/plural) and a plain-text body listing each attached
file. Maritime grupal MUST NOT include planillas from TERRESTRE folders.

## Requirements

### Requirement: Scan only ISO/FLEXI folders for planillas

The system MUST scan only folders classified as MARITIMO ISO or MARITIMO FLEXI
when collecting `PLANILLA DE CARGA*.xlsx` files for the maritime grupal email.
It MUST exclude TERRESTRE and COMPARTIDO folders from this scan.

#### Scenario: Planillas collected from mixed folder types

- GIVEN one TERRESTRE folder with `PLANILLA DE CARGA A.xlsx` and two FLEXI
  folders each with `PLANILLA DE CARGA B.xlsx` and `PLANILLA DE CARGA C.xlsx`
- WHEN the maritime grupal scan runs
- THEN only the two FLEXI planillas are collected
- AND the TERRESTRE planilla is NOT included

#### Scenario: No maritime planillas found

- GIVEN no ISO or FLEXI folders contain `PLANILLA DE CARGA*.xlsx`
- WHEN the maritime grupal scan runs
- THEN no maritime grupal email is produced
- AND the existing CARGA TERRESTRE grupal email is unaffected

### Requirement: Build subject singular or plural by count

When exactly one planilla file is collected, the subject MUST be
`PLANILLA DE CARGA`. When two or more are collected, the subject MUST be
`PLANILLAS DE CARGA`.

#### Scenario: Singular subject for one file

- GIVEN exactly one `PLANILLA DE CARGA*.xlsx` collected from maritime folders
- WHEN the grupal email is assembled
- THEN the subject is `PLANILLA DE CARGA`

#### Scenario: Plural subject for multiple files

- GIVEN two `PLANILLA DE CARGA*.xlsx` files collected from maritime folders
- WHEN the grupal email is assembled
- THEN the subject is `PLANILLAS DE CARGA`

### Requirement: Build body listing attached planillas

The body MUST be plain text with the structure:

```
Estimados,

Se adjunta(n) la(s) planilla(s) de carga correspondiente(s):

  • {name}

Saludos cordiales.
```

Where `{name}` is each planilla filename without extension. Use the singular
form `adjunta la planilla` for one file, plural `adjuntan las planillas` for
two or more.

#### Scenario: Body for two planillas

- GIVEN two planillas named `PLANILLA DE CARGA B.xlsx` and
  `PLANILLA DE CARGA C.xlsx`
- WHEN the grupal email body is assembled
- THEN the body contains both bullet points with the correct names
- AND uses the plural forms `adjuntan` / `las planillas`

#### Scenario: Body for one planilla

- GIVEN exactly one planilla file
- WHEN the grupal email body is assembled
- THEN the body uses the singular forms `adjunta` / `la planilla`

### Requirement: Send to grupal recipients

The system MUST send maritime grupal emails to the same 14 configured
recipients as the CARGA TERRESTRE grupal email (`destinatarios_grupal`).

#### Scenario: Recipient list matches terrestrial grupal

- GIVEN a maritime grupal email ready to send
- WHEN the `To` field is populated
- THEN it contains the same 14 recipients as the CARGA TERRESTRE grupal email
