# Ticket Pesada OCR Specification

## Purpose

Extract structured weigh ticket data from PDF files via OCR pipeline: PDF-to-image conversion (pdf2image), Tesseract OCR in Spanish, and regex parsing into 7 fields (patente camion, patente semi, conductor, DNI, neto, tara, permiso embarque).

## Requirements

### Requirement: Convert PDF to images

The system MUST use `pdf2image` (requiring Poppler binaries) to convert each PDF page to a PIL Image before OCR. If Poppler is not installed or not in PATH, the system MUST surface a clear error.

#### Scenario: Successful PDF conversion

- GIVEN a valid PDF weigh ticket
- WHEN `pdf_a_texto` processes it
- THEN each page is converted to one or more PIL Images
- AND the images are passed to Tesseract OCR

#### Scenario: Missing Poppler binary

- GIVEN Poppler is not installed on the system
- WHEN the panel attempts to open
- THEN an inline error message SHALL appear: "Poppler no encontrado. Descargalo de https://github.com/oschwartz10612/poppler-windows/releases/"
- AND no OCR operations proceed

### Requirement: Run Tesseract OCR with Spanish language

The system MUST invoke `pytesseract.image_to_string()` with `lang='spa'`. If Tesseract or the `spa` language pack is missing, the system MUST surface a clear error.

#### Scenario: Successful OCR extraction

- GIVEN a PIL Image containing a weigh ticket
- WHEN OCR runs with `lang='spa'`
- THEN the raw text string is returned for parsing

#### Scenario: Missing Tesseract binary

- GIVEN Tesseract-OCR is not installed
- WHEN the panel attempts to open
- THEN an inline error message SHALL appear: "Tesseract-OCR no encontrado. Descargalo de https://github.com/UB-Mannheim/tesseract/wiki"
- AND no OCR operations proceed

#### Scenario: Missing Spanish language pack

- GIVEN Tesseract is installed but `spa` language data is missing
- WHEN OCR runs
- THEN the system SHALL report the missing language pack
- AND suggest running `tesseract --list-langs` to verify

### Requirement: Parse extracted text into structured fields

The system MUST extract these 7 fields from OCR output: `patente_camion`, `patente_semi`, `conductor`, `dni`, `neto`, `tara`, `permiso_embarque`. Fields not found SHALL be reported as empty strings.

#### Scenario: All fields extracted from clear ticket

- GIVEN a high-quality scanned weigh ticket with all fields visible
- WHEN `extraer_datos` parses the OCR text
- THEN all 7 fields are populated with non-empty values

#### Scenario: Partial extraction from damaged ticket

- GIVEN a weigh ticket where some fields are smudged or cut off
- WHEN `extraer_datos` parses the OCR text
- THEN only readable fields are populated
- AND unreadable fields are returned as empty strings
