# Ticket Pesada Correo Specification

## Purpose

Fetch weigh ticket PDFs from an IMAP inbox filtered by a configurable sender (`remitente_balanza`), let the user review candidates, and deliver selected PDFs to the OCR pipeline.

## Requirements

### Requirement: Connect via IMAP with existing credentials

The system MUST reuse the existing IMAP connection config (`self.config["correo"]`) to connect. If the connection fails, the system MUST report the error and abort.

#### Scenario: Successful IMAP connection

- GIVEN a valid IMAP configuration with host, port, user, and password
- WHEN the user clicks "Buscar en Correo"
- THEN the system connects to the IMAP server
- AND a progress indicator is shown

#### Scenario: IMAP connection failure

- GIVEN an invalid or unreachable IMAP configuration
- WHEN the system attempts to connect
- THEN the error is logged in the status area
- AND the user is informed that the connection failed

### Requirement: Filter by configured sender

The system MUST search only emails from the sender specified in `self.config["correo"]["remitente_balanza"]`. The user MUST configure this field in Ajustes > Correo before fetching.

#### Scenario: No remitente configured

- GIVEN `remitente_balanza` is empty or not set
- WHEN the user clicks "Buscar en Correo"
- THEN the system SHALL show a warning: "Configurá el Remitente Balanza en Ajustes > Correo primero"
- AND no IMAP search is attempted

### Requirement: Scan recent emails and filter by PDF attachment

The system MUST prompt the user to enter how many recent emails to scan (default: 20). After scanning, it MUST show only those with PDF attachments for the user to select.

#### Scenario: Emails found with PDF attachments

- GIVEN recent emails from `remitente_balanza` contain PDF attachments
- WHEN scanning completes
- THEN a list of matching emails (subject, date, attachment name) is shown
- AND each email has a checkbox for selection

#### Scenario: No emails from sender in range

- GIVEN no emails from `remitente_balanza` exist in the scanned range
- WHEN scanning completes
- THEN a message is displayed: "No se encontraron emails de {remitente} en los últimos {N} mensajes"
- AND no rows are shown

#### Scenario: Emails exist but lack PDF attachments

- GIVEN emails from `remitente_balanza` exist but have no PDF attachments
- WHEN scanning completes
- THEN the list is empty
- AND a message indicates no PDF attachments were found

### Requirement: Download selected attachments

The system MUST download all user-selected PDF attachments to a temporary location and pass the file paths to the OCR pipeline.

#### Scenario: User selects and downloads PDFs

- GIVEN at least one email with a PDF attachment is checked
- WHEN the user confirms download
- THEN each selected attachment is saved to a temp directory
- AND the OCR pipeline is invoked with the file paths
