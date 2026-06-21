# Delta for Threading & UI Responsiveness

## Purpose

Fix UI freezes caused by blocking filesystem operations on the main thread, add a
shared scan helper, cap poll burst processing, add a timeout to `preguntar_reintentar`,
and fix variable access in `_mail_worker`.

## ADDED Requirements

### Requirement: Background filesystem scanning

The system MUST NOT perform `os.listdir()`, `os.scandir()`, or `os.walk()` calls
on the main thread for Desktop folder scanning. All filesystem scans that drive
UI population (import panel, guarda popup, control-final auto-scan) MUST execute
in a background `threading.Thread` (daemon=True). Results MUST be pushed back to
the main thread via `self.after(0, callback)`.

#### Scenario: Import panel scan does not freeze UI

- GIVEN the user clicks the import panel refresh button
- WHEN `_imp_escanear_carpetas` runs
- THEN the filesystem scan MUST execute in a background thread
- AND the UI remains responsive during the scan
- AND results populate the folder list via `self.after(0, ...)`

#### Scenario: Guarda popup scan does not freeze UI

- GIVEN the user opens the "Agregar Guarda" popup
- WHEN `_popup_agregar_guarda` scans for Desktop folders
- THEN the scan MUST execute in a background thread
- AND the popup only appears after the scan completes

#### Scenario: Control-final auto-scan does not freeze UI

- GIVEN auto-mode is ON and the user clicks "Control Final"
- WHEN `_control_final_auto_scan` runs
- THEN the Desktop scan MUST execute in a background thread
- AND the selection popup only appears after the scan completes

### Requirement: Shared scan helper

The system MUST provide a single shared function (e.g. `_scan_desktop_folders`)
that encapsulates the repeated Desktop-folder-scanning logic: walk Desktop 1 level
deep, find folders containing `*CONTENEDORES*.xlsx` or `*.xls`, return a list of
folder metadata dicts. All callers (`_imp_escanear_carpetas`, `_popup_agregar_guarda`,
`_control_final_auto_scan`, `_analizar_planillas`) MUST use this helper instead of
inline `os.listdir` loops.

#### Scenario: Helper returns folder metadata

- GIVEN the Desktop contains 3 folders with CONTENEDORES Excel
- WHEN `_scan_desktop_folders("planillas_carga")` is called
- THEN it returns a list of dicts with keys: `name`, `path`, `excel_path`, `pdf_count`
- AND folders without CONTENEDORES Excel are excluded

#### Scenario: Helper is callable from background thread

- GIVEN a background thread calls `_scan_desktop_folders`
- WHEN the scan runs
- THEN it MUST NOT touch any tkinter widgets
- AND it MUST return results safe to pass via `self.after(0, ...)`

### Requirement: Poll burst cap with adaptive throttle

The system MUST cap `_poll_log_queue` processing at a configurable maximum
(default: 100 messages per tick). When the queue is consistently empty for 3+
consecutive ticks, the poll interval MUST increase from 100ms to 500ms. When
messages resume, the interval MUST reset to 100ms.

#### Scenario: Normal throughput

- GIVEN the log queue has 10 messages
- WHEN `_poll_log_queue` fires
- THEN all 10 messages are processed
- AND the next poll fires after 100ms

#### Scenario: Queue saturated

- GIVEN the log queue has 500 messages
- WHEN `_poll_log_queue` fires
- THEN at most 100 messages are processed this tick
- AND the remaining 400 process over subsequent ticks

#### Scenario: Idle throttle

- GIVEN the log queue has been empty for 3 consecutive ticks
- WHEN `_poll_log_queue` fires
- THEN the next poll interval MUST be 500ms instead of 100ms

### Requirement: preguntar_reintentar timeout

The system MUST add a timeout (default: 30 seconds) to `threading.Event.wait()`
in `preguntar_reintentar`. If the timeout expires, the function MUST return
`False` (cancel) instead of blocking indefinitely.

#### Scenario: Dialog shown successfully

- GIVEN `preguntar_reintentar` is called from a background thread
- WHEN the retry dialog appears and the user clicks "Reintentar"
- THEN the function returns `True`

#### Scenario: Dialog timeout

- GIVEN `preguntar_reintentar` is called from a background thread
- WHEN the dialog fails to show or the user does not respond within 30 seconds
- THEN the function returns `False`
- AND no deadlock occurs

### Requirement: _mail_worker thread-safe variable access

The system MUST NOT access tkinter widget values (e.g. `self._mail_entry_cantidad.get()`)
directly from the `_mail_worker` background thread. The value MUST be read before
thread start and passed as a parameter, or read via `self.after(0, ...)` + Event
synchronization.

#### Scenario: Mail worker reads entry value safely

- GIVEN `_mail_worker` runs in a background thread
- WHEN it needs the "cantidad" value
- THEN the value was captured before thread start or via thread-safe sync
- AND no `TclError` or race condition occurs

## MODIFIED Requirements

### Requirement: preguntar_reintentar thread safety

`preguntar_reintentar` MUST work correctly from both the main thread and
background threads. When called from a background thread with a valid `parent`
widget, it MUST dispatch the dialog to the main thread via `parent.after(0, ...)`
and wait with a timeout. When called from the main thread, it MUST show the
dialog directly. When no parent is available from a background thread, it MUST
return `False` immediately (fail-safe).
(Previously: used `threading.Event.wait()` with no timeout, risking deadlock)

#### Scenario: Called from main thread

- GIVEN `preguntar_reintentar` is called on the main thread
- WHEN the dialog is shown
- THEN it blocks until the user responds
- AND returns the user's choice

#### Scenario: Called from background thread with parent

- GIVEN `preguntar_reintentar` is called from a background thread with a valid parent
- WHEN the dialog is dispatched via `parent.after(0, ...)`
- THEN `threading.Event.wait(timeout=30)` blocks until the user responds or timeout
- AND returns `False` on timeout

#### Scenario: Called from background thread without parent

- GIVEN `preguntar_reintentar` is called from a background thread with `parent=None`
- WHEN the function executes
- THEN it returns `False` immediately without attempting to show a dialog

## Out of Scope

- Changing worker thread architecture or introducing async/await
- Modifying the log queue message protocol
- Adding new UI panels or features
- Changing the number of messages processed per tick (100 is kept as default)
