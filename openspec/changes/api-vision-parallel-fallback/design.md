# Design: API Vision Parallel Fallback with Model Selection

## Technical Approach

Add `api_vision_con_fallback()` in `procesar_tickets.py` that distributes PDFs across multiple enabled API vision models using `concurrent.futures.ThreadPoolExecutor`. Failed PDFs retry with other models before falling to PaddleOCR. UI adds a master toggle + per-model checkboxes in `_ajustes_tab_ocr()`. Workers (`_control_final_worker`, `_cargar_datos_worker`) call the new function when parallel is enabled.

## Architecture Decisions

### Decision: ThreadPoolExecutor for parallelism

**Choice**: `concurrent.futures.ThreadPoolExecutor` (stdlib)
**Alternatives**: `multiprocessing.Pool`, asyncio with aiohttp, manual `threading.Thread`
**Rationale**: Threads avoid pickling issues with tkinter objects already in memory. `ThreadPoolExecutor` provides `as_completed()` for retry coordination. No new dependencies.

### Decision: Round-robin PDF distribution

**Choice**: Cyclic assignment across N models: `pdf[i] → models[i % N]`
**Alternatives**: Weighted distribution, dynamic load balancing, chunk-based splitting
**Rationale**: Simple, predictable, equal distribution. Each thread processes its slice sequentially. Dynamic balancing adds complexity with no benefit since all PDFs are similar size.

### Decision: Retry queue with model exclusion set

**Choice**: Failed PDFs enter a retry queue; each retry excludes models already tried for that PDF
**Alternatives**: Immediate retry on same model, weighted retry, retry-all
**Rationale**: Prevents retrying a model that already failed for a specific PDF (timeout/rate-limit). Exclusion set is per-PDF, not global — different PDFs may still use all models.

### Decision: Thread-safe result collection via Lock + dict

**Choice**: `threading.Lock` protecting a shared `dict` for results
**Alternatives**: `queue.Queue`, concurrent dict, atomic operations
**Rationale**: Simple. Only one thread writes per key (stem is unique per PDF). Lock guards the final merge. Queue adds unnecessary complexity.

## Data Flow

```
┌─────────────────────────────────────────────────────┐
│                   api_vision_con_fallback()         │
│                                                     │
│  1. Distribute PDFs round-robin across enabled      │
│     models → {model: [pdf_paths]}                   │
│                                                     │
│  2. ThreadPoolExecutor (workers = len(enabled_models))│
│     ┌──────────┐ ┌──────────┐ ┌──────────┐         │
│     │ Thread 0 │ │ Thread 1 │ │ Thread N │         │
│     │ Model A  │ │ Model B  │ │ Model N  │         │
│     │ pdf1,pdf4│ │ pdf2,pdf5│ │ pdf3,pdf6│         │
│     └────┬─────┘ └────┬─────┘ └────┬─────┘         │
│          │success      │FAIL        │success         │
│          ▼             ▼            ▼                │
│     datos[stem]   retry_queue   datos[stem]         │
│                                                     │
│  3. Retry loop: pop from queue, try other models    │
│     (exclusion set per PDF)                         │
│     → success → datos[stem]                         │
│     → all tried → paddleocr fallback → textos[stem] │
│                                                     │
│  4. Return {datos, textos, logs}                    │
└─────────────────────────────────────────────────────┘
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `procesar_tickets.py` | Modify | Add `api_vision_con_fallback()` after line 1826 (after `api_vision_extraer_datos`) |
| `ui_app.py` ~10128 | Modify | Add parallel toggle + per-model checkboxes after `_ent_vision_custom_models` textbox |
| `ui_app.py` ~10311-10328 | Modify | Save `parallel_enabled` and `parallel_model_states` in config |
| `ui_app.py` ~7700-7716 | Modify | `_control_final_worker`: branch on `parallel_enabled` |
| `ui_app.py` ~9255-9288 | Modify | `_cargar_datos_worker`: branch on `parallel_enabled` |

## Interfaces / Contracts

```python
# procesar_tickets.py — new function signature
def api_vision_con_fallback(
    rutas_pdfs: list[str],
    api_key: str,
    modelos: list[str],
    temperature: float = 0.1,
    max_tokens: int = 8192,
    timeout: int = 60,
    api_base: str = OPENROUTER_BASE,
    log_callback: Callable[[str], None] | None = None,
) -> dict:
    """Returns: {"datos": {stem: dict}, "textos": {stem: str}, "logs": list[str]}"""
```

```python
# Worker branch pattern (both workers follow same structure)
if parallel_enabled:
    enabled_models = [m for m, on in model_states.items() if on]
    result = procesar_tickets.api_vision_con_fallback(
        tickets_pdf, api_key, enabled_models,
        temperature=temperature, max_tokens=max_tokens,
        timeout=timeout, log_callback=log_fn,
    )
    api_datos_raw = result["datos"]
    textos_por_pdf = result["textos"]
    for msg in result["logs"]:
        self.log_queue.put(f"[{timestamp}] {msg}")
else:
    # existing single-model loop (unchanged)
```

```python
# Config schema addition (inside "api_vision" object)
{
  "parallel_enabled": false,
  "parallel_model_states": {
    "google/gemini-2.0-flash-exp:free": true,
    "meta-llama/llama-3.2-11b-vision:free": true,
    "google/gemini-2.5-flash": true,
    "qwen/qwen3.6-flash": true
  }
}
```

```python
# Thread-safe result dict
import threading

class _ResultCollector:
    def __init__(self):
        self._lock = threading.Lock()
        self._datos = {}
        self._textos = {}

    def add_datos(self, stem: str, data: dict):
        with self._lock:
            self._datos[stem] = data

    def add_texto(self, stem: str, texto: str):
        with self._lock:
            self._textos[stem] = texto

    def get(self) -> dict:
        with self._lock:
            return {"datos": dict(self._datos), "textos": dict(self._textos)}
```

## Thread Safety Considerations

- **Shared dict**: `_ResultCollector` with `threading.Lock` — no concurrent writes
- **Retry queue**: `queue.Queue` is inherently thread-safe; retry threads pop safely
- **log_callback**: Each log call is atomic (single string append to `log_queue` which is `queue.Queue`)
- **PDF read**: Each thread reads its own PDF path — no shared file handles
- **No tkinter calls from threads**: All UI updates go through `log_queue` (already the pattern)

## UI Widget Layout

Insert after line 10128 (`_ent_vision_custom_models` pack), before Temperature:

```
[✓] Habilitar API Vision en Paralelo    (CTkCheckBox → self._chk_parallel_enabled)
    └── When checked, show model panel:
    ┌──────────────────────────────────┐
    │ Modelos para procesamiento paralelo │  (CTkLabel, muted)
    │ [✓] google/gemini-2.0-flash-exp  │  (CTkCheckBox, one per model)
    │ [✓] meta-llama/llama-3.2-11b     │
    │ [✓] google/gemini-2.5-flash      │
    │ [ ] qwen/qwen3.6-flash           │
    └──────────────────────────────────┘
```

The checkbox panel is hidden when the master toggle is off. On toggle, call `_toggle_parallel_panel()`. Checkboxes are dynamically rebuilt from the textbox content when models change (`_sync_modelos_desde_textbox` extension).

## Config Save Logic

In the config save block (~line 10311-10328), add after the `api_cfg` loop:

```python
# Parallel settings
w = _g('_chk_parallel_enabled')
if w is not None:
    api_cfg["parallel_enabled"] = w.get() == 1  # CTkCheckBox returns 0/1

model_states = {}
for attr_name, widget in getattr(self, '_parallel_model_checks', {}).items():
    model_states[attr_name] = widget.get() == 1
if model_states:
    api_cfg["parallel_model_states"] = model_states
```

## Error Handling Strategy

| Scenario | Behavior |
|----------|----------|
| Model timeout | Log failure, add PDF to retry queue, try other models |
| HTTP 429 (rate limit) | Log warning, exclude model for current PDF, retry others |
| HTTP 4xx/5xx | Log error, exclude model for current PDF, retry others |
| All models fail | Fall back to PaddleOCR for that PDF |
| No enabled models | Fall back to PaddleOCR for all PDFs |
| API key missing | Log warning, skip parallel entirely, use PaddleOCR |
| Network error | Same as timeout — retry queue |
| JSON parse error | Log warning, exclude model, retry others |

## Logging Format

All logs use prefix `[API Vision Paralelo]` and go through `log_callback`:

```
[API Vision Paralelo] Distribuyendo 6 PDFs entre 3 modelos
[API Vision Paralelo] google/gemini-2.0-flash-exp: 2 PDFs asignados
[API Vision Paralelo] meta-llama/llama-3.2-11b-vision: 2 PDFs asignados
[API Vision Paralelo] google/gemini-2.5-flash: 2 PDFs asignados
[API Vision Paralelo] google/gemini-2.0-flash-exp FALLÓ para ticket_001: timeout
[API Vision Paralelo] Reintentando ticket_001 con meta-llama/llama-3.2-11b-vision...
[API Vision Paralelo] meta-llama/llama-3.2-11b-vision procesó ticket_001 OK
[API Vision Paralelo] Completado: 5/6 API, 1/6 PaddleOCR
```

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit | `_ResultCollector` thread safety | Spawn 10 threads writing same collector, verify all keys present |
| Unit | Round-robin distribution | Mock 6 PDFs, 3 models → verify assignment |
| Unit | Retry queue exclusion | Mock model failure, verify retry skips failed model |
| Unit | All-fail → PaddleOCR path | Mock all models failing, verify `textos` populated |
| Integration | Full parallel flow | Mock `api_vision_extraer_datos`, run `api_vision_con_fallback`, verify `datos` + `textos` |
| E2E | UI toggle persistence | Enable parallel, save config, reload app, verify state |

## Migration / Rollout

No migration required. New config fields default to `false` / all models enabled when absent. Existing users see no change until they enable the toggle. Config JSON is backward compatible.

## Open Questions

- [ ] Should per-model checkboxes auto-sync when the user edits the models textbox (add/remove models), or keep manual checkbox management? **Recommendation**: auto-sync (add new as enabled, remove deleted).
