# Proposal: API Vision Parallel Fallback with Model Selection

## Intent

When an API vision model fails during "Controlar Tickets" or "Control Final", the system immediately falls back to local PaddleOCR — wasting potentially working alternative models. This change distributes PDFs across multiple enabled API models in parallel, retries failed PDFs with other models, and only falls to PaddleOCR when ALL models fail. The user controls which models participate via checkboxes in OCR settings.

## Scope

### In Scope
- New `api_vision_con_fallback()` helper in `procesar_tickets.py` for parallel model orchestration
- UI toggle "Habilitar API Vision en Paralelo" in OCR settings
- Per-model checkboxes to enable/disable participation in parallel processing
- Config persistence for `parallel_enabled` flag and per-model enabled states
- Logging of model attempts, failures, and retries
- Fallback to PaddleOCR only when ALL enabled models fail for a PDF
- Backward compatible: parallel disabled = current behavior unchanged

### Out of Scope
- Changing the OpenRouter API integration or authentication
- Adding new API providers beyond OpenRouter
- Modifying PaddleOCR/Tesseract engine internals
- Rate limit management beyond respecting timeout settings
- Cost tracking or usage monitoring

## Capabilities

### New Capabilities
- `api-vision-parallel`: Parallel distribution of PDFs across multiple vision API models with per-model enable/disable, retry on failure, and structured logging

### Modified Capabilities
None — this is purely additive functionality layered on top of existing API vision calls.

## Approach

### 1. New function: `api_vision_con_fallback()` in `procesar_tickets.py`

```python
def api_vision_con_fallback(
    rutas_pdfs: list[str],
    api_key: str,
    modelos: list[str],           # ordered: selected first, then others
    temperature: float = 0.1,
    max_tokens: int = 8192,
    timeout: int = 60,
    api_base: str = OPENROUTER_BASE,
    log_callback: callable = None  # for UI logging
) -> dict:
```

- Distributes PDFs across models (round-robin or balanced by queue)
- Launches one thread per model using `concurrent.futures.ThreadPoolExecutor`
- Each thread processes its assigned PDFs sequentially
- If a PDF fails on one model, it goes into a retry queue
- Retry queue attempts failed PDFs with other available models
- Returns: `{"datos": {stem: data}, "textos": {stem: texto}, "logs": [messages]}`
- `datos` = successful API extractions, `textos` = PaddleOCR fallback texts

### 2. UI Controls in `_ajustes_tab_ocr()`

Add after the custom_models textbox:
- Master toggle checkbox: `[✓] Habilitar API Vision en Paralelo`
- When enabled, show checkboxes next to each model in the list
- The selected model in the dropdown is always enabled (can't disable the primary)
- Checkboxes are saved as a dict in config: `parallel_model_states`

### 3. Config additions to `ui_config.json`

```json
{
  "api_vision": {
    "parallel_enabled": false,
    "parallel_model_states": {
      "google/gemini-2.5-flash": true,
      "meta-llama/llama-3.2-11b": true,
      "nvidia/nemotron-nano-2-vl": true,
      "google/gemma-4-26b-a4b-it": false
    }
  }
}
```

### 4. Worker updates

Both `_control_final_worker()` and `_cargar_datos_worker()` check:
```python
if parallel_enabled:
    result = api_vision_con_fallback(pdfs, api_key, enabled_models, ...)
    # process result
else:
    # existing single-model logic (unchanged)
```

### 5. Logging

Log messages follow this pattern:
```
[API Vision Paralelo] Distribuyendo 6 PDFs entre 3 modelos
[API Vision Paralelo] Gemini: 2 PDFs asignados
[API Vision Paralelo] LLaMA: 2 PDFs asignados
[API Vision Paralelo] Nemotron: 2 PDFs asignados
[API Vision Paralelo] Gemini FALLÓ para ticket_001.pdf: timeout
[API Vision Paralelo] Reintentando ticket_001.pdf con LLaMA...
[API Vision Paralelo] LLaMA procesó ticket_001.pdf exitosamente
[API Vision Paralelo] Completado: 5/6 API, 1/6 PaddleOCR
```

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `procesar_tickets.py:1702` | New function | `api_vision_con_fallback()` — parallel orchestration |
| `ui_app.py:10089-10170` | Modified | New UI controls for parallel toggle and model checkboxes |
| `ui_app.py:7704-7714` | Modified | Control Final worker uses new helper when parallel enabled |
| `ui_app.py:9265-9288` | Modified | Controlar Tickets worker uses new helper when parallel enabled |
| `ui_config.json` | Modified | New `parallel_enabled` and `parallel_model_states` fields |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| OpenRouter rate limiting from parallel calls | Medium | Respect per-model timeout; sequential retry on HTTP 429 |
| Slower if most/all models fail | Low | User controls which models participate; can disable slow ones |
| Thread safety in result collection | Low | Use thread-safe dict or locks for shared results |
| Config migration for existing users | None | New fields default to `false` / all-enabled; backward compatible |
| Paid model accidentally triggered | None | Per-model checkboxes default to OFF for all; user explicitly enables |

## Rollback Plan

1. Delete `api_vision_con_fallback()` from `procesar_tickets.py`
2. Revert worker functions to direct `api_vision_extraer_datos()` calls
3. Remove parallel UI controls from `_ajustes_tab_ocr()`
4. Config fields are ignored when absent — no migration needed

## Dependencies

- Python `concurrent.futures` (stdlib) — already available
- OpenRouter API key (existing) — no new credentials needed

## Success Criteria

- [ ] When parallel enabled, PDFs are distributed across enabled models simultaneously
- [ ] Failed PDFs retry with other enabled models before PaddleOCR
- [ ] PaddleOCR fallback only triggers when ALL enabled models fail
- [ ] UI toggle enables/disables parallel mode
- [ ] Per-model checkboxes control which models participate
- [ ] Settings persist across app restarts
- [ ] Log shows model names, failures, and retry attempts
- [ ] When parallel disabled, behavior is identical to current implementation
