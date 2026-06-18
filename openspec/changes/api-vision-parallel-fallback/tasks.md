# Tasks: API Vision Parallel Fallback with Model Selection

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 450-550 |
| 400-line budget risk | Medium |
| Chained PRs recommended | Yes |
| Suggested split | PR 1 (Foundation + UI) → PR 2 (Core Implementation + Workers) → PR 3 (Testing + Config) |
| Delivery strategy | ask-on-risk |
| Chain strategy | stacked-to-main |

Decision needed before apply: Yes
Chained PRs recommended: Yes
Chain strategy: stacked-to-main
400-line budget risk: Medium

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Foundation: New helper classes and utilities | PR 1 | Base branch; tests/docs included |
| 2 | Core implementation: api_vision_con_fallback() function | PR 1 | Depends on PR 1 foundation |
| 3 | UI controls: Parallel toggle + per-model checkboxes | PR 1 | Immediate parent/base branch boundary; depends on PR 1 |
| 4 | Config persistence: Save/load parallel settings | PR 1 | Immediate parent/base branch boundary; depends on PR 1 |
| 5 | Worker integration: _control_final_worker updates | PR 2 | Depends on PR 1 |
| 6 | Worker integration: _cargar_datos_worker updates | PR 2 | Depends on PR 1 |
| 7 | Testing: Unit tests for new functionality | PR 3 | Depends on PR 1-2 |
| 8 | Testing: Integration tests | PR 3 | Depends on PR 1-2 |

## Phase 1: Foundation / Infrastructure

- [x] 1.1 Create `_ResultCollector` class in `procesar_tickets.py` for thread-safe result collection
- [x] 1.2 Create `api_vision_con_fallback()` function signature in `procesar_tickets.py` (lines 1830-1850)
- [x] 1.3 Add imports for `concurrent.futures`, `threading`, and `queue` in `procesar_tickets.py`

## Phase 2: Core Implementation

- [x] 2.1 Implement round-robin PDF distribution logic in `api_vision_con_fallback()`
- [x] 2.2 Implement ThreadPoolExecutor thread management and PDF processing
- [x] 2.3 Implement retry queue logic with model exclusion per PDF
- [x] 2.4 Implement PaddleOCR fallback when all models fail
- [x] 2.5 Implement logging with `[API Vision Paralelo]` prefix
- [x] 2.6 Return structured result: `{"datos": {}, "textos": {}, "logs": []}`

## Phase 3: Integration / Wiring

- [x] 3.1 Add parallel toggle checkbox in `_ajustes_tab_ocr()` after line 10128
- [x] 3.2 Add per-model checkbox panel when parallel enabled
- [x] 3.3 Implement `_toggle_parallel_panel()` function
- [x] 3.4 Add config save logic for `parallel_enabled` and `parallel_model_states`
- [x] 3.5 Modify `_control_final_worker` to branch on `parallel_enabled`
- [x] 3.6 Modify `_cargar_datos_worker` to branch on `parallel_enabled`

## Phase 4: Testing

- [ ] 4.1 Write unit tests for `_ResultCollector` thread safety
- [ ] 4.2 Write unit tests for round-robin PDF distribution
- [ ] 4.3 Write unit tests for retry queue exclusion logic
- [ ] 4.4 Write unit tests for all-fail → PaddleOCR path
- [ ] 4.5 Write integration tests for full parallel flow
- [ ] 4.6 Write E2E tests for UI toggle persistence

## Phase 5: Cleanup

- [x] 5.1 Add comments and docstrings for new functions
- [x] 5.2 Remove any temporary debug code
- [x] 5.3 Update any related documentation

## Implementation Order

The implementation follows a logical dependency order:
1. First, create the foundation (thread-safe result collector and function signature)
2. Then implement the core parallel processing logic
3. Next, wire in the UI controls and configuration
4. Finally, integrate with the existing workers and add comprehensive tests

This ensures each phase builds on the previous one, with clear dependencies and verifiable milestones at each step.

## Review Workload Forecast

- Estimated changed lines: 475
- 400-line budget risk: Medium
- Chained PRs recommended: Yes
- Delivery strategy: ask-on-risk
- Decision needed before apply: Yes
- Suggested work-unit PR split: PR 1 (Foundation + Core + UI + Config) → PR 2 (Workers) → PR 3 (Testing)