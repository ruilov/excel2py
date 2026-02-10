# PROJECT.md

## Project name
Excel2Py AI Translator

## Mission
Translate an Excel workbook into a **single self-contained Python script** for model logic, with a separate automated test suite that includes:
1. inputs
2. calculations
3. tests

The Python output must reproduce the spreadsheet’s computed behavior as closely as practical, with heavy emphasis on automated testing.

---

## Core goals

- Convert workbook logic to pure Python code.
- Use AI to improve code organization and abstraction (function extraction, naming, grouping).
- Keep implementation pragmatic and lightweight (no over-engineering).
- Make every build verifiable through strong tests.
- Keep lineage traceability from generated code back to source workbook cells/ranges.

---

## Non-goals

- Production-grade framework design.
- Maximum runtime performance optimization.
- Strict static typing.
- Complex defensive programming everywhere.
- Complete Excel ecosystem emulation (e.g., VBA runtime).

---

## Key project decisions

1. **Function support strategy**
   - Maintain a library of supported Excel functions.
   - No mandatory predefined function list; implement coverage based on workbook needs.
   - If a workbook requires a function not yet in the library, use AI to generate a Python implementation candidate.
   - Any AI-generated function implementation must pass tests before acceptance.

2. **Expected outputs source**
   - Expected outputs come only from the workbook being translated.
   - Trust workbook Excel-computed values as the baseline source of truth.
   - Tests should validate **every dynamically computed cell** in that workbook.

3. **Numeric tolerance**
   - Use sensible tolerances for floating point comparisons.
   - Exact match for non-floats unless there is a strong reason not to.

4. **Type coercion**
   - Keep generated code simple.
   - Add coercion complexity only when tests prove it is necessary.

5. **Date handling**
   - Use pragmatic date handling.
   - Avoid adding complexity unless tests require it.

6. **Determinism**
   - Generated Python should be deterministic where practical (seed/time/context controls).
   - For volatile functions, tests must support injecting workbook-derived expected values.

7. **Error handling**
   - Keep error machinery simple.
   - Add complexity only when required to pass parity tests.

8. **Dynamic reference limits**
   - No explicit safety limits for dynamic reference resolution.

9. **Circular convergence**
   - Generated code should expose convergence parameters when needed.

10. **Refactor acceptance**
    - If it passes tests, it is acceptable.

11. **Cross-file consistency**
    - No strict consistency requirement across generated scripts.

12. **Token usage**
    - AI/token policies must be configurable (including model selection).

13. **Data sharing**
    - Assume workbook data can be shared with model APIs.

14. **Performance**
    - Performance is not a primary concern right now.

15. **Traceability**
   - Preserve easy mapping from generated code back to workbook origins.

16. **Intermediate artifact storage**
   - For each workbook, write intermediate extraction outputs under `artifacts/<workbook_stem>/raw/`.

17. **Cell export compactness**
   - Store cells in compact JSONL list form instead of verbose dict form.
   - Record schema via metadata to keep decoding explicit.

18. **Cell export filtering**
   - Do not export cells that have neither formula nor value.

19. **Named range filtering**
   - In workbook metadata, keep only named ranges that are referenced by used formulas/values.

20. **Loader defensiveness**
   - Keep loader/export path handling simple and non-defensive where practical; allow file errors to surface naturally.

---

## High-level architecture

Two-pass pipeline:

1. **Deterministic literal translation (truth pass)**
   - Parse workbook structure, formulas, and dependencies.
   - Generate correct literal Python calculations first.
   - Prioritize parity over readability.

2. **AI-assisted refactor (organization pass)**
   - Use AI to group code into coherent functions and improve names/comments.
   - Use formatting/layout + formula grouping as clues for code organization.
   - Semantics must remain unchanged.
   - Final output accepted only if tests pass.

---

## AI usage strategy (token-smart)

### Principle
Use AI for organization and abstraction, not as an unconstrained semantic engine.

### Deterministic first
Before calling AI:
- Canonicalize formulas.
- Cluster repeated/copy-filled formula patterns.
- Build dependency and lineage summaries.
- Extract formatting/grouping metadata.

### Send compact AI payloads
Prefer sending:
- cluster summaries
- representative formulas
- range geometry
- section labels/format hints
- lineage map

Avoid sending entire sheets when summaries are sufficient.

### Configurable AI settings
Expose settings for:
- provider (`openai` / `anthropic`) with support for both from the start
- model names per stage (e.g., planner vs refactor)
- temperature
- max tokens
- retry policy
- cache on/off
- prompt templates

### Caching
Cache AI outputs keyed by deterministic payload hash so unchanged clusters do not consume new tokens.

---

## Feature support matrix (v1 stance)

### Implemented now
- Circular references (iterative calculation with parameters)
- Core dependency extraction and deterministic literal generation
- Runtime helper coverage required by current workbook
- Workbook date handling required by current workbook (`EOMONTH`)

### Planned / not yet implemented
- Dynamic addressing/reference-by-text (`OFFSET`, `INDIRECT`) full support
- Dynamic arrays and spill behavior full support
- Structured references (Excel tables)
- Full named range runtime resolution
- VBA/macros/UDF runtimes
- Power Query/Data Model/OLAP semantics
- Live external link dependency execution (snapshot/frozen values are acceptable)

---

## Single-script output contract

Each workbook outputs one Python file (example: `pricing_model.py`) containing:

- input definitions/defaults
- calculation logic
- helper/runtime utilities used by that workbook
- lineage comments/mapping back to workbook cells/ranges

Tests are separate from main calculation logic and live under `tests/`.
No strict layout standard is required across all generated files, but generated files must remain understandable and testable.

---

## Intermediate extraction artifact contract (current)

For each workbook, write:

- `artifacts/<workbook_stem>/raw/manifest.json`
- `artifacts/<workbook_stem>/raw/workbook_meta.json`
- `artifacts/<workbook_stem>/raw/cells.jsonl`

`cells.jsonl` stores one compact record per exported cell:

- `[sheet_idx, addr, data_type, formula, value]`

Rules:

- Export only cells with a formula or a non-empty value.
- Use `sheet_idx` as 0-based index into `workbook_meta.json -> sheet_names`.
- Publish record field order in metadata/manifest to avoid ambiguity.
- Keep only needed named ranges in `workbook_meta.json`.
- Compute sheet dimensions from actual non-empty formula/value cells (not inflated Excel used-range metadata).

---

## Current implementation status

Implemented:

- Deterministic pipeline: `loader.py -> planner.py -> translator.py`.
- Parser strategy: LALR-first parsing with Earley fallback and parse caching.
- Dependency extraction and calculation ordering artifacts:
  - `derived/formulas.jsonl`
  - `derived/dependencies.jsonl`
  - `derived/calc_order.json`
- Deterministic literal code generation with explicit per-cell functions and cycle iteration parameters.
- Runtime helper coverage for current workbook formulas, including `ABS`, `CEILING`, `CHOOSE`, `EOMONTH`, `IRR`, `SUMIF`, `SUMIFS`.
- Excel-like arithmetic coercion in generated expressions for blanks/numeric ops.
- Automated tests implemented and passing:
  - `tests/test_parity.py`
  - `tests/test_circular.py`
  - `tests/test_determinism.py`

Planned / not yet implemented:

- Shock tests workflow and committed workbook-derived shock snapshots.
- CI pipeline wiring to run tests on push.
- AI refactor stage (`refactor_ai.py`) with acceptance gating.

---

## Testing requirements (mandatory)

Tests are first-class and must run on every modification.

## Required tests

1. **Full-cell parity test**
   - Validate every dynamic/formula cell against workbook-derived expected outputs.

2. **Determinism tests**
   - Fixed seed/time/context should produce repeatable outputs.

3. **Circular convergence tests**
   - Validate convergence behavior and parameterized iteration settings.

Deferred:

- **Shock tests**
  - Keep as planned snapshot-based functionality.
  - Not required for current acceptance gate until re-enabled.

## Test execution policy
- Tests run locally on generation/refactor steps.
- CI-on-push is planned but not yet wired.
- Any failure blocks acceptance.
- Early loader/extractor utilities do not require dedicated unit tests at this stage.

---

## Acceptance gate

A generated/refactored script is accepted if and only if all required tests pass.

If AI-refactored output fails:
- reject refactor
- keep/revert to literal deterministic version
- optionally retry AI with stricter constraints

---

## Coding style rules

- Keep code pragmatic and simple.
- No `jinja2`.
- No `pydantic`.
- No strong typing requirement.
- Avoid over-defensive patterns.
- Sparse documentation:
  - no default docstrings
  - comments only where they add real clarity
  - one-line comments for most non-obvious logic
  - larger comments only for tricky sections

---

## Suggested minimal dependencies

Required:
- `openpyxl`
- `pandas` (recommended for fixtures/reporting)
- `networkx`
- formula parser library (`lark` or `pyparsing`)
- `pytest`
- OpenAI SDK
- Anthropic SDK

Optional:
- `numpy` (if helpful for array operations)
- `pywin32` (or equivalent) for automated Excel recalculation/snapshot generation on Windows
- file watcher for auto-testing during development

---

## Lineage & traceability requirements

Generated code must preserve a mapping from Python logic back to workbook provenance:
- sheet name
- cell/range origin
- cluster/group id (if applicable)

Traceability can be implemented via lightweight comments, sidecar metadata, or internal mapping structures.
Current preference: inline comments in generated Python where practical.

---

## Recommended repo structure

```text
excel2py/
  loader.py
  formula_parser.py
  planner.py
  translator.py
  refactor_ai.py
  emitter.py
  runtime_helpers.py
  lineage.py
  config.py
  cli.py
tests/
  test_parity.py
  test_shocks.py
  test_determinism.py
  test_circular.py
  fixtures/
```

(Structure is guidance, not a strict requirement.)

---

## Agent instructions (Claude/Codex)

When working on this project:

1. Implement deterministic translation first.
2. Preserve correctness before readability.
3. Use AI calls only where they improve abstraction/grouping/refactoring.
4. Minimize token usage via clustering, canonicalization, and caching.
5. Keep generated code lightweight (no unnecessary frameworks or heavy abstractions).
6. Add/update tests with every logic change.
7. Do not accept behavior-changing refactors unless tests pass.
8. Preserve lineage back to workbook cells/ranges.
9. Prefer simple solutions unless tests prove more complexity is needed.
10. Ship incremental, test-passing improvements.

---

## Initial milestone plan

### Milestone 1
- Baseline workbook: `excel_model.xlsx`.
- Parse workbook + formulas + dependencies.
- Build compact intermediate artifact export (`manifest.json`, `workbook_meta.json`, `cells.jsonl`) under `artifacts/<workbook_stem>/raw/`.
- Generate literal single-script Python.
- Implement full-cell parity test against workbook outputs.
- Status: complete.

### Milestone 2
- Add circular iteration support with parameters.
- Add determinism and circular convergence tests.
- Status: deterministic + circular + determinism complete; shock tests deferred; dynamic refs/spill still pending.

### Milestone 3
- Add CI workflow to run active tests on push.
- Add AI pattern grouping + function extraction.
- Add AI refactor pass using formatting/layout cues.
- Add token cache and configurable model/prompt settings.

### Milestone 4
- Improve function support library coverage.
- Add fallback AI generation for missing functions + test gating.
- Strengthen lineage reporting and debugging UX.

---

## Definition of done (per change)

A change is done when:
- code is merged and runnable,
- required tests pass locally (and in CI once CI is configured),
- no regression in active parity/determinism/circular suites,
- lineage mapping remains intact for affected logic.
