# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-script Python tool (`recalc.py`) that audits a company's fixed-asset ledger (Excel) by independently
recalculating depreciation for each asset and flagging rows where the recalculated amount differs from what the
company booked. Built for use in financial-statement audits. All identifiers, comments, and generated Excel output
are in Korean — match that when editing.

## Commands

```powershell
# Install deps
python -m pip install -r requirements.txt

# Run the full test suite (this is the primary correctness gate — always run after touching recalc.py)
python -m pytest test_recalc.py -v

# Run a single test / class
python -m pytest test_recalc.py::TestDoubleDecliningBalance -v
python -m pytest test_recalc.py::TestDoubleDecliningBalance::test_basic_rate_and_floor -v

# Lint (also runs in CI — keep this clean)
python -m ruff check .

# Regenerate the sample ledger (133 rows covering all methods/events)
python generate_sample.py

# Run the tool itself — with no args, reads config.yaml's input_path/output_path
# (defaults to sample_asset_ledger.xlsx / recalc_result.xlsx in this folder)
python recalc.py

# Override the config file, or just the input/output paths, per run
python recalc.py --config config_회사A.yaml
python recalc.py --input other_ledger.xlsx --output other_result.xlsx
```

CI (`.github/workflows/tests.yml`) installs from `requirements.txt`, then runs `ruff check .` and
`pytest test_recalc.py -v` on push/PR. `mypy` and `pytest-cov` are useful for occasional deeper quality passes but
are **not** in `requirements.txt` — install them ad hoc (`pip install mypy pytest-cov`) when you want them.

## Settings live in `config.yaml`, not in `recalc.py`

Company/audit-specific values — `input_path`/`output_path`, `ref_date`, `materiality_threshold`,
`overall_materiality`/`performance_materiality`, `comparison_years`, `yoy_anomaly_threshold_pct`,
`anthropic_model`, and `column_map` — are read from `config.yaml` at the repo root, not hardcoded in `recalc.py`.
`--input`/`--output` CLI args (if passed) win over whatever `config.yaml` says; `--config <path>` swaps which
YAML file is read entirely (defaults to `config.yaml`). Loading is deliberately fault-tolerant: `load_config()`
returns `{}` on a missing file or a YAML parse error (never raises), and `resolve_settings()` fills every key from
`{}` with the same built-in defaults `config.yaml` currently ships with — so tests that `import recalc as R`
without touching `config.yaml` see identical values to what was previously hardcoded. `column_map` specifically is
*merged* over the built-in default dict (`_DEFAULT_COLUMN_MAP`), not replaced, so a config that only overrides one
column still gets sane defaults for the other 19.

**Known trap when wiring a new config-driven global**: Python binds a function's default-argument values *once, at
def time*, not on every call — so a signature like `def f(x, threshold=MATERIALITY_THRESHOLD)` freezes whatever the
global was at module-load time, and a later `--config` override is silently ignored (this was a real bug). The fix
pattern used here: give such a parameter a sentinel default of `None` and resolve the global *inside the body* at
call time. `classify_materiality(diff, threshold=None)` and `detect_yoy_anomalies(df, threshold_pct=None)` both do
`if threshold is None: threshold = <module global>` on entry, so they always see the current value no matter how
they're called (`main()` still also passes the global explicitly, but that's now redundant belt-and-suspenders, not
the sole safeguard). A related variant of the same trap: `fy_ref_date` used to *secretly read* the module global
`FY_YEAR`; it now takes an explicit `current_fy_year` parameter instead (see "Multi-year comparison &
`fy_ref_date`" under Architecture). When you add another `config.yaml`-backed global, prefer the
None-sentinel-resolve-in-body pattern over a live-global default argument, and never let a helper silently read a
config global it could just as easily receive as a parameter.

The `if __name__ == "__main__":` block re-reads whichever config file `--config` points to (even if it's the
default path — the reload is idempotent, so there's no special-casing) and reassigns
`IN_PATH`/`OUT_PATH`/`REF_DATE`/`FY_YEAR`/`MATERIALITY_THRESHOLD`/`OVERALL_MATERIALITY`/
`PERFORMANCE_MATERIALITY`/`ANTHROPIC_MODEL`/`COMPARISON_YEARS`/`YOY_ANOMALY_THRESHOLD_PCT`/`COLUMN_MAP` at module
scope before calling `main()`. Tests never go through this block — `TestResultColumnsAfterRemoval._run_main_with`
(and similar end-to-end tests) monkeypatch `R.IN_PATH`/`R.OUT_PATH` directly and restore them in a `finally` block.
Follow that pattern for any new end-to-end test; use `monkeypatch.setattr(R, "COMPARISON_YEARS", [...])` for
multi-year-sheet tests rather than editing `config.yaml`.

## Architecture

Everything lives in `recalc.py` (~2100 lines). The pipeline: `resolve_columns()` → per-row `parse_asset_row()` →
per-row `validate_asset_inputs()` (bad rows are isolated, never crash the batch) → `recalc_asset()` /
`recalc_accumulated_dep()` / `get_period_formula_meta()` → `main()` assembles the workbook via pandas +
`_inject_recalc_formulas()` + `_write_info_sheet()` + `_format_workbook()`.

### Segment-based depreciation engine

The core abstraction is a **segment**: a dict describing one contiguous stretch of an asset's depreciation life
under one method/basis (`start_idx`, `end_idx`, `method`, `basis`, `salvage`, `life_years`, `rate`,
`origin_start_idx`, `tax_cost`). `build_depreciation_schedule()` produces a list of these by walking capex,
재추정(life/method re-estimate), and disposal events in date order — each event either extends the current segment
or closes it and opens a new one. This is what lets 자본적지출(capex)/재추정/처분/상각중단 compose uniformly across
all 5 methods instead of being handled ad hoc per method.

- 정액법 (straight-line) has its own lightweight fast path (`build_straight_line_segments` +
  `straight_line_current_period_dep`) used when there are no capex/reest/suspension events, because the segment
  engine is unnecessary overhead for the simple case.
- 정률법 (declining balance) and 이중체감법 (double-declining) share `_declining_balance_year_loop` /
  `_book_value_at`'s non-정액법 branch. They differ only in `rate` (tax-table lookup via `get_rate()` for 정률법 vs.
  `2/life_years` for DDB, both behind `_method_rate()`) and floor rule (`tax_cost * 0.05` tax-cliff for 정률법 vs.
  plain `salvage` for DDB, behind `_floor_threshold_for()`). DDB always routes through the full segment engine
  (no fast path) since it has no legacy simple-case behavior to preserve.
- 연수합계법 (sum-of-years-digits) has its own `_syd_year_loop`, structurally parallel to the declining-balance
  loop but with `dep = (basis - salvage) * frac_k`, `frac_k = (life_years - k + 1) / (life_years*(life_years+1)/2)`.
  Unlike DDB, `basis` is fixed per segment (like 정액법) — `dep` depends only on `k` (years since
  `origin_start_idx`), not on accumulated depreciation. Capex does **not** reset `origin_start_idx` (matches
  정액법's "continue over remaining life" behavior); 재추정 **does** reset it to the reest date.
- 생산량비례법 (units-of-production) is handled separately (`units_of_production_current_period_dep`) since it's
  driven by production quantities, not calendar time; `총예정생산량`/`당기실제생산량` are optional columns — if
  absent, no asset is treated as using this method.

`recalc_asset()` routes to the segment engine when
`(method in ("정액법","정률법") and has_extended_events) or method in ("이중체감법","연수합계법")`.

**Universal accumulation identity** (holds for all methods via `_book_value_at`, don't special-case it):
`전기말 감가상각누계액 = (취득원가 + 자본적지출누계) - 전기말장부가액`, and
`당기말 = 전기말 + 당기감가상각비`.

### Excel formula injection (`_inject_recalc_formulas`)

Result cells are written as live Excel formulas (`MONTH()`/`YEAR()`/`VLOOKUP`), not static values, so a reviewer
can click a cell and see how it was derived, or edit an input and watch it recompute. `formula_flags` (one dict
per row: `current_ok`, `accum_ok`, `has_capex`, `has_susp`, `is_units`) gates which cells get formulas vs. stay as
plain computed values — capex, 상각중단 (suspension), and 생산량비례법 rows are excluded from the
경과개월수/당기해당월수/적용내용연수(개월) MONTH()-formula because a single calendar-arithmetic formula can't
safely reproduce what the Python event engine does for those cases (verified via Excel COM
cross-checks — see below). If you change the event engine's date logic, re-verify formula parity, don't assume it
still holds.

### Column-name / method-value matching (3-tier, in `resolve_columns()`)

Company ledgers use inconsistent column names and abbreviated method values, so matching happens in three tiers,
increasing in risk and decreasing in automation:

1. **Exact match** against `COLUMN_MAP` — existing behavior, always wins.
2. **`COLUMN_SYNONYMS`** — a curated dict of common aliases per field, auto-applied. Deliberately conservative:
   e.g. "장부가액" is never listed as a synonym for "취득원가" anywhere, because it may be a genuinely different
   concept (net book value vs. acquisition cost) that only a human should resolve.
3. **AI (`_suggest_columns_ai`, reuses the same lazy-import/API-key-optional/try-except-fallback pattern as
   `get_ai_estimated_cause`) or fuzzy (`_suggest_columns_fuzzy`, `difflib`) suggestion** — **never auto-applied**.
   For missing required columns it's embedded in the raised `KeyError` message and execution still halts; for
   missing optional columns it's a non-fatal `[참고]` hint. Don't change this to auto-apply — a wrong column match
   silently corrupts the entire recalculation, not just one row's comment.

상각방법 values go through `normalize_method()` (`METHOD_ALIASES` dict, whitespace/case-insensitive) at parse time
in `parse_asset_row()`, so every downstream `method == "정액법"`-style comparison keeps working unchanged.
Unrecognized values (after normalization, not in `KNOWN_METHODS`) are caught by `validate_asset_inputs()` and
isolated to the `데이터오류` sheet per-row — they must never reach `recalc_asset()`'s catch-all
`else: raise ValueError`, which exists only as a defensive backstop and would otherwise kill the entire batch run.

### Data-error isolation

Any asset with impossible inputs is caught by `validate_asset_inputs()` and routed to the `데이터오류` sheet with a
specific reason — it never reaches the recalculation functions and never aborts the run for other assets. Two kinds
of checks:

1. **Value sanity**: life < 1, cost ≤ 0, salvage ≥ cost, unrecognized 상각방법 (after `normalize_method`),
   negative `accum_reported`/`prior_period_units`, capex or suspension pairs that are only half-filled, etc.
2. **Event-date ordering against the acquisition date** (`acq`): an error is raised when any event date falls
   *before* `acq` — i.e. `disposal < acq`, `reest_date < acq`, `susp_start < acq`, `susp_end < acq`, or
   `capex_date < acq` (an event dated before the asset was even acquired). These date-order checks only run when
   `acq` is passed — both production call sites (`main()` and `build_multi_year_trend_df()`) pass it, but callers
   that omit `acq` skip them, which keeps the ~30 existing unit-test call sites that don't supply `acq` backward-
   compatible.

When adding a new validation rule, follow this isolate-per-row pattern rather than raising.

### Multi-year comparison & `fy_ref_date`

When `comparison_years` has 2+ years, `build_multi_year_trend_df(df, cols, years, current_fy_year)` recomputes each
asset's depreciation for every listed FY (calling `recalc_asset` per year) and `detect_yoy_anomalies()` flags large
year-over-year swings. `fy_ref_date(year, current_fy_year)` maps each FY to its reference date: it returns the
configured `REF_DATE` for whichever `year` equals `current_fy_year` (so a non-12/31 ref date is respected for the
current year) and `year`-12-31 for every other year. `current_fy_year` is an **explicit parameter, not a read of
the module global `FY_YEAR`** — `recalc_accumulated_dep()` forwards its own `fy_year` argument as
`fy_ref_date(y, fy_year)`, and `build_multi_year_trend_df()` receives `current_fy_year` and passes it through
(`main()` supplies `FY_YEAR` at that single call site). This closes a latent trap where `fy_ref_date` trusting the
global would return the wrong reference date for any caller processing a year other than the global FY (same class
of issue as the "Known trap" above — pass the value, don't secretly read a global).

### AI-assisted cause estimation (optional, separate from column matching)

`get_ai_estimated_cause()` is called only for "유의한 차이" (material difference) assets not already explained by
`get_rule_based_cause()`'s rule-based checks (missed disposal, missed reest, life-ended-but-still-depreciating,
etc.). Gated by `ANTHROPIC_API_KEY` in `.env`; on missing key or API failure it falls back to `"-"` silently — the
rest of the pipeline is unaffected either way. `_suggest_columns_ai` reuses this same optional/fallback pattern.

## Testing conventions

- `test_recalc.py` is organized as one `TestXxx` class per feature/method (`TestStraightLine`,
  `TestDoubleDecliningBalance`, `TestSumOfYearsDigits`, `TestResolveColumns`, etc.) — put new tests in the matching
  existing class or add a new one following that naming.
- Prefer closed-form math checks over hardcoded magic numbers where the method has one (e.g. SYD's fraction series
  sums to exactly `cost - salvage` over `life_years` years — assert against that identity, not a memorized number).
- Tests that exercise `main()` end-to-end must monkeypatch `R.IN_PATH`/`R.OUT_PATH` (see "Settings live in
  `config.yaml`" above) and restore them in `finally`.
- Tests involving the AI path (`_suggest_columns_ai`, `get_ai_estimated_cause`) should monkeypatch the function
  itself or unset `ANTHROPIC_API_KEY` via `monkeypatch.delenv` — don't rely on a real network call.
- When verifying new Excel-formula logic (`_inject_recalc_formulas`), the reliable method used throughout this
  project is PowerShell + Excel COM automation (`New-Object -ComObject Excel.Application`,
  `.CalculateFullRebuild()`, read `.Value2` back) to cross-check formula-evaluated values against
  independently-computed Python values across the full sample file — a naive "the formula looks right" read is not
  sufficient, since Excel's actual evaluation has caught real bugs (e.g. 생산량비례법's dummy `life=1` placeholder,
  suspension's extended effective end date) that weren't visible from reading the formula string alone.
- A `pytest --cov=recalc --cov-report=term-missing` pass has previously turned up real, non-trivial 0%-coverage
  gaps in production paths — not just AI-call bodies or the `__main__` block (both expectedly untested). The
  biggest one found so far: 정률법 재추정 *without* a method switch (`method="정률법"`, `reest_method=None`) routes
  through the legacy `declining_balance_current_period_dep()` fast path, which had zero test coverage even though
  the equivalent 정액법 reest case was tested. When adding a new event combination (reest/capex/suspension ×
  method), check whether it's actually exercised by an existing test before assuming it is — "정액법 has a test for
  X" does not imply "정률법/이중체감법/연수합계법 has one too."

## Sample data

`sample_asset_ledger.xlsx` (133 rows, generated by `generate_sample.py`) is synthetic — no real company or audit
data. It's built to cover all 5 methods crossed with disposal/재추정(incl. method-switch)/자본적지출/상각중단, plus
intentional company-vs-recalculated mismatches and 데이터오류 cases. When adding a new calculation path, add a
corresponding row (or set) to `generate_sample.py` and confirm `python -m pytest` plus a full regeneration still
produces the expected 133건/122정상/11오류/8불일치 split described in `README.md`.

`.gitignore` blocks anything matching `*real*`/`*회사*`/`*company*` patterns — this repo must never contain real
company data.
