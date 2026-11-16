# CHG-51 — Formula-pattern evidence: contract and replay limits

This checkpoint builds **compact saved-state formula evidence** on top of CHG-50's workbook-logic and timeline branch.

## What it does

When a saved XLSX/XLSM version is compared with its previous retained version, groups of **three or more contiguous changed formula cells in the same sheet and column** may be represented by a `formula_patterns` entry. The pattern stores an anchor A1 formula, starting/ending row and exact represented cell count. A1 relative, absolute and mixed references are verified using `openpyxl.formula.translate.Translator` against the anchor.

A formula pattern describes **what was observed in the later saved workbook**, not *how* the user produced it. A fill-down, paste, import, script or other operation can all result in the same state. The two original saved snapshots are retained as the authoritative old/new workbook states.

Nonformula typed values and formula-to-constant overrides are never absorbed into formula patterns. Formula runs separated by an override or other discontinuity are not incorrectly joined.

## Event contract

`snapshot_cell_comparison` retains its existing `total_changes`, `changes_per_sheet`, `changes_per_sheet_column`, `changes`, `unlisted_changes`, `before_snapshot` and `after_snapshot` fields.

Additional fields:
- `formula_patterns`: compact formula runs, each with `anchor_address`, `anchor_formula`, `first_row`, `last_row`, `cell_count` and `encoding=relative_a1_translation`
- `formula_changes_represented_by_patterns`: number of changed cells explicitly encoded in patterns
- `formula_patterns_require_snapshots_for_exact_reconstruction`: true if the event contains formula patterns

**Invariant:** `total_changes == len(changes) + formula_changes_represented_by_patterns + unlisted_changes`.

`details_truncated` means there are unlisted changes, not formula changes represented by patterns. If there are more than 200 pattern groups, remaining formula changes retain ordinary detail/truncation behavior; no false success claim.

The routine's direct-call default remains `compact_formulas=False` for API backward compatibility. Real monitored-folder comparison opts into `compact_formulas=True`.

## Tests

Run `python -m unittest discover -s tests -p test_formula_patterns.py -v` and the full suite `python -m unittest discover -s tests -v`. Verify full before/after snapshots and explicit overrides.

## Not delivered

- Not a guarantee of complete workbook replay or of user intent.
- Not live COM formula event compression.
- Not an AI process reconstruction engine.
- Not a row/key-aware financial business transformation detector.
- Not a substitute for the still-required CHG-50 live Windows Excel pilot.

**Stacked-branch dependency:** PR for CHG-51 targets `CHG-50/excel-workbook-logic-timeline`. Neither is authorized to merge to main without Windows acceptance and relevant independent review.
