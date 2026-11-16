# File Change Log — reproducible public-dataset stress testing

**Checkpoint:** CHG-48 FP&A evidence completeness  
**Verified candidate (at the time of the run):** `b96619d756b717fffff7f446fe30bc10f1427bad`  
**Successful workflow:** https://github.com/lionellmisquitta/file-change-log/actions/runs/43849911614  
**Execution date:** 16 November 2026 (GitHub Actions UTC logs)  
**Evidence source:** the full test stdout in the workflow's Ubuntu job. Windows core tests and Windows packaging passed on the same head.  
**Purpose:** preserve reproducible testing of the Phase 1 *capture* contract, not to claim that a complete Excel user workflow can already be replayed.

## Downloaded input datasets

All datasets come from the public `mwaskom/seaborn-data` GitHub repository, downloaded live by `tests/test_public_dataset_stress.py`. No DataCamp login, account, scraping, or licensed course data is needed. Data sources may change upstream: verify hashes shown by CI before comparing a future run to this recorded one.

| Input | URL | Rows | Raw bytes | SHA-256 on 16 Nov 2026 |
|---|---|---:|---:|---|
| Diamonds | https://raw.githubusercontent.com/mwaskom/seaborn-data/master/diamonds.csv | 53,940 | 2,772,143 | `9574730b03aba241d899c4a97511c5061b19358fab89510774fb6c24168345c4` |
| Taxis | https://raw.githubusercontent.com/mwaskom/seaborn-data/master/taxis.csv | 6,433 | 869,349 | `08d6d71784dbaa2651fee37fc03389754194c05d72d2d19cbc2c799dea6ac09d` |
| Titanic | https://raw.githubusercontent.com/mwaskom/seaborn-data/master/titanic.csv | 891 | 57,018 | `81787d320d7f7b03df935e91de8bd19e11d45c5bbcab86ef4d4a76dc91b7d4f2` |
| Flights | https://raw.githubusercontent.com/mwaskom/seaborn-data/master/flights.csv | 144 | 2,350 | `237d834127d9c6355630d8f443a7a2377b5925923010009b59809ba0b67f4fac` |

## Reproducing the exact dataset preparation ("data massage")

Run `python -m unittest discover -s tests -p test_public_dataset_stress.py -v` from the repository root with `openpyxl`, `watchdog` and the project's normal Python requirements installed. CI also uses `python -m unittest discover -s tests -v` to include the existing suite. **Network access is required only for the downloads**, and the test temporarily stores all data in its own temporary folders; it does not modify user spreadsheets. Tests attempt downloads up to three times and fail (rather than silently substitute a fixture) if the source is unavailable or too small.

The test harness reads each downloaded CSV with `csv.DictReader`, records the raw-byte SHA-256, and verifies minimum row counts. For baseline and file-copy tests, it writes raw CSV bytes unchanged to:

```text
<temporary>/
  source/FP&A/Revenue/2026/November/
    diamonds.csv
    flights.csv
    taxis.csv
    titanic.csv
    copies/titanic.csv       # exact byte-for-byte duplicate
  event-history/             # separate recorder's own local state
```

**Scenario QA-01 — recursive baseline, copy deduplication and restart**

1. Create the five-deep `source/FP&A/Revenue/2026/November` directory.
2. Place the four CSVs there without changing their contents.
3. Run `FolderWorker.scan_once()` twice to allow the configured `settle_seconds=0` in test conditions.
4. Assert four `file_version_captured` events, each classified `baseline`, each with a snapshot.
5. Copy `titanic.csv` byte-for-byte into `copies/titanic.csv`; rescan and assert no increase in snapshot storage.
6. Construct a fresh `FolderWorker` against the same Store, with `reconcile_on_first_scan=True`; rescan and verify no extra saved version for unchanged files.

**Scenario QA-02 — transaction CSV revision**

1. Save the downloaded `taxis.csv` exactly as received (6,433 rows, 869,349 bytes), then scan.
2. Append a single newline byte to make the revised file 869,350 bytes; rescan.
3. Assert two version events and a `snapshot_text_comparison` (or an explicitly signalled `comparison_skipped` if a future configured size limit applies). The recorded run produced `snapshot_text_comparison`.

**Scenario QA-03 — real data to 6,500-row Excel calculation**

1. Take the **first 6,500 data rows** from the original `diamonds.csv`. This selection is deterministic in the dataset's download order; don't shuffle or sample randomly.
2. Create `diamonds_before.xlsx` with one `Diamonds` sheet, headers `cut, carat, price, simulated_margin_formula, owner_override`, and 6,500 rows: cut as string, carat converted to float, price converted to integer; both last columns empty.
3. Save the baseline workbook using `openpyxl`.
4. Fill column D on rows 2 through 6,501 with formula `=C<row>*0.2` (for example, `=C2*0.2`).
5. Write `MANUAL REVIEW` to `E3`; save as `diamonds_after.xlsx`.
6. Run the project's actual `compare_files(before, after, ..., max_cells=200000, max_changes=5000)`, collect the resulting event, and assert 6,501 total changes; 5,000 shown, 1,501 omitted; 6,500 changes in column D; one change in E; `E3` listed first.
7. Note: `openpyxl` writes formula strings but does not calculate them; this scenario checks captured formula text and completeness metadata, **not** the resulting numeric margin calculation or Excel COM callbacks.

**Scenario QA-04 — filter-only saved-file change and export integrity**

1. Take all **144 rows** from `flights.csv`. Create `seasonal_forecast.xlsx` with sheet `Monthly`, columns `year, month, passengers, cost_ratio`: convert year and passengers to integers, keep month as text, and leave cost_ratio empty.
2. Save and scan the baseline.
3. Add an Excel table filter definition via `ws.auto_filter.ref = "A1:D145"`, save and rescan. Do not change any cell value or formula.
4. Assert `no_cell_differences_other_content_changed` was emitted. This proves detection of *some* non-cell workbook byte changes, **not** reconstruction of filter instructions or hidden rows.
5. Call `build_handover(store, include_snapshots=True)`. Verify each `sha256_by_member` checksum against actual ZIP-member bytes, confirm two retained snapshots and the existence of `gaps.jsonl`.

## Actual results captured in GitHub Actions

| Scenario | Verified evidence |
|---|---|
| QA-01 | 4 initial datasets, 5 total version events after one duplicate, `snapshot_bytes=3700860`, no extra snapshot bytes from duplicate; elapsed 0.015 s in the CI job |
| QA-02 | Original 869,349 bytes, revised 869,350 bytes; `snapshot_text_comparison` emitted |
| QA-03 | 6,500 worksheet rows; **6,501 total changes**, 5,000 listed, **1,501 explicitly omitted**; first listed cell `E3`; 444,482-byte serialized comparison event; comparison elapsed 0.342 s |
| QA-04 | 144 input rows; 2 snapshots in ZIP; 5 SHA-256 member hashes checked; `no_cell_differences_other_content_changed` emitted; ZIP 17,832 bytes |

**Full suite:** 66 tests executed; 65 passed, 1 skipped on Ubuntu, same workflow passed core Windows and Windows packaging jobs. The single skipped test is platform-dependent. The measured timings are CI-run observations, **not service-level performance guarantees**.

## Known limits and next adversarial scenarios

- Tests do **not** exercise real Windows Excel COM, OneDrive, scheduled tasks, native mutex, floating notes, file-sharing semantics, or ordering of unsaved user edits.
- No test here proves reconstruction of pivot definitions, sort/filter actions, paste-as-values, user intent, or the numerical equivalence of recalculated Excel formulas.
- The comparison event remains approximately **444 KB** for a 5,000-item display; compression/chunking is a separate potential optimization.
- Datasets are accessible via mutable `master` URLs; CI logs hashes so future investigators can identify upstream content changes.
- A future real-Windows pilot should cover: unsaved formulas and overrides, pivot creation/configuration, filter-only work, insert/delete/shift/sort, repeated saves under file-sharing locks, large workbook sessions, and exact blind replay from an independently preserved starting point.
- Do not merge synthetic QA checks with production telemetry or describe these derived test workbooks as actual users' work.

## Historical audit trail

- The first QA-only branch `CHG-49/public-dataset-qa` ran the four downloaded inputs against `main`. It passed baseline/copy and CSV scenarios but its formula-fill and filter expectations could not pass because the related evidence fields were only on the unmerged CHG-48 branch.
- That first mismatch was test-target setup, **not** evidence that the CHG-48 changes failed. Its PR #10 was closed without merge. The *same public-data test source* was then committed on CHG-48, and the full suite passed in workflow `43849911614`.
- Reference: https://github.com/lionellmisquitta/file-change-log/pull/9
- Verification: https://github.com/lionellmisquitta/file-change-log/actions/runs/43849911614

## How to extend without invalidating past proof

Keep this test file and report under version control. Add a fresh dataset source/URL/hash and a **new** scenario test for each additional workflow. Do not overwrite these recorded hashes or results with data downloaded on a later date. When working on a new build, record the new candidate SHA, test run ID, dataset hashes, failed cases and actual Windows pilot findings alongside this report.
