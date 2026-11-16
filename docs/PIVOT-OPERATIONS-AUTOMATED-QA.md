# Automated pivot-operation recording verification

This testing setup does not require a person to open Excel or manipulate pivot tables.

## What is exercised
A real headless **LibreOffice Calc** spreadsheet engine opens a deterministic 48-row XLSX, creates and updates one native pivot in its own API, and saves actual XLSX files at each step:
1. Pivot created with Region rows and Sum(Revenue)
2. Aggregation switched to Count(Revenue)
3. Product added to columns
4. Month added to page filters
5. Pivot deleted

The Python integration test writes these real engine-generated files to the **same monitored path**, invokes the actual FolderWorker scan/capture pipeline, examines the emitted event log, confirms the pivot definition state transitions, checks displayed computed grand totals (Sum and Count), verifies retained SHA-256-addressed snapshots, and validates every evidence ZIP member checksum.

The expected state labels are: `first_observed`, `definition_changed` (3 times), and `no_longer_present`. The test must **not** report a user pivot creation/deletion command as observed, since it tests saved-state evidence.

## Run without human Excel steps
GitHub Actions: `.github/workflows/real-pivot-operations.yml`.

Locally on Linux with LibreOffice Calc + UNO:
```bash
sudo apt-get install libreoffice-calc python3-uno
python -m pip install -r requirements.txt
RUN_REAL_PIVOT_OPERATIONS=1 python -m unittest discover -s tests -p test_pivot_operations_end_to_end.py -v
```

All generated files live in a disposable temporary directory. Test source is synthetic deterministic data; no confidential customer information and no network download.

## Limits
- A **real spreadsheet engine** executes and saves the pivot operations, but it is *LibreOffice Calc*, not Microsoft Excel.
- The test verifies the recorder's **saved-file observation and parsing path**, not live unsaved Excel COM `OnSheetPivotTableUpdate` events.
- Excel's actual COM event ordering, PivotTable object metadata and filter/slicer UI behavior are **BLOCKED** until a Windows CI runner with installed, licensed desktop Microsoft Excel is available.
- The existing Windows GitHub-hosted runner generally does not provide licensed desktop Excel; don't claim COM acceptance from mock callbacks.
- Pivot recreation in Excel is **not** validated by this integration.
- The test currently uses 48 rows, enabling precise grand-total assertions; separately published public dataset stress tests remain CHG-53.

## Status / interpretation
Passing the automated real-engine suite means pivot saved-state recording works across operations. It does *not* establish all interactive Excel PivotTable commands are captured. Failure is a real QA signal; do not relabel failed execution as skipped or assume an existing pivot proves the operation was recorded.
