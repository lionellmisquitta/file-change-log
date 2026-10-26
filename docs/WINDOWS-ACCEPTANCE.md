# Windows acceptance check — required before operational use

This build's Linux tests cannot establish that a target Windows/Excel combination
raises or exposes all expected COM events. Run these checks on representative copies.
Record Excel version, Windows version, test time, and resulting event IDs.

1. Extract the Windows x64 bundle. Start Recorder.bat must open the local page.
   Verify no pip installation or admin request is required by the application.
2. Add a local test folder. Save a workbook before opening it in Excel. Wait for
   `file_version_captured` with a snapshot path.
3. Open `.xlsx`, `.xlsm`, `.xlsb`, `.xls` and `.csv` test files in desktop Excel.
   Check `excel_workbook_attached` for each and the Recording status panel.
4. Type `100` in a cell, then `10`. Check `excel_range_changed` and values. A blank
   cell outside the initial UsedRange may have unknown old value on its first edit.
5. Enter and edit a formula. Paste a 2x2 range. Clear a range. Verify new values
   and formulas, and old values where `old_value_known` is true.
6. Undo, redo and close without saving. Verify observations are not labelled as
   proven saved business outcomes. Reopen the workbook; baseline must be refreshed.
7. Save, Save As, and cancel a save/close dialog. Check that only explicit save
   results claim success. Confirm source files have not been modified by recorder.
8. Create/delete/rename/activate a sheet. Trigger a PivotTable update and formula
   recalculation. Check the documented metadata-level coverage.
9. Sort and insert/delete rows/cells. Validate limitations of coordinate comparisons
   and old-value caching; do not infer the precise command from a range-change event.
10. Open Excel using a second process (for example `excel /x`). Every attached
    instance must appear or a visible attachment-gap warning must appear. If a
    process is unattachable, activate/save its test workbook and check discovery.
11. Test the largest expected workbook, bulk paste and a file exceeding the
    configured limit. Limits must produce explicit skipped/incomplete records.
12. Add a mapped/shared drive. Disconnect it, wait, reconnect. An unavailable-root
    event must appear without falsely claiming every file was deleted.
13. Add a Drive-for-desktop folder with a locally available test file. Save changes
    and verify snapshots; document any placeholder/download limitations.
14. Pause/resume. Change files while paused. Confirm a gap record and net saved-state
    reconciliation, not invented intermediate events.
15. Enable start at sign-in. Sign out/in. Confirm the process starts quietly; use
    Open Recorder.bat to inspect it. Sleep/resume; recording must recover.
16. Explain an event, export JSONL, and find both the event and explanation. Disable
    sign-in startup, stop the recorder, and verify it does not restart at next sign-in.

## Troubleshooting

- No Excel attachment: confirm desktop Excel, same Windows user and privilege level,
  saved path inside the configured roots, and an active workbook. Do not run Excel
  elevated and the recorder unelevated. The recorder reports unattachable instances.
- No events although attached: workbook code may disable application events. A
  detection record is logged; the recorder never forcibly enables someone else's events.
- No snapshot: inspect root status and `snapshot_retry` / `snapshot_skipped` records.
  Locked files are retried. Make cloud files locally available if required.
- App cannot start: inspect `%LOCALAPPDATA%\FileChangeLog\startup-error.txt` and
  endpoint app-control policy. Do not disable organisational security controls.
- Port 47653 in use by another application: close that application or change the
  single-instance port in the source; the pilot uses this fixed loopback gate.

Windows live integration status remains NOT_VERIFIED until actual results are recorded.
