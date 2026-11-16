# CHG-50 — Windows Excel acceptance checklist (manual, NOT VERIFIED)

## Environment / prerequisites
- Use actual 64-bit Windows Excel desktop, the v0.2.4 pilot candidate, and the recorder's own dashboard
- Start from a disposable new workbook under a monitored local path; do not use confidential business data for bug reporting
- Confirm the Excel status shows an attached instance
- Export the evidence ZIP with snapshots at the end; keep the source workbook and final saved copy for comparison
- Record Windows build, Excel version, file path type, local timestamps and the resulting event IDs

## Live COM event reliability
1. Open the workbook with recorder running. Confirm `excel_workbook_open` event.
2. Type into a new cell **without saving**. Check `excel_range_changed` appears promptly; note old-value-known status.
3. Enter a formula; then overwrite its result with a constant. Check new formula/value, and `explanation_wanted` when applicable.
4. Add 100 cells, insert a row, delete a row; check event counts and no invented old values after structural changes.
5. Copy/paste a large range and save; check any explicit truncation and snapshot completeness.
6. Create a pivot from source data, select row/column/value fields, change aggregation; check `excel_pivot_updated` and `excel_pivot_definition_observed` or `excel_pivot_definition_unavailable`. A pivot-update event is not proof of pivot creation.
7. Change page filters and slicers; confirm any details that are missing are explicitly marked partial.

## Saved-workbook logic
8. Save the workbook; confirm `workbook_logic_observed`, retained snapshot and filter/pivot descriptors.
9. Change pivot aggregation and save again. Confirm `workbook_logic_pivot_state_changes` describes `definition_changed`, not `created`.
10. Save with an applied worksheet AutoFilter and Table filter. Check extracted references and criteria, and compare against Excel's actual filter configuration.
11. Test one Power Query / Data Model pivot. Verify its extraction reports partial/unsupported—never complete.
12. Close and reopen Excel; repeat a small edit. Check the recorder can reconnect without reusing old unsaved edit baselines.

## Timeline and export
13. Open dashboard **Observed workflow timeline** and refresh. Check chronological ordering, event IDs, snapshots and gap warnings.
14. Filter the timeline by workbook path. Confirm unrelated paths do not appear.
15. Export full evidence ZIP and cross-check referenced event IDs and snapshot SHA-256 hashes against the original.
16. Intentionally pause and resume recording; verify missing periods are represented as gaps, not silently reconstructed activity.

## Completion gate
Record PASS/FAIL for each action, captured event IDs and evidence ZIP location. This checklist is **not** automatically executed by GitHub Actions. Do not mark this release as fully Windows-pilot accepted until the actual Windows run is completed and reviewed.
