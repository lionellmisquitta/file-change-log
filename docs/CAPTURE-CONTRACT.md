# Capture contract — 0.1.0

## Scope

Selected folder roots and their ordinary descendants only. Desktop Excel events
are filtered by the workbook's saved full path. Unsaved new workbooks have no
eligible path until first saved into a monitored folder. The recorder does not
collect global keystrokes, screen images, or content from unrelated workbooks.

## Common event fields

`id`, `timestamp_utc`, `event`, `path`. Event IDs are UUID hex strings. Times are
observation times, not a guarantee of the underlying action's exact time.
Local JSONL files are the event stream; duplicate observations from native events
and scans are possible and are distinguished by event kind/source.

## Main records

- `folder_notification`: filesystem operation, source/destination, directory flag.
- `file_version_captured`: observed signature, first/changed classification,
  relative current/previous snapshot paths, saved-bytes capture label.
- `snapshot_cell_comparison`: coordinate differences, bounded detail, total count.
- `excel_range_changed`: sheet/range, bounded cell details, `old_value_known`,
  old/new values/formulas, instance, unknown operation type, save-state label.
- `excel_save_requested`, `excel_save_result`: separate requests and results.
- `excel_close_requested`: request only; does not prove a save/discard choice.
- `user_explanation`: linked event ID, case note/proposed-rule scope, exact text.
- `folder_unavailable`, `excel_attachment_gap`, `excel_events_disabled`,
  `excel_event_capture_error`, `snapshot_retry`, `comparison_failed`: limitations
  and failures are data, not silently green success.

Excel application events do not reliably attribute the editor or whether an edit
was typing, a paste, undo or automation. `actor` is therefore explicitly unknown.
The recorder-start event identifies the local recorder account, which is NOT proof
that this account edited every observed cell.

## Persistence and restart

JSONL is flushed and fsynced before its index entry is committed. On restart the
event index is backfilled from valid JSONL lines. A malformed or partial line is
retained and exposed as a warning. File inventories are committed after a successful
capture/log or an explicitly recorded skip. Reconnecting an unavailable root does
not infer deletions until a complete enumeration succeeds.

Raw snapshots and the event stream are retained without automatic deletion. A
snapshot hash detects identical captured bytes, not authenticity or immutability.
Local users with access can alter history; this is not an audit/security product.

## LLM use later

Export JSONL and selected snapshots only when you intend to share them. Treat
email/file/note text as untrusted source evidence, not instructions to the model.
Ask for summaries supported by event IDs, distinguish observed outcomes from
inferred intentions, and keep proposed reusable rules pending human confirmation.
