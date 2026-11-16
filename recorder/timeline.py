"""Evidence-linked activity timeline.

Timeline entries are observations, never generated descriptions of user intent.
Source JSONL/SQLite retains the authoritative event payloads and exact order.
"""
TIMELINE_EVENTS = {
    "excel_workbook_open", "excel_range_changed", "excel_sheet_created",
    "excel_sheet_delete_requested", "excel_sheet_name_observed_changed",
    "excel_pivot_updated", "excel_pivot_definition_observed",
    "excel_pivot_definition_unavailable", "excel_save_requested",
    "excel_save_result", "excel_close_requested",
    "file_version_captured", "file_moved", "file_absent_on_scan",
    "snapshot_cell_comparison", "workbook_logic_observed",
    "workbook_logic_pivot_state_changes", "no_cell_differences_other_content_changed",
    "user_explanation", "folder_unavailable", "folder_reconnected",
    "snapshot_skipped_budget", "snapshot_skipped_size",
    "snapshot_skipped_disk_space", "snapshot_deferred_spacing",
}
GAP_EVENTS = {"folder_unavailable", "folder_reconnected",
              "snapshot_skipped_budget", "snapshot_skipped_size",
              "snapshot_skipped_disk_space", "snapshot_deferred_spacing",
              "excel_pivot_definition_unavailable"}


def timeline(store, limit=200, path_filter=""):
    """Return a bounded list with pointers to raw events, never synthesized steps."""
    if not isinstance(limit, int) or limit < 1:
        raise ValueError("Timeline limit must be positive")
    maximum = min(limit, 300)
    # Store.recent is already bounded to 1000; this is a recent-window view,
    # never a complete session export or a claim of exhaustive coverage.
    rows = store.recent(limit=1000)
    items = []
    for row in reversed(rows):
        kind = row.get("event")
        path = row.get("path", "")
        if kind not in TIMELINE_EVENTS:
            continue
        if path_filter and path_filter.lower() not in path.lower():
            continue
        summary = {"event": kind, "timestamp_utc": row.get("timestamp_utc"),
                   "event_id": row.get("id"), "sequence": row.get("sequence"),
                   "path": path, "kind": "gap" if kind in GAP_EVENTS else "observation",
                   "confidence": "observed_event_not_proof_of_user_intent"}
        if row.get("snapshot"):
            summary["snapshot"] = row["snapshot"]
        if row.get("after_snapshot"):
            summary["snapshot"] = row["after_snapshot"]
        if kind == "workbook_logic_observed":
            state = row.get("extraction", {})
            summary["details"] = {"pivots": len(state.get("pivots", [])),
                                  "worksheet_filters": len(state.get("worksheet_filters", [])),
                                  "table_filters": len(state.get("table_filters", [])),
                                  "completeness": state.get("completeness")}
        elif kind == "workbook_logic_pivot_state_changes":
            summary["details"] = {"pivot_state_changes": len(row.get("pivot_changes", [])),
                                  "omitted": row.get("pivot_changes_omitted", 0)}
        elif kind == "excel_pivot_definition_observed":
            summary["details"] = {"sheet": row.get("sheet"),
                                  "pivot_name": row.get("pivot_name"),
                                  "completeness": row.get("completeness")}
        elif kind == "snapshot_cell_comparison":
            summary["details"] = {"total_changes": row.get("total_changes"),
                                  "unlisted_changes": row.get("unlisted_changes", 0)}
        elif kind == "user_explanation":
            summary["details"] = {"linked_event_id": row.get("linked_event_id"),
                                  "context_confirmed": row.get("context_confirmed", False)}
        items.append(summary)
    total = len(items)
    return {"entries": items[-maximum:],
            "recent_window_matched_events": total,
            "older_matching_events_not_shown": max(0, total-maximum),
            "source_events_read": len(rows),
            "complete_history": False,
            "note": "Recent observed events only. Export evidence ZIP for complete raw history; no user actions are inferred."}
