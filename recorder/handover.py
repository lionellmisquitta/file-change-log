"""Portable, user-requested local evidence ZIP; no automatic uploads."""
import hashlib
import json
import tempfile
import zipfile
from pathlib import Path

GAP_KINDS = {'excel_monitoring_gap', 'excel_attachment_gap', 'excel_events_disabled',
             'folder_unavailable', 'folder_scan_error', 'snapshot_skipped_budget',
             'snapshot_skipped_size', 'snapshot_skipped_disk_space', 'snapshot_retry',
             'excel_baseline_incomplete', 'excel_workbook_out_of_scope',
             'torn_record_detected', 'monitoring_paused', 'monitoring_resumed'}

def build_handover(store, start='', end='', include_snapshots=False):
    """Return a temporary ZIP path owned by the caller; stream rows rather than buffer."""
    with tempfile.NamedTemporaryFile(prefix='fcl-evidence-', suffix='.zip', delete=False) as temporary:
        output = Path(temporary.name)
    counts = {'events': 0, 'notes': 0, 'gaps': 0, 'malformed_lines': 0, 'snapshots': 0}
    try:
        with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=5,
                             allowZip64=True) as archive:
            with archive.open('events.jsonl', 'w') as events, archive.open('notes.jsonl', 'w') as notes, archive.open('gaps.jsonl', 'w') as gaps:
                # These are an observational timeline, not an exhaustive audit of activity.
                for log in sorted(store.logs.glob('events-*.jsonl')):
                    with log.open('rb') as raw:
                        for line in raw:
                            try:
                                item = json.loads(line)
                            except (UnicodeDecodeError, ValueError):
                                counts['malformed_lines'] += 1
                                continue
                            timestamp = item.get('timestamp_utc', '')
                            if (start and timestamp < start) or (end and timestamp > end):
                                continue
                            payload = (json.dumps(item, ensure_ascii=False, allow_nan=False) + '\\n').encode('utf-8')
                            events.write(payload)
                            counts['events'] += 1
                            if item.get('event') == 'user_explanation':
                                notes.write(payload)
                                counts['notes'] += 1
                            if item.get('event') in GAP_KINDS:
                                gaps.write(payload)
                                counts['gaps'] += 1
            if include_snapshots:
                for snapshot in sorted(store.snapshots.rglob('*')):
                    if snapshot.is_file() and not snapshot.name.endswith('.partial'):
                        archive.write(snapshot, 'snapshots/' + snapshot.relative_to(store.snapshots).as_posix())
                        counts['snapshots'] += 1
            manifest = {'format': 'FileChangeLog Phase 1 evidence', 'schema_version': 1,
                        'source': 'local_observation_only', 'from_utc': start or None, 'to_utc': end or None,
                        'include_snapshots': bool(include_snapshots), 'counts': counts,
                        'limitations': ['Unsaved edits require working desktop Excel COM events.',
                                        'Changes during sleeps, pauses and downtime may be missing.',
                                        'Snapshot budgets may leave versions unavailable.',
                                        'Explanations are supplied by the user, not inferred rules.']}
            archive.writestr('manifest.json', json.dumps(manifest, indent=2, ensure_ascii=False))
        return output
    except Exception:
        output.unlink(missing_ok=True)
        raise
