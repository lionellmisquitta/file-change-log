"""Portable bounded evidence export; private local files are never uploaded."""
import hashlib
import json
import tempfile
import zipfile
from pathlib import Path

GAP_KINDS = {
    'monitoring_gap', 'excel_monitoring_gap', 'excel_attachment_gap',
    'excel_events_disabled', 'folder_unavailable', 'folder_scan_error',
    'snapshot_skipped_budget', 'snapshot_skipped_size', 'snapshot_skipped_disk_space',
    'snapshot_budget_reached', 'baseline_allowance_reached', 'storage_warning',
    'snapshot_retry', 'snapshot_delayed', 'snapshot_deferred_spacing',
    'file_not_settling', 'comparison_failed',
    'excel_baseline_incomplete', 'excel_workbook_out_of_scope',
    'torn_record_detected', 'monitoring_paused', 'monitoring_resumed'
}
SNAPSHOT_FIELDS = ('snapshot', 'previous_snapshot', 'last_snapshot')


def build_handover(store, start='', end='', include_snapshots=False):
    """Make an explicit-range ZIP. Snapshots must be referenced by included events."""
    with tempfile.NamedTemporaryFile(prefix='fcl-evidence-', suffix='.zip', delete=False) as handle:
        output = Path(handle.name)
    counts = {'events': 0, 'notes': 0, 'gaps': 0, 'malformed_lines': 0, 'snapshots': 0}
    hashes = {}
    snapshot_paths = set()
    try:
        with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_DEFLATED,
                             compresslevel=5, allowZip64=True) as archive:
            for member, kind in (('events.jsonl', 'events'),
                                 ('notes.jsonl', 'notes'), ('gaps.jsonl', 'gaps')):
                digest = hashlib.sha256()
                with archive.open(member, 'w') as destination:
                    for log in sorted(store.logs.glob('events-*.jsonl')):
                        with log.open('rb') as raw:
                            for line in raw:
                                try:
                                    item = json.loads(line)
                                except (UnicodeDecodeError, ValueError):
                                    if kind == 'events':
                                        counts['malformed_lines'] += 1
                                    continue
                                timestamp = item.get('timestamp_utc', '')
                                if (start and timestamp < start) or (end and timestamp > end):
                                    continue
                                if kind == 'notes' and item.get('event') != 'user_explanation':
                                    continue
                                if kind == 'gaps' and item.get('event') not in GAP_KINDS:
                                    continue
                                if kind == 'events' and include_snapshots:
                                    for field in SNAPSHOT_FIELDS:
                                        value = item.get(field)
                                        if isinstance(value, str):
                                            snapshot_paths.add(value)
                                payload = (json.dumps(item, ensure_ascii=False, allow_nan=False) +
                                           '\n').encode('utf-8')
                                destination.write(payload)
                                digest.update(payload)
                                counts[kind] += 1
                hashes[member] = digest.hexdigest()

            if include_snapshots:
                root = store.snapshots.resolve()
                for relative in sorted(snapshot_paths):
                    candidate = (store.directory / relative).resolve()
                    if not candidate.is_file() or root not in candidate.parents:
                        continue
                    archive_name = 'snapshots/' + candidate.relative_to(root).as_posix()
                    digest = hashlib.sha256()
                    with candidate.open('rb') as source, archive.open(archive_name, 'w') as destination:
                        for chunk in iter(lambda: source.read(1024*1024), b''):
                            digest.update(chunk)
                            destination.write(chunk)
                    hashes[archive_name] = digest.hexdigest()
                    counts['snapshots'] += 1
            manifest = {
                'format': 'FileChangeLog Phase 1 evidence', 'schema_version': 2,
                'source': 'local_observation_only', 'from_utc': start or None,
                'to_utc': end or None, 'include_snapshots': bool(include_snapshots),
                'counts': counts, 'sha256_by_member': hashes,
                'limitations': [
                    'Unsaved Excel edits require working desktop Excel COM events.',
                    'Changes during sleeps, pauses and downtime may be missing.',
                    'Storage limits may prevent some source versions from being available.',
                    'Only snapshots referenced in the selected interval are exported.',
                    'Explanations are human supplied; no business rules are inferred.'
                ]
            }
            archive.writestr('manifest.json',
                             json.dumps(manifest, indent=2, ensure_ascii=False))
        return output
    except Exception:
        output.unlink(missing_ok=True)
        raise
