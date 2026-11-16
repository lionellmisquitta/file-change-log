import csv
import json
import zipfile
from collections import Counter
from pathlib import Path

def workbook_cells(path, limit=200000):
    import openpyxl
    # Read copies only. Never save originals or execute workbook macros/links.
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False, keep_links=False)
    result = {}
    count = 0
    try:
        for ws in wb.worksheets:
            cells = {}
            for row in ws.iter_rows():
                for cell in row:
                    count += 1
                    if count > limit:
                        raise ValueError(f'Comparison exceeds {limit} scanned cells; snapshots are retained.')
                    if cell.value is not None:
                        cells[cell.coordinate] = cell.value
            result[ws.title] = cells
    finally:
        wb.close()
    return result

def workbook_features(path):
    result = dict(has_pivot_tables=False, has_power_query=False, has_macros=False,
                  has_external_workbook_links=False)
    if Path(path).suffix.lower() not in {'.xlsx', '.xlsm'}:
        return result
    with zipfile.ZipFile(path) as archive:
        names = [x.lower() for x in archive.namelist()]
        result['has_pivot_tables'] = any(x.startswith('xl/pivottables/') for x in names)
        result['has_power_query'] = any('microsoft.mashup' in x or x.startswith('xl/connections') for x in names)
        result['has_macros'] = any(x.endswith('/vbaproject.bin') for x in names)
        result['has_external_workbook_links'] = any(x.startswith('xl/externallinks/') for x in names)
    return result


def compare_files(before, after, emit, path, max_cells=200000, max_changes=5000):
    """Supplemental coordinate comparison, not an inferred action or commercial rule."""
    ext = Path(after).suffix.lower()
    if ext in {'.xlsx', '.xlsm'}:
        old, new = workbook_cells(before, max_cells), workbook_cells(after, max_cells)
        sheet_counts = Counter()
        column_counts = Counter()
        details = []
        count = 0
        for sheet in sorted(set(old) | set(new)):
            if sheet not in old or sheet not in new:
                emit('snapshot_sheet_added' if sheet not in old else 'snapshot_sheet_removed', path, sheet=sheet)
            for address in sorted(set(old.get(sheet, {})) | set(new.get(sheet, {}))):
                prior, value = old.get(sheet, {}).get(address), new.get(sheet, {}).get(address)
                if prior == value:
                    continue
                count += 1
                sheet_counts[sheet] += 1
                column = ''.join(x for x in address if x.isalpha())
                column_counts[(sheet, column)] += 1
                wanted = isinstance(prior, str) and prior.startswith('=') and not (isinstance(value, str) and value.startswith('='))
                details.append(dict(sheet=sheet, address=address, old_value=prior, new_value=value,
                                    explanation_wanted=wanted))
        # Typed values and manual overrides are more important than uniform formula fills.
        if len(details) > max_changes:
            details.sort(key=lambda x: (0 if x['explanation_wanted'] else
                         2 if isinstance(x['new_value'], str) and x['new_value'].startswith('=') else 1,
                         x['sheet'], x['address']))
            details = details[:max_changes]
        emit('snapshot_cell_comparison', path, changes=details, total_changes=count,
             changes_per_sheet=dict(sheet_counts),
             changes_per_sheet_column={sheet: {column: n for (name, column), n in column_counts.items() if name == sheet}
                                       for sheet in sorted(sheet_counts)},
             affected_sheets=sorted(sheet_counts), unlisted_changes=count-len(details),
             details_truncated=count > max_changes,
             explanation_wanted=any(x['explanation_wanted'] for x in details),
             no_cell_differences_other_content_changed=count == 0,
             comparison='cell_coordinates_not_business_row_keys',
             before_snapshot=str(before), after_snapshot=str(after))
        if count == 0:
            emit('no_cell_differences_other_content_changed', path,
                 before_snapshot=str(before), after_snapshot=str(after),
                 reason='Workbook bytes differ without changes in the compared cells',
                 filter_pivot_format_or_metadata_change_possible=True)
    elif ext in {'.csv', '.tsv', '.txt', '.md', '.json', '.xml', '.yaml', '.yml', '.sql', '.log', '.ini', '.cfg'}:
        if max(Path(before).stat().st_size, Path(after).stat().st_size) > 2 * 1024 * 1024:
            emit('comparison_skipped', path, reason='Text comparison limit: 2 MiB; snapshots retained.')
            return
        import difflib
        a = Path(before).read_text(encoding='utf-8-sig', errors='replace').splitlines()
        b = Path(after).read_text(encoding='utf-8-sig', errors='replace').splitlines()
        diff = []
        truncated = False
        for line in difflib.unified_diff(a, b, fromfile='before', tofile='after', n=2):
            if len(diff) >= max_changes:
                truncated = True
                break
            diff.append(line[:20000])
        emit('snapshot_text_comparison', path, diff=diff, details_truncated=truncated,
             before_snapshot=str(before), after_snapshot=str(after))
    else:
        emit('comparison_not_available', path, reason='Original snapshots retained; detailed saved-file comparison is not implemented for this format.',
             before_snapshot=str(before), after_snapshot=str(after))
