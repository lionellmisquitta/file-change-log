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
        count = 0
        details = []
        for sheet in sorted(set(old) | set(new)):
            if sheet not in old or sheet not in new:
                emit('snapshot_sheet_added' if sheet not in old else 'snapshot_sheet_removed', path, sheet=sheet)
            for address in sorted(set(old.get(sheet, {})) | set(new.get(sheet, {}))):
                a, b = old.get(sheet, {}).get(address), new.get(sheet, {}).get(address)
                if a != b:
                    count += 1
                    if len(details) < max_changes:
                        details.append(dict(sheet=sheet, address=address, old_value=a, new_value=b))
        emit('snapshot_cell_comparison', path, changes=details, total_changes=count,
             details_truncated=count > max_changes, comparison='cell_coordinates_not_business_row_keys',
             before_snapshot=str(before), after_snapshot=str(after))
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
