"""Windows desktop COM event collector; no Excel automation from a service."""
import os
import re
import threading
import time
from .common import cell_address, canonical, clean, inside, now

def range_bounds(address):
    """Return (min_col, min_row, max_col, max_row) for an A1 rectangle."""
    parts = re.findall(r'\$?([A-Z]+)\$?(\d+)', str(address).upper())
    if not parts:
        return None
    points = []
    for letters, row in parts:
        col = 0
        for character in letters:
            col = col * 26 + ord(character) - 64
        points.append((col, int(row)))
    return (min(p[0] for p in points), min(p[1] for p in points),
            max(p[0] for p in points), max(p[1] for p in points))


def is_growth_outside_old_range(old_range, new_range, changed_addresses):
    """Only preserve known-empty history when changes are entirely outside old extent."""
    old = range_bounds(old_range)
    new = range_bounds(new_range)
    if not old or not new or not changed_addresses:
        return False
    # Growing the rectangle must contain all of its previous cells.
    if new[0] > old[0] or new[1] > old[1] or new[2] < old[2] or new[3] < old[3]:
        return False
    for address in changed_addresses:
        cell = range_bounds(address)
        if not cell or (old[0] <= cell[0] <= old[2] and old[1] <= cell[1] <= old[3]):
            return False
    return True


class CellCache:
    """Bounded, conservative last-observed cell state; not an Excel undo log."""
    def __init__(self):
        self.sheets = {}
        self.complete = set()

    def seed(self, key, cells, complete=True):
        # Sparse cache: a fully scanned sheet has known-empty cells elsewhere.
        self.sheets[key] = {k: v for k, v in cells.items()
                            if v.get('value') is not None or v.get('formula') is not None}
        if complete:
            self.complete.add(key)
        else:
            self.complete.discard(key)

    def invalidate(self, key):
        self.sheets.pop(key, None)
        self.complete.discard(key)

    def change(self, key, cells):
        cache = self.sheets.get(key)
        changes = []
        empty = {'value': None, 'formula': None}
        for address, state in cells.items():
            known = cache is not None and (address in cache or key in self.complete)
            prior = cache.get(address, empty) if known else empty
            was_empty = prior.get('value') is None and prior.get('formula') is None
            is_empty = state.get('value') is None and state.get('formula') is None
            operation = ('observed_unknown_prior' if not known else
                         'added' if was_empty and not is_empty else
                         'cleared' if not was_empty and is_empty else 'updated')
            changes.append(dict(address=address, operation=operation,
                                old_value=prior.get('value') if known else None,
                                old_formula=prior.get('formula') if known else None,
                                old_value_known=known,
                                old_value_source='last_observed_at_address' if known else 'unknown',
                                new_value=state.get('value'), new_formula=state.get('formula')))
        if cache is None:
            cache = self.sheets.setdefault(key, {})
        for address, state in cells.items():
            if state.get('value') is None and state.get('formula') is None:
                cache.pop(address, None)
            else:
                cache[address] = state
        return changes

def matrix(value, rows, cols):
    if rows == 1 and cols == 1:
        return [[value]]
    return value

class ExcelListener:
    def __init__(self, store, settings, roots, enabled=lambda: True):
        self.store, self.settings, self.roots, self.enabled = store, settings, roots, enabled
        self.cache = CellCache()
        self.stop = threading.Event()
        self.health = dict(state='starting', attached_instances=0, last_event=None, detail='')
        self.connections = {}
        self.reset = threading.Event()
        self.known_books = set()
        self.sheet_names = {}
        self.throttles = {}
        self.warned_paths = set()
        self.used_ranges = {}
        self.reseed_needed = set()
        self.baseline_failure_reasons = {}
        self.thread = threading.Thread(target=self.run, daemon=True, name='Excel events')

    def scope(self, path):
        return self.enabled() and bool(path) and any(inside(path, root) for root in self.roots()) and not inside(path, self.store.directory)

    def emit(self, kind, path, **fields):
        self.health['last_event'] = now()
        return self.store.emit(kind, path, source='excel_desktop_event', actor='not_attributed_by_excel_event',
                               **fields)

    def path_of(self, workbook):
        try:
            return str(workbook.FullName) if workbook.Path else ''
        except Exception:
            return ''

    def sheet_key(self, sh):
        return canonical(self.path_of(sh.Parent)), str(sh.CodeName)

    def cells(self, target, limit):
        count = int(target.CountLarge)
        if count > limit:
            raise ValueError(f'Range has {count} cells; capture limit is {limit}')
        result = {}
        for area in target.Areas:
            rows, cols = int(area.Rows.Count), int(area.Columns.Count)
            values, formulas = matrix(area.Value2, rows, cols), matrix(area.Formula, rows, cols)
            for r in range(rows):
                for c in range(cols):
                    f = formulas[r][c]
                    result[cell_address(int(area.Row)+r, int(area.Column)+c)] = {
                        'value': clean(values[r][c]), 'formula': f if isinstance(f, str) and f.startswith('=') else None}
        return result

    def prime(self, workbook, instance):
        path = self.path_of(workbook)
        if not self.scope(path):
            if path and path not in self.warned_paths:
                self.warned_paths.add(path)
                # Do not ingest workbook content outside the user's roots.
                self.store.emit('excel_workbook_out_of_scope',
                                reason='An open Excel workbook did not match any selected monitored folder',
                                path_is_web_url=path.lower().startswith(('https://', 'http://')),
                                workbook_path_not_retained=True,
                                live_edits_not_captured=True)
            return
        book_key = (instance, canonical(path))
        if book_key in self.known_books:
            return
        self.known_books.add(book_key)
        self.emit('excel_workbook_attached', path, instance=instance,
                  previous_edits_not_recoverable=True, saved_status='current_in_memory_state_baseline')
        for sh in workbook.Worksheets:
            self.seed_sheet(sh, instance)

    def opened(self, workbook, instance):
        # A reopened path may have discarded the previous session's edits.
        path = canonical(self.path_of(workbook))
        self.known_books.discard((instance, path))
        for key in list(self.cache.sheets):
            if key[0] == instance and key[1] == path:
                self.cache.invalidate(key)
        self.prime(workbook, instance)

    def seed_sheet(self, sh, instance):
        path = self.path_of(sh.Parent)
        key = (instance, *self.sheet_key(sh))
        self.sheet_names[key] = str(sh.Name)
        try:
            baseline = self.cells(sh.UsedRange, int(self.settings.get('excel_cache_cells', 200000)))
            # A newly empty cell may lie outside UsedRange; a later change to it
            # must be classified as an addition rather than lost as unknown.
            budget = int(self.settings.get('excel_cache_cells', 200000))
            used = sum(len(v) for v in self.cache.sheets.values())
            if used + len(baseline) > budget:
                self.cache.invalidate(key)
                self.reseed_needed.discard(key)
                self.baseline_failure_reasons[key] = 'total_cache_budget'
                self.emit('excel_baseline_incomplete', path, sheet=str(sh.Name),
                          reason='Total cache budget reached', old_values_may_be_unknown=True)
            else:
                self.cache.seed(key, baseline)
                self.used_ranges[key] = str(sh.UsedRange.Address)
                self.reseed_needed.discard(key)
                self.baseline_failure_reasons.pop(key, None)
        except Exception as e:
            self.cache.invalidate(key)
            reason = str(e)
            non_retryable = isinstance(e, ValueError) and 'capture limit' in reason
            if non_retryable:
                self.reseed_needed.discard(key)
            else:
                self.reseed_needed.add(key)
            if self.baseline_failure_reasons.get(key) != reason:
                self.emit('excel_baseline_incomplete', path, sheet=str(sh.Name),
                          reason=reason, old_values_may_be_unknown=True,
                          retry_scheduled=not non_retryable)
                self.baseline_failure_reasons[key] = reason

    def changed(self, sh, target, instance):
        path = self.path_of(sh.Parent)
        if not self.scope(path):
            return
        key = (instance, *self.sheet_key(sh))
        maximum = int(self.settings.get('excel_event_cells', 10000))
        try:
            cells = self.cells(target, maximum)
        except Exception as e:
            self.cache.invalidate(key)
            self.reseed_needed.add(key)
            self.emit('excel_range_changed', path, sheet=str(sh.Name), range=str(target.Address),
                      details_captured=False, reason=str(e), old_baseline_invalidated=True,
                      save_state='not_confirmed_saved', action_type='unknown')
            return
        # A user typing in a new cell expands UsedRange. That is not by
        # itself structural editing: empty cells outside a complete seed are
        # known empty and should be correctly reported as additions.
        try:
            current_range = str(sh.UsedRange.Address)
            previous_range = self.used_ranges.get(key)
            range_changed = previous_range is not None and current_range != previous_range
            growth = (range_changed and
                      is_growth_outside_old_range(previous_range, current_range, cells))
            # A full-row/column edit is potentially structural; do not infer old values.
        except Exception as e:
            # A broken parser must never silently make stale values look reliable.
            self.cache.invalidate(key)
            self.reseed_needed.add(key)
            self.used_ranges.pop(key, None)
            self.emit('excel_range_structure_check_failed', path, sheet=str(sh.Name),
                      reason=str(e), old_values_may_be_unknown=True)
            current_range = None
            range_changed = False
            growth = False
        try:
            entire_axis = bool(target.EntireRow.CountLarge == sh.Cells.CountLarge or
                               target.EntireColumn.CountLarge == sh.Cells.CountLarge)
        except Exception:
            entire_axis = False
        if (range_changed and not growth) or entire_axis:
            self.cache.invalidate(key)
            self.reseed_needed.add(key)
            self.used_ranges.pop(key, None)
            self.emit('excel_old_values_invalidated', path, sheet=str(sh.Name),
                      reason='worksheet_extent_or_structure_changed')
        elif current_range is not None:
            self.used_ranges[key] = current_range
        changes = self.cache.change(key, cells)
        explanation_reasons = []
        if len(changes) >= 100:
            explanation_reasons.append('large_range_change')
        if entire_axis:
            explanation_reasons.append('whole_row_or_column_change')
        for change in changes:
            previous_formula = change.get('old_formula')
            current_formula = change.get('new_formula')
            if (change.get('old_value_known') and previous_formula and not current_formula
                    and change.get('new_value') is not None):
                change['explanation_wanted'] = True
                change['explanation_reason'] = 'formula_replaced_by_constant'
                if 'formula_replaced_by_constant' not in explanation_reasons:
                    explanation_reasons.append('formula_replaced_by_constant')
        self.emit('excel_range_changed', path, sheet=str(sh.Name), range=str(target.Address),
                  explanation_wanted=bool(explanation_reasons),
                  explanation_reasons=explanation_reasons, changes=changes,
                  details_captured=True, save_state='not_confirmed_saved', instance=instance,
                  action_type='not_inferred_from_range_change')

    def simple(self, kind, wb, instance, **fields):
        path = self.path_of(wb)
        if self.scope(path):
            self.emit(kind, path, instance=instance, **fields)

    def sheet_event(self, kind, sh, instance, throttle=False):
        path = self.path_of(sh.Parent)
        if not self.scope(path):
            return
        key = (instance, *self.sheet_key(sh))
        previous = self.sheet_names.get(key)
        current = str(sh.Name)
        if previous and previous != current:
            self.emit('excel_sheet_name_observed_changed', path, previous_name=previous, new_name=current)
        self.sheet_names[key] = current
        tag = (kind, key)
        if throttle and time.monotonic()-self.throttles.get(tag, 0) < 10:
            return
        self.throttles[tag] = time.monotonic()
        self.emit(kind, path, sheet=current, instance=instance, individual_cell_results_not_implied=True)

    def pivot_updated(self, sh, pivot, instance):
        """Capture observable COM pivot state, not a purported creation command."""
        path = self.path_of(sh.Parent)
        if not self.scope(path):
            return
        sheet = str(sh.Name)
        self.emit('excel_pivot_updated', path, sheet=sheet, instance=instance,
                  individual_cell_results_not_implied=True,
                  pivot_definition_observation_attempted=True)
        try:
            name = str(pivot.Name)[:160]
            source = None
            source_error = None
            try:
                source = str(pivot.SourceData)[:512]
            except Exception as exc:
                source_error = type(exc).__name__
            def axis_fields(axis):
                result = []
                for field in pivot.PivotFields():
                    if len(result) >= 200:
                        break
                    try:
                        if int(field.Orientation) == axis:
                            result.append(str(field.Name)[:160])
                    except Exception:
                        continue
                return result
            values = []
            for field in pivot.DataFields:
                if len(values) >= 100:
                    break
                try:
                    values.append(dict(name=str(field.Name)[:160],
                                       source_name=str(field.SourceName)[:160],
                                       function_code=int(field.Function)))
                except Exception:
                    continue
            self.emit('excel_pivot_definition_observed', path, sheet=sheet,
                      pivot_name=name, source_data=source,
                      source_read_error=source_error,
                      row_fields=axis_fields(1), column_fields=axis_fields(2),
                      page_fields=axis_fields(3), data_fields=values,
                      source='excel_desktop_event', completeness='partial_com_observation',
                      observation='pivot_updated_not_creation_proof')
        except Exception as exc:
            self.emit('excel_pivot_definition_unavailable', path, sheet=sheet,
                      reason=type(exc).__name__,
                      pivot_update_observed=True)

    def handlers(self, instance):
        collector = self
        class Events:
            def safe(self, func, *args, **kwargs):
                try:
                    func(*args, **kwargs)
                except Exception as e:
                    collector.health.update(detail=f'Event capture failed: {e}')
                    try:
                        collector.store.emit('excel_event_capture_error', reason=str(e), instance=instance)
                    except Exception:
                        pass
            def OnWorkbookOpen(self, Wb):
                self.safe(collector.opened, Wb, instance)
                self.safe(collector.simple, 'excel_workbook_open', Wb, instance)
            def OnWorkbookActivate(self, Wb):
                self.safe(collector.prime, Wb, instance)
            def OnWorkbookBeforeSave(self, Wb, SaveAsUI, Cancel):
                self.safe(collector.simple, 'excel_save_requested', Wb, instance, save_as_dialog=bool(SaveAsUI), success_not_implied=True)
            def OnWorkbookAfterSave(self, Wb, Success):
                self.safe(collector.prime, Wb, instance)
                self.safe(collector.simple, 'excel_save_result', Wb, instance, success=bool(Success))
            def OnWorkbookBeforeClose(self, Wb, Cancel):
                self.safe(collector.simple, 'excel_close_requested', Wb, instance, completed_not_implied=True)
                # Release associated in-memory caches; no deleted-content inference.
                path = canonical(collector.path_of(Wb))
                for key in list(collector.cache.sheets):
                    if key[0] == instance and key[1] == path:
                        collector.cache.invalidate(key)
                        collector.used_ranges.pop(key, None)
                        collector.reseed_needed.discard(key)
                collector.known_books.discard((instance, path))
            def OnSheetChange(self, Sh, Target):
                self.safe(collector.changed, Sh, Target, instance)
            def OnWorkbookNewSheet(self, Wb, Sh):
                self.safe(collector.sheet_event, 'excel_sheet_created', Sh, instance)
                self.safe(collector.seed_sheet, Sh, instance)
            def OnSheetBeforeDelete(self, Sh):
                self.safe(collector.sheet_event, 'excel_sheet_delete_requested', Sh, instance)
            def OnSheetActivate(self, Sh):
                self.safe(collector.sheet_event, 'excel_sheet_activated', Sh, instance)
            def OnSheetCalculate(self, Sh):
                self.safe(collector.sheet_event, 'excel_sheet_calculated', Sh, instance, throttle=True)
            def OnSheetPivotTableUpdate(self, Sh, Target):
                self.safe(collector.pivot_updated, Sh, Target, instance)
        return Events

    def discover(self):
        import pythoncom
        import win32com.client
        import win32process
        apps = {}
        def add(app):
            try:
                hwnd = int(app.Hwnd)
                if hwnd:
                    apps[win32process.GetWindowThreadProcessId(hwnd)[1]] = app
            except Exception:
                pass
        try:
            add(win32com.client.GetActiveObject('Excel.Application'))
        except Exception:
            pass
        rot = pythoncom.GetRunningObjectTable()
        context = pythoncom.CreateBindCtx(0)
        for moniker in rot.EnumRunning():
            try:
                name = moniker.GetDisplayName(context, None)
                if any(ext in name.lower() for ext in ('.xlsx', '.xlsm', '.xlsb', '.xls', '.csv')):
                    obj = win32com.client.Dispatch(rot.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch))
                    if str(obj.Application.Name).lower() == 'microsoft excel':
                        add(obj.Application)
            except Exception:
                continue
        return apps

    def run(self):
        if os.name != 'nt':
            self.health.update(state='unavailable', detail='Live Excel events require Windows desktop Excel. Folder monitoring remains available.')
            return
        try:
            import pythoncom
            import win32com.client
            import win32gui
            import win32process
            pythoncom.CoInitialize()
        except Exception as e:
            self.health.update(state='unavailable', detail=f'Excel COM dependency unavailable: {e}')
            self.store.emit('excel_listener_unavailable', reason=str(e))
            return
        previous_gaps = None
        last_tick = time.monotonic()
        try:
            while not self.stop.is_set():
                current_tick = time.monotonic()
                if current_tick - last_tick > 30:
                    self.store.emit('excel_monitoring_gap', reason='monitor_thread_delayed_possible_sleep_or_suspend', gap_seconds=round(current_tick-last_tick, 1), intermediate_edits_not_recoverable=True)
                    self.cache.sheets.clear()
                    self.known_books.clear()
                    self.connections.clear()
                last_tick = current_tick
                try:
                    if self.reset.is_set():
                        self.cache.sheets.clear()
                        self.known_books.clear()
                        self.reset.clear()
                    apps = self.discover()
                    for hwnd, app in apps.items():
                        if hwnd not in self.connections:
                            sink = win32com.client.WithEvents(app, self.handlers(str(hwnd)))
                            self.connections[hwnd] = (app, sink)
                        for wb in app.Workbooks:
                            self.prime(wb, str(hwnd))
                            for sh in wb.Worksheets:
                                key = (str(hwnd), *self.sheet_key(sh))
                                if key in self.reseed_needed:
                                    self.seed_sheet(sh, str(hwnd))
                        if not bool(app.EnableEvents):
                            self.health['detail'] = 'Excel has disabled events; edits may be missed.'
                            if self.throttles.get(('disabled', hwnd), 0) + 30 < time.monotonic():
                                self.store.emit('excel_events_disabled', instance=hwnd, capture_gap=True)
                                self.throttles[('disabled', hwnd)] = time.monotonic()
                    for hwnd in list(self.connections):
                        try:
                            _ = self.connections[hwnd][0].Workbooks.Count
                        except Exception:
                            self.connections.pop(hwnd)
                            self.known_books = {k for k in self.known_books if k[0] != str(hwnd)}
                            self.store.emit('excel_instance_disconnected', instance=hwnd)
                    visible_pids = set()
                    def window(hwnd, _):
                        if win32gui.GetClassName(hwnd) == 'XLMAIN':
                            visible_pids.add(win32process.GetWindowThreadProcessId(hwnd)[1])
                    win32gui.EnumWindows(window, None)
                    attached_pids = set(self.connections)
                    gaps = sorted(visible_pids - attached_pids)
                    if gaps != previous_gaps:
                        if gaps:
                            self.store.emit('excel_attachment_gap', process_ids=gaps,
                                            reason='Excel instance visible but not discoverable through COM. Edits in that instance are not captured.')
                        previous_gaps = gaps
                    self.health.update(state='partial' if gaps else ('listening' if self.connections else 'waiting_for_excel'),
                                       attached_instances=len(self.connections), unattached_processes=gaps)
                    # Pump COM messages in the same apartment that subscribed.
                    deadline = time.monotonic()+2
                    while time.monotonic()<deadline and not self.stop.is_set():
                        pythoncom.PumpWaitingMessages()
                        self.stop.wait(.05)
                except Exception as e:
                    self.health.update(state='retrying', detail=str(e))
                    self.stop.wait(2)
        finally:
            self.connections.clear()
            pythoncom.CoUninitialize()
