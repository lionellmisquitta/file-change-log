"""Windows desktop COM event collector; no Excel automation from a service."""
import os
import threading
import time
from .common import cell_address, canonical, clean, inside, now

class CellCache:
    def __init__(self):
        self.sheets = {}

    def seed(self, key, cells):
        self.sheets[key] = dict(cells)

    def invalidate(self, key):
        self.sheets.pop(key, None)

    def change(self, key, cells):
        cache = self.sheets.get(key)
        changes = []
        for address, state in cells.items():
            known = cache is not None and address in cache
            prior = cache.get(address) if known else None
            changes.append(dict(address=address, old_value=prior.get('value') if prior else None,
                                old_formula=prior.get('formula') if prior else None, old_value_known=known,
                                new_value=state.get('value'), new_formula=state.get('formula')))
        if cache is None:
            cache = self.sheets.setdefault(key, {})
        cache.update(cells)
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
            self.cache.seed(key, self.cells(sh.UsedRange, int(self.settings.get('excel_cache_cells', 200000))))
        except Exception as e:
            self.cache.invalidate(key)
            self.emit('excel_baseline_incomplete', path, sheet=str(sh.Name), reason=str(e), old_values_may_be_unknown=True)

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
            self.emit('excel_range_changed', path, sheet=str(sh.Name), range=str(target.Address),
                      details_captured=False, reason=str(e), old_baseline_invalidated=True,
                      save_state='not_confirmed_saved', action_type='unknown')
            return
        changes = self.cache.change(key, cells)
        self.emit('excel_range_changed', path, sheet=str(sh.Name), range=str(target.Address), changes=changes,
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
                self.safe(collector.sheet_event, 'excel_pivot_updated', Sh, instance, throttle=True)
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
        try:
            while not self.stop.is_set():
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
