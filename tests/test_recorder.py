import json
import os
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import patch
import urllib.error
import urllib.request
from http.server import ThreadingHTTPServer

from recorder.app import Recorder, make_handler, validate_config
from recorder.common import inside, cell_address, canonical
from recorder.compare import compare_files
from recorder.excel import CellCache, ExcelListener
from recorder.folders import FolderWorker
from recorder.store import Store

class Workspace(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.base = Path(self.temp.name)
        self.root = self.base/'watched'; self.root.mkdir()
        self.store = Store(self.base/'history')
        self.worker = FolderWorker(self.root, self.store, {'settle_seconds': 0, 'snapshot_limit_mb': 512})
    def tearDown(self):
        self.store.close(); self.temp.cleanup()
    def capture(self):
        self.worker.scan_once(); self.worker.scan_once()
    def kinds(self):
        return [e['event'] for e in self.store.recent(limit=1000)]

class FileCaptureTests(Workspace):
    def test_saved_versions_and_text_changes_preserve_original(self):
        p=self.root/'orders.csv';p.write_text('id,qty\nA,100\n')
        self.capture(); original=p.read_bytes()
        p.write_text('id,qty\nA,10\n');self.capture()
        versions=[e for e in self.store.recent() if e['event']=='file_version_captured']
        self.assertEqual(len(versions),2)
        self.assertEqual((self.store.directory/versions[-1]['snapshot']).read_bytes(), original)
        self.assertIn('snapshot_text_comparison',self.kinds())
        self.assertEqual(p.read_text(),'id,qty\nA,10\n')
    def test_other_binary_type_is_logged_not_copied(self):
        (self.root/'video.mp4').write_bytes(b'test')
        self.capture()
        self.assertIn('snapshot_not_selected',self.kinds())
        self.assertFalse(list(self.store.snapshots.rglob('*.mp4')))
    def test_document_snapshot_without_claiming_content_diff(self):
        p=self.root/'report.pdf';p.write_bytes(b'version1');self.capture()
        p.write_bytes(b'version2');self.capture()
        self.assertIn('comparison_not_available',self.kinds())
        self.assertEqual(len(list(self.store.snapshots.rglob('*.pdf'))),2)
    def test_offline_root_never_causes_mass_deletions(self):
        (self.root/'a.txt').write_text('keep');self.capture()
        moved=self.base/'disconnected';self.root.rename(moved)
        self.worker.scan_once()
        self.assertIn('folder_unavailable',self.kinds())
        self.assertNotIn('file_absent_on_scan',self.kinds())
        moved.rename(self.root);self.worker.scan_once()
        self.assertIn('folder_reconnected',self.kinds())
    def test_partial_enumeration_is_not_deletion(self):
        (self.root/'a.txt').write_text('keep');self.capture()
        with patch('recorder.folders.walk_files',side_effect=OSError('permission denied')):
            self.worker.scan_once()
        self.assertNotIn('file_absent_on_scan',self.kinds())
    def test_deleted_file_keeps_snapshot(self):
        p=self.root/'a.txt';p.write_text('keep');self.capture();p.unlink();self.worker.scan_once()
        self.assertIn('file_absent_on_scan',self.kinds())
        self.assertEqual(len(list(self.store.snapshots.rglob('*.txt'))),1)
    def test_source_changed_during_copy_does_not_commit_snapshot(self):
        p=self.root/'a.txt';p.write_text('keep')
        from recorder.folders import signature
        sig=signature(p)
        with patch('recorder.folders.signature',return_value=[9,9,9,9]):
            with self.assertRaises(OSError):self.worker.capture(str(p),sig)
        self.assertFalse(list(self.store.snapshots.rglob('*.partial')))
        self.assertFalse(list(self.store.snapshots.rglob('*.txt')))
    def test_locked_file_is_retried_without_replacing_baseline(self):
        p=self.root/'a.txt';p.write_text('old');self.capture()
        before=self.store.inventory(self.worker.root)
        p.write_text('new');self.worker.scan_once()
        with patch.object(self.worker,'capture',side_effect=PermissionError('locked')):
            self.worker.scan_once()
        self.assertEqual(self.store.inventory(self.worker.root),before)
        self.assertIn('snapshot_retry',self.kinds())
        self.worker.scan_once()
        self.assertNotEqual(self.store.inventory(self.worker.root),before)
    def test_temporary_files_are_ignored(self):
        for name in ['~$orders.xlsx','incomplete.tmp']:(self.root/name).write_text('ignored')
        self.capture();self.assertFalse(self.store.inventory(self.worker.root))
    def test_no_duplicate_capture_after_unchanged_restart(self):
        (self.root/'a.txt').write_text('a');self.capture()
        fresh=FolderWorker(self.root,self.store,{'settle_seconds':0})
        fresh.scan_once();fresh.scan_once()
        self.assertEqual(self.kinds().count('file_version_captured'),1)
    def test_paused_worker_does_not_capture(self):
        (self.root/'a.txt').write_text('a');self.worker.enabled=lambda:False
        self.capture();self.assertEqual(self.kinds(),[])
    def test_symlink_does_not_escape_scope(self):
        outside=self.base/'outside';outside.mkdir();(outside/'secret.txt').write_text('secret')
        try:(self.root/'link').symlink_to(outside,target_is_directory=True)
        except OSError:self.skipTest('Symlink privilege unavailable')
        self.capture();self.assertFalse(self.store.inventory(self.worker.root))

class ComparisonTests(Workspace):
    def test_xlsx_values_formulas_and_sheet_changes(self):
        import openpyxl
        before=self.base/'before.xlsx';after=self.base/'after.xlsx'
        wb=openpyxl.Workbook();wb.active['A1']=100;wb.active['B1']='=A1*2';wb.save(before)
        wb.active['A1']=10;wb.active['B1']='=A1*3';wb.create_sheet('New');wb.save(after);wb.close()
        original=before.read_bytes()
        compare_files(before,after,self.store.emit,'source.xlsx')
        e=next(e for e in self.store.recent() if e['event']=='snapshot_cell_comparison')
        self.assertEqual(e['total_changes'],2);self.assertFalse(e['details_truncated'])
        self.assertIn('snapshot_sheet_added',self.kinds());self.assertEqual(before.read_bytes(),original)
    def test_first_snapshot_reports_existing_values_without_inventing_edits(self):
        import openpyxl
        p = self.root/'initial.xlsx'
        wb = openpyxl.Workbook()
        wb.active['A1'] = 12345
        wb.active['B3'] = 'new entry'
        wb.save(p)
        wb.close()
        self.capture()
        event = next(e for e in self.store.recent() if e['event'] == 'snapshot_initial_cells')
        self.assertEqual(event['total_populated_cells'], 2)
        self.assertEqual({(c['address'], c['value']) for c in event['cells']},
                         {('A1', 12345), ('B3', 'new entry')})
        self.assertEqual(event['observation'],
                         'first_saved_version_not_proof_of_when_cells_were_entered')
        self.assertNotIn('snapshot_cell_comparison', self.kinds())

    def test_newly_populated_cell_in_second_version_is_reported(self):
        import openpyxl
        p = self.root/'new-entry.xlsx'
        wb = openpyxl.Workbook()
        wb.active['A1'] = 100
        wb.save(p)
        self.capture()
        wb.active['A2'] = 99999
        wb.save(p)
        wb.close()
        self.capture()
        event = next(e for e in self.store.recent() if e['event'] == 'snapshot_cell_comparison')
        self.assertIn(dict(sheet='Sheet', address='A2', old_value=None, new_value=99999),
                      event['changes'])

    def test_comparison_limit_is_explicit(self):
        import openpyxl
        paths=[self.base/'a.xlsx',self.base/'b.xlsx']
        for p in paths:
            wb=openpyxl.Workbook();wb.active['B2']=1;wb.save(p);wb.close()
        with self.assertRaises(ValueError):compare_files(*paths,self.store.emit,'source',max_cells=1)

class CacheTests(unittest.TestCase):
    def test_unknown_old_is_not_invented(self):
        c=CellCache();e=c.change('sheet',{'A1':{'value':10,'formula':None}})[0]
        self.assertFalse(e['old_value_known']);self.assertIsNone(e['old_value'])
    def test_newly_filled_and_cleared_existing_cells(self):
        c = CellCache()
        c.seed('sheet', {'A1': {'value': None, 'formula': None}, 'A2': {'value': 123, 'formula': None}})
        filled = c.change('sheet', {'A1': {'value': 99999, 'formula': None}})[0]
        cleared = c.change('sheet', {'A2': {'value': None, 'formula': None}})[0]
        self.assertEqual(filled['operation'], 'added')
        self.assertTrue(filled['old_value_known'])
        self.assertEqual(cleared['operation'], 'cleared')
        self.assertEqual(cleared['old_value'], 123)

    def test_old_and_new_formulas_are_distinct(self):
        c=CellCache();c.seed('sheet',{'A1':{'value':2,'formula':'=1+1'}})
        e=c.change('sheet',{'A1':{'value':3,'formula':'=1+2'}})[0]
        self.assertTrue(e['old_value_known']);self.assertEqual(e['old_formula'],'=1+1');self.assertEqual(e['new_formula'],'=1+2')
        c.invalidate('sheet');self.assertFalse(c.change('sheet',{'A1':{'value':4}})[0]['old_value_known'])
    def test_cell_addresses(self):
        self.assertEqual(cell_address(2,28),'AB2')

class StoreTests(Workspace):
    def test_concurrent_events_are_complete_json(self):
        def write(n):
            for i in range(10):self.store.emit('concurrent',thread=n,item=i)
        ts=[threading.Thread(target=write,args=(n,)) for n in range(4)]
        for t in ts:t.start()
        for t in ts:t.join()
        self.assertEqual(len(self.store.recent()),40)
        lines=list(self.store.logs.glob('*.jsonl'))[0].read_text().splitlines()
        self.assertEqual(len({json.loads(x)['id'] for x in lines}),40)
    def test_literal_search_does_not_expand_sql_wildcards(self):
        self.store.emit('file','100%_done');self.store.emit('other','100Xdone')
        self.assertEqual(len(self.store.recent(search='%_')),1)
    def test_jsonl_rebuilds_lost_index_entry(self):
        event=self.store.emit('recoverable','x')
        with self.store.lock:self.store.db.execute('DELETE FROM events');self.store.db.commit()
        recovered=Store(self.store.directory)
        try:self.assertEqual(recovered.recent()[0]['id'],event['id'])
        finally:recovered.close()

class PhaseOneRegressionTests(Workspace):
    def test_torn_jsonl_record_does_not_swallow_next_event(self):
        self.store.emit('before_torn', 'a')
        log = next(self.store.logs.glob('events-*.jsonl'))
        with log.open('ab') as handle:
            handle.write(b'{"truncated":')
        recovered = Store(self.store.directory)
        try:
            recovered.emit('after_torn', 'b')
        finally:
            recovered.close()
        replay = Store(self.store.directory)
        try:
            kinds = [event['event'] for event in replay.recent()]
            self.assertIn('torn_record_detected', kinds)
            self.assertIn('after_torn', kinds)
        finally:
            replay.close()

    def test_snapshot_deduplicated_across_folders_and_renames(self):
        p = self.root/'one.txt'
        p.write_bytes(b'hello')
        self.capture()
        used = self.store.snapshot_bytes
        (self.root/'sub').mkdir()
        (self.root/'sub'/'two.txt').write_bytes(b'hello')
        self.capture()
        self.assertEqual(self.store.snapshot_bytes, used)
        (self.root/'sub'/'two.txt').rename(self.root/'sub'/'three.txt')
        self.capture()
        self.assertEqual(self.store.snapshot_bytes, used)

    def test_rename_preserves_file_identity_without_false_deletion(self):
        source = self.root/'before.csv'
        destination = self.root/'after.csv'
        source.write_text('record,total\\n1,15\\n')
        self.capture()
        old_snapshot = self.store.inventory(self.worker.root)[canonical(source)][1]
        source.rename(destination)
        self.worker.scan_once()
        new_inventory = self.store.inventory(self.worker.root)
        self.assertNotIn(canonical(source), new_inventory)
        self.assertEqual(new_inventory[canonical(destination)][1], old_snapshot)
        self.assertIn('file_moved', self.kinds())
        self.assertNotIn('file_absent_on_scan', self.kinds())

    def test_budget_skips_with_explicit_evidence(self):
        self.worker.settings['snapshot_budget_mb'] = 10
        self.worker.settings['baseline_budget_mb'] = 0
        p = self.root/'orders.txt'
        p.write_bytes(b'something')
        self.capture()
        self.assertIn('snapshot_skipped_budget', self.kinds())
        self.assertIsNone(self.store.inventory(self.worker.root)[canonical(p)][1])

    def test_recursive_paste_inside_new_nested_directory(self):
        deep = self.root/'year'/'quarter'/'batch'
        deep.mkdir(parents=True)
        (deep/'payments.csv').write_text('invoice,value\n1,300\n')
        self.capture()
        self.assertIn(canonical(deep/'payments.csv'), self.store.inventory(self.worker.root))

    def test_sparse_seeded_sheet_classifies_new_outside_used_range(self):
        cache = CellCache()
        cache.seed('sheet', {'A1': {'value': 3, 'formula': None}})
        result = cache.change('sheet', {'Z10': {'value': 7, 'formula': None}})[0]
        self.assertTrue(result['old_value_known'])
        self.assertEqual(result['operation'], 'added')

    def test_invalidated_sheet_never_invents_old_value(self):
        cache = CellCache()
        cache.seed('sheet', {'A2': {'value': 'old', 'formula': None}})
        cache.invalidate('sheet')
        result = cache.change('sheet', {'A2': {'value': 'new', 'formula': None}})[0]
        self.assertFalse(result['old_value_known'])
        self.assertEqual(result['operation'], 'observed_unknown_prior')

class PilotFindingsRegressionTests(Workspace):
    def test_reconciliation_labels_changes_made_while_down(self):
        p = self.root/'working.csv'
        p.write_text('before')
        self.capture()
        p.write_text('after')
        worker = FolderWorker(self.root, self.store,
                              {'settle_seconds': 0}, reconcile_on_first_scan=True)
        worker.scan_once()
        worker.scan_once()
        events = [e for e in self.store.recent() if e['event'] == 'file_version_captured']
        self.assertEqual(events[0]['classification'], 'changed_while_not_monitoring')

    def test_baseline_and_total_snapshot_limits_have_distinct_events(self):
        self.worker.settings.update(snapshot_budget_mb=10, baseline_budget_mb=0)
        (self.root/'initial.txt').write_text('initial')
        self.capture()
        self.assertIn('baseline_allowance_reached', self.kinds())
        self.assertNotIn('snapshot_budget_reached', self.kinds())
        self.worker.settings['snapshot_budget_mb'] = .000001
        (self.root/'later.txt').write_text('later')
        self.capture()
        self.assertIn('snapshot_budget_reached', self.kinds())

    def test_high_storage_delays_intermediate_snapshots_without_losing_last(self):
        p = self.root/'busy.csv'
        p.write_text('first')
        self.worker.settings.update(snapshot_budget_mb=10, baseline_budget_mb=5)
        self.capture()
        initial_count = self.kinds().count('file_version_captured')
        self.store.snapshot_bytes = 7*1024*1024
        p.write_text('second')
        self.capture()
        self.assertEqual(self.kinds().count('file_version_captured'), initial_count)
        self.assertIn(canonical(p), self.worker.pending)
        self.worker.last_snapshot_at[canonical(p)] -= 901
        self.worker.scan_once()
        self.assertEqual(self.kinds().count('file_version_captured'), initial_count+1)

    def test_identical_baseline_workbooks_do_not_repeat_cell_dumps(self):
        import openpyxl
        a = self.root/'a.xlsx'
        b = self.root/'b.xlsx'
        book = openpyxl.Workbook()
        book.active['A1'] = 'same'
        book.save(a)
        book.close()
        b.write_bytes(a.read_bytes())
        self.capture()
        self.assertEqual(self.kinds().count('snapshot_initial_cells'), 1)

    def test_used_range_growth_keeps_known_empty_excel_cells(self):
        from types import SimpleNamespace as N
        listener = ExcelListener(self.store, {'excel_cache_cells': 100,
                            'excel_event_cells': 10}, lambda:[str(self.root)])
        wb = N(Path=str(self.root), FullName=str(self.root/'book.xlsx'))
        area = N(Row=1, Column=1, Rows=N(Count=1), Columns=N(Count=1),
                 Value2=42, Formula=42)
        original = N(CountLarge=1, Areas=[area], Address='$A$1')
        sh = N(Parent=wb, CodeName='Sheet1', Name='Sheet', UsedRange=original)
        wb.Worksheets = [sh]
        listener.opened(wb, 'excel1')
        sh.UsedRange = N(Address='$A$1:$Z$10')
        target_area = N(Row=10, Column=26, Rows=N(Count=1),
                        Columns=N(Count=1), Value2='new note', Formula='new note')
        target = N(CountLarge=1, Areas=[target_area], Address='$Z$10')
        listener.changed(sh, target, 'excel1')
        first = next(e for e in self.store.recent() if e['event']=='excel_range_changed')
        self.assertEqual(first['changes'][0]['operation'], 'added')
        self.assertTrue(first['changes'][0]['old_value_known'])
        self.assertNotIn('excel_old_values_invalidated', self.kinds())
        target_area.Row = 7
        target_area.Column = 3
        target_area.Value2 = 'next'
        target_area.Formula = 'next'
        target.Address = '$C$7'
        listener.changed(sh, target, 'excel1')
        after = next(e for e in self.store.recent() if e['event']=='excel_range_changed')
        self.assertEqual(after['changes'][0]['operation'], 'added')
        self.assertTrue(after['changes'][0]['old_value_known'])

class NovemberV2RetestRegressions(Workspace):
    def test_a1_range_parser_and_growth_predicate(self):
        from recorder.excel import range_bounds, is_growth_outside_old_range
        self.assertEqual(range_bounds('$A$1:$Z$10'), (1, 1, 26, 10))
        self.assertEqual(range_bounds('$C$7'), (3, 7, 3, 7))
        self.assertTrue(is_growth_outside_old_range('$A$1:$C$6',
                                                   '$A$1:$Z$10', {'Z10': {}}))
        self.assertFalse(is_growth_outside_old_range('$A$1:$C$6',
                                                    '$A$1:$C$10', {'B4': {}}))

    def test_inserting_cells_invalidates_old_value_at_shifted_address(self):
        from types import SimpleNamespace as N
        listener = ExcelListener(self.store, {'excel_cache_cells': 1000,
                                  'excel_event_cells': 20}, lambda: [str(self.root)])
        wb = N(Path=str(self.root), FullName=str(self.root/'royalty.xlsx'))
        area = N(Row=1, Column=1, Rows=N(Count=1), Columns=N(Count=1),
                 Value2='seed', Formula='seed')
        target = N(CountLarge=1, Areas=[area], Address='$A$1')
        sh = N(Parent=wb, CodeName='Sheet1', Name='Royalties',
               UsedRange=N(Address='$A$1:$C$10', CountLarge=1,
                           Areas=[area]))
        wb.Worksheets = [sh]
        listener.opened(wb, 'excel')
        key = ('excel', *listener.sheet_key(sh))
        listener.cache.seed(key, {'B7': {'value': 'v7_2', 'formula': None}})
        sh.UsedRange.Address = '$A$1:$C$11'
        area.Row = 4
        area.Value2 = 'inserted'
        area.Formula = 'inserted'
        target.Address = '$A$4'
        listener.changed(sh, target, 'excel')
        self.assertIn('excel_old_values_invalidated', self.kinds())
        area.Row = 7
        area.Column = 2
        area.Value2 = 'v6_2'
        area.Formula = 'v6_2'
        target.Address = '$B$7'
        listener.changed(sh, target, 'excel')
        latest = next(e for e in self.store.recent()
                      if e['event']=='excel_range_changed')
        self.assertFalse(latest['changes'][0]['old_value_known'])
        self.assertEqual(latest['changes'][0]['operation'], 'observed_unknown_prior')

    def test_oversized_sheet_baseline_does_not_schedule_poll_retry(self):
        from types import SimpleNamespace as N
        listener = ExcelListener(self.store, {'excel_cache_cells': 200000,
                                  'excel_event_cells': 10000}, lambda: [str(self.root)])
        wb = N(Path=str(self.root), FullName=str(self.root/'large.xlsx'))
        huge = N(CountLarge=250000, Address='$A$1:$AX$5000')
        sh = N(Parent=wb, CodeName='Sheet1', Name='Large', UsedRange=huge)
        wb.Worksheets = [sh]
        listener.prime(wb, 'excel')
        key = ('excel', *listener.sheet_key(sh))
        self.assertNotIn(key, listener.reseed_needed)
        self.assertEqual(self.kinds().count('excel_baseline_incomplete'), 1)
        listener.seed_sheet(sh, 'excel')
        self.assertEqual(self.kinds().count('excel_baseline_incomplete'), 1)

    def test_snapshot_spacing_is_logged_once_then_final_version_captured(self):
        p = self.root/'sales.csv'
        p.write_text('one')
        self.worker.settings.update(snapshot_budget_mb=10, baseline_budget_mb=5)
        self.capture()
        self.store.snapshot_bytes = 7*1024*1024
        p.write_text('two')
        self.capture()
        self.worker.scan_once()
        self.assertEqual(self.kinds().count('snapshot_deferred_spacing'), 1)
        self.assertIn(canonical(p), self.worker.pending)
        self.worker.last_snapshot_at[canonical(p)] -= 901
        self.worker.scan_once()
        self.assertEqual(self.kinds().count('snapshot_deferred_spacing'), 1)
        self.assertEqual(self.kinds().count('file_version_captured'), 2)

class ConfigurationTests(Workspace):
    @unittest.skipUnless(os.name == 'nt', 'Windows short-path alias check')
    def test_windows_short_path_cannot_bypass_history_exclusion(self):
        import win32api
        short=win32api.GetShortPathName(str(self.base))
        with self.assertRaises(ValueError):validate_config({'roots':[short]},self.store.directory)
    def test_reject_output_within_monitored_root(self):
        with self.assertRaises(ValueError):validate_config({'roots':[str(self.base)]},self.store.directory)
    def test_reject_nested_roots(self):
        with self.assertRaises(ValueError):validate_config({'roots':[str(self.root),str(self.root/'sub')]},self.store.directory)
    def test_accept_unavailable_absolute_network_mount(self):
        config=validate_config({'roots':[str(self.base/'offline')]},self.store.directory)
        self.assertEqual(len(config['roots']),1)
    def test_reject_invalid_intervals(self):
        with self.assertRaises(ValueError):validate_config({'scan_seconds':0},self.store.directory)
    def test_scope_blocks_sibling_prefix(self):
        self.assertFalse(inside(self.base/'watched-other',self.root))

class HttpTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.recorder=Recorder(Path(self.temp.name)/'history')
        self.server=ThreadingHTTPServer(('127.0.0.1',0),make_handler(self.recorder,'secret',lambda:None))
        self.thread=threading.Thread(target=self.server.serve_forever,daemon=True);self.thread.start()
        self.url=f'http://127.0.0.1:{self.server.server_port}'
    def tearDown(self):
        self.server.shutdown();self.server.server_close();self.recorder.store.close();self.temp.cleanup()
    def request(self,path,body=None,token='secret',host=None):
        headers={'X-Recorder-Token':token,'Content-Type':'application/json'}
        if host:headers['Host']=host
        return urllib.request.urlopen(urllib.request.Request(self.url+path,headers=headers,data=json.dumps(body).encode() if body is not None else None),timeout=3)
    def test_no_token_cannot_read_content(self):
        with self.assertRaises(urllib.error.HTTPError) as e:self.request('/api/events',token='')
        self.assertEqual(e.exception.code,403)
    def test_untrusted_host_is_rejected(self):
        with self.assertRaises(urllib.error.HTTPError) as e:self.request('/api/status',host='evil.example')
        self.assertEqual(e.exception.code,403)
    def test_note_roundtrip_and_export(self):
        text='<script>alert(1)</script> — only Q3'
        with self.request('/api/note',{'text':text,'event_id':'abc','scope':'case_note'}) as r:self.assertEqual(r.status,200)
        with self.request('/api/events') as r:data=json.load(r)
        self.assertEqual(data[0]['text'],text);self.assertTrue(data[0]['context_confirmed']);self.assertNotIn('reusable_rule_not_approved',data[0])
        with self.request('/api/export') as r:self.assertIn('only Q3',r.read().decode())
    def test_note_empty_rejected(self):
        with self.assertRaises(urllib.error.HTTPError) as e:self.request('/api/note',{'text':''})
        self.assertEqual(e.exception.code,400)
    def test_evidence_zip_contains_complete_history_notes_and_manifest(self):
        import io
        import zipfile
        self.recorder.store.emit('file_version_captured', 'a.csv', snapshot='snapshots/abc.csv')
        self.recorder.store.emit('user_explanation', text='manual royalty adjustment', scope='case_note')
        with self.request('/api/handover') as response:
            body = response.read()
        with zipfile.ZipFile(io.BytesIO(body)) as archive:
            names = set(archive.namelist())
            self.assertTrue({'events.jsonl', 'notes.jsonl', 'gaps.jsonl', 'manifest.json'} <= names)
            manifest = json.loads(archive.read('manifest.json'))
            self.assertEqual(manifest['counts']['notes'], 1)
            self.assertIn(b'manual royalty adjustment', archive.read('notes.jsonl'))
            self.assertEqual(manifest['counts']['events'], 2)
            self.assertFalse(any(name.startswith('snapshots/') for name in names))

    def test_silent_and_explanation_modes_persist(self):
        with self.request('/api/mode', {'mode': 'explanation', 'floating': True}) as response:
            self.assertEqual(response.status, 200)
        config = json.loads(self.recorder.config_path.read_text(encoding='utf-8'))
        self.assertEqual(config['explanation_mode'], 'explanation')
        self.assertTrue(config['floating_notes'])
        with self.request('/api/mode', {'mode': 'silent', 'floating': False}) as response:
            self.assertEqual(response.status, 200)
        self.assertEqual(self.recorder.status()['settings']['explanation_mode'], 'silent')

    def test_pause_setting_persists(self):
        with self.request('/api/pause', {'paused': True}) as response:
            self.assertEqual(response.status, 200)
        self.assertTrue(self.recorder.paused)
        newer = Recorder(self.recorder.store.directory)
        try:
            self.assertTrue(newer.paused)
        finally:
            newer.store.close()

    def test_evidence_range_references_only_selected_snapshots_and_hashes(self):
        import io
        import hashlib
        import zipfile
        a = self.recorder.store.snapshots/'old.txt'
        b = self.recorder.store.snapshots/'current.txt'
        a.write_bytes(b'old-data')
        b.write_bytes(b'current-data')
        with patch('recorder.store.now', return_value='2026-10-01T00:00:00.000+00:00'):
            self.recorder.store.emit('file_version_captured', 'old.txt', snapshot='snapshots/old.txt')
        with patch('recorder.store.now', return_value='2026-11-16T10:00:00.000+00:00'):
            self.recorder.store.emit('file_version_captured', 'current.txt', snapshot='snapshots/current.txt')
            self.recorder.store.emit('monitoring_gap', cause='restart')
        url = '/api/handover?start=2026-11-16T00:00:00.000%2B00:00&end=2026-11-17T00:00:00.000%2B00:00&snapshots=1'
        with self.request(url) as response:
            payload = response.read()
        with zipfile.ZipFile(io.BytesIO(payload)) as archive:
            names = set(archive.namelist())
            self.assertIn('snapshots/current.txt', names)
            self.assertNotIn('snapshots/old.txt', names)
            self.assertIn(b'monitoring_gap', archive.read('gaps.jsonl'))
            manifest = json.loads(archive.read('manifest.json'))
            digest = hashlib.sha256(archive.read('snapshots/current.txt')).hexdigest()
            self.assertEqual(manifest['sha256_by_member']['snapshots/current.txt'], digest)

    def test_page_has_no_external_scripts(self):
        with self.request('/') as r:
            html=r.read().decode();self.assertIn("textContent=content",html);self.assertNotIn('<script src=',html)
            self.assertEqual(r.headers['Referrer-Policy'],'no-referrer')

class ExcelEventTests(Workspace):
    def make_sheet(self, value=100, formula=100, path=None):
        from types import SimpleNamespace as N
        area=N(Row=1,Column=1,Rows=N(Count=1),Columns=N(Count=1),Value2=value,Formula=formula)
        target=N(CountLarge=1,Areas=[area],Address='$A$1')
        wb=N(Path=str(self.root),FullName=str(path or self.root/'book.xlsb'))
        sh=N(Parent=wb,CodeName='Sheet1',Name='Orders',UsedRange=target)
        wb.Worksheets=[sh]
        return wb,sh,target,area
    def listener(self):
        return ExcelListener(self.store,{'excel_cache_cells':100,'excel_event_cells':10},lambda:[str(self.root)])
    def test_live_event_records_observed_old_new_and_unsaved_status(self):
        listener=self.listener();wb,sh,target,area=self.make_sheet()
        listener.opened(wb,'instance');area.Value2=10;area.Formula=10
        listener.handlers('instance')().OnSheetChange(sh,target)
        event=self.store.recent()[0]
        self.assertEqual(event['event'],'excel_range_changed')
        self.assertEqual(event['changes'][0]['old_value'],100)
        self.assertEqual(event['changes'][0]['new_value'],10)
        self.assertEqual(event['save_state'],'not_confirmed_saved')
    def test_outside_workbook_content_is_not_logged(self):
        listener=self.listener();wb,sh,target,area=self.make_sheet(path=self.base/'private.xlsx')
        listener.changed(sh,target,'instance');self.assertEqual(self.kinds(),[])
    def test_oversized_live_event_is_explicit_and_invalidates_old_cache(self):
        listener=self.listener();wb,sh,target,area=self.make_sheet();listener.prime(wb,'instance')
        target.CountLarge=100000;listener.changed(sh,target,'instance')
        e=self.store.recent()[0];self.assertFalse(e['details_captured']);self.assertTrue(e['old_baseline_invalidated'])
    def test_reopened_workbook_does_not_reuse_discarded_edit_baseline(self):
        listener=self.listener();wb,sh,target,area=self.make_sheet();listener.opened(wb,'instance')
        area.Value2=10;area.Formula=10;listener.changed(sh,target,'instance')
        area.Value2=100;area.Formula=100;listener.opened(wb,'instance')
        area.Value2=20;area.Formula=20;listener.changed(sh,target,'instance')
        self.assertEqual(self.store.recent()[0]['changes'][0]['old_value'],100)
    def test_save_request_and_result_have_separate_semantics(self):
        listener=self.listener();wb,sh,target,area=self.make_sheet();events=listener.handlers('instance')()
        events.OnWorkbookBeforeSave(wb,True,False);events.OnWorkbookAfterSave(wb,False)
        logs=self.store.recent();self.assertFalse(next(e for e in logs if e['event']=='excel_save_result')['success'])
        self.assertTrue(next(e for e in logs if e['event']=='excel_save_requested')['success_not_implied'])

if __name__=='__main__':unittest.main()
