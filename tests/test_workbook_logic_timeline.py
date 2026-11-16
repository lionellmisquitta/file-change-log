"""CHG-50 independently reproducible workbook-logic and timeline assertions."""
import json
import tempfile
import unittest
import zipfile
from pathlib import Path

import openpyxl

from recorder.workbook_logic import extract_workbook_logic, compare_workbook_logic
from recorder.timeline import timeline
from recorder.store import Store
from recorder.folders import FolderWorker


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"


def add_pivot_fixture(source, dest, name="RevenuePivot", aggregation="sum", filter_value="EU"):
    """Inject a minimal standards-shaped OOXML pivot *metadata fixture*.

    It is deliberately not advertised as an Excel-created or fully functioning
    pivot workbook. Real Windows Excel roundtrip remains a separate acceptance gate.
    """
    with zipfile.ZipFile(source) as inp:
        files = {x.filename: inp.read(x) for x in inp.infolist()}
    sheet = "xl/worksheets/sheet1.xml"
    original = files[sheet].decode()
    original = original.replace("</worksheet>", '<autoFilter ref="A1:C4"><filterColumn colId="0">'
        '<filters><filter val="'+filter_value+'"/></filters></filterColumn></autoFilter></worksheet>')
    files[sheet] = original.encode()
    relpath = "xl/worksheets/_rels/sheet1.xml.rels"
    files[relpath] = ('<?xml version="1.0" encoding="UTF-8"?>'
       '<Relationships xmlns="'+PKG+'">'
       '<Relationship Id="rPivot" Type="'+REL+'/pivotTable" Target="../pivotTables/pivotTable1.xml"/>'
       '</Relationships>').encode()
    files["xl/pivotTables/pivotTable1.xml"] = ('<?xml version="1.0"?>'
        '<pivotTableDefinition xmlns="'+MAIN+'" name="'+name+'" cacheId="1">'
        '<location ref="E1:F4" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>'
        '<pivotFields count="3"><pivotField/><pivotField/><pivotField/></pivotFields>'
        '<rowFields count="1"><field x="0"/></rowFields>'
        '<dataFields count="1"><dataField name="Total revenue" fld="2" subtotal="'+aggregation+'"/></dataFields>'
        '</pivotTableDefinition>').encode()
    files["xl/pivotTables/_rels/pivotTable1.xml.rels"] = ('<?xml version="1.0"?>'
       '<Relationships xmlns="'+PKG+'">'
       '<Relationship Id="rCache" Type="'+REL+'/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/>'
       '</Relationships>').encode()
    files["xl/pivotCache/pivotCacheDefinition1.xml"] = ('<?xml version="1.0"?>'
       '<pivotCacheDefinition xmlns="'+MAIN+'">'
       '<cacheSource type="worksheet"><worksheetSource sheet="Data" ref="A1:C4"/></cacheSource>'
       '<cacheFields count="3"><cacheField name="Region"/><cacheField name="Month"/>'
       '<cacheField name="Revenue"/></cacheFields></pivotCacheDefinition>').encode()
    with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as out:
        for path, value in files.items():
            out.writestr(path, value)


class WorkbookLogicTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.baseline = self.root / "baseline.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        for row in [["Region","Month","Revenue"],["EU","Jan",100],
                    ["US","Jan",120],["EU","Feb",140]]:
            ws.append(row)
        wb.save(self.baseline)
        wb.close()

    def tearDown(self):
        self.tmp.cleanup()

    def test_empty_workbook_has_no_invented_pivots(self):
        state = extract_workbook_logic(self.baseline)
        self.assertEqual(state["pivots"], [])
        self.assertEqual(state["worksheet_filters"], [])
        self.assertFalse(state["reconstruction_verified"])

    def test_snapshot_filter_and_pivot_definition(self):
        target = self.root / "with-pivot.xlsx"
        add_pivot_fixture(self.baseline, target)
        state = extract_workbook_logic(target)
        self.assertEqual(len(state["pivots"]), 1)
        pivot = state["pivots"][0]
        self.assertEqual(pivot["name"], "RevenuePivot")
        self.assertEqual(pivot["rows"], ["Region"])
        self.assertEqual(pivot["values"][0]["aggregation"], "sum")
        self.assertEqual(pivot["source"]["reference"], "A1:C4")
        self.assertEqual(state["worksheet_filters"][0]["criteria"][0]["values"], ["EU"])
        self.assertFalse(state["reconstruction_verified"])
        delta = compare_workbook_logic(extract_workbook_logic(self.baseline), state)
        self.assertEqual(delta["pivot_changes"][0]["observation"], "first_observed")

    def test_existing_pivot_configuration_change_is_not_called_creation(self):
        one = self.root / "one.xlsx"
        two = self.root / "two.xlsx"
        add_pivot_fixture(self.baseline, one, aggregation="sum")
        add_pivot_fixture(self.baseline, two, aggregation="average")
        delta = compare_workbook_logic(extract_workbook_logic(one), extract_workbook_logic(two))
        self.assertEqual(delta["pivot_changes"][0]["observation"], "definition_changed")

    def test_bad_relationship_does_not_escape_package(self):
        original = self.root / "one.xlsx"
        changed = self.root / "bad.xlsx"
        add_pivot_fixture(self.baseline, original)
        with zipfile.ZipFile(original) as inp:
            files = {f.filename: inp.read(f) for f in inp.infolist()}
        files["xl/worksheets/_rels/sheet1.xml.rels"] = files[
            "xl/worksheets/_rels/sheet1.xml.rels"].replace(b"../pivotTables/pivotTable1.xml",
                                                           b"../../../../outside.xml")
        with zipfile.ZipFile(changed, "w") as out:
            for name, value in files.items():
                out.writestr(name, value)
        with self.assertRaises(ValueError):
            extract_workbook_logic(changed)

    def test_folder_worker_emits_pivot_definition_and_state_delta(self):
        watched = self.root / "watched"
        watched.mkdir()
        store = Store(self.root/"history")
        worker = FolderWorker(watched, store, {"settle_seconds":0,"snapshot_limit_mb":50})
        try:
            item = watched / "case.xlsx"
            add_pivot_fixture(self.baseline, item, aggregation="sum")
            worker.scan_once()
            worker.scan_once()
            self.assertTrue(any(e["event"] == "workbook_logic_observed" for e in store.recent(100)))
            add_pivot_fixture(self.baseline, item, aggregation="average")
            worker.scan_once()
            worker.scan_once()
            rows = store.recent(200)
            delta = next(e for e in rows if e["event"] == "workbook_logic_pivot_state_changes")
            self.assertEqual(delta["pivot_changes"][0]["observation"], "definition_changed")
        finally:
            store.close()


class TimelineTests(unittest.TestCase):
    def test_observation_order_linkage_and_gap_semantics(self):
        with tempfile.TemporaryDirectory() as root:
            store = Store(Path(root) / "history")
            try:
                x = store.emit("excel_workbook_open", "C:/Finance.xlsx")
                y = store.emit("excel_pivot_updated", "C:/Finance.xlsx", sheet="Pivot")
                z = store.emit("snapshot_skipped_budget", "C:/Finance.xlsx")
                store.emit("user_explanation", linked_event_id=y["id"], text="revised")
                store.emit("excel_workbook_open", "C:/Other.xlsx")
                summary = timeline(store, limit=3, path_filter="Finance.xlsx")
                self.assertFalse(summary["complete_history"])
                self.assertEqual(summary["recent_window_matched_events"], 3)
                self.assertEqual([e["event_id"] for e in summary["entries"]], [x["id"],y["id"],z["id"]])
                self.assertEqual(summary["entries"][-1]["kind"], "gap")
                self.assertTrue(all(e["confidence"] == "observed_event_not_proof_of_user_intent"
                                    for e in summary["entries"]))
            finally:
                store.close()


class LivePivotMetadataTests(unittest.TestCase):
    def test_com_pivot_update_records_state_and_not_creation(self):
        from types import SimpleNamespace as N
        from recorder.excel import ExcelListener
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            store = Store(root / "history")
            try:
                workbook = N(Path=str(root), FullName=str(root/"finance.xlsx"))
                sheet = N(Parent=workbook, Name="Pivot", CodeName="Sheet1")
                def fields():
                    return [N(Orientation=1, Name="Region"),
                            N(Orientation=2, Name="Month"),
                            N(Orientation=3, Name="Currency"),
                            N(Orientation=4, Name="Revenue")]
                pivot = N(Name="Sales", SourceData="Data!A1:C100",
                          PivotFields=fields,
                          DataFields=[N(Name="Total Revenue", SourceName="Revenue", Function=-4157)])
                listener = ExcelListener(store, {}, lambda: [str(root)])
                listener.handlers("excel1")().OnSheetPivotTableUpdate(sheet, pivot)
                events = store.recent(30)
                p = next(x for x in events if x["event"] == "excel_pivot_definition_observed")
                self.assertEqual(p["row_fields"], ["Region"])
                self.assertEqual(p["column_fields"], ["Month"])
                self.assertEqual(p["page_fields"], ["Currency"])
                self.assertEqual(p["data_fields"][0]["function_code"], -4157)
                self.assertEqual(p["observation"], "pivot_updated_not_creation_proof")
                self.assertEqual(p["completeness"], "partial_com_observation")
            finally:
                store.close()

    def test_com_failure_does_not_fabricate_definition(self):
        from types import SimpleNamespace as N
        from recorder.excel import ExcelListener
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            store = Store(root / "history")
            try:
                workbook = N(Path=str(root), FullName=str(root/"finance.xlsx"))
                sheet = N(Parent=workbook, Name="Pivot", CodeName="Sheet1")
                listener = ExcelListener(store, {}, lambda: [str(root)])
                listener.handlers("excel1")().OnSheetPivotTableUpdate(sheet, object())
                kinds = [e["event"] for e in store.recent(10)]
                self.assertIn("excel_pivot_updated", kinds)
                self.assertIn("excel_pivot_definition_unavailable", kinds)
                self.assertNotIn("excel_pivot_definition_observed", kinds)
            finally:
                store.close()


if __name__ == "__main__":
    unittest.main()
