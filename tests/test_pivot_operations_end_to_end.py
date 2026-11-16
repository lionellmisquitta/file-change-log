"""Real spreadsheet-engine pivot-operation recording end-to-end QA.

Requires LibreOffice Calc + python3-uno, enabled only by RUN_REAL_PIVOT_OPERATIONS=1.
No user interface clicking. No Microsoft Excel COM is involved.
"""
import hashlib
import json
import os
import shutil
import socket
import subprocess
import tempfile
import time
import unittest
import zipfile
from pathlib import Path

import openpyxl

from recorder.folders import FolderWorker
from recorder.handover import build_handover
from recorder.store import Store
from recorder.workbook_logic import extract_workbook_logic


@unittest.skipUnless(os.environ.get("RUN_REAL_PIVOT_OPERATIONS") == "1",
                     "Real LibreOffice UNO integration is run by dedicated pivot-operation CI")
class PivotOperationsE2ETests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.base = Path(self.tmp.name)
        self.source = self.base / "source.xlsx"
        self.stage_dir = self.base / "engine-generated"
        self.stage_dir.mkdir()
        self.watch = self.base / "watched"
        self.watch.mkdir()
        self.recorder_file = self.watch / "Pivot_Operations.xlsx"
        self.store = Store(self.base / "history")
        self.worker = FolderWorker(self.watch, self.store,
                                   {"settle_seconds": 0, "snapshot_limit_mb": 50,
                                    "comparison_cells": 200000})
        self.office = None
        self._source_fixture()
        self._run_spreadsheet_engine()

    def tearDown(self):
        if self.office is not None:
            self.office.terminate()
            try:
                self.office.wait(timeout=8)
            except subprocess.TimeoutExpired:
                self.office.kill()
                self.office.wait(timeout=3)
        self.store.close()
        self.tmp.cleanup()

    def _source_fixture(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "SourceData"
        ws.append(("Region", "Product", "Month", "Revenue"))
        total = 0
        for month_index, month in enumerate(("Jan", "Feb", "Mar", "Apr")):
            for reg_index, region in enumerate(("North", "South", "East", "West")):
                for product_index, product in enumerate(("Alpha", "Beta", "Gamma")):
                    revenue = 1000 + reg_index * 30 + product_index * 15 + month_index * 7
                    ws.append((region, product, month, revenue))
                    total += revenue
        self.expected_sum = total
        self.expected_count = 48
        wb.save(self.source)
        wb.close()

    def _run_spreadsheet_engine(self):
        if not shutil.which("soffice") or not Path("/usr/bin/python3").exists():
            self.fail("LibreOffice UNO runtime absent; cannot claim pivot operation execution")
        with socket.socket() as listener:
            listener.bind(("127.0.0.1", 0))
            port = listener.getsockname()[1]
        profile = (self.base / "lo-profile").as_uri()
        self.office = subprocess.Popen([
            "soffice", "-env:UserInstallation=" + profile,
            "--headless",
            "--accept=socket,host=localhost,port=%d;urp;StarOffice.ServiceManager" % port,
            "--norestore", "--nodefault", "--nofirststartwizard"
        ], stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
        for _ in range(100):
            if self.office.poll() is not None:
                self.fail("LibreOffice exited before UNO connection was ready")
            try:
                with socket.create_connection(("127.0.0.1", port), timeout=0.3):
                    break
            except OSError:
                time.sleep(0.15)
        else:
            self.fail("LibreOffice UNO socket did not become available")
        runner = Path(__file__).with_name("uno_pivot_operations.py")
        command = ["/usr/bin/python3", str(runner), str(self.source),
                   str(self.stage_dir), str(port)]
        executed = subprocess.run(command, capture_output=True, text=True, timeout=90)
        if executed.returncode != 0:
            self.fail("UNO generation failed:\n%s\n%s" % (executed.stdout, executed.stderr))
        self.assertEqual(len(list(self.stage_dir.glob("*.xlsx"))), 5)
        print("REAL_ENGINE_OPERATIONS " + executed.stdout.strip(), flush=True)

    def _record(self, data):
        # Each version goes through the actual FolderWorker saved-file
        # observation path: the same monitored path is replaced five times.
        self.recorder_file.write_bytes(data)
        self.worker.scan_once()
        self.worker.scan_once()

    def test_real_pivot_operations_and_recorded_state_deltas(self):
        raw = self.source.read_bytes()
        self._record(raw)
        expected = [
            ("01_pivot_created_sum.xlsx", "first_observed", "sum", 0, 0),
            ("02_aggregation_changed_count.xlsx", "definition_changed", "count", 0, 0),
            ("03_product_column_added.xlsx", "definition_changed", "count", 1, 0),
            ("04_month_page_filter_added.xlsx", "definition_changed", "count", 1, 1),
            ("05_pivot_deleted.xlsx", "no_longer_present", None, 0, 0),
        ]
        for name, expected_action, expected_aggregate, expected_columns, expected_page in expected:
            stage = self.stage_dir / name
            self.assertTrue(stage.is_file(), "UNO stage output missing " + name)
            with zipfile.ZipFile(stage) as archive:
                pivot_parts = [item for item in archive.namelist()
                               if item.startswith("xl/pivotTables/pivotTable")
                               and item.endswith(".xml")]
                self.assertEqual(len(pivot_parts), 1 if expected_aggregate else 0)
            state = extract_workbook_logic(stage)
            if expected_aggregate is not None:
                self.assertEqual(len(state["pivots"]), 1)
                pivot = state["pivots"][0]
                self.assertEqual(pivot["name"], "RevenuePivot")
                self.assertEqual(pivot["rows"], ["Region"])
                self.assertEqual(pivot["values"][0]["aggregation"], expected_aggregate)
                self.assertEqual(len(pivot["columns"]), expected_columns)
                self.assertEqual(len(pivot["page_filters"]), expected_page)
                self.assertEqual(pivot["source"]["type"], "worksheet")
            else:
                self.assertEqual(state["pivots"], [])
            if name.startswith("01"):
                workbook = openpyxl.load_workbook(stage, read_only=True, data_only=True)
                try:
                    rows = list(workbook["PivotTest"].values)
                    totals = [cell for row in rows if row[0] == "Total Result"
                              for cell in row[1:] if isinstance(cell, (int, float))]
                    self.assertEqual(totals, [self.expected_sum])
                finally:
                    workbook.close()
            if name.startswith("02"):
                workbook = openpyxl.load_workbook(stage, read_only=True, data_only=True)
                try:
                    rows = list(workbook["PivotTest"].values)
                    totals = [cell for row in rows if row[0] == "Total Result"
                              for cell in row[1:] if isinstance(cell, (int, float))]
                    self.assertEqual(totals, [self.expected_count])
                finally:
                    workbook.close()
            self._record(stage.read_bytes())
            current = self.store.recent(limit=1000)
            deltas = [x for x in current
                      if x["event"] == "workbook_logic_pivot_state_changes"]
            self.assertTrue(deltas, "Recorded saved-version pivot delta absent for " + name)
            self.assertEqual(deltas[0]["pivot_changes"][0]["observation"], expected_action)
            self.assertEqual(len(deltas[0]["pivot_changes"]), 1)
        all_events = self.store.recent(limit=1000)
        versions = [e for e in all_events if e["event"] == "file_version_captured"]
        self.assertEqual(len(versions), 6)
        self.assertEqual(len([e for e in all_events
                              if e["event"] == "workbook_logic_observed"]), 6)
        self.assertEqual(len([e for e in all_events
                              if e["event"] == "workbook_logic_pivot_state_changes"]), 5)
        for event in versions:
            snapshot = self.store.directory / event["snapshot"]
            self.assertTrue(snapshot.exists())
            # All retained versions are content-addressed SHA-256 bytes.
            self.assertEqual(snapshot.stem, hashlib.sha256(snapshot.read_bytes()).hexdigest())
        archive = build_handover(self.store, include_snapshots=True)
        try:
            with zipfile.ZipFile(archive) as z:
                manifest = json.loads(z.read("manifest.json"))
                for path, sha in manifest["sha256_by_member"].items():
                    self.assertEqual(hashlib.sha256(z.read(path)).hexdigest(), sha)
                print("PIVOT_OPERATION_RECORDING_SUMMARY " + json.dumps({
                    "engine": "libreoffice_uno",
                    "monitored_versions": len(versions),
                    "pivot_state_deltas": 5,
                    "state_transitions": [x[1] for x in expected],
                    "original_and_all_five_versions_sha256_verified": True,
                    "export_sha256_members": len(manifest["sha256_by_member"]),
                    "excel_com_events_exercised": False,
                }, sort_keys=True), flush=True)
        finally:
            archive.unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
