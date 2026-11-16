"""Independent public-data stress checks for Phase 1 recorder.

These tests download reproducible raw public CSV datasets on the GitHub CI runner.
The datasets are not invented and the recorder runs against actual downloaded bytes.
They do not simulate Excel's Windows COM adapter or establish user action intent.
"""
import csv
import hashlib
import io
import json
import tempfile
import time
import unittest
import urllib.request
import zipfile
from collections import Counter
from pathlib import Path

import openpyxl

from recorder.compare import compare_files
from recorder.folders import FolderWorker
from recorder.handover import build_handover
from recorder.store import Store

PUBLIC_DATASETS = {
    "flights": ("https://raw.githubusercontent.com/mwaskom/seaborn-data/master/flights.csv", 140),
    "titanic": ("https://raw.githubusercontent.com/mwaskom/seaborn-data/master/titanic.csv", 850),
    "taxis": ("https://raw.githubusercontent.com/mwaskom/seaborn-data/master/taxis.csv", 6000),
    "diamonds": ("https://raw.githubusercontent.com/mwaskom/seaborn-data/master/diamonds.csv", 50000),
}
SHA256 = {}


def download_dataset(name):
    url, minimum = PUBLIC_DATASETS[name]
    error = None
    for attempt in range(3):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers={
                    "User-Agent": "FileChangeLog-QA-Public-Dataset/1.0"}), timeout=35) as response:
                data = response.read(15 * 1024 * 1024)
            text = data.decode("utf-8-sig")
            rows = list(csv.DictReader(io.StringIO(text)))
            if len(rows) < minimum:
                raise AssertionError("%s had %d rows, expected >= %d" % (name, len(rows), minimum))
            SHA256[name] = hashlib.sha256(data).hexdigest()
            return data, rows
        except Exception as exc:
            error = exc
            time.sleep(attempt + 1)
    raise AssertionError("Public dataset download failed %s: %s" % (url, error))


class PublicDatasetStress(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.datasets = {}
        for name in PUBLIC_DATASETS:
            cls.datasets[name] = download_dataset(name)
        print("DATASET_INPUTS " + json.dumps({
            name: {"url": PUBLIC_DATASETS[name][0], "rows": len(cls.datasets[name][1]),
                   "bytes": len(cls.datasets[name][0]), "sha256": SHA256[name]}
            for name in PUBLIC_DATASETS}, sort_keys=True), flush=True)

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.home = Path(self.tmp.name)
        self.root = self.home / "source" / "FP&A" / "Revenue" / "2026" / "November"
        self.root.mkdir(parents=True)
        self.store = Store(self.home / "event-history")
        self.settings = {
            "settle_seconds": 0,
            "scan_seconds": 1,
            "snapshot_limit_mb": 50,
            "snapshot_budget_mb": 500,
            "baseline_budget_mb": 250,
            "comparison_cells": 200000,
            "comparison_details": 5000,
        }
        self.worker = FolderWorker(self.home / "source", self.store, self.settings)

    def tearDown(self):
        self.store.close()
        self.tmp.cleanup()

    def scan(self):
        self.worker.scan_once()
        self.worker.scan_once()

    def events(self, kind=None):
        result = self.store.recent(limit=1000)
        return [x for x in result if x["event"] == kind] if kind else result

    def test_four_real_datasets_baseline_copy_and_restart(self):
        for name, (data, _) in self.datasets.items():
            (self.root / (name + ".csv")).write_bytes(data)
        start = time.monotonic()
        self.scan()
        baseline = self.events("file_version_captured")
        self.assertEqual(len(baseline), 4)
        self.assertTrue(all(x["classification"] == "baseline" for x in baseline))
        self.assertEqual(len([x for x in baseline if x.get("snapshot")]), 4)
        source = self.root / "titanic.csv"
        duplicate = self.root / "copies" / "titanic.csv"
        duplicate.parent.mkdir()
        before_bytes = self.store.snapshot_bytes
        duplicate.write_bytes(source.read_bytes())
        self.scan()
        self.assertEqual(self.store.snapshot_bytes, before_bytes)
        restarted = FolderWorker(self.home / "source", self.store, self.settings,
                                 reconcile_on_first_scan=True)
        restarted.scan_once()
        restarted.scan_once()
        self.assertEqual(len(self.events("file_version_captured")), 5)
        print("BASELINE_AND_DEDUPE " + json.dumps({
            "dataset_count": 4, "snapshots_after_duplicate": self.store.snapshot_bytes,
            "elapsed_seconds": round(time.monotonic() - start, 3),
            "versions": len(self.events("file_version_captured"))}), flush=True)

    def test_real_csv_revision_and_metadata_only_large_diff(self):
        raw, rows = self.datasets["taxis"]
        csv_file = self.root / "taxis.csv"
        csv_file.write_bytes(raw)
        self.scan()
        revised = raw + b"\n"
        csv_file.write_bytes(revised)
        self.scan()
        self.assertEqual(len(self.events("file_version_captured")), 2)
        self.assertTrue(any(x["event"] == "snapshot_text_comparison" or
                            x["event"] == "comparison_skipped" for x in self.events()))
        print("CSV_REVISION " + json.dumps({
            "source_rows": len(rows), "bytes_initial": len(raw),
            "bytes_revised": len(revised), "comparison_kinds":
                sorted({x["event"] for x in self.events()
                        if "comparison" in x["event"]})}), flush=True)

    def test_actual_diamonds_6500_row_excel_formula_fill_and_exception(self):
        rows = self.datasets["diamonds"][1][:6500]
        before = self.root / "diamonds_before.xlsx"
        after = self.root / "diamonds_after.xlsx"
        book = openpyxl.Workbook()
        ws = book.active
        ws.title = "Diamonds"
        ws.append(["cut", "carat", "price", "simulated_margin_formula", "owner_override"])
        for row in rows:
            ws.append([row["cut"], float(row["carat"]), int(row["price"]), None, None])
        book.save(before)
        for ix in range(2, len(rows) + 2):
            ws.cell(ix, 4, "=C%d*0.2" % ix)
        ws["E3"] = "MANUAL REVIEW"
        book.save(after)
        book.close()
        captured = []
        started = time.monotonic()
        compare_files(before, after,
                      lambda kind, path, **values: captured.append(dict(event=kind, **values)),
                      str(after), max_cells=200000, max_changes=5000)
        event = next(x for x in captured if x["event"] == "snapshot_cell_comparison")
        self.assertEqual(event["total_changes"], 6501)
        self.assertEqual(event["unlisted_changes"], 1501)
        self.assertEqual(event["changes_per_sheet"]["Diamonds"], 6501)
        self.assertEqual(event["changes_per_sheet_column"]["Diamonds"]["D"], 6500)
        self.assertEqual(event["changes_per_sheet_column"]["Diamonds"]["E"], 1)
        self.assertEqual(event["changes"][0]["address"], "E3")
        self.assertIn("Diamonds", event["affected_sheets"])
        print("LARGE_FORMULA_FILL " + json.dumps({
            "rows": len(rows), "changed_total": event["total_changes"],
            "listed": len(event["changes"]), "omitted": event["unlisted_changes"],
            "first_listed": event["changes"][0]["address"],
            "elapsed_seconds": round(time.monotonic()-started,3),
            "event_json_bytes": len(json.dumps(event, default=str).encode())}), flush=True)

    def test_real_finance_timeseries_filter_only_and_integrity_export(self):
        rows = self.datasets["flights"][1]
        work = self.root / "seasonal_forecast.xlsx"
        book = openpyxl.Workbook()
        ws = book.active
        ws.title = "Monthly"
        ws.append(["year", "month", "passengers", "cost_ratio"])
        for row in rows:
            ws.append([int(row["year"]), row["month"], int(row["passengers"]), None])
        book.save(work)
        self.scan()
        ws.auto_filter.ref = "A1:D145"
        book.save(work)
        book.close()
        self.scan()
        kinds = {x["event"] for x in self.events()}
        self.assertIn("no_cell_differences_other_content_changed", kinds)
        evidence = build_handover(self.store, include_snapshots=True)
        try:
            with zipfile.ZipFile(evidence) as archive:
                names = archive.namelist()
                manifest = json.loads(archive.read("manifest.json"))
                checks = manifest["sha256_by_member"]
                for name, digest in checks.items():
                    self.assertEqual(hashlib.sha256(archive.read(name)).hexdigest(), digest)
                self.assertGreaterEqual(manifest["counts"]["snapshots"], 2)
                self.assertIn("gaps.jsonl", names)
                print("FILTER_AND_EXPORT " + json.dumps({
                    "flight_rows": len(rows), "snapshots": manifest["counts"]["snapshots"],
                    "hashed_members": len(checks),
                    "no_cell_difference_event": True,
                    "zip_bytes": evidence.stat().st_size}), flush=True)
        finally:
            evidence.unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
