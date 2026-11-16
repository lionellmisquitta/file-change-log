"""Independent QA: TWO genuinely public, previously-unused data sources.

(1) Real pre-existing Excel pivot/slicer workbook, not synthetic pivot OOXML.
(2) Public supermarket transaction CSV, converted into test workbooks.

No production app code is changed. The test captures native source bytes, SHA256,
input sizes, raw OOXML pivot counts, replay boundaries and recorder event evidence.
"""
import csv
import hashlib
import io
import json
import tempfile
import unittest
import urllib.request
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl

from recorder.compare import workbook_features
from recorder.folders import FolderWorker
from recorder.store import Store
from recorder.workbook_logic import extract_workbook_logic, compare_workbook_logic
from recorder.handover import build_handover
from recorder.timeline import timeline

SOURCES = {
    "real_excel_pivots": "https://raw.githubusercontent.com/danawoodruff/Excel/main/Pivots.xlsx",
    "supermarket_transactions": "https://raw.githubusercontent.com/selva86/datasets/master/supermarket_sales.csv",
}
SS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def downloaded(url, minimum, maximum=20_000_000):
    req = urllib.request.Request(url, headers={"User-Agent": "FileChangeLog-Real-Pivot-Adversarial-QA/1.0"})
    with urllib.request.urlopen(req, timeout=50) as stream:
        value = stream.read(maximum + 1)
    if not minimum <= len(value) <= maximum:
        raise AssertionError("Unexpected public fixture byte size: %d" % len(value))
    return value


def raw_pivot_xml(archive):
    names = archive.namelist()
    return [x for x in names if x.lower().startswith("xl/pivottables/")
            and x.lower().endswith(".xml") and "/_rels/" not in x.lower()]


def prepare_retail_csv(rows, workbook, with_formulas=False, with_filter=False):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Transactions"
    ws.append(["Branch", "City", "Product line", "Quantity", "Total", "cogs",
               "Gross income", "Manual adjustment", "Calculated variance"])
    for i, row in enumerate(rows, 2):
        ws.append([row["Branch"], row["City"], row["Product line"],
                   int(row["Quantity"]), float(row["Total"]), float(row["cogs"]),
                   float(row["gross income"]), None, None])
        if with_formulas:
            ws.cell(i, 9, "=E%d-F%d" % (i,i))
    if with_formulas:
        ws["H3"] = "FINANCE REVIEW"
        ws["I4"] = 123.45  # deliberately interrupts the repeated formula pattern
    if with_filter:
        ws.auto_filter.ref = "A1:I%d" % (len(rows)+1)
        ws.auto_filter.add_filter_column(0, ["A"])
    wb.save(workbook)
    wb.close()


def patch_pivot_state(source, target, pivot_part):
    """Change only one existing pivot's saved definition in a derived ZIP copy.
    This is intentionally an OOXML mutation, NOT a simulated user action in Excel.
    """
    with zipfile.ZipFile(source) as archive:
        members = {info.filename: archive.read(info) for info in archive.infolist()}
    root = ET.fromstring(members[pivot_part])
    changed = False
    for elem in root.findall(".//"+SS+"dataField"):
        existing = elem.get("subtotal", "sum")
        elem.set("subtotal", "count" if existing != "count" else "sum")
        changed = True
        break
    if not changed:
        # Change a pivot definition attribute that is a saved-state observation;
        # no claims about actual UI replay are made.
        root.set("showHeaders", "0" if root.get("showHeaders", "1") != "0" else "1")
        changed = True
    members[pivot_part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in members.items():
            archive.writestr(name, data)
    return "pivot_saved_definition_modified"


class RealPublicPivotStress(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.raw = {name: downloaded(url, 1000) for name, url in SOURCES.items()}
        cls.retail_rows = list(csv.DictReader(io.StringIO(
            cls.raw["supermarket_transactions"].decode("utf-8-sig"))))
        if len(cls.retail_rows) < 900:
            raise AssertionError("Public retail dataset unexpectedly small")
        print("PUBLIC_INPUTS " + json.dumps({
            name: {"url": SOURCES[name], "bytes": len(data),
                   "sha256": hashlib.sha256(data).hexdigest()}
            for name, data in cls.raw.items()}, sort_keys=True), flush=True)
        print("SUPERMARKET_ROW_COUNT %d" % len(cls.retail_rows), flush=True)

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.watch = self.root/"watch"/"Finance"/"Nov"
        self.watch.mkdir(parents=True)
        self.store = Store(self.root/"history")
        self.worker = FolderWorker(self.root/"watch", self.store,
                                   {"settle_seconds":0, "snapshot_limit_mb":50,
                                    "comparison_cells":200000})
        self.pivots = self.watch/"real_public_pivots.xlsx"
        self.pivots.write_bytes(self.raw["real_excel_pivots"])

    def tearDown(self):
        self.store.close()
        self.tmp.cleanup()

    def scan(self):
        self.worker.scan_once(); self.worker.scan_once()

    def test_real_excel_pivots_extract_all_actual_xml_definitions(self):
        with zipfile.ZipFile(self.pivots) as archive:
            names = raw_pivot_xml(archive)
            self.assertGreater(len(names), 0,
                               "The source is not a genuine persisted-pivot workbook")
            native_names = [ET.fromstring(archive.read(part)).get("name") for part in names]
            has_slicers = any(x.lower().startswith("xl/slicers/") or
                              x.lower().startswith("xl/slicercaches/") for x in archive.namelist())
        state = extract_workbook_logic(self.pivots)
        flags = workbook_features(self.pivots)
        print("REAL_PIVOT_DIAGNOSTIC " + json.dumps({
            "pivot_xml_count":len(names), "raw_names":native_names,
            "extracted_count":len(state["pivots"]),
            "extracted_names":[p["name"] for p in state["pivots"]],
            "warnings":state["warnings"],
            "completeness":state["completeness"],
            "slicers_present":has_slicers,
            "flags":flags,
        }, sort_keys=True), flush=True)
        self.assertTrue(flags["has_pivot_tables"])
        self.assertEqual(len(state["pivots"]), len(names),
                         "Genuine Excel pivot definitions were silently missed")
        self.assertCountEqual([p["name"] for p in state["pivots"]], native_names)
        self.assertFalse(state["reconstruction_verified"])
        if has_slicers:
            self.assertIn("slicers_not_fully_reconstructible", state["warnings"])
            self.assertEqual(state["completeness"],"partial_or_unsupported")
        for pivot in state["pivots"]:
            self.assertTrue(pivot["name"])
            self.assertIn("rows", pivot)
            self.assertIn("columns", pivot)
            self.assertIn("values", pivot)

    def test_pivot_metadata_change_detected_without_invented_creation(self):
        with zipfile.ZipFile(self.pivots) as z:
            parts = raw_pivot_xml(z)
        self.assertTrue(parts)
        modified = self.watch/"pivot_edited.xlsx"
        patch_pivot_state(self.pivots, modified, parts[0])
        a = extract_workbook_logic(self.pivots)
        b = extract_workbook_logic(modified)
        delta = compare_workbook_logic(a,b)
        print("REAL_PIVOT_STATE_DELTA "+json.dumps({
            "changes":len(delta["pivot_changes"]),
            "observations":[x["observation"] for x in delta["pivot_changes"]],
            "before_pivots":len(a["pivots"]),
            "after_pivots":len(b["pivots"]),
        }, sort_keys=True), flush=True)
        self.assertGreaterEqual(len(delta["pivot_changes"]),1)
        self.assertTrue(all(x["observation"]=="definition_changed" for x in delta["pivot_changes"]))

    def test_recorder_captures_genuine_pivot_workbook_and_export_integrity(self):
        pristine = self.pivots.read_bytes()
        self.scan()
        self.assertEqual(self.pivots.read_bytes(), pristine)
        events = self.store.recent(200)
        versions = [e for e in events if e["event"] == "file_version_captured"
                    and e["path"].lower().endswith("real_public_pivots.xlsx")]
        self.assertEqual(len(versions),1)
        self.assertTrue(versions[0]["snapshot"])
        logic = [e for e in events if e["event"] == "workbook_logic_observed"
                 and e["path"].lower().endswith("real_public_pivots.xlsx")]
        self.assertEqual(len(logic),1)
        self.assertGreater(len(logic[0]["extraction"]["pivots"]),0)
        trail = timeline(self.store,path_filter="real_public_pivots.xlsx")
        self.assertTrue(any(e["event"]=="workbook_logic_observed" for e in trail["entries"]))
        export = build_handover(self.store, include_snapshots=True)
        try:
            with zipfile.ZipFile(export) as zipped:
                manifest = json.loads(zipped.read("manifest.json"))
                for name,digest in manifest["sha256_by_member"].items():
                    self.assertEqual(hashlib.sha256(zipped.read(name)).hexdigest(), digest)
                print("REAL_PIVOT_EXPORT "+json.dumps({
                    "snapshots":manifest["counts"]["snapshots"],
                    "verified_member_hashes":len(manifest["sha256_by_member"]),
                    "evidence_zip_bytes":export.stat().st_size,
                }, sort_keys=True), flush=True)
        finally:
            export.unlink(missing_ok=True)

    def test_public_retail_formula_patterns_override_and_filters(self):
        before, after = self.watch/"retail_before.xlsx", self.watch/"retail_after.xlsx"
        rows = self.retail_rows
        prepare_retail_csv(rows,before)
        prepare_retail_csv(rows,after,with_formulas=True,with_filter=True)
        captured=[]
        from recorder.compare import compare_files
        compare_files(before,after,
                      lambda kind,path,**kw:captured.append(dict(event=kind,**kw)),
                      str(after),max_changes=30,compact_formulas=True)
        result = next(x for x in captured if x["event"]=="snapshot_cell_comparison")
        state = extract_workbook_logic(after)
        self.assertEqual(result["total_changes"],len(rows)+1)
        self.assertGreater(result["formula_changes_represented_by_patterns"],len(rows)-10)
        self.assertEqual(result["total_changes"],
                         result["formula_changes_represented_by_patterns"]+
                         len(result["changes"])+result["unlisted_changes"])
        self.assertEqual(result["unlisted_changes"],0)
        self.assertEqual(state["worksheet_filters"][0]["criteria"][0]["values"],["A"])
        print("RETAIL_FORMULA_FILTER "+json.dumps({
            "rows":len(rows), "total_changes":result["total_changes"],
            "formula_pattern_cells":result["formula_changes_represented_by_patterns"],
            "individual_change_details":len(result["changes"]),
            "unlisted":result["unlisted_changes"],
            "filter":state["worksheet_filters"][0]["criteria"][0]["values"],
            "formula_patterns":len(result["formula_patterns"]),
        }, sort_keys=True), flush=True)

    def test_retail_revised_workbook_scan_and_manual_exception_survives(self):
        workbook=self.watch/"retail.xlsx"
        prepare_retail_csv(self.retail_rows,workbook)
        self.scan()
        prepare_retail_csv(self.retail_rows,workbook,with_formulas=True,with_filter=True)
        self.scan()
        events = self.store.recent(200)
        comparisons = [e for e in events if e["event"]=="snapshot_cell_comparison"
                       and e["path"].lower().endswith("retail.xlsx")]
        self.assertEqual(len(comparisons),1)
        result=comparisons[0]
        self.assertEqual(result["total_changes"],len(self.retail_rows)+1)
        self.assertEqual(result["unlisted_changes"],0)
        details = {d["address"]:d for d in result["changes"]}
        self.assertEqual(details["H3"]["new_value"],"FINANCE REVIEW")
        self.assertEqual(details["I4"]["new_value"],123.45)
        events_by_type=Counter(e["event"] for e in events)
        print("RETAIL_SNAPSHOT_CHECK "+json.dumps({
            "changes":result["total_changes"],
            "formula_patterns":len(result["formula_patterns"]),
            "manual_cells":sorted(set(details)&{"H3","I4"}),
            "snapshot_versions":events_by_type["file_version_captured"],
            "pivot_logic_events":events_by_type["workbook_logic_observed"],
            "gap_events":events_by_type["snapshot_skipped_budget"],
        }, sort_keys=True), flush=True)


if __name__ == "__main__":
    unittest.main()
