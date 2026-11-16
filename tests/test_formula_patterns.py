"""CHG-51: compact formula saved-state regressions.

Tests deliberately verify round-trip translation of each represented formula,
and guard against interpreting a pattern as a user fill command.
"""
import tempfile
import unittest
from pathlib import Path

import openpyxl
from openpyxl.formula.translate import Translator

from recorder.formula_patterns import pack_formula_runs
from recorder.compare import compare_files
from recorder.folders import FolderWorker
from recorder.store import Store


class FormulaPatternTests(unittest.TestCase):
    def test_translation_roundtrip_relative_absolute_and_mixed(self):
        cases = [
            ("=C2*0.2", "=C3*0.2"),
            ("=SUM($A2,B$1)", "=SUM($A3,B$1)"),
            ("=IF(C2>0,C2,0)", "=IF(C3>0,C3,0)"),
        ]
        for first, second in cases:
            with self.subTest(first=first):
                data = [
                    dict(sheet="Data",address="D2",new_value=first,old_value=None),
                    dict(sheet="Data",address="D3",new_value=second,old_value=None),
                    dict(sheet="Data",address="D4",
                         new_value=Translator(first,origin="D2").translate_formula("D4"), old_value=None)]
                patterns, removed = pack_formula_runs(data)
                self.assertEqual(len(patterns), 1)
                self.assertEqual(len(removed), 3)
                pattern = patterns[0]
                for row in range(pattern["first_row"], pattern["last_row"]+1):
                    reconstructed = Translator(pattern["anchor_formula"],
                                               origin=pattern["anchor_address"]).translate_formula("D"+str(row))
                    self.assertEqual(reconstructed, data[row-2]["new_value"])
                self.assertTrue(pattern["exact_prior_and_final_state_requires_snapshots"])

    def test_discontinuities_and_manual_overrides_not_compressed(self):
        changes = [
            dict(sheet="Finance",address="D2",new_value="=C2+1"),
            dict(sheet="Finance",address="D3",new_value="=C3+1"),
            dict(sheet="Finance",address="D4",new_value=120,old_value="=C4+1",
                 explanation_wanted=True),
            dict(sheet="Finance",address="D5",new_value="=C5+1"),
            dict(sheet="Finance",address="D6",new_value="=C6+1"),
        ]
        patterns, omitted = pack_formula_runs(changes)
        self.assertEqual(patterns, [])
        self.assertFalse(omitted)

    def test_saved_comparison_reduces_event_and_preserves_exceptions(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            before, after = base/"before.xlsx", base/"after.xlsx"
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.title="Finance"
            sheet["C1"]="Revenue"; sheet["D1"]="Margin"
            for row in range(2,502):
                sheet.cell(row,3,row*100)
            sheet["D3"]="=C3*0.3"  # prior formula that becomes a typed override
            wb.save(before)
            for row in range(2,502):
                sheet.cell(row,4,f"=C{row}*0.2")
            sheet["D3"]=999
            wb.save(after)
            wb.close()
            records=[]
            compare_files(before, after, lambda kind, path, **kw:records.append(
                dict(event=kind,**kw)), str(after), compact_formulas=True, max_changes=30)
            output=next(x for x in records if x["event"]=="snapshot_cell_comparison")
            self.assertEqual(output["total_changes"], 500)
            self.assertEqual(output["formula_changes_represented_by_patterns"], 498)
            self.assertEqual(len(output["formula_patterns"]), 1)
            self.assertEqual(output["formula_patterns"][0]["first_row"], 4)
            self.assertEqual(output["formula_patterns"][0]["last_row"], 501)
            self.assertEqual(output["unlisted_changes"], 0)
            override = next(x for x in output["changes"] if x["address"]=="D3")
            self.assertTrue(override["explanation_wanted"])
            self.assertEqual(override["old_value"],"=C3*0.3")
            self.assertEqual(override["new_value"],999)
            self.assertIn("D2",[x["address"] for x in output["changes"]])
            self.assertTrue(output["formula_patterns_require_snapshots_for_exact_reconstruction"])

    def test_real_folder_worker_emits_compaction_without_touching_workbook(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); watched=root/"watched";watched.mkdir()
            store=Store(root/"history")
            try:
                worker=FolderWorker(watched,store,{"settle_seconds":0,"snapshot_limit_mb":50})
                p=watched/"finance.xlsx"
                wb=openpyxl.Workbook()
                ws=wb.active
                for row in range(2,102):
                    ws.cell(row,3,row*10)
                wb.save(p)
                worker.scan_once();worker.scan_once()
                for row in range(2,102):
                    ws.cell(row,4,f"=C{row}*0.2")
                wb.save(p)
                prior=p.read_bytes()
                worker.scan_once();worker.scan_once()
                self.assertEqual(p.read_bytes(),prior)
                e=next(x for x in store.recent(200) if x["event"]=="snapshot_cell_comparison")
                self.assertEqual(e["total_changes"],100)
                self.assertEqual(e["formula_changes_represented_by_patterns"],100)
                self.assertEqual(e["unlisted_changes"],0)
                self.assertEqual(len(e["changes"]),0)
                self.assertEqual(len(e["formula_patterns"]),1)
                self.assertTrue(e["before_snapshot"])
                self.assertTrue(e["after_snapshot"])
                self.assertTrue(Path(e["before_snapshot"]).is_file())
                self.assertTrue(Path(e["after_snapshot"]).is_file())
            finally:
                store.close()


if __name__=="__main__":
    unittest.main()
