"""CHG-52: saved-byte lineage and evidence-linked prompt priority."""
from pathlib import Path
import tempfile
import unittest
from recorder.folders import FolderWorker
from recorder.common import canonical
from recorder.store import Store
from recorder.timeline import timeline


class ExactContentLineageTests(unittest.TestCase):
    def test_identical_arrival_has_evidence_link_not_unproven_copy_command(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            watched = root / "watch"
            watched.mkdir()
            store = Store(root / "history")
            try:
                worker = FolderWorker(watched, store, {"settle_seconds": 0, "snapshot_limit_mb": 50})
                original = watched / "original.csv"
                original.write_text("id,value\nA,100\n", encoding="utf-8")
                worker.scan_once();worker.scan_once()
                duplicate = watched / "new_copy.csv"
                duplicate.write_bytes(original.read_bytes())
                worker.scan_once();worker.scan_once()
                events = store.recent(100)
                e = next(x for x in events if x["event"] == "file_identical_content_observed")
                self.assertEqual(e["path"], canonical(duplicate))
                self.assertEqual(e["matching_paths"], [canonical(original)])
                self.assertEqual(e["observation"], "same_saved_bytes_not_proof_of_copy_operation")
                captures = [x for x in events if x["event"] == "file_version_captured"]
                self.assertIn(e["capture_event_id"], [x["id"] for x in captures])
                self.assertEqual(e["snapshot"], next(x for x in captures if x["path"] == canonical(original))["snapshot"])
                t = timeline(store, path_filter="new_copy.csv")
                linked = next(x for x in t["entries"] if x["event"] == "file_identical_content_observed")
                self.assertEqual(linked["details"]["capture_event_id"], e["capture_event_id"])
            finally:
                store.close()

    def test_new_different_content_never_gets_lineage(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            watched = root / "watch"
            watched.mkdir()
            store = Store(root / "history")
            try:
                worker = FolderWorker(watched, store, {"settle_seconds": 0, "snapshot_limit_mb": 50})
                (watched/"a.csv").write_text("alpha")
                worker.scan_once();worker.scan_once()
                (watched/"b.csv").write_text("beta")
                worker.scan_once();worker.scan_once()
                self.assertFalse(any(x["event"] == "file_identical_content_observed"
                                     for x in store.recent(100)))
            finally:
                store.close()

    def test_prompt_selector_is_optional_and_priority_ordered(self):
        html = (Path(__file__).resolve().parents[1] / "web" / "index.html").read_text(encoding='utf-8')
        self.assertIn("e.event==='excel_range_changed'&&e.explanation_wanted===true", html)
        self.assertIn("e.event==='snapshot_cell_comparison'&&e.explanation_wanted===true", html)
        self.assertIn("e.event==='excel_sheet_delete_requested'", html)
        self.assertIn("settings.explanation_mode==='explanation'", html)
        self.assertIn("candidate.id!==dismissedPromptId", html)


if __name__ == "__main__":
    unittest.main()
