"""CHG-58 regressions: Silent Mode, window focus and explanation controls."""
import json
import os
import time
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import patch

from recorder.app import Recorder, validate_config
from recorder.notes_window import NotesWindow


class FakeStore:
    def __init__(self):
        self.events = []
    def emit(self, event, **detail):
        self.events.append((event, detail))
    def recent(self, limit=30):
        return [{"id": "E1", "event": "excel_range_changed",
                 "path": r"C:\\Temp\\Example.xlsx"}]


class FakeWindow:
    def __init__(self):
        self.started = 0
        self.stopped = 0
    def start(self):
        self.started += 1
    def stop(self):
        self.stopped += 1


class ExplanationModeTests(unittest.TestCase):
    def test_legacy_config_disables_floating_when_silent(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg = validate_config(
                {"explanation_mode": "silent", "floating_notes": True},
                Path(tmp))
            self.assertEqual(cfg["explanation_mode"], "silent")
            self.assertFalse(cfg["floating_notes"])
            explanation = validate_config(
                {"explanation_mode": "explanation", "floating_notes": True},
                Path(tmp))
            self.assertTrue(explanation["floating_notes"])

    def test_api_mode_transition_clears_floating_preference(self):
        with tempfile.TemporaryDirectory() as tmp:
            recorder = Recorder.__new__(Recorder)
            recorder.config = {"explanation_mode": "explanation",
                               "floating_notes": True}
            recorder.config_path = Path(tmp) / "settings.json"
            recorder.lock = threading.RLock()
            recorder.notes_window = FakeWindow()
            recorder.store = FakeStore()
            recorder.set_mode("silent", True)
            self.assertFalse(recorder.config["floating_notes"])
            self.assertEqual(recorder.notes_window.stopped, 1)
            saved = json.loads(recorder.config_path.read_text(encoding="utf-8"))
            self.assertFalse(saved["floating_notes"])
            self.assertEqual(saved["explanation_mode"], "silent")
            self.assertEqual(recorder.store.events[-1][1]["floating_notes"], False)
            recorder.set_mode("explanation", True)
            self.assertEqual(recorder.notes_window.started, 1)
            self.assertTrue(recorder.config["floating_notes"])

    def test_invalid_mode_not_saved(self):
        with tempfile.TemporaryDirectory() as tmp:
            recorder = Recorder.__new__(Recorder)
            recorder.config = {"explanation_mode": "silent", "floating_notes": False}
            recorder.config_path = Path(tmp) / "settings.json"
            recorder.lock = threading.RLock()
            recorder.notes_window = FakeWindow()
            recorder.store = FakeStore()
            with self.assertRaises(ValueError):
                recorder.set_mode("unknown", True)
            self.assertFalse(recorder.config_path.exists())
            self.assertEqual(recorder.notes_window.started, 0)

    def test_dashboard_hides_and_disables_option_in_silent_mode(self):
        markup = (Path(__file__).resolve().parents[1] / "web" /
                  "index.html").read_text(encoding="utf-8")
        self.assertIn('id="floatingOption"', markup)
        self.assertIn("el('floatingOption').hidden=settings.explanation_mode!=='explanation'", markup)
        self.assertIn("el('floating').disabled=settings.explanation_mode!=='explanation'", markup)
        self.assertIn("if(mode==='silent')el('floating').checked=false", markup)

    def test_native_window_uses_activating_show_and_own_thread_focus(self):
        import inspect
        source = inspect.getsource(NotesWindow)
        self.assertIn("win32con.SW_SHOWNORMAL", source)
        self.assertIn("win32gui.SetFocus(controls['edit'])", source)
        self.assertIn("win32gui.PostMessage(", source)
        self.assertNotIn("SW_SHOWNOACTIVATE", source.replace(
            "# SW_SHOWNOACTIVATE was previously used", ""))
        self.assertIn("'floating_notes_window_failed'", source)



@unittest.skipUnless(os.name == "nt", "Windows desktop Win32 UI smoke test")
class NativeNotesWindowSmoke(unittest.TestCase):
    def test_editor_control_created_accepts_text_and_saves(self):
        import win32con
        import win32gui
        store = FakeStore()
        window = NotesWindow(store)
        try:
            window.start()
            deadline = time.monotonic() + 8
            while time.monotonic() < deadline and not window.hwnd:
                time.sleep(.05)
            self.assertTrue(window.hwnd, "Win32 window failed to start: %r" %
                            store.events)
            hwnd = window.hwnd
            edit = win32gui.GetDlgItem(hwnd, 1001)
            self.assertTrue(edit and win32gui.IsWindow(edit),
                            "Editable child control was not created")
            win32gui.SendMessage(edit, win32con.WM_SETTEXT, 0,
                                 "Automated explanation")
            self.assertEqual(win32gui.GetWindowText(edit),
                             "Automated explanation")
            win32gui.SendMessage(hwnd, win32con.WM_COMMAND, 1002, 0)
            explained = [details for kind, details in store.events
                         if kind == "user_explanation"]
            self.assertEqual(len(explained), 1)
            self.assertEqual(explained[0]["text"], "Automated explanation")
            self.assertEqual(win32gui.GetWindowText(edit), "")
        finally:
            if window.hwnd and win32gui.IsWindow(window.hwnd):
                win32gui.PostMessage(window.hwnd, win32con.WM_CLOSE, 0, 0)
            if window.thread:
                window.thread.join(timeout=4)


if __name__ == "__main__":
    unittest.main()
