"""Optional lightweight always-on-top Windows explanation entry window."""
import os
import threading

class NotesWindow:
    def __init__(self, store, on_close=None):
        self.store = store
        self.on_close = on_close
        self.hwnd = None
        self.thread = None
        self._closing_programmatically = False

    def start(self):
        if os.name != 'nt' or (self.thread and self.thread.is_alive()):
            return
        self.thread = threading.Thread(target=self._run, daemon=True, name='Floating notes')
        self.thread.start()

    def stop(self):
        if self.hwnd:
            self._closing_programmatically = True
            import win32con
            import win32gui
            win32gui.PostMessage(self.hwnd, win32con.WM_CLOSE, 0, 0)

    def _run(self):
        import win32api
        import win32con
        import win32gui
        from .common import now
        save_id = 1002
        edit_id = 1001
        controls = {}
        owner = self
        def handler(hwnd, message, wparam, lparam):
            if message == win32con.WM_CREATE:
                controls['edit'] = win32gui.CreateWindowEx(
                    win32con.WS_EX_CLIENTEDGE, 'EDIT', '',
                    win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.WS_VSCROLL |
                    win32con.ES_LEFT | win32con.ES_MULTILINE | win32con.ES_AUTOVSCROLL,
                    12, 40, 325, 105, hwnd, edit_id, None, None)
                controls['save'] = win32gui.CreateWindowEx(
                    0, 'BUTTON', 'Save explanation',
                    win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.BS_PUSHBUTTON,
                    185, 156, 150, 29, hwnd, save_id, None, None)
                win32gui.CreateWindowEx(0, 'STATIC', 'Type what you did and why (optional)',
                                        win32con.WS_CHILD | win32con.WS_VISIBLE,
                                        12, 12, 320, 22, hwnd, 0, None, None)
                return 0
            if message == win32con.WM_COMMAND and win32api.LOWORD(wparam) == save_id:
                text = win32gui.GetWindowText(controls['edit']).strip()
                if text:
                    try:
                        recent = owner.store.recent(limit=30)
                        context = next((dict(event_id=e['id'], path=e.get('path', ''),
                                             event=e['event'])
                                        for e in recent if e.get('path') and e['event'] != 'user_explanation'), None)
                        owner.store.emit('user_explanation', text=text[:20000],
                                         scope='case_note', linked_event_id=context['event_id'] if context else '',
                                         context_suggested=context, context_confirmed=False,
                                         entered_using='floating_window')
                        win32gui.SetWindowText(controls['edit'], '')
                    except Exception as exc:
                        win32gui.MessageBox(hwnd, str(exc), 'Explanation not saved', win32con.MB_ICONERROR)
                return 0
            if message == win32con.WM_CLOSE:
                if not owner._closing_programmatically and owner.on_close:
                    owner.on_close()
                win32gui.DestroyWindow(hwnd)
                return 0
            if message == win32con.WM_DESTROY:
                owner.hwnd = None
                win32gui.PostQuitMessage(0)
                return 0
            return win32gui.DefWindowProc(hwnd, message, wparam, lparam)
        class_name = 'FileChangeLogNotesWindow'
        wc = win32gui.WNDCLASS()
        wc.lpszClassName = class_name
        wc.lpfnWndProc = handler
        wc.hInstance = win32api.GetModuleHandle(None)
        try:
            win32gui.RegisterClass(wc)
        except Exception:
            # Another instance of the class in this process may have registered it.
            pass
        x = max(0, win32api.GetSystemMetrics(win32con.SM_CXSCREEN) - 380)
        y = max(0, win32api.GetSystemMetrics(win32con.SM_CYSCREEN) - 280)
        hwnd = win32gui.CreateWindowEx(
            win32con.WS_EX_TOPMOST | win32con.WS_EX_TOOLWINDOW,
            class_name, 'File Change Log - Explanation',
            win32con.WS_OVERLAPPED | win32con.WS_CAPTION |
            win32con.WS_SYSMENU | win32con.WS_MINIMIZEBOX,
            x, y, 360, 235, 0, 0, wc.hInstance, None)
        self.hwnd = hwnd
        win32gui.ShowWindow(hwnd, win32con.SW_SHOWNOACTIVATE)
        win32gui.PumpMessages()
