"""Optional Windows explanation entry window.

Only shown by the Recorder in Explanation Mode. All native controls are
created and focused on the window-owning UI thread, never from an HTTP worker.
"""
import os
import threading


class NotesWindow:
    FOCUS_MESSAGE_OFFSET = 37

    def __init__(self, store, on_close=None):
        self.store = store
        self.on_close = on_close
        self.hwnd = None
        self.thread = None

    def start(self):
        if os.name != 'nt':
            return
        import win32con
        import win32gui
        if self.thread and self.thread.is_alive():
            if self.hwnd and win32gui.IsWindow(self.hwnd):
                # The UI thread owns keyboard focus. Ask it to show and focus,
                # rather than trying SetFocus from the HTTP/server thread.
                win32gui.PostMessage(
                    self.hwnd, win32con.WM_APP + self.FOCUS_MESSAGE_OFFSET, 0, 0)
            return
        self.thread = threading.Thread(
            target=self._run, daemon=True, name='Floating notes')
        self.thread.start()

    def stop(self):
        if os.name != 'nt':
            return
        if self.hwnd:
            import win32con
            import win32gui
            if win32gui.IsWindow(self.hwnd):
                win32gui.ShowWindow(self.hwnd, win32con.SW_HIDE)

    def _run(self):
        import win32api
        import win32con
        import win32gui

        save_id = 1002
        edit_id = 1001
        controls = {}
        owner = self

        def focus_editor(hwnd):
            # SW_SHOWNOACTIVATE was previously used, which could leave the
            # window visible but without a usable keyboard input target.
            win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
            if controls.get('edit') and win32gui.IsWindow(controls['edit']):
                win32gui.SetFocus(controls['edit'])

        def create_controls(hwnd):
            # Build children *after* the parent CreateWindowEx returns.
            # The pywin32 class callback cannot be relied on to initialize
            # children during WM_CREATE on all supported Windows runtimes.
            controls['edit'] = win32gui.CreateWindowEx(
                win32con.WS_EX_CLIENTEDGE, 'EDIT', '',
                win32con.WS_CHILD | win32con.WS_VISIBLE |
                win32con.WS_TABSTOP | win32con.WS_VSCROLL |
                win32con.ES_LEFT | win32con.ES_MULTILINE |
                win32con.ES_AUTOVSCROLL | win32con.ES_WANTRETURN,
                12, 40, 325, 105, hwnd, edit_id, None, None)
            if not controls['edit']:
                raise RuntimeError('Floating explanation edit control could not be created')
            controls['save'] = win32gui.CreateWindowEx(
                0, 'BUTTON', 'Save explanation',
                win32con.WS_CHILD | win32con.WS_VISIBLE |
                win32con.WS_TABSTOP | win32con.BS_PUSHBUTTON,
                185, 156, 150, 29, hwnd, save_id, None, None)
            if not controls['save']:
                raise RuntimeError('Floating explanation save button could not be created')
            win32gui.CreateWindowEx(
                0, 'STATIC', 'Type what you did and why (optional)',
                win32con.WS_CHILD | win32con.WS_VISIBLE,
                12, 12, 320, 22, hwnd, 0, None, None)

        def handler(hwnd, message, wparam, lparam):
            if message == win32con.WM_APP + owner.FOCUS_MESSAGE_OFFSET:
                focus_editor(hwnd)
                return 0
            if message == win32con.WM_COMMAND and win32api.LOWORD(wparam) == save_id:
                value = win32gui.GetWindowText(controls['edit']).strip()
                if value:
                    try:
                        recent = owner.store.recent(limit=30)
                        context = next((
                            dict(event_id=e['id'], path=e.get('path', ''),
                                 event=e['event'])
                            for e in recent
                            if e.get('path') and e['event'] != 'user_explanation'
                        ), None)
                        owner.store.emit(
                            'user_explanation', text=value[:20000],
                            scope='case_note',
                            linked_event_id=context['event_id'] if context else '',
                            context_suggested=context,
                            context_confirmed=False,
                            entered_using='floating_window')
                        win32gui.SetWindowText(controls['edit'], '')
                    except Exception as exc:
                        win32gui.MessageBox(
                            hwnd, str(exc), 'Explanation not saved',
                            win32con.MB_ICONERROR)
                return 0
            if message == win32con.WM_CLOSE:
                if owner.on_close:
                    owner.on_close()
                win32gui.DestroyWindow(hwnd)
                return 0
            if message == win32con.WM_DESTROY:
                owner.hwnd = None
                win32gui.PostQuitMessage(0)
                return 0
            return win32gui.DefWindowProc(hwnd, message, wparam, lparam)

        import uuid
        class_name = 'FileChangeLogNotesWindow' + uuid.uuid4().hex
        try:
            wc = win32gui.WNDCLASS()
            wc.lpszClassName = class_name
            wc.lpfnWndProc = handler
            wc.hInstance = win32api.GetModuleHandle(None)
            win32gui.RegisterClass(wc)
            x = max(0, win32api.GetSystemMetrics(win32con.SM_CXSCREEN) - 380)
            y = max(0, win32api.GetSystemMetrics(win32con.SM_CYSCREEN) - 280)
            hwnd = win32gui.CreateWindowEx(
                win32con.WS_EX_TOPMOST | win32con.WS_EX_TOOLWINDOW,
                class_name, 'File Change Log - Explanation',
                win32con.WS_OVERLAPPED | win32con.WS_CAPTION |
                win32con.WS_SYSMENU | win32con.WS_MINIMIZEBOX,
                x, y, 360, 235, 0, 0, wc.hInstance, None)
            if not hwnd:
                raise RuntimeError('Floating explanation parent window could not be created')
            create_controls(hwnd)
            owner.hwnd = hwnd
            focus_editor(hwnd)
            win32gui.PumpMessages()
        except Exception as exc:
            # Failure must not silently leave a useless floating rectangle or
            # interrupt background file/excel monitoring.
            owner.store.emit('floating_notes_window_failed',
                             error_type=type(exc).__name__,
                             message=str(exc)[:300])
        finally:
            owner.hwnd = None
