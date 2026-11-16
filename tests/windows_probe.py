"""Explicit target-host probe; exit 2 is NOT_VERIFIED, not a passing substitute."""
import os
import sys
if os.name != 'nt':
    print('NOT_VERIFIED: run this probe on Windows with desktop Excel open.')
    sys.exit(2)
try:
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject('Excel.Application')
    print('Excel COM accessible:', app.Name, app.Version)
    print('Complete docs/WINDOWS-ACCEPTANCE.md to verify live events and sign-in startup.')
except Exception as e:
    print('NOT_VERIFIED:', e)
    sys.exit(2)
finally:
    try: pythoncom.CoUninitialize()
    except NameError: pass
