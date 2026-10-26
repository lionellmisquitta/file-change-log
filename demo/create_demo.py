"""Create synthetic, disposable fixtures. Never use real vendor data for a demo."""
from pathlib import Path
import os
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import openpyxl

def main():
    parent = Path.home() / 'Documents'
    parent.mkdir(exist_ok=True)
    folder = parent / 'FileChangeLog Demo'
    index = 2
    while folder.exists():
        folder = parent / f'FileChangeLog Demo {index}'
        index += 1
    folder.mkdir()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Demo Orders'
    ws.append(['Order', 'Material', 'Quantity', 'Rate', 'Illustrative amount'])
    ws.append(['DEMO-001', 'SAMPLE-SYSTEM', 100, 2, '=C2*D2'])
    ws.append(['DEMO-002', 'SAMPLE-SEAT', 25, 1, '=C3*D3'])
    for col, width in [('A',18),('B',24),('C',15),('D',15),('E',24)]:
        ws.column_dimensions[col].width=width
    ws.freeze_panes='A2'
    wb.save(folder/'Demo Orders.xlsx')
    wb.close()
    (folder/'Demo Notes.txt').write_text('Synthetic demonstration only. Change quantity C2 from 100 to 10.\n',encoding='utf-8')
    print('Created synthetic demo files in:')
    print(folder)
    print('\nAdd this folder in the recorder, then open Demo Orders.xlsx in desktop Excel.')
    print('Wait for Excel attachment. Change C2 from 100 to 10 and save. Never use real vendor data for this demo.')

if __name__=='__main__':main()
