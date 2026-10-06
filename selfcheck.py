# -*- coding: utf-8 -*-
"""Onellenorzes: Excel COM-mal kitolti probaszamokkal, ellenorzi a kepleteket,
majd MENTES NELKUL bezarja. A leadott fajl ures marad."""
import os
import win32com.client as win32

PATH = os.path.abspath('Megterulesi-kimutatas-idegenrendeszet.xlsx')

xl = win32.DispatchEx('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(PATH)
    d, s = wb.Worksheets('Adatok'), wb.Worksheets('Összegzés')

    # --- probaadat: 2 ugytipus
    # 1. sor (5): 10 fo, 1 eves ciklus, teny 10 db, 100 000 Ft/ugy
    d.Cells(5, 2).Value, d.Cells(5, 3).Value = 10, 1
    d.Cells(5, 5).Value, d.Cells(5, 6).Value = 10, 100000
    # 2. sor (6): 9 fo, 3 eves ciklus -> 3 ugy/ev, teny 3 db, 50 000 Ft/ugy
    d.Cells(6, 2).Value, d.Cells(6, 3).Value = 9, 3
    d.Cells(6, 5).Value, d.Cells(6, 6).Value = 3, 50000

    # --- berkoltseg: 500 000 brutto x 12 x 1,13 + 0 = 6 780 000
    s.Cells(5, 2).Value = 500000     # brutto
    s.Cells(6, 2).Value = 12         # honap
    s.Cells(7, 2).Value = 0.13       # szocho
    s.Cells(8, 2).Value = 0          # egyeb
    s.Cells(25, 2).Value = 0.10      # berindex
    s.Cells(26, 2).Value = 0.10      # dijindex
    xl.CalculateFullRebuild()

    got = {
        'evesitett_ugyszam': d.Cells(20, 4).Value,
        'elkerult_evesitett': d.Cells(20, 7).Value,
        'elkerult_teny': d.Cells(20, 8).Value,
        'munkaltatoi_koltseg': s.Cells(9, 2).Value,
        'netto_evesitett': s.Cells(14, 2).Value,
        'fedezet_pct': s.Cells(16, 2).Value,
        'netto_teny': s.Cells(22, 2).Value,
        '3ev_netto': s.Cells(32, 4).Value,
    }
    exp = {
        'evesitett_ugyszam': 13.0,              # 10 + 9/3
        'elkerult_evesitett': 1150000.0,        # 10*100k + 3*50k
        'elkerult_teny': 1150000.0,             # 10*100k + 3*50k
        'munkaltatoi_koltseg': 6780000.0,
        'netto_evesitett': 1150000.0 - 6780000.0,
        'fedezet_pct': 1150000.0 / 6780000.0,
        'netto_teny': 1150000.0 - 6780000.0,
        # 3 ev, mindket index 10%: (1,15M - 6,78M) * (1 + 1,1 + 1,21)
        '3ev_netto': (1150000.0 - 6780000.0) * (1 + 1.1 + 1.21),
    }
    for k, e in exp.items():
        g = got[k]
        assert g is not None, k + ' = None (keplet hiba)'
        assert abs(float(g) - e) < 1.0, '{0}: kapott {1}, vart {2}'.format(k, g, e)
        print('OK  {0:22s} {1}'.format(k, g))
    print('\nMinden keplet helyes. A fajl mentes nelkul bezarva - ures marad.')
finally:
    wb.Close(SaveChanges=False)
    xl.Quit()
