#!/usr/bin/env python3
"""Regenerate website_data.json for the spare-parts dashboard from the workbook.

Reads  workbook/All_Machines_Spareparts_Final.xlsm  (repo root) and writes
website_data.json (repo root) with the exact schema the dashboard expects.

Run from the repo root:  python scripts/build_data.py
Dependencies: openpyxl (pandas is installed by the workflow but not needed here).

Logic is a faithful port of the original export_data.py that built the first
website_data.json: same machine list, same grouping (by New SAP, else by
lower-cased name + vendor), same most-common value picks, same status rules,
same sort order, same compact usage rows. Only the file paths and the
`generated` date (today instead of a hard-coded date) differ.
"""
import json
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
WORKBOOK_DIR = ROOT / 'workbook'
EXACT_NAME = 'All_Machines_Spareparts_Final.xlsm'
OUT = ROOT / 'website_data.json'


def find_workbook():
    """Pick the workbook to read: the exact expected name if present, else the
    single .xlsm/.xlsx in workbook/ (uploads from a phone often arrive renamed,
    e.g. with spaces or a -1 suffix). Returns None when nothing usable exists."""
    exact = WORKBOOK_DIR / EXACT_NAME
    if exact.exists():
        return exact
    cands = sorted(WORKBOOK_DIR.glob('*.xlsm')) + sorted(WORKBOOK_DIR.glob('*.xlsx'))
    if len(cands) == 1:
        return cands[0]
    return None

MACHINES = ['Mdo1', 'Mdo2', 'PE01', 'PE02', 'Mdo3', 'Vista1', 'Vista2', 'Mira1', 'Mira2', 'Novo',
            'Compound', 'Debagger', 'INK', 'Plastic System', 'Renzmann', 'NGR', 'KC01', 'KC02',
            'KC03', 'KC04', 'KH01', 'Bimec S01', 'Ashe S02', 'Kampf S03', 'Bimec S04', 'Kampf S06',
            'Elastic EC1', 'Elastic EC2', 'MTA Chiller', 'Consumables']
VH = {'W&H Code': 'W&H', 'Compound Code': 'Compound', 'P.S Code': 'P.S', 'NGR Code': 'NGR',
      'Lemo Code': 'Lemo', 'Kampf Code': 'Kampf'}
# Extra sheets read by the export (besides the 30 machine sheets).
EXTRA_SHEETS = ['Stock_Analysis', 'Data_Integrity_Check', 'Sync Preview',
                'Change Log', 'Watch_List', 'SearchHistory']


def nc(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s if s else None


def nn(v):
    if v is None or v == '':
        return None
    try:
        f = float(v)
        return int(f) if f.is_integer() else f
    except Exception:
        return None


def clean(s):
    if s is None:
        return None
    s = re.sub(r'\s+', ' ', str(s)).strip()
    return s or None


def dashboard_meta(wb):
    """Last Refresh / Last Sync info from the workbook's Dashboard sheet.

    Each label sits two columns left of its value (G1 'Last Refresh' -> I1
    '11/21/2026  12:44:24 PM'; A10 'Last Sync Time' -> C10 datetime;
    A11 'Last Sync Result' -> C11 'No Changes'; A12 'Last Run Changes' -> C12 0).
    Returns {} when the sheet or the labels are missing.
    """
    if 'Dashboard' not in wb.sheetnames:
        return {}
    ws = wb['Dashboard']
    labels = {'last_refresh': 'Last Refresh', 'last_sync_time': 'Last Sync Time',
              'last_sync_result': 'Last Sync Result', 'last_run_changes': 'Last Run Changes'}
    found = {}
    for row in ws.iter_rows(min_row=1, max_row=20):
        for c in row:
            if isinstance(c.value, str):
                for key, lab in labels.items():
                    if c.value.strip() == lab and key not in found:
                        found[key] = ws.cell(row=c.row, column=c.column + 2).value

    def iso(v):
        if v is None:
            return None
        if hasattr(v, 'isoformat'):
            return v.isoformat()
        s = re.sub(r'\s+', ' ', str(v)).strip()
        for fmt in ('%m/%d/%Y %I:%M:%S %p', '%m/%d/%Y %H:%M:%S',
                    '%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
            try:
                return datetime.strptime(s, fmt).isoformat()
            except Exception:
                pass
        return s

    out = {}
    if 'last_refresh' in found:
        out['last_refresh'] = iso(found['last_refresh'])
    if 'last_sync_time' in found:
        out['last_sync_time'] = iso(found['last_sync_time'])
    if 'last_sync_result' in found:
        out['last_sync_result'] = str(found['last_sync_result']).strip()
    if 'last_run_changes' in found:
        out['last_run_changes'] = nn(found['last_run_changes'])
    return out


def main():
    src = find_workbook()
    if src is None:
        print('ERROR: no workbook found in the workbook/ folder.', file=sys.stderr)
        print('Upload All_Machines_Spareparts_Final.xlsm to the workbook/ folder.',
              file=sys.stderr)
        return 1
    print('reading workbook:', src.name)

    wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
    missing = [s for s in list(MACHINES) + EXTRA_SHEETS if s not in wb.sheetnames]
    if missing:
        print(f'ERROR: workbook is missing these sheets: {", ".join(missing)}',
              file=sys.stderr)
        print('The uploaded file must be the full spare-parts workbook.',
              file=sys.stderr)
        return 1

    rows = []
    for mi, m in enumerate(MACHINES):
        ws = wb[m]
        rws = list(ws.iter_rows(values_only=True))
        header = [(str(h).strip() if h is not None else None) for h in rws[0]]
        idx = {h: i for i, h in enumerate(header) if h}
        for r in rws[1:]:
            if not any(v is not None and str(v).strip() != '' for v in r):
                continue
            vendor = vtype = None
            for h, t in VH.items():
                if h in idx and nc(r[idx[h]]):
                    vendor, vtype = nc(r[idx[h]]), t
            name = clean(r[idx['Spare Parts']]) if 'Spare Parts' in idx else None
            if not name:
                code = nc(r[idx['New SAP']]) if 'New SAP' in idx else None
                if not code:
                    code = nc(r[idx['SAP Code']]) if 'SAP Code' in idx else None
                if not code:
                    code = vendor
                if not code:
                    continue
                name = f"بدون اسم ({code})"
            rows.append({'m': mi, 'name': name, 'vendor': vendor, 'vtype': vtype,
                         'sap': nc(r[idx['SAP Code']]) if 'SAP Code' in idx else None,
                         'newsap': nc(r[idx['New SAP']]) if 'New SAP' in idx else None,
                         'pcs': nn(r[idx['PCS']]) if 'PCS' in idx else None,
                         'stock': nn(r[idx['Stock']]) if 'Stock' in idx else None,
                         'ext': nn(r[idx['EXT Stock']]) if 'EXT Stock' in idx else None,
                         'succ': nc(r[idx['Successor']]) if 'Successor' in idx else None,
                         'pr': nc(r[idx['PR.No']]) if 'PR.No' in idx else None,
                         'label': clean(r[idx['Label']]) if 'Label' in idx else None,
                         'cls': clean(r[idx['Class']]) if 'Class' in idx else None})

    # ---- Stock_Analysis: min stock by code (any code type) ----
    min_by_code = {}
    sa_rows = []
    ws = wb['Stock_Analysis']
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if r[0] is None:
            continue
        code = nc(r[0])
        sa_rows.append({'code': code, 'name': clean(r[1]), 'pcs': nn(r[2]),
                        'stock': nn(r[3]), 'min': nn(r[4])})
        if nn(r[4]) is not None and code not in min_by_code:
            min_by_code[code] = nn(r[4])

    def status(stock, mn):
        if mn is None:
            return 'No Min'
        if stock in (None, 0):
            return 'Critical'
        if stock <= mn:
            return 'Below Min'
        if stock <= mn * 1.2:
            return 'Reorder Soon'
        return 'Healthy'

    for x in sa_rows:
        x['status'] = status(x['stock'], x['min'])

    # ---- Master catalog ----
    groups = defaultdict(list)
    for r in rows:
        key = ('SAP', r['newsap']) if r['newsap'] else ('NAME', r['name'].lower(), r['vendor'])
        groups[key].append(r)

    def most_common(vals):
        vals = [v for v in vals if v not in (None, '')]
        if not vals:
            return None
        return Counter(vals).most_common(1)[0][0]

    catalog = []
    key_to_idx = {}
    for key, g in groups.items():
        newsap = key[1] if key[0] == 'SAP' else None
        name = most_common([x['name'] for x in g])
        sap = most_common([x['sap'] for x in g])
        vendor = most_common([x['vendor'] for x in g])
        vtype = most_common([x['vtype'] for x in g])
        stock = most_common([x['stock'] for x in g])
        ext = most_common([x['ext'] for x in g])
        total_pcs = sum(x['pcs'] or 0 for x in g)
        machines = sorted({x['m'] for x in g})
        succ = most_common([x['succ'] for x in g])
        prs = sorted({x['pr'] for x in g if x['pr']})
        labels = sorted({x['label'] for x in g if x['label']})
        # min stock: by newsap, then vendor, then sap
        mn = None
        for c in (newsap, vendor, sap):
            if c and c in min_by_code:
                mn = min_by_code[c]
                break
        st = status(stock, mn)
        idx = len(catalog)
        key_to_idx[key] = idx
        catalog.append({
            'name': name, 'newsap': newsap, 'sap': sap, 'vendor': vendor, 'vtype': vtype,
            'pcs': total_pcs, 'stock': stock, 'ext': ext, 'min': mn, 'status': st,
            'machines': machines, 'n_machines': len(machines), 'rows': len(g),
            'succ': succ, 'prs': prs, 'labels': labels,
        })
    # sort catalog: status severity, then name
    sev = {'Critical': 0, 'Below Min': 1, 'Reorder Soon': 2, 'No Min': 3, 'Healthy': 4}
    catalog_sorted = sorted(catalog, key=lambda c: (sev[c['status']], (c['name'] or '').lower()))
    # rebuild index map after sort
    old_to_new = {key_to_idx[k]: i for i, k in enumerate(
        sorted(groups.keys(), key=lambda k: (sev[catalog[key_to_idx[k]]['status']],
                                             (catalog[key_to_idx[k]]['name'] or '').lower())))}
    catalog = catalog_sorted

    # ---- Usage rows (compact): [catIdx, machineIdx, pcs, stock, ext, name?, vendor?, sap?, succ?, pr?] ----
    usage = []
    name_mismatch = 0
    for r in rows:
        key = ('SAP', r['newsap']) if r['newsap'] else ('NAME', r['name'].lower(), r['vendor'])
        ci = old_to_new[key_to_idx[key]]
        cat = catalog[ci]
        u = [ci, r['m'], r['pcs'], r['stock'], r['ext']]
        nm = r['name'] if r['name'] != cat['name'] else None
        vn = r['vendor'] if r['vendor'] != cat['vendor'] else None
        sp = r['sap'] if r['sap'] != cat['sap'] else None
        sc = r['succ'] if r['succ'] != cat['succ'] else None
        pr = r['pr'] if r['pr'] not in (None, *cat['prs']) else None
        if nm:
            name_mismatch += 1
        u += [nm, vn, sp, sc, pr]
        usage.append(u)

    # ---- Integrity (full, compact) ----
    ws = wb['Data_Integrity_Check']
    integrity = []
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if r[0] is None:
            continue
        integrity.append([r[0], r[1], r[2], nc(r[3]), r[4]])
    # ---- Sync preview ----
    ws = wb['Sync Preview']
    sync_prev = []
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if r[0] is None:
            continue
        sync_prev.append([nc(r[0]), r[1], r[2], r[3], r[4], str(r[5]) if r[5] is not None else None,
                          str(r[6]) if r[6] is not None else None, r[7], r[8]])
    # ---- Change log recent 300 ----
    ws = wb['Change Log']
    cl = [r for r in list(ws.iter_rows(values_only=True))[1:] if r[0] is not None]
    cl_sorted = sorted(cl, key=lambda r: r[0] if hasattr(r[0], 'year') else datetime(1900, 1, 1),
                       reverse=True)[:300]
    changelog = [[r[0].isoformat() if hasattr(r[0], 'isoformat') else str(r[0]), r[1], r[2], r[3],
                  nc(r[4]), r[5], r[6], r[7]] for r in cl_sorted]
    # ---- Watch list ----
    ws = wb['Watch_List']
    watch = []
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if r[0] is None:
            continue
        watch.append([nc(r[0]), r[1].isoformat() if hasattr(r[1], 'isoformat') else str(r[1]),
                      clean(r[2]), nn(r[3]), nn(r[4]), r[5], r[6]])
    # ---- Search history ----
    ws = wb['SearchHistory']
    search_hist = [[r[0], r[1].isoformat() if hasattr(r[1], 'isoformat') else str(r[1])]
                   for r in list(ws.iter_rows(values_only=True))[1:] if r[0] is not None]

    data = {
        'generated': date.today().isoformat(),
        'dashboard_meta': dashboard_meta(wb),
        'machines': MACHINES,
        'catalog': catalog,
        'usage': usage,
        'stock_analysis': sa_rows,
        'integrity': integrity,
        'sync_preview': sync_prev,
        'changelog_recent': changelog,
        'watchlist': watch,
        'search_history': search_hist,
    }
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
    print('catalog entries:', len(catalog))
    print('  with newsap:', sum(1 for c in catalog if c['newsap']))
    print('  without newsap (name+vendor key):', sum(1 for c in catalog if not c['newsap']))
    print('usage rows:', len(usage), 'name overrides:', name_mismatch)
    print('status counts:', dict(Counter(c['status'] for c in catalog)))
    print('catalog stock=0:', sum(1 for c in catalog if c['stock'] == 0))
    print('catalog stock blank:', sum(1 for c in catalog if c['stock'] is None))
    print('catalog with min:', sum(1 for c in catalog if c['min'] is not None))
    print('catalog with PR:', sum(1 for c in catalog if c['prs']))
    print('SA rows:', len(sa_rows), 'integrity:', len(integrity), 'sync:', len(sync_prev),
          'changelog:', len(changelog))
    print('website_data.json size MB:', round(os.path.getsize(OUT) / 1e6, 2))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
