#!/usr/bin/env python3
"""Regenerate website_data.json for the spare-parts dashboard from the workbook.

Reads  workbook/All_Machines_Spareparts_Final.xlsm  (repo root) and writes
website_data.json (repo root) with the exact schema the dashboard expects.
Also generates changelog.json (repo root) for the dashboard's "Update log"
tab: one entry per run (newest first), diffing the final catalog against the
previous website_data.json.

Run from the repo root:  python scripts/build_data.py
Dependencies: openpyxl (pandas is installed by the workflow but not needed here).

Logic is a faithful port of the original export_data.py that built the first
website_data.json: same machine list, same grouping (by New SAP, else by
lower-cased name + vendor), same most-common value picks, same status rules,
same sort order, same compact usage rows. Only the file paths and the
`generated` date (today instead of a hard-coded date) differ.
"""
import csv
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
CSV_DIR = ROOT / 'csv-exports'


def _s(v):
    """Render a value for CSV: None -> '', otherwise str()."""
    return '' if v is None else str(v)


def write_csvs(catalog, sa_rows):
    """Write the three CSV exports the dashboard's download buttons produce,
    so the team can download them as files (the in-page download is blocked
    by the browser sandbox). Files are UTF-8 with BOM so Arabic opens
    correctly in Excel.

    - stock-alerts.csv: the dashboard's Stock Alerts tab, exactly. Rule:
      stock == 0, or status Critical / Below Min, or stock unrecorded with
      PCS > 0 (unrecorded stock counted as 0, per Adora's 2026-11-24 choice,
      option A). Sorted by PCS descending (stable, so ties keep catalog
      order, the same order the dashboard shows). Unrecorded stock is
      written as 0.
    - catalog.csv: the full unified catalog, in catalog order.
    - stock-analysis.csv: the Stock_Analysis sheet as-is, in sheet order.
    """
    CSV_DIR.mkdir(exist_ok=True)

    alerts = [c for c in catalog
              if c['stock'] == 0 or c['status'] in ('Critical', 'Below Min')
              or (c['stock'] is None and (c['pcs'] or 0) > 0)]
    alerts = sorted(alerts, key=lambda c: -(c['pcs'] or 0))
    with open(CSV_DIR / 'stock-alerts.csv', 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['Name', 'New SAP', 'Vendor Code', 'Vendor Type', 'PCS', 'Stock',
                    'EXT Stock', 'Min Stock', 'Status', 'PR Numbers', 'Machines Count'])
        for c in alerts:
            w.writerow([_s(c['name']), _s(c['newsap']), _s(c['vendor']), _s(c['vtype']),
                        _s(c['pcs']), '0' if c['stock'] is None else _s(c['stock']),
                        _s(c['ext']), _s(c['min']), _s(c['status']),
                        '; '.join(c['prs']), _s(c['n_machines'])])

    with open(CSV_DIR / 'catalog.csv', 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['Name', 'New SAP', 'SAP', 'Vendor Code', 'Vendor Type', 'PCS', 'Stock',
                    'EXT Stock', 'Min Stock', 'Status', 'Machines Count', 'Rows'])
        for c in catalog:
            w.writerow([_s(c['name']), _s(c['newsap']), _s(c['sap']), _s(c['vendor']),
                        _s(c['vtype']), _s(c['pcs']), _s(c['stock']), _s(c['ext']),
                        _s(c['min']), _s(c['status']), _s(c['n_machines']), _s(c['rows'])])

    with open(CSV_DIR / 'stock-analysis.csv', 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['Code', 'Name', 'PCS', 'Stock', 'Min Stock', 'Status'])
        for x in sa_rows:
            w.writerow([_s(x['code']), _s(x['name']), _s(x['pcs']), _s(x['stock']),
                        _s(x['min']), _s(x['status'])])
    return len(alerts)


# ---- Changelog (dashboard "Update log" tab, idea 10) ----
# The dashboard's "سجل التحديثات" tab reads changelog.json (repo root) via the
# same channels as website_data.json. It expects a JSON array, newest first,
# of entries: {date, items, alerts, added, critical, resolved,
# stock_changes, note}. added/critical/resolved are [{name, newsap}];
# stock_changes are [{name, newsap, old, new}] for |new-old| >= 10.
# The diff is computed against the PREVIOUS website_data.json catalog, AFTER
# the stock-changes.json overrides are applied (i.e. the final catalog).
CHANGELOG = ROOT / 'changelog.json'
CHANGELOG_MAX = 50


def _cat_key(c):
    """Match key between catalog items, same as the grouping in main():
    New SAP when present, else (lower-cased name, vendor)."""
    if c.get('newsap'):
        return ('SAP', c['newsap'])
    return ('NAME', (c.get('name') or '').lower(), c.get('vendor'))


def _item_ref(c):
    """The {name, newsap} reference the dashboard's changelog tab shows."""
    return {'name': c.get('name'), 'newsap': c.get('newsap')}


def _alert_count(catalog):
    """Alert count using the SAME rule as the stock-alerts CSV in write_csvs:
    stock == 0, or status Critical / Below Min, or unrecorded stock with
    PCS > 0 (unrecorded stock counted as 0, per Adora's option A)."""
    return sum(1 for c in catalog
               if c['stock'] == 0 or c['status'] in ('Critical', 'Below Min')
               or (c['stock'] is None and (c['pcs'] or 0) > 0))


def _load_prev_catalog(path):
    """Read the catalog from a previous website_data.json. Returns None when
    the file is absent (first run -> baseline entry), else the catalog list
    (or [] when the file is unreadable / has no catalog)."""
    if not path.exists():
        return None
    try:
        prev = json.loads(path.read_text(encoding='utf-8'))
        cat = prev.get('catalog') if isinstance(prev, dict) else None
        return cat if isinstance(cat, list) else []
    except Exception as e:
        print(f'WARNING: could not read previous {path.name}: {e}', file=sys.stderr)
        return []


def compute_changelog_entry(new_catalog, prev_catalog, run_date):
    """Build one changelog entry by diffing the final new catalog against the
    previous website_data.json catalog. prev_catalog is None on the first run
    (no previous file), which yields a baseline entry with empty lists."""
    entry = {
        'date': run_date,
        'items': len(new_catalog),
        'alerts': _alert_count(new_catalog),
        'added': [],
        'critical': [],
        'resolved': [],
        'stock_changes': [],
        'note': '',
    }
    if prev_catalog is None:
        entry['note'] = 'baseline'
        return entry
    old_by_key = {}
    for c in prev_catalog:
        old_by_key[_cat_key(c)] = c
    for c in new_catalog:
        key = _cat_key(c)
        old = old_by_key.get(key)
        if old is None:
            entry['added'].append(_item_ref(c))
            continue
        if old.get('status') != 'Critical' and c.get('status') == 'Critical':
            entry['critical'].append(_item_ref(c))
        if old.get('status') in ('Critical', 'Below Min') and \
                c.get('status') not in ('Critical', 'Below Min'):
            entry['resolved'].append(_item_ref(c))
        os_ = old.get('stock')
        ns = c.get('stock')
        if os_ is not None and ns is not None and abs(ns - os_) >= 10:
            entry['stock_changes'].append({'name': c.get('name'),
                                           'newsap': c.get('newsap'),
                                           'old': os_, 'new': ns})
    return entry


def update_changelog(entry):
    """Prepend one entry to changelog.json (repo root), capping the history
    at CHANGELOG_MAX entries. Returns the number of entries kept."""
    history = []
    if CHANGELOG.exists():
        try:
            raw = json.loads(CHANGELOG.read_text(encoding='utf-8'))
            if isinstance(raw, list):
                history = raw
            else:
                print(f'WARNING: {CHANGELOG.name} is not a JSON array; starting fresh',
                      file=sys.stderr)
        except Exception as e:
            print(f'WARNING: could not read {CHANGELOG.name}: {e}; starting fresh',
                  file=sys.stderr)
    history = [entry] + history
    history = history[:CHANGELOG_MAX]
    with open(CHANGELOG, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=2)
    return len(history)


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


# ---- Dashboard edit overrides (stock-changes.json) ----
# The dashboard's Adora-only edit mode downloads stock-changes.json: absolute
# stock / min-stock values keyed by New SAP (or by name), plus withdrawal
# entries carrying a stock_after. Adora uploads that file to workbook/ (or
# the repo root) and this applies it on top of the workbook values, so the
# team's live data reflects his edits until the workbook itself is updated.
# Values are absolute, so re-applying the same file is idempotent.
CHANGES_CANDIDATES = [WORKBOOK_DIR / 'stock-changes.json',
                      ROOT / 'stock-changes.json']
_MISSING = object()


def _change_key(c):
    """Normalise one change entry to (newsap, name, vendor, stock, min).
    stock/min are _MISSING when the key is absent, None when explicitly null
    (meaning 'unrecorded'), else a number."""
    if not isinstance(c, dict):
        return None
    newsap = nc(c.get('newsap') or c.get('new_sap') or c.get('newSap')
                or c.get('code') or c.get('key'))
    name = clean(c.get('name'))
    vendor = nc(c.get('vendor'))
    stock = _MISSING
    for k in ('stock', 'new_stock', 'newStock', 'stock_after'):
        if k in c:
            stock = nn(c[k]) if c[k] is not None else None
            break
    mn = _MISSING
    for k in ('min', 'min_stock', 'minStock', 'new_min'):
        if k in c:
            mn = nn(c[k]) if c[k] is not None else None
            break
    if not newsap and not name:
        return None
    return (newsap, name, vendor, stock, mn)


def load_stock_changes():
    """Read stock-changes.json (workbook/ first, else repo root). Returns a
    list of normalised change tuples. Empty list when the file is absent."""
    path = None
    for p in CHANGES_CANDIDATES:
        if p.exists():
            path = p
            break
    if path is None:
        return []
    try:
        raw = json.loads(path.read_text(encoding='utf-8-sig'))
    except Exception as e:
        print(f'WARNING: could not read {path.name}: {e}', file=sys.stderr)
        return []
    entries = []
    if isinstance(raw, dict):
        for key in ('overrides', 'changes', 'items', 'entries', 'withdrawals'):
            v = raw.get(key)
            if isinstance(v, list):
                entries.extend(v)
            elif isinstance(v, dict):
                for k, val in v.items():
                    if isinstance(val, dict):
                        d = dict(val)
                        if k and str(k)[0].isdigit():
                            d.setdefault('newsap', k)
                        else:
                            d.setdefault('name', k)
                        entries.append(d)
        if not entries:
            entries = [raw]
    elif isinstance(raw, list):
        entries = raw
    out = []
    for c in entries:
        t = _change_key(c)
        if t:
            out.append(t)
    return out


def apply_stock_changes(catalog, status_fn):
    """Apply absolute stock/min overrides to the catalog in place, recomputing
    each affected item's status. Returns (applied, skipped) counts."""
    changes = load_stock_changes()
    if not changes:
        return 0, 0
    applied = 0
    skipped = 0
    for newsap, name, vendor, stock, mn in changes:
        target = None
        if newsap:
            for c in catalog:
                if c['newsap'] and str(c['newsap']).strip() == str(newsap).strip():
                    target = c
                    break
        if target is None and name:
            cands = [c for c in catalog if c['name'] and c['name'].lower() == name.lower()]
            if vendor:
                for c in cands:
                    if c['vendor'] and str(c['vendor']).strip() == str(vendor).strip():
                        target = c
                        break
            if target is None and len(cands) == 1:
                target = cands[0]
            elif target is None and len(cands) > 1:
                print(f"WARNING: name '{name}' matches {len(cands)} catalog items; "
                      f"skipping (give a New SAP or vendor)", file=sys.stderr)
                skipped += 1
                continue
        if target is None:
            print(f"WARNING: no catalog item for change "
                  f"(newsap={newsap}, name={name})", file=sys.stderr)
            skipped += 1
            continue
        if stock is not _MISSING:
            target['stock'] = stock
        if mn is not _MISSING:
            target['min'] = mn
        target['status'] = status_fn(target['stock'], target['min'])
        applied += 1
    return applied, skipped

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
    # ---- Dashboard edit overrides (stock-changes.json) ----
    # Apply Adora's dashboard edits on top of the workbook values, before the
    # sort, so the new statuses drive the ordering and the CSV exports.
    n_applied, n_skipped = apply_stock_changes(catalog, status)
    if n_applied or n_skipped:
        print(f'stock-changes.json: applied {n_applied}, skipped {n_skipped}')

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
    # ---- Changelog for the dashboard "Update log" tab (idea 10) ----
    # Diff the FINAL catalog (after stock-changes.json overrides) against the
    # previous website_data.json, BEFORE overwriting it.
    prev_catalog = _load_prev_catalog(OUT)
    entry = compute_changelog_entry(catalog, prev_catalog, data['generated'])
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
    n_alerts = write_csvs(catalog, sa_rows)
    n_hist = update_changelog(entry)
    print('changelog entries:', n_hist, '| added', len(entry['added']),
          '| critical', len(entry['critical']), '| resolved', len(entry['resolved']),
          '| stock changes', len(entry['stock_changes']), '| note:', entry['note'] or '-')
    print('CSV exports:', 'stock-alerts', n_alerts, '| catalog', len(catalog),
          '| stock-analysis', len(sa_rows))
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
