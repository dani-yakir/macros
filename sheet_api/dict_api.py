"""
CLI client for the macros spreadsheet Apps Script web app (stdlib only).

Credentials: env vars MACROS_SHEET_URL / MACROS_SHEET_TOKEN, or KEY=value lines in
sheet_api/.env (gitignored).

Generic commands (any tab; columns by header name, or letter if no header matches):
  tabs
  read SHEET [--from N] [--to N] [--formulas] [--grep TEXT]
  set SHEET ROW "Header=value" ...       value: number, text, =formula, or empty to clear
  append SHEET "Header=value" ...
  delete SHEET ROW [ROW ...]
  add-column SHEET HEADER                 always added after the last column
  rename-column SHEET COLUMN NEW_HEADER
Formulas may use {row} and {col:<header>} placeholders, e.g.
  "וולידציה==4*{col:פחמימה}{row} + 4*{col:חלבון}{row} + 9*{col:שומן}{row}"

dict shortcuts:
  list [--grep TEXT]
  add NAME --unit 100 --carbs X --protein X --fat X --calories X [--fiber X] [--overwrite] [--dry-run]

Every write prints the server's JSON, including the previous contents of changed cells.
"""

import argparse
import json
import os
import sys
import urllib.request

sys.stdout.reconfigure(encoding='utf-8')

DICT = 'dict'
NAME, UNIT, CARBS, PROTEIN, FAT, KCAL, VALID, FIBER = (
    'אוכל', 'יחידה', 'פחמימה', 'חלבון', 'שומן', 'קלוריות', 'וולידציה', 'סיבים')


def load_env_file():
    """Load KEY=value lines from sheet_api/.env without overriding real env vars."""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
    if not os.path.exists(path):
        return
    with open(path, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith('#') and '=' in line:
                k, v = line.split('=', 1)
                os.environ.setdefault(k.strip(), v.strip())


load_env_file()


def call(action, **payload):
    url = os.environ.get('MACROS_SHEET_URL')
    token = os.environ.get('MACROS_SHEET_TOKEN')
    if not url or not token:
        sys.exit('MACROS_SHEET_URL and MACROS_SHEET_TOKEN must be set')
    body = json.dumps({'token': token, 'action': action, **payload}).encode('utf-8')
    req = urllib.request.Request(url, data=body, method='POST',
                                 headers={'Content-Type': 'text/plain;charset=utf-8'})
    # Apps Script answers with a 302 to googleusercontent; urllib follows it as GET.
    with urllib.request.urlopen(req, timeout=60) as resp:
        res = json.loads(resp.read().decode('utf-8'))
    if not res.get('ok'):
        sys.exit(f"error: {json.dumps(res, ensure_ascii=False)}")
    return res


def show(res):
    print(json.dumps(res, ensure_ascii=False))


def parse_value(s):
    if s == '':
        return None
    if s.startswith('='):
        return s
    try:
        f = float(s)
        return int(f) if f.is_integer() and '.' not in s else f
    except ValueError:
        return s


def parse_pairs(pairs):
    values = {}
    for p in pairs:
        if '=' not in p:
            sys.exit(f'expected Header=value, got: {p}')
        k, v = p.split('=', 1)
        values[k.strip()] = parse_value(v)
    return values


def fmt_row(r):
    return f"{r['row']:>4}  " + '  | '.join(f'{k}={v}' for k, v in r['values'].items())


# ---------- generic ----------

def cmd_tabs(a):
    for t in call('tabs')['tabs']:
        print(f"{t['name']}: {t['rows']} rows x {t['cols']} cols | {t['headers']}")


def cmd_read(a):
    res = call('read', sheet=a.sheet, **{'from': a.from_, 'to': a.to, 'formulas': a.formulas})
    for r in res['rows']:
        line = fmt_row(r)
        if a.grep and a.grep not in line:
            continue
        print(line)
        if a.formulas and r.get('formulas'):
            print('      formulas: ' + json.dumps(r['formulas'], ensure_ascii=False))


def cmd_set(a):
    show(call('set', sheet=a.sheet, row=a.row, values=parse_pairs(a.pairs)))


def cmd_append(a):
    show(call('append', sheet=a.sheet, values=parse_pairs(a.pairs)))


def cmd_delete(a):
    show(call('delete_rows', sheet=a.sheet, rows=a.rows))


def cmd_add_column(a):
    show(call('add_column', sheet=a.sheet, header=a.header))


def cmd_rename_column(a):
    show(call('rename_column', sheet=a.sheet, column=a.column, to=a.new_header))


# ---------- dict shortcuts ----------

def dict_rows():
    res = call('read', sheet=DICT)
    return res['headers'], res['rows']


def cmd_list(a):
    _, rows = dict_rows()
    for r in rows:
        if a.grep and a.grep not in str(r['values'].get(NAME, '')):
            continue
        print(fmt_row(r))


def validation_formula(has_fiber):
    f = '=4*{col:%s}{row} + 4*{col:%s}{row} + 9*{col:%s}{row}' % (CARBS, PROTEIN, FAT)
    if has_fiber:
        f += ' + 2*{col:%s}{row}' % FIBER
    return f


def cmd_add(a):
    v = 4 * a.carbs + 4 * a.protein + 9 * a.fat + 2 * (a.fiber or 0)
    diff = (v - a.calories) / a.calories * 100 if a.calories else 0
    print(f"{a.name}: unit={a.unit} carbs={a.carbs} protein={a.protein} fat={a.fat} "
          f"fiber={a.fiber} kcal={a.calories} | check={v:.1f} ({diff:+.1f}%)")
    if a.dry_run:
        return
    headers, rows = dict_rows()
    has_fiber = FIBER in headers
    if a.fiber is not None and not has_fiber:
        sys.exit(f'dict has no {FIBER} column; add it first')
    values = {NAME: a.name.strip(), UNIT: a.unit, CARBS: a.carbs, PROTEIN: a.protein,
              FAT: a.fat, KCAL: a.calories, VALID: validation_formula(has_fiber)}
    if has_fiber:
        values[FIBER] = a.fiber
    existing = next((r for r in rows if str(r['values'].get(NAME, '')).strip() == a.name.strip()), None)
    if existing:
        if not a.overwrite:
            sys.exit(f"exists: {fmt_row(existing)}  (use --overwrite)")
        show(call('set', sheet=DICT, row=existing['row'], values=values))
    else:
        show(call('append', sheet=DICT, values=values))


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest='cmd', required=True)

    sub.add_parser('tabs').set_defaults(func=cmd_tabs)

    p = sub.add_parser('read')
    p.add_argument('sheet')
    p.add_argument('--from', dest='from_', type=int)
    p.add_argument('--to', type=int)
    p.add_argument('--formulas', action='store_true')
    p.add_argument('--grep')
    p.set_defaults(func=cmd_read)

    p = sub.add_parser('set')
    p.add_argument('sheet')
    p.add_argument('row', type=int)
    p.add_argument('pairs', nargs='+')
    p.set_defaults(func=cmd_set)

    p = sub.add_parser('append')
    p.add_argument('sheet')
    p.add_argument('pairs', nargs='+')
    p.set_defaults(func=cmd_append)

    p = sub.add_parser('delete')
    p.add_argument('sheet')
    p.add_argument('rows', type=int, nargs='+')
    p.set_defaults(func=cmd_delete)

    p = sub.add_parser('add-column')
    p.add_argument('sheet')
    p.add_argument('header')
    p.set_defaults(func=cmd_add_column)

    p = sub.add_parser('rename-column')
    p.add_argument('sheet')
    p.add_argument('column')
    p.add_argument('new_header')
    p.set_defaults(func=cmd_rename_column)

    p = sub.add_parser('list')
    p.add_argument('--grep')
    p.set_defaults(func=cmd_list)

    p = sub.add_parser('add')
    p.add_argument('name')
    for k in ('unit', 'carbs', 'protein', 'fat', 'calories'):
        p.add_argument(f'--{k}', type=float, required=True)
    p.add_argument('--fiber', type=float)
    p.add_argument('--overwrite', action='store_true')
    p.add_argument('--dry-run', action='store_true')
    p.set_defaults(func=cmd_add)

    a = ap.parse_args()
    a.func(a)


if __name__ == '__main__':
    main()
