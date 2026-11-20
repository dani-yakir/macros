"""
CLI client for the macros 'dict' sheet Apps Script web app (stdlib only).

Requires env vars:
  MACROS_SHEET_URL    the web app /exec URL
  MACROS_SHEET_TOKEN  the TOKEN script property

Usage:
  python sheet_api/dict_api.py list [--grep TEXT]
  python sheet_api/dict_api.py add NAME --unit 100 --carbs 10 --protein 5 --fat 2 --calories 78 [--overwrite] [--dry-run]
"""

import argparse
import json
import os
import sys
import urllib.request

sys.stdout.reconfigure(encoding='utf-8')


def call(payload):
    url = os.environ.get('MACROS_SHEET_URL')
    token = os.environ.get('MACROS_SHEET_TOKEN')
    if not url or not token:
        sys.exit('MACROS_SHEET_URL and MACROS_SHEET_TOKEN must be set')
    body = json.dumps({'token': token, **payload}).encode('utf-8')
    req = urllib.request.Request(url, data=body, method='POST',
                                 headers={'Content-Type': 'text/plain;charset=utf-8'})
    # Apps Script answers with a 302 to googleusercontent; urllib follows it as GET.
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode('utf-8'))


def validation(c, p, f):
    return 4 * c + 4 * p + 9 * f


def cmd_list(args):
    res = call({'action': 'list'})
    if not res.get('ok'):
        sys.exit(f"error: {res}")
    for r in res['rows']:
        if args.grep and args.grep not in r['name']:
            continue
        print(f"{r['row']:>4}  {r['name']}  | unit={r['unit']} carbs={r['carbs']} "
              f"protein={r['protein']} fat={r['fat']} kcal={r['calories']}")


def cmd_add(args):
    v = validation(args.carbs, args.protein, args.fat)
    diff = (v - args.calories) / args.calories * 100 if args.calories else 0
    print(f"{args.name}: unit={args.unit} carbs={args.carbs} protein={args.protein} "
          f"fat={args.fat} kcal={args.calories} | 4/4/9 check={v:.1f} ({diff:+.1f}%)")
    if args.dry_run:
        return
    res = call({'action': 'upsert', 'name': args.name, 'unit': args.unit,
                'carbs': args.carbs, 'protein': args.protein, 'fat': args.fat,
                'calories': args.calories, 'overwrite': args.overwrite})
    print(json.dumps(res, ensure_ascii=False))
    if not res.get('ok'):
        sys.exit(1)


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest='cmd', required=True)
    lp = sub.add_parser('list')
    lp.add_argument('--grep')
    lp.set_defaults(func=cmd_list)
    a = sub.add_parser('add')
    a.add_argument('name')
    for k in ('unit', 'carbs', 'protein', 'fat', 'calories'):
        a.add_argument(f'--{k}', type=float, required=True)
    a.add_argument('--overwrite', action='store_true')
    a.add_argument('--dry-run', action='store_true')
    a.set_defaults(func=cmd_add)
    args = ap.parse_args()
    args.func(args)


if __name__ == '__main__':
    main()
