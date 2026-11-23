# Macros tracker

The source of truth is a Google Sheet with tabs `food` (daily log), `calc` and `dict`
(known foods). `macros.xlsx`, `food_dict.csv` and `macros/*.csv` are local snapshots
produced by `export_snapshots.py`, so they may be stale. Never treat them as the live dict.

## Adding a food from a label photo

When the user sends a photo of a nutrition label (usually Hebrew), add it to the live `dict`
tab via `sheet_api/dict_api.py` (reads `MACROS_SHEET_URL` / `MACROS_SHEET_TOKEN` from
`sheet_api/.env` or env vars;
setup in `sheet_api/README.md`).

`dict` columns: `אוכל` (name) | `יחידה` (unit) | `פחמימה` carbs g | `חלבון` protein g |
`שומן` fat g | `קלוריות` kcal | `וולידציה` (formula, written by the script) |
`סיבים` fiber g (only once that column exists; `add` detects it).
Macro values are **per `unit`**: `unit=100` means values per 100 g (weighed foods);
`unit=1` means values per 1 piece/serving (counted foods, e.g. `ביצה`, `אבקה` = 1 scoop).

Steps:
1. Read the label. Prefer the "ל-100 גרם" column. Map: `פחמימות` → carbs (total, not
   just sugars `מתוכן סוכרים`), `חלבונים` → protein, `שומנים` → fat (total),
   `אנרגיה (קלוריות)` → calories (kcal, not kJ). If something is unreadable, ask. Don't guess.
2. Unit: default `100` with per-100 g values. If the user says it's eaten by the piece/
   serving (or it's clearly a single-unit item like a bar or a yogurt cup), use `unit=1` with
   per-serving values (take the "למנה" column, or scale per-100 g by the serving weight).
3. Name: short Hebrew name in the style of existing entries; use the user's name if given.
   Run `python sheet_api/dict_api.py list --grep <word>` to check for duplicates/similar items.
4. Sanity check: compare 4·carbs + 4·protein + 9·fat + 2·fiber to kcal (labels exclude fiber
   from carbs). Over ~10% off → point it out (polyols or misreads are common causes) before adding.
5. Show the user the proposed row and **wait for confirmation**, then:
   `python sheet_api/dict_api.py add "<name>" --unit 100 --carbs X --protein X --fat X --calories X [--fiber X]`
   Always read fiber (`סיבים תזונתיים`) off the label when present.
   If it returns `exists`, ask before re-running with `--overwrite`.
6. Report the row number it was written to.

Keep a reply short; the user is on a phone.

## General sheet editing

`dict_api.py` also edits any tab (`tabs`, `read`, `set`, `append`, `delete`, `add-column`,
`rename-column`; see the docstring). Columns are addressed by header name.
- Show the user what will change and wait for an OK before structural edits, deletes or
  edits outside `dict`. Every write prints the previous contents; keep them to undo with.
- `food`/`calc` rows look up `dict` by header name:
  `VLOOKUP($A3, dict!$A:$Z, MATCH("<header>", dict!$1:$1, 0), false) * $B3 / <same for יחידה>`.
  Keep that pattern for new lookup cells, and never rename `dict` headers without updating
  them. `add-column` appends at the end; to reorder, rename headers and rewrite the cells,
  then compare every value against a backup taken first.
- Don't overwrite formula cells with plain values unless asked.
- Recovery of last resort: File → Version history in Google Sheets.
