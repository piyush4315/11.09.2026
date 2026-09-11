#!/usr/bin/env python3
"""
tools/preview_buyer_groups_search.py -- preview what the Buyer Groups sheet shows.

There is no Excel in this sandbox, so this replays the sheet's own formula chain
in Python.  Everything it needs (which Buyer Groups column is fed by which Final
Calculation Sheet column, and which source row each displayed row is linked to)
is parsed out of the workbook itself, so the preview cannot drift away from the
sheet:

    SEARCH row 4 (B4:AI4)
      -> AJ6:AJ42   per-lot match flag          (1 = that source lot matches)
      -> AK6:AK42   running rank of the matches
      -> AL7:AL43   MATCH(rank, AK6:AK42)        strip row -> source row
      -> A7:AI43    the strip itself (INDEX of 'Final Calculation Sheet'!..)
      -> CF on A45:AI96 greys every lot row whose lot is not in the strip

usage
    python3 tools/preview_buyer_groups_search.py 11.09.2026.xlsm --search C4=NATIONAL
    python3 tools/preview_buyer_groups_search.py 11.09.2026.xlsm \\
            --search B4=1875 --search H4=1181
    python3 tools/preview_buyer_groups_search.py 11.09.2026.xlsm --search C4="(blank)"
"""

import argparse
import re
import sys

import openpyxl
from openpyxl.utils import column_index_from_string as cidx
from openpyxl.utils import get_column_letter as gcl

SHEET, FCS = "Buyer Groups", "Final Calculation Sheet"
SRC_FIRST, SRC_LAST = 4, 40                    # data rows on Final Calculation Sheet
BAR, STRIP, TABLE = 6, (7, 43), (45, 94)       # strip bar, strip, buyer blocks
SEARCH_ROW = 4
QREF = re.compile(r"'(?:[^']+)'\$?!\$?([A-Z]{1,3})\$?(\d+)")


def col_map(ws):
    """BG column -> FCS column, read off the first table row's formulas."""
    out = {}
    for col in range(1, 36):
        v = ws.cell(row=TABLE[0], column=col).value
        if not isinstance(v, str):
            continue
        m = re.search(r"'Final Calculation Sheet'!\$([A-Z]{1,3})\$(\d+)", v)
        if m:
            out[col] = cidx(m.group(1))
    return out


def row_source_rows(ws):
    """{display row -> FCS source row} for every lot row of the grouped table."""
    out = {}
    for r in range(TABLE[0], TABLE[1] + 1):
        if not isinstance(ws.cell(row=r, column=1).value, (int, float)):
            continue                                        # the ∑ subtotal row
        v = ws.cell(row=r, column=2).value
        m = re.search(r"'Final Calculation Sheet'!\$F\$(\d+)", v or "")
        if m:
            out[r] = int(m.group(1))
    return out


def matches(value, needle):
    """The rule the sheet uses: (blank)/(nonblank) sentinels, else contains-text."""
    txt = "" if value is None else str(value)
    needle = str(needle)
    if needle == "(blank)":
        return txt == ""
    if needle == "(nonblank)":
        return txt != ""
    return needle.lower() in txt.lower()                    # SEARCH(): contains, case-insensitive


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsm")
    ap.add_argument("--search", action="append", default=[],
                    help="SEARCH-row cell and value, e.g. C4=NATIONAL ENTERPRISES (repeatable)")
    ap.add_argument("--show", type=int, default=10, help="how many strip rows to print")
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.xlsm, data_only=True)          # values (FCS holds data)
    wf = openpyxl.load_workbook(args.xlsm, data_only=False)         # formulas (row links)
    ws, fcs = wb[SHEET], wb[FCS]
    wsf = wf[SHEET]
    cmap = col_map(wsf)
    if len(cmap) < 30:
        sys.exit(f"could not read the column links from row {TABLE[0]} ({len(cmap)} found) "
                 "- is this the un-converted sheet?")

    search = {}
    for item in args.search:
        cell, _, value = item.partition("=")
        col = cidx(re.sub(r"\d", "", cell.upper()))
        if not 2 <= col <= 35:
            sys.exit(f"--search {item}: the SEARCH row covers B4:AI4 only")
        search[col] = value
    if not search:
        sys.exit("pass at least one --search C4=SOMETHING")

    # AJ6:AJ42 -> AK6:AK42 -> AL7:AL43, exactly as the sheet does it
    flags, rank, cur = [], [], 0
    for src in range(SRC_FIRST, SRC_LAST + 1):
        ok = all(matches(fcs.cell(row=src, column=cmap[c]).value, needle)
                 for c, needle in search.items() if c in cmap)
        flags.append(1 if ok else 0)
        cur += flags[-1]
        rank.append(cur if flags[-1] else None)
    order = [src for src, f in zip(range(SRC_FIRST, SRC_LAST + 1), flags) if f]

    print("search:  " + ", ".join(f"{gcl(c)}{SEARCH_ROW} = {v!r}" for c, v in search.items()))
    print(f"match flags AJ6:AJ42 -> {sum(flags)} of {len(flags)} lots match\n")

    print(f"the strip (rows {STRIP[0]}:{STRIP[1]}) - 'MATCHED LOTS' pinned above the buyer blocks")
    cols = [2, 3, 4, 6, 8]
    print("  row   " + "".join(f"{wsf.cell(row=5, column=c).value!s:24.24}" for c in cols))
    shown = 0
    for k in range(1, STRIP[1] - STRIP[0] + 1):
        row = STRIP[0] + k - 1
        src = order[k - 1] if k <= len(order) else None
        if src is None:
            continue
        vals = [fcs.cell(row=src, column=cmap[c]).value if c in cmap else "" for c in cols]
        print(f"  {row:>4}   " + "".join(f"{str(v):24.24}" for v in vals) + f" <- FCS row {src}")
        shown += 1
        if shown >= args.show:
            if shown < len(order):
                print(f"  ...   {len(order) - shown} more matching row(s) further down the strip")
            break
    if not order:
        print(f"  (nothing matches, so rows {STRIP[0]}:{STRIP[1]} stay blank and the bar reads 0 of 37)")

    # bar text
    print(f"\nrow {BAR} bar reads:  'MATCHED LOTS - {sum(flags)} of 37 lot(s) match'")

    # grey-out: a lot row stays in colour only when its source row is in the strip
    src_of = row_source_rows(wsf)
    dim = [r for r, s in sorted(src_of.items()) if s not in set(order)]
    live = [r for r, s in sorted(src_of.items()) if s in set(order)]

    def brief(rows):
        out, start, prev = [], None, None
        for r in rows:
            if start is None:
                start = prev = r
            elif r == prev + 1:
                prev = r
            else:
                out.append(f"{start}" if start == prev else f"{start}-{prev}")
                start = prev = r
        if start is not None:
            out.append(f"{start}" if start == prev else f"{start}-{prev}")
        return ", ".join(out) or "none"

    print(f"\ngrey-out over rows {TABLE[0]}:{TABLE[1]} (lots that do NOT match):")
    print(f"  stay in colour : rows {brief(live)}   ({len(live)} lot rows)")
    print(f"  greyed out     : rows {brief(dim)}   ({len(dim)} lot rows)")
    print("\n(the ∑ subtotal rows, TOTAL and GRAND TOTAL never grey out; nothing is hidden - "
          "the totals still cover every lot)")


if __name__ == "__main__":
    main()
