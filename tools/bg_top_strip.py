#!/usr/bin/env python3
"""
tools/bg_top_strip.py -- pin the Buyer Groups "MATCHED LOTS" strip to the TOP.

The workbook is hand-written OOXML, so this script rewrites
xl/worksheets/sheet3.xml ('Buyer Groups') and swaps that single part back into
the .xlsm - every other member of the zip is copied through byte for byte.

    OLD layout                                   NEW layout
    -------------------------------------------  -------------------------------------------
    rows 1..5  title / how-to / counts /          rows 1..5    unchanged (SEARCH row 4 and
               SEARCH row 4 / header row 5                    header row 5 stay exactly put)
    rows 6..55 13 buyer blocks + their           row  6       the ▸ MATCHED LOTS bar
               shaded ∑ subtotal rows           rows 7..43   the strip: matching lots, kept
    row  56    TOTAL                                          open while a SEARCH box holds a
    row  57    GRAND TOTAL                                    value, hidden while they are empty
    row  58    blank                             row  44      spacer / the strip's +/- control
    rows 59..96 blank + MATCHED LOTS strip        rows 45..94  13 buyer blocks + subtotals
                                                 rows 95..96  TOTAL / GRAND TOTAL

What the transform guarantees
  * Only DISPLAY columns move: A:AI, plus the strip's own AL rank pointer.
  * Helper columns stay glued to their rows, because the rest of the workbook
    reads them row-for-row:
        AJ / AK   rows 6:42   per-lot match flag + running rank
        AM : DB   rows 4:42   cascading SEARCH dropdown values
        DG : EN   rows 6:42   "BG match helpers", read by Lists!LE:QJ
  * Every unqualified same-sheet reference inside a moved formula is remapped
    with the same row function, so each block's SUM ranges, the ROWS() lot
    counters and the TOTAL row's list of ∑ cells keep pointing where they belong.
  * Cross-sheet references ('Final Calculation Sheet'!.., Lists!..) are never touched.

Two formula bugs get fixed on the way, because the strip cannot work without them:
  * AJ6:AJ42 returned TRUE/FALSE.  SUM() ignores booleans inside a range and
    COUNTIF(range,1) cannot see them either, so the rank chain AK -> AL never
    fired and the strip listed nothing.  They now return 1/0.
  * Non-matching lots are greyed out while a search is active (new conditional
    format at priority 1, reusing the grey dxf that Live Search already uses).

    python3 tools/bg_top_strip.py 11.09.2026.xlsm           # dry run + report
    python3 tools/bg_top_strip.py 11.09.2026.xlsm --apply   # rewrite in place
"""

import argparse
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

from openpyxl.utils import column_index_from_string as cidx
from openpyxl.utils import get_column_letter as gcl

SHEET_PART = "xl/worksheets/sheet3.xml"            # 'Buyer Groups'

DISPLAY_COLS = set(range(1, 36)) | {38}           # A:AI plus AL
OLD_TABLE = (6, 57)                               # blocks, TOTAL, GRAND TOTAL
OLD_STRIP = (59, 96)                              # strip bar + its 37 result rows
TABLE_SHIFT, STRIP_SHIFT = 39, -53
NEW_BAR, NEW_STRIP, NEW_TABLE = 6, (7, 43), (45, 94)
AJ_COL, AL_COL = 36, 38

CF_OPEN, CF_CLOSE = "<conditionalFormatting", "</conditionalFormatting>"

# ---------------------------------------------------------------- row mapping


def map_row(row):
    """Old row number -> new row number, for display cells and refs to them."""
    if OLD_TABLE[0] <= row <= OLD_TABLE[1]:
        return row + TABLE_SHIFT
    if OLD_STRIP[0] <= row <= OLD_STRIP[1]:
        return row + STRIP_SHIFT
    return row


# ---------------------------------------------------------------- ref rewriting

TOKEN = re.compile(r"""
     (?P<str>"(?:[^"]|"")*")                                    # text literal: left alone
   | (?P<qual>(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!              # Sheet!A1  /  Sheet!A1:B2
              \$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)
   | (?P<ref>(?<![A-Za-z0-9_.$])\$?[A-Z]{1,3}\$?\d+(?![A-Za-z0-9_(.]))
""", re.X)
REF_ONLY = re.compile(r"^(?P<ac>\$?)(?P<col>[A-Z]{1,3})(?P<ar>\$?)(?P<row>\d+)$")
F_RE = re.compile(r"<f>(.*?)</f>", re.S)
QREF_RE = re.compile(r"'[^']+'!\$?[A-Z]{1,3}\$?\d+")

stats = {"moved": 0, "kept": 0, "other_sheet": 0}


def rewrite_refs(text, fn):
    """Apply fn(col_number, row) -> new_row | None to every bare same-sheet ref."""
    out, pos = [], 0
    for m in TOKEN.finditer(text):
        out.append(text[pos:m.start()])
        pos = m.end()
        body = m.group(0)
        if m.lastgroup == "qual":
            stats["other_sheet"] += 1
            out.append(body)
            continue
        if m.lastgroup != "ref":
            out.append(body)
            continue
        r = REF_ONLY.match(body)
        col, row = cidx(r["col"]), int(r["row"])
        if col not in DISPLAY_COLS:                 # helper column: pinned in place
            stats["kept"] += 1
            out.append(body)
            continue
        new = fn(col, row)
        if new is None or new == row:
            stats["kept"] += 1
            out.append(body)
            continue
        stats["moved"] += 1
        out.append(body[: len(body) - len(r["row"])] + str(new))
    out.append(text[pos:])
    return "".join(out)


def remap(text):
    """Remap the unqualified same-sheet cell references inside one formula."""
    return rewrite_refs(text, lambda col, row: map_row(row))


def map_sqref(sqref):
    areas = []
    for area in sqref.split():
        m = re.match(r"^([A-Z]{1,3})(\d+):([A-Z]{1,3})(\d+)$", area)
        if not m:
            areas.append(area)
            continue
        areas.append(f"{m.group(1)}{map_row(int(m.group(2)))}:{m.group(3)}{map_row(int(m.group(4)))}")
    return " ".join(areas)

# ---------------------------------------------------------------- xml plumbing

ROW_RE = re.compile(r'<row r="(\d+)"([^>]*?)(?:/>|>(.*?)</row>)', re.S)
CELL_RE = re.compile(r'<c r="([A-Z]{1,3})(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)


def parse_sheet(xml, report):
    i = xml.index("<sheetData>") + len("<sheetData>")
    j = xml.rindex("</sheetData>")
    rows = {}
    for m in ROW_RE.finditer(xml[i:j]):
        num, attrs, inner = int(m.group(1)), m.group(2).strip(), m.group(3) or ""
        cells = {}
        for cm in CELL_RE.finditer(inner):
            letters, crow, cattrs, cinner = cm.groups()
            if int(crow) != num:
                sys.exit(f"cell {letters}{crow} sits inside row element {num}")
            col = cidx(letters)
            if col in cells:
                sys.exit(f"duplicate cell {letters}{num}")
            cells[col] = (cattrs, cinner or "")
        if CELL_RE.sub("", inner).strip():
            sys.exit(f"unparsed content in row {num}")
        rows[num] = (attrs, cells)
    if not rows:
        sys.exit("no rows parsed")
    report.append(f"parsed {len(rows)} rows / {sum(len(c) for _, c in rows.values())} cells")
    return xml[:i], xml[j:], rows


def render(num, attrs, cells):
    head = f'<row r="{num}"' + (f" {attrs}" if attrs else "")
    if not cells:
        return head + "/>"
    parts = []
    for col in sorted(cells):
        cattrs, cinner = cells[col]
        ref = f"{gcl(col)}{num}"
        parts.append(f'<c r="{ref}"{cattrs}>{cinner}</c>' if cinner else f'<c r="{ref}"{cattrs}/>')
    return head + ">" + "".join(parts) + "</row>"

# ---------------------------------------------------------------- the move


def moved_formula(inner):
    return F_RE.sub(lambda fm: f"<f>{remap(fm.group(1))}</f>", inner) if inner else ""


def convert(rows, report):
    if AL_COL in rows.get(7, ("", {}))[1]:
        sys.exit("row 7 already holds an AL cell - this file looks like it was converted already")

    new = {r: rows[r] for r in range(1, 6)}

    # the strip moves up: old 59..96 -> new 6..43.  Each of those new rows also keeps
    # the helper cells that were already on it (AJ/AK/AM:DB/DG:EN live on rows 6..42
    # and are read row-for-row by Lists, so they must not budge).
    for old in range(OLD_STRIP[0], OLD_STRIP[1] + 1):
        target = old + STRIP_SHIFT
        cells = {}
        for col, (cattrs, cinner) in sorted(rows.get(old, ("", {}))[1].items()):
            if col in DISPLAY_COLS:
                cells[col] = (cattrs, moved_formula(cinner))
        for col, val in sorted(rows.get(target, ("", {}))[1].items()):
            if col not in DISPLAY_COLS:
                if col in cells:
                    sys.exit(f"helper cell {gcl(col)}{target} would clash with the strip")
                cells[col] = val
        attrs = f'outlineLevel="1" hidden="1"' if NEW_STRIP[0] <= target <= NEW_STRIP[1] else ""
        new[target] = (attrs, cells)

    # the grouped table moves down: old 6..57 -> new 45..96
    for old in range(OLD_TABLE[0], OLD_TABLE[1] + 1):
        target = old + TABLE_SHIFT
        attrs, cells_old = rows[old]
        cells = {}
        for col, (cattrs, cinner) in sorted(cells_old.items()):
            if col in DISPLAY_COLS:
                cells[col] = (cattrs, moved_formula(cinner))
        new[target] = (attrs, cells)

    report.append("rows remapped: strip -> 6..43, buyer blocks -> 45..96, row 44 left as the spacer")
    return new

# ---------------------------------------------------------------- polish

BAR_FORMULA = (
    'IF(COUNTIF($B$4:$AI$4,"&lt;&gt;")=0,'
    '"▸ MATCHED LOTS  —  rows 7:43 sit above the buyer blocks: type in any yellow SEARCH box and '
    'every matching lot is pinned here, while the lots that do not match go grey below.  '
    'Double-click this row to lift the matching buyers’ whole blocks to the top of the list too.",'
    '"▸ MATCHED LOTS  —  "&amp;SUM($AJ$6:$AJ$42)&amp;" of 37 lot(s) match   ·   non-matching lots are '
    'greyed out in the blocks below   ·   double-click this row to lift the matching buyers’ blocks '
    'up, double-click again to put them back"'
    '&amp;IF(IFERROR($EP$1=1,FALSE),"   ·   THE BLOCKS ARE LIFTED - double-click here to put them back",""))'
)

HINT_OLD = "matching lots will list here"
HINT_NEW = "the SEARCH row above is what fills this list"


def polish(new, report):
    # column A is hidden and C is the first scrolling column, so the bar message goes
    # into C6 and D6:AI6 stay blank to let the text spill across the whole row.
    _, cells = new[NEW_BAR]
    cells[3] = (' s="117"', f"<f>{BAR_FORMULA}</f>")
    for col in range(4, 36):
        cells[col] = (' s="117" t="n"', "")

    # AJ6:AJ42 -> 1/0, otherwise SUM() and COUNTIF(...,1) never see a match
    fixed = 0
    for r in range(6, 43):
        hcells = new.get(r, ("", {}))[1]
        if AJ_COL not in hcells:
            continue
        cattrs, cinner = hcells[AJ_COL]
        m = F_RE.match(cinner)
        if m and m.group(1).startswith("AND("):
            hcells[AJ_COL] = (cattrs, f"<f>IF({m.group(1)},1,0)</f>")
            fixed += 1
    if fixed != 37:
        sys.exit(f"expected to fix 37 AJ match flags, fixed {fixed}")

    # the strip's own hint line now sits under the SEARCH row instead of above it
    first = new[NEW_STRIP[0]][1].get(4)
    if first and HINT_OLD in first[1]:
        new[NEW_STRIP[0]][1][4] = (first[0], first[1].replace(HINT_OLD, "↓ " + HINT_NEW))
    report.append(f"strip bar text written, AJ6:AJ42 match flags now 1/0 ({fixed} cells), "
                  "hint line re-worded for its new place at the top")

# ---------------------------------------------------------------- checks


def verify_exact_remap(old_rows, new, report):
    """Before polish(): a moved cell must be byte-identical to the old cell except for
    remapped row numbers, so nothing can be quietly mangled."""
    n = 0
    for old in list(range(OLD_TABLE[0], OLD_TABLE[1] + 1)) + list(range(OLD_STRIP[0], OLD_STRIP[1] + 1)):
        target = old + (TABLE_SHIFT if old <= OLD_TABLE[1] else STRIP_SHIFT)
        for col, (cattrs, cinner) in sorted(old_rows[old][1].items()):
            if col not in DISPLAY_COLS:
                continue
            if moved_formula(cinner) != new[target][1][col][1]:
                sys.exit(f"{gcl(col)}{old} -> {gcl(col)}{target}: unexpected edit")
            if cattrs != new[target][1][col][0]:
                sys.exit(f"{gcl(col)}{old} -> {gcl(col)}{target}: style/attrs changed")
            n += 1
    report.append(f"{n} moved cells are byte-identical apart from their remapped row numbers")


def verify(old_rows, new, report):
    def flat(rws):
        return {f"{gcl(c)}{r}": v for r, (_, cells) in rws.items()
                for c, v in cells.items() if c in DISPLAY_COLS}

    o, n = flat(old_rows), flat(new)
    # polish() deliberately rewrites these cells afterwards, so their own references
    # are allowed to differ from the old ones (cross-sheet links must not).
    POLISHED = {f"{gcl(c)}{OLD_STRIP[0]}" for c in DISPLAY_COLS} | {f"D{OLD_STRIP[0] + 1}"}
    if len(o) != len(n):
        sys.exit(f"display cell count changed {len(o)} -> {len(n)}")

    def ref_pairs(inner):
        out = []
        for fm in F_RE.finditer(inner):
            for rm in TOKEN.finditer(fm.group(1)):
                if rm.lastgroup == "ref":
                    r = REF_ONLY.match(rm.group(0))
                    out.append((cidx(r["col"]), int(r["row"])))
        return out

    for key in o:
        letters, row = re.match(r"([A-Z]+)(\d+)", key).groups()
        tgt = f"{letters}{map_row(int(row))}"
        if tgt not in n:
            sys.exit(f"{key} lost - expected it at {tgt}")
        if QREF_RE.findall(o[key][1]) != QREF_RE.findall(n[tgt][1]):
            sys.exit(f"cross-sheet refs changed in {key}")
        if key in POLISHED:
            continue          # the bar text and the strip hint are re-worded by polish()
        expect = [(c, map_row(r) if c in DISPLAY_COLS else r) for c, r in ref_pairs(o[key][1])]
        if expect != ref_pairs(n[tgt][1]):
            sys.exit(f"{key} -> {tgt} references {expect}, got {ref_pairs(n[tgt][1])}")
    report.append(f"{len(n)} display cells carried over; every 'Final Calculation Sheet' / Lists "
                  "reference preserved cell for cell and every same-sheet ref mapped exactly")

    for r in range(1, 97):
        ohelp = {c: v for c, v in old_rows.get(r, ("", {}))[1].items() if c not in DISPLAY_COLS}
        nhelp = {c: v for c, v in new.get(r, ("", {}))[1].items() if c not in DISPLAY_COLS}
        if r >= 43 and nhelp:
            sys.exit(f"helper cells drifted below row 42 (row {r})")
        if set(ohelp) - {AJ_COL} != set(nhelp) - {AJ_COL}:
            sys.exit(f"helper columns changed shape at row {r}")
        for c in set(ohelp) - {AJ_COL}:
            if ohelp[c] != nhelp[c]:
                sys.exit(f"helper cell {gcl(c)}{r} was modified")
    report.append("helper cells AJ/AK, AM:DB and DG:EN still sit on rows 4..42 untouched "
                  "(only AJ's TRUE/FALSE became 1/0)")

    for probe_old, probe_new in ((6, 45), (17, 56), (30, 69), (56, 95), (57, 96), (60, 7), (96, 43)):
        if QREF_RE.findall(o[f"B{probe_old}"][1]) != QREF_RE.findall(n[f"B{probe_new}"][1]):
            sys.exit(f"B{probe_old} -> B{probe_new} lot-number source changed")
        if bool(o[f"B{probe_old}"][1]) != bool(n[f"B{probe_new}"][1]):
            sys.exit(f"B{probe_old} -> B{probe_new} formula presence changed")
    report.append("spot checks: table rows 6/17/30/56/57 -> 45/56/69/95/96 and strip rows "
                  "60/96 -> 7/43 keep pulling the very same source cells")
    report.append(f"references: {stats['moved']} remapped, {stats['kept']} pinned, "
                  f"{stats['other_sheet']} cross-sheet left alone")

# ---------------------------------------------------------------- tail: CF etc.

DIM_RULE = (
    '<conditionalFormatting sqref="A45:AI96"><cfRule type="expression" priority="1" dxfId="50">'
    '<formula>AND(ISNUMBER($A45),$B45&lt;&gt;"",COUNTIF($B$4:$AI$4,"&lt;&gt;")&gt;0,'
    'COUNTIFS($B$7:$B$43,$B45,$C$7:$C$43,$C45)=0)</formula></cfRule></conditionalFormatting>'
)
PER_COL = re.compile(r"^([A-Z]{1,3})6:([A-Z]{1,3})55$")


def shift_anchor_row6(text):
    """Clone of a table rule for the strip: those formulas are anchored on row 6 and
    the strip starts one row lower.  $B$4 and the helper columns are left alone."""
    return rewrite_refs(text, lambda col, row: row + 1 if row == 6 else None)


def convert_tail(tail, report):
    """Move every conditional format onto the new rows, and clone each SEARCH highlight
    rule for the strip as its own single-area block so the anchor of the relative
    formula can never be read the wrong way round."""
    blocks = re.findall(r'<conditionalFormatting sqref="([^"]+)">(.*?)</conditionalFormatting>',
                        tail, flags=re.S)
    if not blocks:
        sys.exit("no conditional formatting found - unexpected file shape")

    out = tail
    cloned = 0
    max_prio = max(int(x) for x in re.findall(r'priority="(\d+)"', tail))
    for sqref, inner in blocks:
        old_block = f'<conditionalFormatting sqref="{sqref}">{inner}</conditionalFormatting>'
        new_inner = re.sub(r"<formula>(.*?)</formula>",
                           lambda fm: f"<formula>{remap(fm.group(1))}</formula>", inner, flags=re.S)
        block = f'<conditionalFormatting sqref="{map_sqref(sqref)}">{new_inner}</conditionalFormatting>'
        out = out.replace(old_block, block, 1)

        m = PER_COL.match(sqref)
        if m and m.group(1) == m.group(2) and cidx(m.group(1)) in DISPLAY_COLS:
            col = m.group(1)
            strip_inner = re.sub(r"<formula>(.*?)</formula>",
                                 lambda fm: f"<formula>{shift_anchor_row6(fm.group(1))}</formula>",
                                 inner, flags=re.S)
            max_prio += 1
            strip_inner = re.sub(r'priority="\d+"', 'priority="%d"' % max_prio, strip_inner, count=1)
            clone = (f'<conditionalFormatting sqref="{col}{NEW_STRIP[0]}:{col}{NEW_STRIP[1]}">'
                     + strip_inner + "</conditionalFormatting>")
            out = out.replace(block, block + clone, 1)
            cloned += 1

    if cloned != 34:
        sys.exit(f"expected to clone 34 highlight rules for the strip, cloned {cloned}")
    if 'sqref="A45:AI96"' not in out:
        out = out.replace(CF_OPEN, DIM_RULE + CF_OPEN, 1)
    report.append(f"conditional formats remapped onto the new rows; {cloned} SEARCH-highlight rules "
                  "cloned over the strip; grey-out rule for non-matching lots added (dxf 50, priority 1)")
    return out


HOWTO_NEW = (
    "MATCHED LOTS ARE PINNED TO THE TOP — as soon as any SEARCH box holds a value, the strip in "
    "rows 7:43 opens above the buyer blocks and lists every matching lot there (same rule as the "
    "yellow highlight: contains-text per column, AND across columns, (blank)/(nonblank) supported), "
    "while the lots that do NOT match go grey inside the blocks below; the strip hides itself again "
    "once the SEARCH row is cleared, and its +/- control on row 44 opens or closes it by hand too. "
    "For multi-value ANY (comma) or wildcards use Live Search.  •  LIFT MATCHING BLOCKS — double-click "
    "the ▸ MATCHED LOTS bar (row 6) and every buyer block holding a matching lot moves to the top of "
    "the list, matched lots first inside the block; double-click the bar again to restore the plain "
    "alphabetical order. Typing in the SEARCH row puts the blocks back in that plain order first, so "
    "every search starts from the same layout. Blocks are physically moved, never rewritten, so "
    "shading, row groups, subtotals and every live link stay intact."
)


def convert_howto(xml, report):
    pat = re.compile(r"MATCHED LOTS — the block beneath GRAND TOTAL lists every lot.*?use Live Search\.",
                     re.S)
    if not pat.search(xml):
        sys.exit("could not find the MATCHED LOTS bullet in row 2")
    report.append("row 2 'HOW TO READ' now documents the top strip, the grey-out and the float")
    return pat.sub(lambda m: HOWTO_NEW, xml, count=1)

# ---------------------------------------------------------------- zip i/o


def read_part(path, name):
    with zipfile.ZipFile(path) as z:
        return z.read(name).decode("utf-8")


def rewrite_part(src, name, text, dst):
    with zipfile.ZipFile(src) as zin:
        infos = zin.infolist()
        data = {i.filename: zin.read(i.filename) for i in infos}
    data[name] = text.encode("utf-8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            zout.writestr(zi, data[info.filename])

# ---------------------------------------------------------------- main


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsm")
    ap.add_argument("--apply", action="store_true", help="rewrite the workbook in place")
    ap.add_argument("--out", default=None, help="write the result to a different file")
    ap.add_argument("--dump", action="store_true", help="print the new row map")
    args = ap.parse_args()

    report = []
    xml = read_part(args.xlsm, SHEET_PART)
    head, tail, old_rows = parse_sheet(xml, report)
    tail = tail[len("</sheetData>"):]

    new = convert(old_rows, report)
    verify_exact_remap(old_rows, new, report)
    polish(new, report)
    verify(old_rows, new, report)

    out_xml = head + "".join(render(r, *new[r]) for r in sorted(new)) + "</sheetData>"
    out_xml += convert_tail(tail, report)
    out_xml = convert_howto(out_xml, report)

    if not (0.5 * len(xml) < len(out_xml) < 1.5 * len(xml)):
        sys.exit(f"sheet xml size implausible: {len(xml)} -> {len(out_xml)}")
    ET.fromstring(out_xml)
    report.append("new sheet3.xml is well-formed XML")

    for line in report:
        print("  •", line)
    if args.dump:
        for r in sorted(new):
            attrs, cells = new[r]
            print(f"  row {r:>3} {attrs:<28} {' '.join(gcl(c) for c in sorted(cells))}")

    dst = args.out or (args.xlsm if args.apply else None)
    if not dst:
        print("\ndry run - pass --apply to write the change")
        return
    if dst == args.xlsm:
        tmp = dst + ".tmp"
        rewrite_part(args.xlsm, SHEET_PART, out_xml, tmp)
        os.replace(tmp, dst)
    else:
        rewrite_part(args.xlsm, SHEET_PART, out_xml, dst)
    print(f"\nwrote {dst}")


if __name__ == "__main__":
    main()
