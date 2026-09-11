#!/usr/bin/env python3
"""
tools/preview_fcs_search.py -- prove the FCS search row is wired up, and preview it.

Two jobs:

  --check-wiring (default, also runs with a search)
      walks all 34 column families across the three places they live and insists
      they all name the same parts, so a search box above column X can only ever
      test column X:

          FCS  X2                 the box
          FCS  AL..BS on each row  the flag IF($X$2=""…, SEARCH($X$2, X4)…)
          FCS  BT4                COUNTIF over the 34 flags
          Lists QP..             flag / present / count / list, reading that flag column
          workbook               SUGF_X -> that list column, over those rows
          FCS                    dataValidation X2 -> SUGF_X
          FCS                    CF X4:X40 -> AND($X$2<>"", <its flag>4=1)
          FCS                    CF A4:AH40 -> AND(…, $BT4=0)

  --search COL=TEXT (repeatable)
      replays the rule the sheet uses - empty box passes, (blank)/(nonblank) test
      emptiness, otherwise SEARCH() = contains text, any case - and prints which
      FCS rows match, which go grey, which cells light up, and what each dropdown
      would be offering at that moment (the cascade), for the columns whose values
      are typed data rather than formulas.

    python3 tools/preview_fcs_search.py 11.09.2026.xlsm
    python3 tools/preview_fcs_search.py 11.09.2026.xlsm --search G4=NATIONAL
    python3 tools/preview_fcs_search.py 11.09.2026.xlsm --search F=1875 --search A=3000
"""

import argparse
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

from openpyxl.utils import column_index_from_string as cidx
from openpyxl.utils import get_column_letter as gcl

SHEET = "Final Calculation Sheet"
FCS_PART, LISTS_PART, WB_PART = ("xl/worksheets/sheet1.xml", "xl/worksheets/sheet4.xml",
                                 "xl/workbook.xml")
SEARCH_ROW, DATA, N_COLS = 2, (4, 40), 34
HELP0 = 38
MASTER = HELP0 + N_COLS
NAME_PREFIX = "SUGF_"
CELL_RE = re.compile(r'<c r="([A-Z]{1,3})\d+"([^>/]*?)(?:/>|>(.*?)</c>)', re.S)
ROW_RE = re.compile(r'<row r="(\d+)"([^>/]*)(?:/>|>(.*?)</row>)', re.S)
F_RE = re.compile(r"<f>(.*?)</f>", re.S)


def unesc(t):
    return t.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&amp;", "&")


def rows_of(xml):
    body = xml.split("<sheetData>", 1)[1].split("</sheetData>", 1)[0]
    out = {}
    for m in ROW_RE.finditer(body):
        cells = {}
        for c in CELL_RE.finditer(m.group(3) or ""):
            cells[c.group(1)] = unesc(c.group(3) or "")
        out[m.group(1)] = cells
    return out


def formula_of(cells, ref):
    letters = re.match(r"([A-Z]+)", ref).group(1)
    row = re.search(r"\d+", ref).group(0)
    m = F_RE.search(cells.get(row, {}).get(letters, ""))
    return unesc(m.group(1)) if m else ""


# ------------------------------------------------------------------ wiring audit

def check_wiring(fcs_xml, lst_xml, wb_xml, report):
    fcs, lst = rows_of(fcs_xml), rows_of(lst_xml)
    pools = {}
    for m in re.finditer(r'<c r="([A-Z]{1,3})3"[^>]*t="inlineStr"><is><t[^>]*>([^<]*)</t></is></c>',
                         lst_xml):
        hit = re.search(r"\((?:Final|payment) ([A-Z]{1,2})\)", m.group(2))
        if m.group(2).strip().endswith(")") and re.search(r"\(Final ([A-Z]{1,2})\)\s*$", m.group(2)):
            pools[re.search(r"\(Final ([A-Z]{1,2})\)\s*$", m.group(2))[1]] = gcl(cidx(m.group(1)) + 4)
    problems = []
    names = dict(re.findall(r'<definedName name="([^"]+)">([^<]*)</definedName>', wb_xml))
    dvs = dict((re.search(r"([A-Z]+)\d+", m.group(1)).group(1), m.group(2))
               for m in re.finditer(r'<dataValidation sqref="([A-Z]+\d+)"[^>]*type="list">'
                                    r"<formula1>([^<]+)</formula1>", fcs_xml))
    cf_hit, cf_dim = {}, ("", "")
    for m in re.finditer(r'<conditionalFormatting sqref="([A-Z]+\d+:[A-Z]+\d+)">'
                         r'<cfRule[^>]*dxfId="(\d+)"[^>]*><formula>(.*?)</formula>', fcs_xml, re.S):
        rng, dxf, f = unesc(m.group(1)), m.group(2), unesc(m.group(3))
        if dxf == "50":
            cf_dim = (rng, f)
        elif dxf == "3":
            cf_hit[re.match(r"([A-Z]+)\d+", rng).group(1)] = (rng, f)
    for i in range(N_COLS):
        col = gcl(1 + i)
        helpcol = gcl(HELP0 + i)
        f_flag = formula_of(fcs, f"{helpcol}{DATA[0]}")
        want = [f"${col}${SEARCH_ROW}", f"{col}{DATA[0]}"]
        if not all(w in f_flag for w in want):
            problems.append(f"{helpcol}{DATA[0]} should test {col}{DATA[0]} against ${col}${SEARCH_ROW}"
                            f" but reads: {f_flag[:90]}")
        if "ISNUMBER(SEARCH(" not in f_flag or '"(blank)"' not in f_flag or '"(nonblank)"' not in f_flag:
            problems.append(f"{helpcol}{DATA[0]} lost the (blank)/(nonblank)/SEARCH rule")
        f_master = formula_of(fcs, f"{gcl(MASTER)}{DATA[0]}")
        if f"${gcl(HELP0)}{DATA[0]}:${gcl(HELP0 + N_COLS - 1)}{DATA[0]}" not in f_master:
            problems.append(f"master {gcl(MASTER)}{DATA[0]} does not span the 34 flags: {f_master}")

        name = names.get(f"{NAME_PREFIX}{col}", "")
        m = re.match(r"Lists!\$([A-Z]+)\$(\d+):INDEX\(Lists!\$[A-Z]+\$\d+:\$[A-Z]+\$?(\d+),"
                     r"MAX\(Lists!\$([A-Z]+)\$\d+:\$[A-Z]+\$?(\d+)\)\+2\)", name)
        if not m:
            problems.append(f"{NAME_PREFIX}{col} is not a dynamic INDEX/MAX range: {name[:80]}")
            continue
        list_col, first_row, last_row, count_col, count_end = m.groups()
        if (int(first_row), int(last_row)) != (DATA[0], DATA[1] + 2):
            problems.append(f"{NAME_PREFIX}{col} spans rows {first_row}..{last_row}")
        if dvs.get(col) != f"{NAME_PREFIX}{col}":
            problems.append(f"the box {col}{SEARCH_ROW} offers {dvs.get(col)!r}, not {NAME_PREFIX}{col}")
        f_list = formula_of(lst, f"{list_col}{DATA[0]}")
        f_count = formula_of(lst, f"{count_col}{DATA[0] + 1}") or formula_of(lst, f"{count_col}{DATA[0]}")
        if f"INDEX(${list_col}" in f_list:
            problems.append(f"{list_col} list formula looks self-referential")
        if "MATCH(ROW()-3" not in f_list or "MAX(" not in f_list:
            problems.append(f"{list_col}{DATA[0]} is not a ranked list: {f_list[:80]}")
        # find this column's cascade trio by the flag column it reads
        present_col = gcl(cidx(list_col) - 2)
        flag_col = gcl(cidx(list_col) - 3)
        f_present = formula_of(lst, f"{present_col}{DATA[0]}")
        pool_expected = pools.get(col)
        f_other = formula_of(lst, f"{flag_col}{DATA[0]}")
        if f"'{SHEET}'!${col}${DATA[0]}:${col}${DATA[1]}" not in f_present:
            problems.append(f"{present_col} does not test column {col} of {SHEET}: {f_present[:100]}")
        if f"${flag_col}$" not in f_present:
            problems.append(f"{present_col} does not read the {flag_col} all-other-filters column")
        if f"${pool_expected}$" not in f_present:
            problems.append(f"{present_col} is not fed by the {col} value pool (${pool_expected}$)")
        own = {helpcol}
        for a, b in re.findall(rf"'{SHEET}'!\$?([A-Z]{{1,3}}){DATA[0]}:\$?([A-Z]{{1,3}}){DATA[0]}", f_other):
            if any(own >= {gcl(c)} for c in ()):
                pass
            if cidx(a) <= cidx(helpcol) <= cidx(b):
                problems.append(f"{flag_col}{DATA[0]} counts its own column {helpcol} as an 'other' filter")
        if f_other.count("COUNTIF(") != (1 if i in (0, N_COLS - 1) else 2):
            problems.append(f"{flag_col}{DATA[0]} should hold back exactly its own column: {f_other[:110]}")
        rng, f_hit = cf_hit.get(col, ("-", "-"))
        if rng != f"{col}{DATA[0]}:{col}{DATA[1]}":
            problems.append(f"the yellow rule for column {col} covers {rng}")
        if f"${col}${SEARCH_ROW}" not in f_hit or f"${helpcol}{DATA[0]}" not in f_hit:
            problems.append(f"the yellow rule for {col} does not pair the box with {helpcol}: {f_hit[:90]}")
        if not f_count:
            problems.append(f"{count_col} has no running-count formula on row {DATA[0]}")
    if "A4:AH40" not in cf_dim[0] or f"${gcl(MASTER)}{DATA[0]}=0" not in cf_dim[1]:
        problems.append(f"the grey-out rule is missing or wrong: {cf_dim}")
    report.append(f"wiring: {N_COLS} families checked, "
                  f"{len(cf_hit)} per-column highlight rules, grey-out over A4:AH40")
    return problems


# ------------------------------------------------------------------ replay

def matches(value, needle):
    txt = "" if value is None else str(value)
    needle = str(needle)
    if needle == "(blank)":
        return txt == ""
    if needle == "(nonblank)":
        return txt != ""
    if "*" in needle or "?" in needle:
        # SEARCH() matches the pattern anywhere inside the text, so both ends are
        # implicitly wildcards - that is what "SCRAP*" / "*COPPER" / "18?0" mean here
        rx = ".*" + re.escape(needle).replace("\\*", ".*").replace("\\?", ".") + ".*"
        return re.search(rx, txt, re.I) is not None
    return needle.lower() in txt.lower()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsm")
    ap.add_argument("--search", action="append", default=[],
                    help="FCS column and value, e.g. --search G=NATIONAL (repeatable)")
    ap.add_argument("--dropdowns", type=int, default=4, help="how many candidate lists to print")
    args = ap.parse_args()

    with zipfile.ZipFile(args.xlsm) as z:
        fcs_xml = z.read(FCS_PART).decode("utf-8")
        lst_xml = z.read(LISTS_PART).decode("utf-8")
        wb_xml = z.read(WB_PART).decode("utf-8")
    report = []
    problems = check_wiring(fcs_xml, lst_xml, wb_xml, report)
    for line in report:
        print("  •", line)
    if problems:
        print(f"\n  {len(problems)} wiring problem(s):")
        for p in problems[:12]:
            print("   -", p)
        sys.exit(1)
    print("  • every box, flag, pool, name, validation and conditional format agree")

    import openpyxl
    wb = openpyxl.load_workbook(args.xlsm, data_only=True)
    fcs = wb[SHEET]
    search = {}
    for item in args.search:
        col, _, value = item.partition("=")
        col = col.strip().upper().rstrip("2")
        if len(col) > 2 or col < "A" or cidx(col) > N_COLS:
            sys.exit(f"--search {item}: FCS columns run A..AH")
        search[col] = value
    if not search:
        print(f"\nnothing typed in row {SEARCH_ROW} - all {DATA[1] - DATA[0] + 1} rows show, "
              f"no grey, no highlight, every dropdown full")
        return

    tested = []
    for col in search:
        v = fcs[f"{col}{DATA[0]}"].value
        tested.append((col, v is not None))
    for col, ok in tested:
        if not ok:
            print(f"  note: column {col} holds formulas whose values are not cached in this file "
                  f"- it is still searched by the sheet, only this preview cannot read it")

    hits, greyed = [], []
    for r in range(DATA[0], DATA[1] + 1):
        ok = all(matches(fcs[f"{col}{r}"].value, text) for col, text in search.items())
        (hits if ok else greyed).append(r)

    print(f"\nsearch typed: " + ", ".join(f"{col}{SEARCH_ROW}={text!r}" for col, text in search.items()))
    print(f"  BT (master flag) is 1 on {len(hits)} of {DATA[1] - DATA[0] + 1} rows "
          f"-> {len(hits)} row(s) stay in colour, {len(greyed)} go grey")
    print(f"  matching rows  : {brief(hits)}")
    if greyed:
        print(f"  greyed rows    : {brief(greyed)}")
    lit = []
    for col, text in search.items():
        for r in range(DATA[0], DATA[1] + 1):
            if r in hits and matches(fcs[f"{col}{r}"].value, text):
                lit.append(f"{col}{r}")
    print(f"  cells that light up (yellow): {', '.join(lit[:14])}{' …' if len(lit) > 14 else ''} "
          f"({len(lit)})")

    # what the dropdowns would offer: unique values of that column among the rows
    # that survive every OTHER box, in the order the Lists engine produces
    print(f"\ncascading dropdown contents (from the same rule the Lists blocks implement):")
    shown = 0
    for col in list(search) + [gcl(c) for c in range(1, N_COLS + 1) if gcl(c) not in search]:
        others = {c: t for c, t in search.items() if c != col}
        vals = []
        for r in range(DATA[0], DATA[1] + 1):
            if any(not matches(fcs[f"{c}{r}"].value, t) for c, t in others.items()):
                continue
            v = fcs[f"{col}{r}"].value
            if v is None:
                continue
            key = str(v)
            if key not in [str(x) for x in vals]:
                vals.append(v)
        if not vals:
            continue
        print(f"   {col}{SEARCH_ROW} ({fcs[f'{col}3'].value}): "
              f"{', '.join(str(v)[:18] for v in vals[:9])}{' …' if len(vals) > 9 else ''}")
        shown += 1
        if shown >= max(args.dropdowns, len(search)):
            break


def brief(rows):
    out, start, prev = [], None, None
    for r in rows:
        if start is None:
            start = prev = r
        elif r == prev + 1:
            prev = r
        else:
            out.append(str(start) if start == prev else f"{start}-{prev}")
            start = prev = r
    if start is not None:
        out.append(str(start) if start == prev else f"{start}-{prev}")
    return ", ".join(out) or "none"


if __name__ == "__main__":
    main()
