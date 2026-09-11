#!/usr/bin/env python3
"""
tools/fcs_search_row.py -- give 'Final Calculation Sheet' the Live Search treatment.

On the source sheet itself, without moving a single row and without touching a
line of VBA:

  row 2 (already inside the frozen pane, so it rides along while you scroll)
      A2:AH2   one yellow SEARCH box per column - all 34 of them
      AI2      a live read-out: how many of the 37 rows match, and what to type
  hidden helper columns on FCS
      AL4:BS40 one 1/0 flag per source row per column ("this row passes this box")
      BT4:BT40 the master flag: 1 = the row passes every filled box
  conditional formats on FCS
      X4:X40   the cells a box actually matched light up   (dxfId 3, the yellow the
               other two sheets use)
      A4:AH40  a row that matches nothing goes grey         (dxfId 50, the grey of
               Live Search and Buyer Groups)
      Nothing is hidden: the totals in rows 41/42 and the AutoFilter keep covering
      all 37 lots, and GRAND TOTAL on the other sheets stays cross-checked.
  cascading dropdowns
      one list dataValidation per box -> defined name SUGF_<col> -> a dynamic range
      on the hidden Lists sheet holding exactly those values of that column that
      still survive the other 33 boxes.  That is the same FY/FZ/GA/GB machinery
      SUG_* uses for Live Search and SUGN_* uses for Buyer Groups, cloned.  The one
      column with no value pool yet (Doc./Invoice Date) gets a pool block built in
      the shape of the existing ones.

Why FCS is not sorted the way Buyer Groups can be: Live Search, Buyer Groups and
Lists reference its rows one by one ('Final Calculation Sheet'!$F$4 and friends),
so reordering them would scramble the buyer blocks and every subtotal.  Getting
matches at the top is what Live Search and the Buyer Groups float are for; this
makes the source itself searchable.

    python3 tools/fcs_search_row.py 11.09.2026.xlsm              # dry run -> /tmp
    python3 tools/fcs_search_row.py 11.09.2026.xlsm --apply      # into the workbook
    python3 tools/fcs_search_row.py 11.09.2026.xlsm --out /tmp/x.xlsm --dump
"""

import argparse
import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

from openpyxl.utils import column_index_from_string as cidx
from openpyxl.utils import get_column_letter as gcl

FCS_PART, LISTS_PART, WB_PART = ("xl/worksheets/sheet1.xml", "xl/worksheets/sheet4.xml",
                                 "xl/workbook.xml")
LS_PART, BG_PART, STYLES_PART = "xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml", "xl/styles.xml"

SHEET = "Final Calculation Sheet"
SEARCH_ROW = 2                                 # the row that gets the boxes
DATA = (4, 40)                                 # the 37 lot rows
N_COLS = 34                                    # A..AH
READOUT = 35                                   # AI
HELP0 = 38                                     # AL; AJ/AK stay clear as a buffer
MASTER = HELP0 + N_COLS                        # BT: 1 = the row passes every box
LIST_ROWS = DATA                                 # rows the Lists engine uses
NAME_PREFIX = "SUGF_"
DXF_HIT, DXF_DIM = 3, 50
LISTS_FIRST_NEW = 453                            # Lists is used up to QJ (452)
LIST_NAME_END = LIST_ROWS[1] + 2                 # …:$42, the padding the other families use

CELL_RE = re.compile(r'<c r="([A-Z]{1,3})\d+"([^>/]*?)(?:/>|>(.*?)</c>)', re.S)
ROW_RE = re.compile(r'<row r="(\d+)"([^>/]*)(?:/>|>(.*?)</row>)', re.S)
LABEL_RE = re.compile(r'<c r="([A-Z]{1,3})3"[^>]*t="inlineStr"><is><t[^>]*>([^<]*)</t></is></c>')


def esc(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def die(msg):
    sys.exit(f"ABORT: {msg}")


def formula(text):
    return f"<f>{esc(text)}</f>"


# ------------------------------------------------------------------ sheet plumbing

def parse_rows(body):
    """{row: (row attrs, {column index: (cell attrs, inner or None)})}"""
    out = {}
    for m in ROW_RE.finditer(body):
        cells = {}
        for c in CELL_RE.finditer(m.group(3) or ""):
            cells[cidx(c.group(1))] = (c.group(2) or "", c.group(3))
        out[int(m.group(1))] = (m.group(2) or "", cells)
    return out


def render_row(num, attrs, cells):
    if not cells:
        return f'<row r="{num}"{attrs}/>'
    inner = "".join(
        f'<c r="{gcl(col)}{num}"{a}/>' if text is None else f'<c r="{gcl(col)}{num}"{a}>{text}</c>'
        for col, (a, text) in sorted(cells.items()))
    return f'<row r="{num}"{attrs}>{inner}</row>'


def split_sheet(xml):
    head, rest = xml.split("<sheetData>", 1)
    body, tail = rest.split("</sheetData>", 1)
    return head, body, tail


def rebuild(xml, rows):
    head, body, tail = split_sheet(xml)
    if len(ROW_RE.findall(body)) != len(rows):
        die("the row count changed under me - refusing to write a half-parsed sheet")
    return (head + "<sheetData>"
            + "".join(render_row(r, *rows[r]) for r in sorted(rows))
            + "</sheetData>" + tail)


# ------------------------------------------------------------------ the formulas

def flag(col, row):
    """Does source row `row` pass the box above column `col`?  Empty box = everything
    passes; (blank)/(nonblank) test emptiness; otherwise SEARCH() = contains, any
    case - the identical rule Live Search and Buyer Groups use, so a search behaves
    the same wherever it is typed."""
    box, here = f"${col}${SEARCH_ROW}", f"{col}{row}"
    return (f'IF({box}="",1,IF({box}="(blank)",--({here}=""),'
            f'IF({box}="(nonblank)",--({here}<>""),--ISNUMBER(SEARCH({box},{here})))))')


def master(row, first, last):
    return f'IF(COUNTIF(${gcl(first)}{row}:${gcl(last)}{row},0)>0,0,1)'


def other_filters(row, own, first, last):
    """Lists: does source row `row` still stand once every OTHER box is applied?"""
    s = f"'{SHEET}'!"
    if own == first:
        return f'IF(COUNTIF({s}${gcl(own + 1)}{row}:${gcl(last)}{row},0)>0,0,1)'
    if own == last:
        return f'IF(COUNTIF({s}${gcl(first)}{row}:${gcl(own - 1)}{row},0)>0,0,1)'
    return (f'IF(COUNTIF({s}${gcl(first)}{row}:${gcl(own - 1)}{row},0)'
            f'+COUNTIF({s}${gcl(own + 1)}{row}:${gcl(last)}{row},0)>0,0,1)')


def candidate_present(pool_col, fcs_col, flag_col, row):
    a, b = LIST_ROWS
    src = f"'{SHEET}'!${fcs_col}${a}:${fcs_col}${b}"
    return (f'IF(${pool_col}${row}="",FALSE,SUMPRODUCT(({src}=${pool_col}${row})'
            f'*(${flag_col}${a}:${flag_col}${b}=1))>0)')


def running_count(present_col, count_col, row):
    if row == LIST_ROWS[0]:
        return f'--(${present_col}${row})'
    return f'=${count_col}{row - 1}+--(${present_col}{row})'


def dropdown_list(pool_col, count_col):
    a, b = LIST_ROWS
    return (f'IF(ROW()-3>MAX(${count_col}${a}:${count_col}${b}),"",'
            f'INDEX(${pool_col}${a}:${pool_col}${b},'
            f'MATCH(ROW()-3,${count_col}${a}:${count_col}${b},0)))')


def pool_block(base, fcs_col):
    """A value pool in the shape of the existing ones: raw column, rank, sorted
    values, first-of-value flag, deduplicated list (the fifth column is the pool)."""
    a, b = LIST_ROWS
    raw, rank, uniq, cnt, _list = (gcl(base + k) for k in range(5))
    src = f"'{SHEET}'!${fcs_col}"
    out = {}
    for r in range(a, b + 1):
        out[(base + 0, r)] = f'IF({src}${r}="","",{src}${r})'
        out[(base + 1, r)] = (f'IF({raw}{r}="","",COUNTIFS({raw}${a}:{raw}${b},"<"&{raw}{r},'
                              f'{raw}${a}:{raw}${b},"<>"&"")+COUNTIF({raw}${a}:{raw}{r},{raw}{r}))')
        out[(base + 2, r)] = (f'IF(ROW()-3>COUNTIF({raw}${a}:{raw}${b},"<>"&""),"",'
                              f'INDEX({raw}${a}:{raw}${b},MATCH(ROW()-3,{rank}${a}:{rank}${b},0)))')
        first_flag = f'IF(AND({uniq}{r}<>"",COUNTIF({uniq}${a}:{uniq}{r},{uniq}{r})=1),1,0)'
        out[(base + 3, r)] = first_flag if r == a else f'={cnt}{r - 1}+{first_flag[1:]}'
        out[(base + 4, r)] = (f'IF(ROW()-3>MAX({cnt}${a}:{cnt}${b}),"",'
                              f'INDEX({uniq}${a}:{uniq}${b},MATCH(ROW()-3,{cnt}${a}:{cnt}${b},0)))')
    return out


def sug_name(col, list_col, count_col):
    """Same shape as SUG_*/SUGN_*: an absolute dynamic range, padded to row 42."""
    a, b = LIST_ROWS
    return (f'<definedName name="{NAME_PREFIX}{col}">Lists!${list_col}${a}'
            f':INDEX(Lists!${list_col}${a}:${list_col}${LIST_NAME_END},'
            f'MAX(Lists!${count_col}${a}:${count_col}${b})+2)</definedName>')


# ------------------------------------------------------------------ the build

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsm")
    ap.add_argument("--apply", action="store_true", help="rewrite the workbook in place")
    ap.add_argument("--out", help="where the dry run writes (default /tmp/fcs_search.xlsm)")
    ap.add_argument("--dump", action="store_true", help="print the new row shapes")
    args = ap.parse_args()

    report = []
    with zipfile.ZipFile(args.xlsm) as z:
        raw = {n: z.read(n).decode("utf-8") for n in (FCS_PART, LISTS_PART, WB_PART, BG_PART, LS_PART)}
    fcs, lst, wbx, bgx, lsx = (raw[n] for n in (FCS_PART, LISTS_PART, WB_PART, BG_PART, LS_PART))

    # ---- guards: right file, not already converted
    if "<dataValidations" in fcs:
        die(f"{SHEET} already carries dataValidations - looks converted already")
    if NAME_PREFIX in wbx:
        die(f"{NAME_PREFIX}* defined names already exist - looks converted already")
    if ">FCS " in lst:
        die("an 'FCS …' label already exists on Lists - looks converted already")

    head_f, body_f, tail_f = split_sheet(fcs)
    rows_f = parse_rows(body_f)
    rows_l = parse_rows(split_sheet(lst)[1])
    for r in range(DATA[0], DATA[1] + 1):
        if r not in rows_f or len([c for c in rows_f[r][1] if c <= N_COLS]) != N_COLS:
            die(f"FCS row {r} does not hold all {N_COLS} columns - not the layout expected")
    if SEARCH_ROW not in rows_f:
        die(f"FCS row {SEARCH_ROW} is missing - that is where the boxes go")
    for c in range(1, N_COLS + 1):
        if (rows_f[SEARCH_ROW][1].get(c, ("", None))[1]) not in (None, ""):
            die(f"{gcl(c)}{SEARCH_ROW} is not empty - the search row would eat your data")
    if re.search(r'<col[^>]*min="(3[89]|4\d|[5-7]\d)"', head_f):
        die("something already occupies the helper columns AL..BT")

    # borrow the look of the Buyer Groups search row (styles + row height)
    m_box, m_lab = re.search(r'<c r="B4" s="(\d+)"', bgx), re.search(r'<c r="A4" s="(\d+)"', bgx)
    if not (m_box and m_lab):
        die("could not read the search-box styles from Buyer Groups row 4")
    S_BOX, S_LABEL = f' s="{m_box[1]}"', f' s="{m_lab[1]}"'
    # the height of Live Search's own search row, so the dropdowns have room to breathe
    m_h = re.search(r'<row r="4"[^>]*\sht="([\d.]+)"', lsx)
    style = re.search(r'\ss="(\d+)"', rows_f[SEARCH_ROW][0])
    height = f' ht="{m_h[1]}" customHeight="1"' if m_h else ""
    height += f' s="{style[1]}"' if style else ""

    # which Lists pool feeds which column?  read off the Labels in Lists row 3
    pools = {}
    for m in LABEL_RE.finditer(split_sheet(lst)[1]):
        hit = re.search(r"\(Final ([A-Z]{1,2})\)\s*$", m.group(2).strip())
        if hit:
            pools[hit.group(1)] = cidx(m.group(1)) + 4
    missing = [gcl(c) for c in range(1, N_COLS + 1) if gcl(c) not in pools]
    report.append(f"{N_COLS - len(missing)} of {N_COLS} columns already have a value pool on Lists; "
                  f"a fresh one is built for {'/'.join(missing) or 'nobody'}")

    # ---- 1. the boxes and the read-out
    rows_f[SEARCH_ROW] = (height, dict(rows_f[SEARCH_ROW][1]))
    for c in range(1, N_COLS + 1):
        rows_f[SEARCH_ROW][1][c] = (S_BOX, None)
    read_out = (
        f'IF(COUNTIF($A${SEARCH_ROW}:$AH${SEARCH_ROW},"<>")=0,'
        f'"\U0001f50d SEARCH - nothing typed yet, all 37 rows are shown.   ·   type in any yellow box '
        f'above a column, or pick from its dropdown: every column is searched, the boxes are ANDed '
        f'together, (blank)/(nonblank) and * / ? work as usual",'
        f'"\U0001f50d MATCHED "&SUM(${gcl(MASTER)}${DATA[0]}:${gcl(MASTER)}${DATA[1]})&" of 37 row(s)'
        f'   ·   the rows that miss are greyed out, nothing is hidden and the totals still cover all 37'
        f'   ·   clear a box to widen the dropdowns again")')
    rows_f[SEARCH_ROW][1][READOUT] = (S_LABEL, formula(read_out))
    report.append(f"row {SEARCH_ROW}: {N_COLS} search boxes A{SEARCH_ROW}:AH{SEARCH_ROW} in the yellow "
                  f"style of the other sheets' search rows, {height.strip() or 'height left as is'}"
                  f", live read-out in AI{SEARCH_ROW}")

    # ---- 2. the hidden flags
    first, last = HELP0, HELP0 + N_COLS - 1
    for r in range(DATA[0], DATA[1] + 1):
        rows_f[r] = (rows_f[r][0], dict(rows_f[r][1]))
        for i in range(N_COLS):
            rows_f[r][1][first + i] = (' s="137"', formula(flag(gcl(1 + i), r)))
        rows_f[r][1][MASTER] = (' s="137"', formula(master(r, first, last)))
    report.append(f"helpers: {N_COLS} flag columns {gcl(first)}:{gcl(last)} + master {gcl(MASTER)} "
                  f"on rows {DATA[0]}..{DATA[1]}, hidden and outline-grouped")

    # ---- 3. conditional formats
    dim_formula = (f'AND(COUNTIF($A${SEARCH_ROW}:$AH${SEARCH_ROW},"<>")>0,'
                   f'${gcl(MASTER)}{DATA[0]}=0)')
    cf = [f'<conditionalFormatting sqref="A{DATA[0]}:AH{DATA[1]}">'
          f'<cfRule type="expression" priority="1" dxfId="{DXF_DIM}">'
          f'<formula>{esc(dim_formula)}</formula></cfRule></conditionalFormatting>']
    for i in range(N_COLS):
        col, helpcol = gcl(1 + i), gcl(first + i)
        hit_formula = f'AND(${col}${SEARCH_ROW}<>"",${helpcol}{DATA[0]}=1)'
        cf.append(f'<conditionalFormatting sqref="{col}{DATA[0]}:{col}{DATA[1]}">'
                  f'<cfRule type="expression" priority="{2 + i}" dxfId="{DXF_HIT}">'
                  f'<formula>{esc(hit_formula)}</formula></cfRule></conditionalFormatting>')
    old_cf = re.findall(r"<conditionalFormatting .*?</conditionalFormatting>", tail_f, re.S)
    keep = "".join(re.sub(r'priority="(\d+)"', lambda m: f'priority="{int(m[1]) + 1 + N_COLS}"', b)
                   for b in old_cf)
    tail_new = tail_f.replace("".join(old_cf), "".join(cf) + keep, 1)
    if old_cf and tail_new == tail_f:
        die("could not slot the new conditional formats in next to the existing ones")
    dvs = ('<dataValidations count="%d">' % N_COLS + "".join(
        f'<dataValidation sqref="{gcl(1 + i)}{SEARCH_ROW}" showDropDown="0" showInputMessage="0" '
        f'showErrorMessage="0" allowBlank="0" type="list"><formula1>{NAME_PREFIX}{gcl(1 + i)}'
        f'</formula1><formula2>0</formula2></dataValidation>' for i in range(N_COLS)) + "</dataValidations>")
    if "<pageMargins" not in tail_new:
        die("unexpected FCS tail - no <pageMargins> to place the dataValidations before")
    tail_new = tail_new.replace("<pageMargins", dvs + "<pageMargins", 1)

    head_new = re.sub(r'<dimension ref="[^"]*"\s*/>', f'<dimension ref="A1:{gcl(MASTER)}46"/>',
                      head_f, count=1)
    coldef = (f'<col min="{first}" max="{MASTER}" width="11" customWidth="1" hidden="1" '
              f'outlineLevel="1" style="137"/>')
    if "</cols>" not in head_new:
        die("FCS has no <cols> block to hide the helper columns in")
    head_new = head_new.replace("</cols>", coldef + "</cols>", 1)
    fcs_out = (head_new + "<sheetData>"
               + "".join(render_row(r, *rows_f[r]) for r in sorted(rows_f))
               + "</sheetData>" + tail_new)

    # ---- 4. Lists: the missing pool, then one cascade per column
    nxt = LISTS_FIRST_NEW
    for col in missing:
        for (c, r), text in pool_block(nxt, col).items():
            rows_l.setdefault(r, ("", {}))[1][c] = (' t="str"', formula(text))
        rows_l.setdefault(3, ("", {}))[1][nxt] = (
            ' t="inlineStr"', f'<is><t xml:space="preserve">FCS {col} (Final {col})</t></is>')
        pools[col] = nxt + 4
        report.append(f"Lists: value pool for column {col} at {gcl(nxt)}:{gcl(nxt + 4)}")
        nxt += 5
    cbase = nxt
    names = {}
    labels = ("FY (all-other-filters flag)", "FZ (candidate present flag)",
              "GA (running count)", "GB (dynamic dropdown list)")
    for i in range(N_COLS):
        col = gcl(1 + i)
        pool, flagc = gcl(pools[col]), gcl(cbase + i * 4)
        presc, cntc, listc = gcl(cbase + i * 4 + 1), gcl(cbase + i * 4 + 2), gcl(cbase + i * 4 + 3)
        names[col] = (listc, cntc)
        for k, lab in enumerate(labels):
            rows_l.setdefault(3, ("", {}))[1][cbase + i * 4 + k] = (
                ' t="inlineStr"', f'<is><t xml:space="preserve">FCS {col} cascade: {lab}</t></is>')
        for r in range(LIST_ROWS[0], LIST_ROWS[1] + 1):
            cells = rows_l.setdefault(r, ("", {}))[1]
            cells[cbase + i * 4 + 0] = (' t="str"', formula(other_filters(r, first + i, first, last)))
            cells[cbase + i * 4 + 1] = (' t="str"', formula(candidate_present(pool, col, flagc, r)))
            cells[cbase + i * 4 + 2] = (' t="str"', formula(running_count(presc, cntc, r)))
            cells[cbase + i * 4 + 3] = (' t="str"', formula(dropdown_list(pool, cntc)))
    lst_end = cbase + N_COLS * 4 - 1
    lst_out = rebuild(lst, rows_l)
    lst_out = re.sub(r'<dimension ref="[^"]*"\s*/>', f'<dimension ref="A1:{gcl(lst_end)}114"/>',
                     lst_out, count=1)
    report.append(f"Lists: {N_COLS} cascade blocks {gcl(cbase)}:{gcl(lst_end)} "
                  f"(other-filters flag / candidate present / running count / dropdown list)")

    # ---- 5. the names the dropdowns read
    dn = "".join(sug_name(gcl(1 + i), *names[gcl(1 + i)]) for i in range(N_COLS))
    if "<definedNames>" in wbx:
        wb_new = wbx.replace("</definedNames>", dn + "</definedNames>", 1)
        wb_new = re.sub(r'<definedNames count="(\d+)">',
                        lambda m: f'<definedNames count="{int(m[1]) + N_COLS}">', wb_new, count=1)
    else:
        wb_new = wbx.replace("</sheets>", "</sheets><definedNames>" + dn + "</definedNames>", 1)
    report.append(f"workbook.xml: {N_COLS} names {NAME_PREFIX}A..{NAME_PREFIX}AH, built exactly "
                  f"like the SUG_*/SUGN_* families")

    verify(args.xlsm, fcs_out, lst_out, wb_new, rows_f, rows_l, cbase, lst_end, report,
           {FCS_PART: fcs_out, LISTS_PART: lst_out, WB_PART: wb_new}, len(missing))
    if args.dump:
        for r in (SEARCH_ROW, DATA[0], DATA[1]):
            print(f"  FCS row {r:>2} ({len(rows_f[r][1])} cells): "
                  + " ".join(gcl(c) for c in sorted(rows_f[r][1])[-8:]))
        print(f"  Lists row {LIST_ROWS[0]} tail: "
              + " ".join(gcl(c) for c in sorted(rows_l[LIST_ROWS[0]][1])[-6:]))
    out = args.xlsm if args.apply else (args.out or "/tmp/fcs_search.xlsm")
    write(args.xlsm, out, {FCS_PART: fcs_out, LISTS_PART: lst_out, WB_PART: wb_new}, report)
    for line in report:
        print("  •", line)
    print(f"\nwrote {out}" + ("" if args.apply else "   (dry run - pass --apply for the workbook)"))


# ------------------------------------------------------------------ verification

def verify(xlsm, fcs_out, lst_out, wb_new, rows_f, rows_l, cbase, lst_end, report, repl, n_pool):
    for tag, xml in (("FCS", fcs_out), ("Lists", lst_out), ("workbook", wb_new)):
        try:
            ET.fromstring(xml)
        except ET.ParseError as exc:
            die(f"{tag} is not well-formed XML: {exc}")
    report.append("all three rewritten parts parse as well-formed XML")

    if len(re.findall(r"<dataValidation ", fcs_out)) != N_COLS:
        die("the dataValidations block is the wrong size")
    for i in range(N_COLS):
        if f'name="{NAME_PREFIX}{gcl(1 + i)}"' not in wb_new:
            die(f"defined name {NAME_PREFIX}{gcl(1 + i)} was not written")
    got = "".join(re.findall(rf'<definedName name="{NAME_PREFIX}[^"]+">[^<]+</definedName>', wb_new))
    if not got:
        die("no SUGF_ names found after building them")
    for m in re.finditer(rf'<definedName name="{NAME_PREFIX}([A-Z]+)">'
                         rf'Lists!\$([A-Z]+)\$(\d+):INDEX\(Lists!\$([A-Z]+)\$(\d+):\$([A-Z]+)(\d+),'
                         rf'MAX\(Lists!\$([A-Z]+)\$(\d+):\$([A-Z]+)(\d+)\)\+2\)</definedName>', wb_new):
        col, a1, r1, a2, r2, a3, r3, a4, r4, a5, r5 = m.groups()
        if a1 != a2 or a4 != a5 or (int(r1), int(r2)) != LIST_ROWS or int(r3) != LIST_NAME_END:
            die(f"{NAME_PREFIX}{col} does not span the rows the other families use: {m.group(0)[:110]}")
        if not (cbase <= cidx(a1) <= lst_end and cbase <= cidx(a4) <= lst_end):
            die(f"{NAME_PREFIX}{col} points outside the new Lists columns")
    report.append(f"the {N_COLS} dropdowns each resolve to a name spanning exactly the new Lists "
                  f"cascade columns, rows {LIST_ROWS[0]}..{LIST_NAME_END} like SUG_*/SUGN_*")

    with zipfile.ZipFile(xlsm) as z:
        old_f = parse_rows(split_sheet(z.read(FCS_PART).decode("utf-8"))[1])
        old_l = parse_rows(split_sheet(z.read(LISTS_PART).decode("utf-8"))[1])
        _old_head, _old_body, old_tail = split_sheet(z.read(FCS_PART).decode("utf-8"))
        n_old_cf = len(re.findall(r"<cfRule", old_tail))
        dxfs = len(re.findall("<dxf>", z.read(STYLES_PART).decode("utf-8")))

    for r, (attrs, cells) in old_f.items():
        for col, (a, inner) in cells.items():
            got_cell = rows_f.get(r, ("", {}))[1].get(col)
            if got_cell is None:
                die(f"FCS {gcl(col)}{r} vanished")
            if (got_cell[1] or "") != (inner or ""):
                die(f"FCS {gcl(col)}{r} content changed - only row {SEARCH_ROW} (which is empty) is touched")
    for r in range(DATA[0], DATA[1] + 1):
        cells = rows_f[r][1]
        for c in range(HELP0, MASTER + 1):
            if "<f>" not in (cells.get(c, ("", ""))[1] or ""):
                die(f"FCS row {r} is missing its helper formula in {gcl(c)}")
        for c in range(1, N_COLS + 1):
            if c not in cells:
                die(f"FCS {gcl(c)}{r} lost its own cell")
    report.append(f"every pre-existing FCS cell kept its content; "
                  f"{N_COLS * (DATA[1] - DATA[0] + 1) + DATA[1] - DATA[0] + 1} helper formulas live only "
                  f"in the hidden columns")

    for r, (attrs, cells) in old_l.items():
        for col, (a, inner) in cells.items():
            got_cell = rows_l.get(r, ("", {}))[1].get(col)
            if got_cell is None or (got_cell[1] or "") != (inner or ""):
                die(f"Lists {gcl(col)}{r} was modified - only new columns may change")
    for r, (attrs, cells) in rows_l.items():
        for col in cells:
            if col < LISTS_FIRST_NEW and col not in old_l.get(r, ("", {}))[1]:
                die(f"a cell appeared inside the existing Lists area at {gcl(col)}{r}")
    report.append(f"Lists: the {cidx('QJ')} columns that were there are byte-identical; "
                  f"{n_pool * 5 + N_COLS * 4} new columns start at {gcl(LISTS_FIRST_NEW)}")

    pr = [int(v) for v in re.findall(r'priority="(\d+)"', fcs_out)]
    if len(pr) != len(set(pr)):
        die(f"conditional-format priorities collide: {sorted(pr)}")
    if min(pr) != 1 or len(pr) != N_COLS + 1 + n_old_cf:
        die(f"expected {N_COLS + 1} new rules on top of the {n_old_cf} existing ones, got {len(pr)}")
    dxf_used = [int(v) for v in re.findall(r'<cfRule[^>]*dxfId="(\d+)"', fcs_out)]
    if max(dxf_used) >= dxfs:
        die("a new rule points past the end of the dxf list")
    if DXF_DIM not in dxf_used or DXF_HIT not in dxf_used:
        die("the grey/yellow dxfIds are not in the new rules")
    report.append(f"{len(pr)} conditional formats, unique priorities 1..{max(pr)}, reusing dxfIds "
                  f"{DXF_HIT} (yellow) and {DXF_DIM} (grey) from the existing {dxfs} - styles.xml untouched")

    for probe in ('<autoFilter ref="A3:AH43"', 'xSplit="7" ySplit="3"', 'codeName="Sheet1"'):
        if probe not in fcs_out:
            die(f"FCS lost {probe}")
    if f'<dimension ref="A1:{gcl(MASTER)}46"/>' not in fcs_out:
        die("the FCS dimension was not extended over the new columns")
    report.append("the frozen pane (rows 1-3, so the boxes stay on screen), the AutoFilter and the "
                  "sheet code-name are untouched; the dimension covers the new columns")

    try:
        import openpyxl
        probe = "/tmp/fcs_search_openpyxl.xlsm"
        write(xlsm, probe, repl, None)
        wb = openpyxl.load_workbook(probe)
        ws = wb[SHEET]
        ndv = len(ws.data_validations.dataValidation)
        ncf = len(list(ws.conditional_formatting))
        nsm = len([n for n in wb.defined_names if n.startswith(NAME_PREFIX)])
        if ndv != N_COLS:
            die(f"openpyxl counts {ndv} validations on FCS, expected {N_COLS}")
        if ncf != N_COLS + 3:          # 34 highlights + 1 grey-out + the 2 blocks already there
            die(f"openpyxl counts {ncf} conditional-format blocks, expected {N_COLS + 3}")
        if nsm != N_COLS:
            die(f"openpyxl sees {nsm} {NAME_PREFIX} names, expected {N_COLS}")
        if not ws.row_dimensions[SEARCH_ROW].height:
            die(f"row {SEARCH_ROW} has no height - the boxes would be squashed")
        if not ws.column_dimensions[gcl(HELP0)].hidden:
            die("the helper columns did not come out hidden")
        other = wb["Buyer Groups"]
        if len(other.data_validations.dataValidation) != 34:
            die("Buyer Groups lost its dropdowns - something spilled across sheets")
        report.append(f"openpyxl re-opens the result cleanly: {ndv} FCS validations, {ncf} CF blocks, "
                      f"{nsm} names, row {SEARCH_ROW} height {ws.row_dimensions[SEARCH_ROW].height}, "
                      f"helpers hidden; Buyer Groups keeps its 34 dropdowns and "
                      f"{len(list(other.conditional_formatting))} CF blocks")
    except ImportError:
        report.append("(openpyxl not importable here - skipped the re-open check)")


# ------------------------------------------------------------------ zip i/o

def write(xlsm, out, replacements, report):
    if replacements is None:
        import shutil
        shutil.copyfile(xlsm, out)
        return
    tmp = out + ".tmp"
    with zipfile.ZipFile(xlsm) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in replacements:
                data = replacements[item.filename].encode("utf-8")
            zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            zi.compress_type = item.compress_type
            zi.external_attr = item.external_attr
            zout.writestr(zi, data)
    os.replace(tmp, out)
    with zipfile.ZipFile(xlsm) as a, zipfile.ZipFile(out) as b:
        if a.namelist() != b.namelist():
            die("the part list changed")
        untouched = [n for n in a.namelist() if n not in replacements]
        bad = [n for n in untouched if a.read(n) != b.read(n)]
        if bad:
            die(f"these parts must not change: {bad}")
    if report is not None:
        report.append(f"{len(untouched)} of {len(a.namelist())} parts copied byte for byte, including "
                      f"xl/vbaProject.bin (no macro touched), xl/styles.xml and the other sheets")


if __name__ == "__main__":
    main()
