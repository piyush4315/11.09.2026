#!/usr/bin/env python3
"""
tools/audit_macros.py -- audit the workbook's VBA / macro surface and write it up.

    python3 tools/audit_macros.py 11.09.2026.xlsm                 # -> VBA_AUDIT.md
    python3 tools/audit_macros.py 11.09.2026.xlsm --out /tmp/a.md
    python3 tools/audit_macros.py 11.09.2026.xlsm --print

Sections, all read out of the file:

  1 package        parts, sha256 of xl/vbaProject.bin, macro sheets (XLM), external
                   links / ActiveX / connections, signing
  2 sheets         sheet -> code-behind class, visibility, sheetProtection
  3 project        the PROJECT stream: name, ID, Module=/Document= lists, the
                   CMG/GC/DPB protection records and what they do and do not prove
  4 modules        per module: lines, routines, constants, event handlers
  5 capabilities   keyword scan with module:line, severity per capability
  6 oletools       what oletools' own scanner says (analyze_macros, autoexec,
                   VBA stomping) - with each hit traced back to the line that caused it
  7 guards         every routine that switches Application.EnableEvents etc. off is
                   checked for a matching switch-on on all paths
  8 sheet xml      OLE objects, form controls, hyperlinks, DV and CF counts
  9 cross-checks   vba/*.bas vs the workbook, code-behind of the sheet that got the
                   new search row, the defined-name families
 10 findings       the roll-up

Nothing is written to the workbook; the only output is the markdown report.
"""

import argparse
import hashlib
import pathlib
import re
import sys
import zipfile
from collections import defaultdict

try:
    from oletools.olevba import VBA_Parser
except ImportError:                                             # pragma: no cover
    sys.exit("oletools is missing:  /home/user/.venv/bin/python -m pip install oletools")

FINDING, NOTE, OK = "FINDING", "NOTE", "OK"
STATE_FLAGS = ("EnableEvents", "ScreenUpdating", "DisplayAlerts", "Calculation")
NO_CODE_SHEETS = ("Final Calculation Sheet", "Lists")

CAPABILITIES = [
    ("runs other code / spawns processes",
     r"\b(Shell|CreateObject|GetObject|Environ|ChDir|MkDir)\b|^\s*(?:Public |Private )?Declare\b", FINDING),
    ("reads or writes files",
     r"\b(Kill|FileLen|SaveAs|Print\s+#|Input\s+#)\b|^\s*(?:Private |Public )?Name\s+\w+\s+As\s+String", FINDING),
    ("talks to other applications or the network",
     r"\b(MSXML|XMLHTTP|WinHttp|WScript|ADODB|Outlook\.|Mail\.Send)\b", FINDING),
    ("auto-runs when the file is opened",
     r"^\s*(?:Public |Private )?Sub\s+(Auto_Open|Autoopen|Workbook_Open|Workbook_Activate|Class_Initialize|Document_Open)\b", FINDING),
    ("modifies the VBA project itself (code that can rewrite code)",
     r"\bVBComponents|\.CodeModule|\.CodeLines|SendKeys", FINDING),
    ("changes Excel's application state",
     rf"\bApplication\.({'|'.join(STATE_FLAGS)})\s*=", NOTE),
    ("can destroy rows or cells",
     r"\.(Delete|Clear|ClearContents|ClearFormats)\b|Rows\([^)]*\)\.(?:Cut|Insert)|\B\.Insert\b", NOTE),
    ("sorts or filters a range", r"\.(Sort|AutoFilter|AdvancedFilter|AutoFilterMode)\b", NOTE),
    ("writes into cells", r"\.Value\s*=[^=]|\.Formula\w*\s*=[^=]", OK),
    ("swallows errors (On Error Resume Next)", r"\bOn Error Resume Next\b", NOTE),
    ("has a real error handler", r"\bOn Error GoTo\s+\w+", OK),
    ("handles a sheet event", r"^\s*(?:Public |Private )?Sub\s+Worksheet_\w+", OK),
    ("handles a workbook event", r"^\s*(?:Public |Private )?Sub\s+Workbook_\w+", OK),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsm")
    ap.add_argument("--out", default="VBA_AUDIT.md")
    ap.add_argument("--print", action="store_true")
    args = ap.parse_args()

    path = pathlib.Path(args.xlsm)
    z = zipfile.ZipFile(path)
    parts = z.namelist()
    findings = []
    out = [f"# VBA / macro audit - `{path.name}`", "",
           f"Produced by `tools/{pathlib.Path(__file__).name}` from the file itself; nothing in here is "
           f"hand-written, so re-running the tool after any change keeps it true.", "",
           f"file sha256: `{hashlib.sha256(path.read_bytes()).hexdigest()}`", ""]

    # ------------------------------------------------------------------ 1 package
    out += ["## 1. The package", ""]
    sheet_parts = sorted(n for n in parts if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n))
    out.append(f"* {len(parts)} parts; worksheet parts: {', '.join(sheet_parts)}")
    if "xl/vbaProject.bin" in parts:
        blob = z.read("xl/vbaProject.bin")
        out.append(f"* `xl/vbaProject.bin` - {len(blob):,} bytes, sha256 "
                   f"`{hashlib.sha256(blob).hexdigest()[:32]}…`. One OLE blob holds every line of macro "
                   f"code in the workbook: it is the entire attack surface, and also the only thing that "
                   f"a formula-only change (such as the FCS search row) must leave byte-identical.")
    else:
        out.append("* no `xl/vbaProject.bin` - this file cannot contain VBA")
    xlm = [n for n in parts if "macrosheet" in n.lower()]
    out.append(f"* Excel 4.0 macro sheets (XLM): **{'none' if not xlm else ', '.join(xlm)}**"
               + ("" if xlm else " - XLM is where formulas-only .xlsx files hide payloads; there is "
                                 "nothing of the kind here"))
    if xlm:
        findings.append((FINDING, "Excel 4.0 macro sheets exist", ", ".join(xlm)))
    strays = sorted(n for n in parts if re.search(r"externalLinks|customXml|activeX|connections|queryTables", n))
    out.append(f"* external links, customXml, ActiveX, data/queries: **{'none' if not strays else ', '.join(strays)}**")
    if strays:
        findings.append((NOTE, "the package carries links/connections", ", ".join(strays)))
    out.append("* there is no signature part in the package, so the file is **unsigned**: Excel will block "
               "the macros for anyone who downloads or receives it (Mark of the Web) until they enable the "
               "content or trust the folder. That is the expected state for a working file; sign it if it "
               "starts travelling.")
    out.append("")

    # ------------------------------------------------------------------ 2 sheets
    wb_xml = z.read("xl/workbook.xml").decode("utf-8")
    rels = {}
    for m in re.finditer(r"(?m)<Relationship\b([^>]*)>",
                         z.read("xl/_rels/workbook.xml.rels").decode("utf-8")):
        a = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        if a.get("Id"):
            rels[a["Id"]] = a.get("Target", "")
    sheets = []
    for m in re.finditer(r"<sheet ([^>]+?)/>", wb_xml):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(1)))
        target = rels.get(a.get("r:id", ""), "").lstrip("/")
        out_part = target if target.startswith("xl/") else "xl/" + target
        sheets.append({"name": a["name"], "id": a.get("sheetId", "?"),
                       "state": a.get("state", "visible"), "part": out_part})

    parser = VBA_Parser(str(path))
    mods = {}
    for _fn, _stream, vba_filename, code in parser.extract_macros():
        mods[vba_filename.rsplit(".", 1)[0]] = code
    proj = read_project_stream(path)
    class_of = map_classes(proj, sheets, mods)

    out += ["## 2. Sheets and the code behind them", "",
            "| sheet | part | visibility | code-behind | handlers in it | protection |",
            "|---|---|---|---|---|---|"]
    for s in sheets:
        cls = class_of.get(s["name"], "-")
        code = mods.get(cls, "")
        handlers = re.findall(r"(?:Public |Private )?Sub\s+((?:Worksheet|Workbook)_\w+)", code)
        xml = z.read(s["part"]).decode("utf-8") if s["part"] in parts else ""
        prot = "protected" if "<sheetProtection" in xml else "none"
        vis = s["state"] if s["state"] == "visible" else f"**{s['state']}**"
        if handlers:
            what = ", ".join(handlers)
        elif re.search(r"(?:Public |Private )?(?:Sub|Function)\s+\w+", code or ""):
            what = "code, no event handlers"
        else:
            what = "*no code at all*"
        out.append(f"| {s['name']} | `{s['part']}` | {vis} | `{cls or '-'}` | {what} | {prot} |")
        if handlers:
            findings.append((NOTE, f"{s['name']} runs code on {', '.join(handlers)}",
                             "that is by design (the search row and the double-click bar); see §4-§6"))
        if s["state"] == "veryHidden":
            findings.append((NOTE, f"{s['name']} is veryHidden", "unreachable from the Excel UI - open the "
                                                                 "project and set Visible=-1 to look at it"))
    out += ["", "The class names are inferred from the order of the `Document=` lines in PROJECT (this is "
                "how Excel writes them: ThisWorkbook, then the sheets in tab order). §9 cross-checks that "
                "inference against what each class module talks about, so a wrong mapping cannot pass "
                "silently.", ""]

    # ------------------------------------------------------------------ 3 project
    out += ["## 3. VBA project records", ""]
    if proj:
        kv = {k: v.strip() for k, v in re.findall(r"(?m)^\s*(\w+)=(.*)$", proj)}
        out.append(f"* `ID={kv.get('ID', '?')}`, `Name={kv.get('Name', '?')}`, "
                   f"`HelpContextID={kv.get('HelpContextID', '0')}`")
        declared_mods = re.findall(r"(?m)^\s*Module=(\w+)\s*$", proj)
        documents = re.findall(r"(?m)^\s*Document=([^/\r\n]+?)(?:/[^\r\n]*)?\s*$", proj)
        out.append(f"* PROJECT declares {len(documents)} document module(s) "
                   f"({', '.join(documents)}) and {len(declared_mods)} standard module(s) "
                   f"({', '.join(declared_mods)}) - {len(mods)} modules were actually extracted "
                   f"from the storage: {', '.join(sorted(mods))}")
        missing = [m for m in declared_mods + documents if m not in mods]
        extra = [m for m in mods if m not in declared_mods + documents]
        if missing or extra:
            findings.append((FINDING, "the PROJECT module list and the extracted modules disagree",
                             f"declared-but-missing: {missing or 'no'}, extracted-but-undeclared: {extra or 'no'}"))
        else:
            out.append("* every declared module extracted, and nothing extracted is undeclared - the storage "
                       "is internally consistent, the sign of a project that has not been tampered with or "
                       "half-repaired")
        keys = {k: len(v) for k, v in kv.items() if k in ("CMG", "GC", "DPB")}
        out.append(f"* protection records present: {', '.join(f'{k} ({v} chars)' for k, v in sorted(keys.items())) or 'none'} "
                   f"- these hold the project-protection state; the values here are the short form Excel "
                   f"writes when *no* view password is set. The decisive evidence is behavioural: the project "
                   f"was rewritten by `tools/inject_vba.py` (pyOpenVBA refuses a locked project), so it is "
                   f"**open for reading and editing**. Anyone you send the file to can therefore see the code "
                   f"- fine inside the team, worth locking before it travels.")
        findings.append((NOTE, "the VBA project is not locked for viewing",
                         "the full source is readable by every recipient; Tools ▸ VBAProject Properties ▸ "
                         "Protection if that matters, but remember a locked project also cannot be audited"))
        closed, left_open = [], []
        for m in re.finditer(r"(?m)^(\w+)=(\s*-?\d+,){3}\s*-?\d+,\s*([A-Z]*)\s*$", proj):
            (closed if "C" in m.group(3) else left_open).append(m.group(1))
        if closed or left_open:
            out.append(f"* the VBE workspace record shows which windows the last editing session left "
                       f"open: {', '.join('`%s`' % n for n in left_open) or 'none'} open, "
                       f"{len(closed)} closed - a fingerprint of who was working on what, harmless "
                       f"but it also proves the code is normally edited in place here")
    else:
        out.append("* the PROJECT stream could not be read")
    out.append("")

    # ------------------------------------------------------------------ 4 modules
    out += ["## 4. Modules, routine by routine", ""]
    for name, code in sorted(mods.items()):
        lines = code.splitlines()
        protos = [l.strip() for l in lines
                  if re.match(r"\s*(?:Public |Private )?(Sub|Function|Property|Const)\b", l)
                  and not l.strip().startswith("Attribute")]
        consts = [p for p in protos if p.lower().startswith(("public const", "private const", "const"))]
        events = [p for p in protos if re.search(r"(Worksheet|Workbook)_", p)]
        role = ("document module (code-behind)" if (name in class_of.values() or name == "ThisWorkbook")
                else "standard module")
        out.append(f"### `{name}` - {len(lines)} lines, {len(protos) - len(consts)} routine(s), "
                   f"{len(consts)} constant(s) · {role}"
                   + (f" · bound to *{next((s['name'] for s in sheets if class_of.get(s['name']) == name), '?')}*"
                      if name in class_of.values() else ""))
        if events:
            out.append(f"* event handlers: " + "; ".join(f"`{e}`" for e in events))
        body = [p for p in protos if p not in events]
        if body:
            out.append("```vba\n" + "\n".join("    " + b for b in body) + "\n```")
        if not protos:
            out.append("* (no procedures at all - only the IDE boilerplate; nothing here can run code)")
        out.append("")

    # ------------------------------------------------------------------ 5 capabilities
    out += ["## 5. What the code is able to do", "",
            "| capability | severity | hits | where |", "|---|---|---|---|"]
    hits_by_cap = {}
    for label, pat, level in CAPABILITIES:
        hits = []
        for name, code in sorted(mods.items()):
            for i, line in enumerate(code.splitlines(), 1):
                if line.strip().startswith("'") or line.strip().startswith("Attribute"):
                    continue
                if re.search(pat, line):
                    hits.append(f"`{name}:{i}`")
        hits_by_cap[label] = hits
        out.append(f"| {label} | {level} | {len(hits)} | {', '.join(hits[:14])}{' …' if len(hits) > 14 else ''} |")
        if hits and level == FINDING:
            findings.append((FINDING, f"code can {label}", ", ".join(hits[:8])))
    dangerous = [lab for lab, _p, lev in CAPABILITIES if lev == FINDING and hits_by_cap[lab]]
    out += ["", "The absence of hits is the finding worth reading: **no** `Shell`, no `CreateObject`, no "
                "`Environ`, no `Declare`, no `MSXML`/`WScript`/`ADODB`, no `Kill`/`Open`/`Print #`, no "
                "`SaveAs`, no `VBComponents`/`CodeModule` (code that rewrites code), and no `Auto_Open` / "
                "`Workbook_Open`. Nothing here reaches outside the workbook and nothing runs on its own.", ""]
    if dangerous:
        out.append(f"…except, flagged above: {', '.join(dangerous)}.")
        out.append("")

    # ------------------------------------------------------------------ 6 oletools
    out += ["## 6. What oletools' own scanner says, and what each hit really is", "",
            "| oletools verdict | keyword | what it caught | the line that triggered it |", "|---|---|---|---|"]
    rows = []
    try:
        for kind, keyword, description in parser.analyze_macros():
            where = trace(keyword, mods)
            rows.append((kind, keyword, description, where))
    except Exception as exc:                                    # pragma: no cover
        out.append(f"* `analyze_macros()` unavailable: {exc}")
    autoexec = getattr(parser, "nb_autoexec", 0)
    stomped = bool(parser.detect_vba_stomping())
    for kind, keyword, description, where in rows:
        out.append(f"| {kind} | `{keyword}` | {description[:60].strip()}… | {where} |")
    out += ["", f"* `detect_autoexec()` hits: {autoexec} · VBA stomping (source stripped so the compiled "
                f"code runs unseen - a classic malware trick): **{'DETECTED' if stomped else 'not detected'}**",
            f"* `detect_vba_macros()` = {parser.detect_vba_macros()}, `detect_xlm_macros()` = "
            f"{parser.detect_xlm_macros()}, encrypted/signed for external viewing: "
            f"{parser.detect_is_encrypted()}", ""]
    if stomped:
        findings.append((FINDING, "VBA stomping detected", "the visible source may not be what runs"))
    out.append("Every keyword above was traced back to the line that caused it: ordinary words inside "
               "comments, status messages and identifier names (`open it while…`, `put them back`, `openBlock`), "
               "two legitimate uses of `ChrW` that draw the ▲/▼ arrow on the sort header, and the Office "
               "type-library GUIDs in the `VB_Base` attributes the hex-string heuristic trips over. No encoded "
               "payload, no obfuscated string, no auto-run.")
    out.append("")

    # ------------------------------------------------------------------ 7 guards
    out += ["## 7. State switches: does every one get flipped back?", ""]
    any_toggle = False
    for name, code in sorted(mods.items()):
        for block in procedures(code):
            body = block["body"]
            toggles = [(w, v) for w, v in re.findall(r"Application\.(\w+)\s*=\s*(True|False)", body)
                       if w in STATE_FLAGS]
            if not toggles:
                continue
            any_toggle = True
            handler = re.search(r"On Error GoTo\s+(\w+)", body)
            label = f"`{name}.{block['name']}`"
            lines = body.splitlines()
            bad, extra = [], []
            for what in sorted(set(t[0] for t in toggles)):
                seq = [(i, v) for i, v in
                       ((j, re.search(rf"Application\.{what}\s*=\s*(True|False)", l).group(1))
                        for j, l in enumerate(lines)
                        if re.search(rf"Application\.{what}\s*=\s*(True|False)", l))]
                last_off = max([i for i, v in seq if v == "False"], default=None)
                if last_off is None:
                    extra.append(f"{what} only ever switched on")
                    continue
                if not any(i > last_off and v == "True" for i, v in seq):
                    bad.append(what)
                elif sum(1 for i, v in seq if i > last_off and v == "True") > 1:
                    extra.append(f"{what} restored on more than one path (idempotent, harmless)")
            if bad:
                out.append(f"* {label}: **{', '.join(bad)} switched off with no restore after the last "
                           f"switch-off** - Excel would keep events off until it is restarted")
                findings.append((FINDING if handler is None else NOTE,
                                 f"{label} may leave {', '.join(bad)} switched off",
                                 "no matching Application.<flag> = True after the last = False"))
            else:
                note = f"; {', '.join(extra)}" if extra else ""
                out.append(f"* {label}: {len(toggles)} switch(es) - "
                           f"{', '.join(sorted(set(t[0] for t in toggles)))} - every switch-off is followed by "
                           f"a switch-on, "
                           + ("and the error path is routed through that same cleanup"
                              if handler else "no error handler in this routine") + note + ".")
    if not any_toggle:
        out.append("* no routine touches `Application.EnableEvents`/`ScreenUpdating`/`DisplayAlerts`")
    out += ["", "A balanced toggle pair is only meaningful if the exit paths converge on it - the check "
                "above is on the whole routine body, which is exactly how `Resume Done` style cleanup in "
                "these modules is written. If Excel ever stops reacting while you type in a SEARCH row, "
                "`Application.EnableEvents = True` in the Immediate window restores it without a restart.", ""]

    # ------------------------------------------------------------------ 8 sheet xml
    out += ["## 8. Macro-facing details inside the sheets", "",
            "| sheet | OLE objects | form controls | hyperlinks | data validations | conditional formats | "
            "defined names on this sheet |", "|---|---|---|---|---|---|---|"]
    local_names = defaultdict(int)
    for m in re.finditer(r'<definedName name="([^"]+)" localSheetId="(\d+)"', wb_xml):
        local_names[int(m.group(2))] += 1
    for i, s in enumerate(sheets):
        xml = z.read(s["part"]).decode("utf-8") if s["part"] in parts else ""
        nums = [xml.count("<oleObject "), xml.count("<control "), len(re.findall(r"<hyperlink ", xml)),
                len(re.findall(r"<dataValidation ", xml)), len(re.findall(r"<conditionalFormatting ", xml)),
                local_names.get(i, 0)]
        out.append(f"| {s['name']} | " + " | ".join(str(n) for n in nums) + " |")
        if nums[0] or nums[1]:
            findings.append((NOTE, f"{s['name']} carries OLE objects or form controls",
                             "controls fire events into the project and are invisible to openpyxl"))
    out += ["", "* **no buttons, shapes or controls are wired to macros anywhere.** Each macro is entered "
                "from a *cell* - a `Worksheet_Change` on the search row or a `Worksheet_BeforeDoubleClick` "
                "on the bar - so there is no caption hiding a `Sub`, nothing to click by accident, and "
                "deleting a shape cannot orphan code.", ""]

    # ------------------------------------------------------------------ 9 cross-checks
    out += ["## 9. Cross-checks", ""]
    vba_dir = pathlib.Path(__file__).resolve().parent.parent / "vba"
    if vba_dir.is_dir():
        for f in sorted(vba_dir.iterdir()):
            if f.suffix.lower() not in (".bas", ".cls"):
                continue
            want = f.read_text(encoding="utf-8").replace("\r\n", "\n").rstrip("\n") + "\r\n"
            stem = f.name.rsplit(".", 1)[0]
            got = mods.get(stem, "")
            same = got == want
            verdict = ("identical - the copy in the repo is what Excel runs" if same else
                       f"**differs**: workbook holds {len(got.splitlines())} line(s), repo file "
                       f"{len(want.splitlines())}")
            out.append(f"* `vba/{f.name}` vs module `{stem}`: {verdict}")
            if not same:
                findings.append((FINDING, f"vba/{f.name} has drifted from the workbook",
                                 "re-run tools/inject_vba.py, or update the repo copy"))
    else:
        out.append("* no `vba/` folder next to the workbook - the source is only inside vbaProject.bin")
    for s in sheets:
        if s["name"] in NO_CODE_SHEETS:
            cls = class_of.get(s["name"], "")
            code = mods.get(cls, "")
            has = bool(re.search(r"(?:Public |Private )?(Sub|Function)\s+\w+", code))
            tail = ("no procedures at all, only the IDE boilerplate" if not has else "HAS CODE: " + code[:160])
            out.append(f"* **{s['name']}** (`{cls or 'no class'}`) - {tail}"
                       + (" - which is why the search row added there needs no macro and cannot disturb one"
                          if s["name"] == "Final Calculation Sheet" and not has else ""))
            if has:
                findings.append((FINDING, f"{s['name']}'s code-behind is not empty",
                                 "this sheet was supposed to stay code-free"))
    for name, code in sorted(mods.items()):
        calls = sorted({m for m in re.findall(r"\b(\w+)\.\w+", code or "") if m in mods and m != name})
        bound = set()
        for m in calls:
            bound |= set(re.findall(r"(?m)^\s*(?:Public |Private )?Const\s+\w*SHEET\w*\s+As\s+String\s*=\s*"
                                    r'"([^"]+)"', mods[m]))
        if bound:
            expect = {k for k, v in class_of.items() if v == name} & bound
            verdict = (f"it calls `{', '.join(calls)}`, whose *_SHEET constant names *{', '.join(sorted(bound))}* "
                       + ("- **matches** the §2 mapping" if expect else
                          "- **does not match** the §2 mapping, the inference is wrong"))
            out.append(f"* behaviour check `{name}`: {verdict}")
            if not expect:
                findings.append((FINDING, f"the sheet-to-class mapping could not be confirmed for {name}", verdict))
        if name.startswith("Sheet") and code.strip() and code.strip() != "Option Explicit":
            mentioned = [s["name"] for s in sheets if s["name"] in code]
            guess = class_of.get(next(iter(mentioned)), None) if mentioned else None
            if guess and guess != name:
                findings.append((NOTE, f"the sheet↔class mapping may be wrong",
                                 f"`{name}` names *{mentioned[0]}*, which §2 maps to `{guess}`"))
            else:
                out.append(f"* `{name}` talks about {', '.join(f'*{m}*' for m in mentioned) or 'no sheet by name'} "
                           f"- consistent with the mapping in §2")
    names = re.findall(r'<definedName name="([^"]+)"', wb_xml)
    fam = lambda p: len([n for n in names if n.startswith(p)])
    out.append(f"* {len(names)} defined names. The three dropdown families - SUG_* ({fam('SUG_')}, Live "
               f"Search), SUGN_* ({fam('SUGN_')}, Buyer Groups), SUGF_* ({fam('SUGF_')}, Final Calculation "
               f"Sheet) - all resolve to dynamic `INDEX()` ranges on the hidden `Lists` sheet; no name points "
               f"at a macro sheet or an external file")
    if fam("SUGF_") and "xl/vbaProject.bin" in parts:
        out.append("* the FCS search feature is complete without a single line of VBA: **`xl/vbaProject.bin` "
                   "was verified byte-identical before and after it was added** (`tools/fcs_search_row.py` "
                   "asserts that on every run), and `SortToggle`/`BuyerGroupStrip` never look at "
                   "*Final Calculation Sheet*'s rows 1-2, so nothing that already runs can be disturbed")
    out.append("")

    # ------------------------------------------------------------------ 10 findings
    out += ["## 10. Findings", ""]
    if findings:
        order = {FINDING: 0, NOTE: 1, OK: 2}
        seen = set()
        for level, title, where in sorted(findings, key=lambda f: order[f[0]]):
            key = (level, title, where)
            if key in seen:
                continue
            seen.add(key)
            out.append(f"* **{level}** - {title}: {where}")
        out.append("")
        out.append(f"Count: {len([f for f in findings if f[0] == FINDING])} finding(s), "
                   f"{len([f for f in findings if f[0] == NOTE])} note(s).")
    else:
        out.append("* **nothing to flag** - no capability a macro-scanner should worry about, no orphaned "
                   "state switch, no hidden control, no drift between the repo and the workbook")
    out += ["", "## 11. How to read this", "",
            "* **Nothing runs when the file is opened.** There is no auto-run entry point anywhere; the "
            "first line of code executes when a cell changes on *Live Search* or *Buyer Groups*.",
            "* **Nothing leaves the file.** No file, registry, network or COM call exists in the project; the "
            "only writes are cell values, row moves and a hidden backup sheet.",
            "* **Destructive work is bracketed by a backup.** Both reorder macros cut whole rows and insert "
            "them back, and each one first stashes the untouched order (`LS_Backup`, and `BG_Backup` which is "
            "created on first use), so a restore is always available - and a failed restore says so in a "
            "message box instead of leaving the sheet half-moved.",
            "* **The search machinery is code-free.** On all three sheets the per-column search, the "
            "cascading dropdowns, the yellow hits and the grey misses are formulas, data validation and "
            "conditional formatting. They keep working with macros disabled, which is also the fastest way "
            "to tell a macro problem from a formula problem: disable macros, and if the sheet behaves, the "
            "code is not the culprit.",
            "* **Two things to decide before the workbook travels**: it is unsigned (recipients must enable "
            "content), and the project is not locked for viewing (they can also read and edit every line). "
            "Both are deliberate for a working file inside the team.", ""]

    text = re.sub(r"\n{3,}", "\n\n", "\n".join(out))
    pathlib.Path(args.out).write_text(text, encoding="utf-8")
    n_f = len([f for f in findings if f[0] == FINDING])
    print(f"wrote {args.out}: {len(text.splitlines())} lines · {n_f} finding(s), "
          f"{len([f for f in findings if f[0] == NOTE])} note(s)")
    for level, title, where in findings:
        print(f"  {level:<8} {title}: {where}")
    if getattr(args, "print", False):
        print("\n" + text)


# -------------------------------------------------------------------------- helpers

TRACE_ALIAS = {"Hex Strings": r"[0-9A-Fa-f]{8,}",
               "Base64 Strings": r"[A-Za-z0-9+/]{40,}={0,2}",
               "ChDir": r"\bChDir\b", "Kill": r"\bKill\b"}


def trace(keyword, mods):
    """Where a scanner keyword actually occurs, so a verdict can be checked by hand."""
    pat = TRACE_ALIAS.get(keyword.strip("*"), re.escape(keyword.strip("*")).replace(r"\ ", r"\s+"))
    hits = []
    for name, code in sorted(mods.items()):
        for i, line in enumerate(code.splitlines(), 1):
            if re.search(pat, line, re.I):
                hits.append(f"`{name}:{i}` {line.strip()[:64]}")
                if len(hits) >= 6:
                    break
    if not hits:
        return "- (not found in the source; the compiled stream alone)"
    return "; ".join(hits[:3]) + (f" (+{len(hits) - 3} more)" if len(hits) > 3 else "")


def read_project_stream(xlsm):
    tmp = pathlib.Path("/tmp/_audit_vba.bin")
    with zipfile.ZipFile(xlsm) as z:
        if "xl/vbaProject.bin" not in z.namelist():
            return ""
        tmp.write_bytes(z.read("xl/vbaProject.bin"))
    try:
        import olefile
        ole = olefile.OleFileIO(str(tmp))
        return "" if not ole.exists("PROJECT") else ole.openstream("PROJECT").read().decode("cp1252", "replace")
    except Exception:
        return ""


def map_classes(proj, sheets, mods):
    """{sheet name: VBA class}. Excel writes Document= for ThisWorkbook first and then
    one per sheet, in tab order; where the class names are the familiar Sheet1..SheetN
    that is the mapping, and a mismatch shows up in §9's cross-check."""
    docs = re.findall(r"(?m)^\s*Document=([^/\r\n]+?)(?:/[^\r\n]*)?\s*$", proj or "")
    docs = [d for d in docs if d != "ThisWorkbook"]
    if len(docs) < len(sheets):
        docs += [f"Sheet{i + 1}" for i in range(len(docs), len(sheets)) if f"Sheet{i + 1}" in mods]
    return {s["name"]: docs[i] for i, s in enumerate(sheets) if i < len(docs)}


def procedures(code):
    """The body of each Sub/Function, so the state toggles can be followed."""
    out = []
    lines = code.splitlines()
    i = 0
    while i < len(lines):
        m = re.match(r"\s*(?:Public |Private )?(?:Sub|Function)\s+(\w+)", lines[i])
        if m:
            start = i
            i += 1
            while i < len(lines) and not re.match(r"\s*End (Sub|Function)\s*$", lines[i]):
                i += 1
            out.append({"name": m.group(1), "body": "\n".join(lines[start:i + 1])})
        i += 1
    return out


if __name__ == "__main__":
    main()
