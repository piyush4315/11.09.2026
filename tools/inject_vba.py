#!/usr/bin/env python3
"""
tools/inject_vba.py -- put the VBA sources in vba/ back into the workbook.

The macro text lives in the repo as plain .bas / .cls files so it can be read and
diffed like any other source; this writes those files into xl/vbaProject.bin with
pyOpenVBA (https://github.com/pyspread/py-openvba), leaving every other part of
the workbook untouched.  Without --apply it writes a copy next to the input so
the result can be inspected first.

    python3 tools/inject_vba.py 11.09.2026.xlsm            # dry run -> /tmp
    python3 tools/inject_vba.py 11.09.2026.xlsm --apply    # into the workbook

Requires py-openvba:  /home/user/.venv/bin/python -m pip install py-openvba
"""

import argparse
import pathlib
import sys

try:
    from pyopenvba import ExcelFile, VBAModuleKind
except ImportError:                                    # pragma: no cover
    sys.exit("py-openvba is missing:  python3 -m venv /home/user/.venv && "
             "/home/user/.venv/bin/python -m pip install py-openvba olevba")


def normalise(text):
    """VBA stores CRLF and needs a trailing newline."""
    return text.replace("\r\n", "\n").rstrip("\n") + "\r\n"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsm")
    ap.add_argument("--dir", default="vba", help="folder holding the .bas/.cls sources")
    ap.add_argument("--apply", action="store_true", help="write into the workbook itself")
    ap.add_argument("--out", help="where the dry run copies to (default /tmp/<name>.vba.xlsm)")
    args = ap.parse_args()

    src_dir = pathlib.Path(args.dir)
    files = sorted(p for p in src_dir.iterdir() if p.suffix in (".bas", ".cls"))
    if not files:
        sys.exit(f"no .bas/.cls sources in {src_dir}/")

    out = args.xlsm if args.apply else (args.out or f"/tmp/{pathlib.Path(args.xlsm).stem}.vba.xlsm")
    print(f"injecting {len(files)} module(s) from {src_dir}/ into {args.xlsm}")
    with ExcelFile(args.xlsm) as wb:
        existing = list(wb.module_names())
        for f in files:
            name, text = f.stem, normalise(f.read_text(encoding="utf-8"))
            if any(ord(ch) > 126 for ch in text):
                sys.exit(f"{f} holds non-ASCII characters - the VBA project's code page "
                         f"cannot store them; write them as ChrW(...) instead")
            if name in existing:
                wb.set_module(name, text)
                print(f"  updated  {name:<20} {len(text.splitlines()):>3} lines")
            elif f.suffix == ".bas":
                wb.vba_project().add_module(name, text, kind=VBAModuleKind.standard)
                print(f"  added    {name:<20} {len(text.splitlines()):>3} lines")
            else:
                sys.exit(f"{f.name} is a class/sheet module that does not exist in the project "
                         f"yet - only .bas modules can be added")
        for problem in wb.validate():
            print("  validate:", problem)
        wb.save(out)
    with ExcelFile(out) as check:
        print(f"  modules now: " + ", ".join(sorted(check.module_names())))
    print(f"\nwrote {out}" + ("" if args.apply else "  (dry run - pass --apply for the workbook)"))


if __name__ == "__main__":
    main()
