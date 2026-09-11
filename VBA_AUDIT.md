# VBA / macro audit - `11.09.2026.xlsm`

Produced by `tools/audit_macros.py` from the file itself; nothing in here is hand-written, so re-running the tool after any change keeps it true.

file sha256: `0dfc93fc9c5ea261246fa1e69aba24fe954bd5f44c296bbe715fc149d3135d60`

## 1. The package

* 14 parts; worksheet parts: xl/worksheets/sheet1.xml, xl/worksheets/sheet2.xml, xl/worksheets/sheet3.xml, xl/worksheets/sheet4.xml, xl/worksheets/sheet5.xml
* `xl/vbaProject.bin` - 43,520 bytes, sha256 `3e570d613a64af326375b5a3c37195f8…`. One OLE blob holds every line of macro code in the workbook: it is the entire attack surface, and also the only thing that a formula-only change (such as the FCS search row) must leave byte-identical.
* Excel 4.0 macro sheets (XLM): **none** - XLM is where formulas-only .xlsx files hide payloads; there is nothing of the kind here
* external links, customXml, ActiveX, data/queries: **none**
* there is no signature part in the package, so the file is **unsigned**: Excel will block the macros for anyone who downloads or receives it (Mark of the Web) until they enable the content or trust the folder. That is the expected state for a working file; sign it if it starts travelling.

## 2. Sheets and the code behind them

| sheet | part | visibility | code-behind | handlers in it | protection |
|---|---|---|---|---|---|
| Final Calculation Sheet | `xl/worksheets/sheet1.xml` | visible | `Sheet1` | *no code at all* | none |
| Live Search | `xl/worksheets/sheet2.xml` | visible | `Sheet2` | Worksheet_SelectionChange, Worksheet_BeforeDoubleClick | none |
| Buyer Groups | `xl/worksheets/sheet3.xml` | visible | `Sheet3` | Worksheet_Change, Worksheet_BeforeDoubleClick, Worksheet_Activate, Worksheet_SelectionChange | none |
| Lists | `xl/worksheets/sheet4.xml` | **hidden** | `Sheet4` | *no code at all* | none |
| LS_Backup | `xl/worksheets/sheet5.xml` | **hidden** | `Sheet5` | *no code at all* | none |

The class names are inferred from the order of the `Document=` lines in PROJECT (this is how Excel writes them: ThisWorkbook, then the sheets in tab order). §9 cross-checks that inference against what each class module talks about, so a wrong mapping cannot pass silently.

## 3. VBA project records

* `ID="{29937148-7B5A-4EA6-AEB0-0A4040C508AD}"`, `Name="VBAProject"`, `HelpContextID="0"`
* PROJECT declares 6 document module(s) (ThisWorkbook, Sheet1, Sheet2, Sheet3, Sheet4, Sheet5) and 2 standard module(s) (SortToggle, BuyerGroupStrip) - 8 modules were actually extracted from the storage: BuyerGroupStrip, Sheet1, Sheet2, Sheet3, Sheet4, Sheet5, SortToggle, ThisWorkbook
* every declared module extracted, and nothing extracted is undeclared - the storage is internally consistent, the sign of a project that has not been tampered with or half-repaired
* protection records present: CMG (26 chars), DPB (24 chars), GC (18 chars) - these hold the project-protection state; the values here are the short form Excel writes when *no* view password is set. The decisive evidence is behavioural: the project was rewritten by `tools/inject_vba.py` (pyOpenVBA refuses a locked project), so it is **open for reading and editing**. Anyone you send the file to can therefore see the code - fine inside the team, worth locking before it travels.
* the VBE workspace record shows which windows the last editing session left open: `Sheet2` open, 7 closed - a fingerprint of who was working on what, harmless but it also proves the code is normally edited in place here

## 4. Modules, routine by routine

### `BuyerGroupStrip` - 489 lines, 22 routine(s), 21 constant(s) · standard module
```vba
    Private Const BG_SHEET As String = "Buyer Groups"
    Private Const BG_BACKUP As String = "BG_Backup"
    Private Const BG_SEARCH_ROW As Long = 4
    Private Const BG_BAR_ROW As Long = 6
    Private Const BG_STRIP_FIRST As Long = 7
    Private Const BG_STRIP_LAST As Long = 43
    Private Const BG_TABLE_FIRST As Long = 45
    Private Const BG_TABLE_LAST As Long = 94              ' last subtotal row
    Private Const BG_COL_LAST As Long = 35                ' AI
    Private Const BG_NUM_COL As Long = 1                  ' A  lot number, or the sum sigil
    Private Const BG_LOT_COL As Long = 2                  ' B  Lot No.
    Private Const BG_BUYER_COL As Long = 3                ' C  Buyer
    Private Const BG_LABEL_COL As Long = 4                ' D  "SUBTOTAL - <buyer>" on a subtotal row
    Private Const BG_FLAG_COL As Long = 36                ' AJ match flag
    Private Const BG_FLAG_ROW0 As Long = 2                ' AJ row = source row + 2
    Private Const BG_SRC_FIRST As Long = 4                ' first data row on Final Calculation Sheet
    Private Const BG_SRC_MARK As String = "'Final Calculation Sheet'!$F$"
    Private Const SUM_SIGIL As Long = 8721                ' the subtotal marker in column A
    Private Const STATE_COL As Long = 146                 ' EP, far off column: 1 = blocks are lifted
    Private Const MAX_BLOCKS As Long = 40
    Private Const MAX_LOTS As Long = 40
    Public Sub OnSearchRowChange(ByVal Target As Range)
    Public Sub OnBarDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)
    Public Sub SyncStrip(ByVal ws As Worksheet)
    Public Sub UnfloatIfNeeded(ByVal ws As Worksheet)
    Public Sub ToggleFloat(ByVal ws As Worksheet)
    Private Function FloatMatches(ByVal ws As Worksheet) As Long
    Private Function RestoreOriginal(ByVal ws As Worksheet) As Long
    Private Sub Rescan(ByVal ws As Worksheet)
    Private Function Prepare(ByVal ws As Worksheet) As Boolean
    Private Function FindBlock(ByVal key As String) As Long
    Private Function Better(ByVal a As Long, ByVal b As Long) As Boolean
    Private Function SourceRow(ByVal ws As Worksheet, ByVal r As Long) As Long
    Private Function Matched(ByVal ws As Worksheet, ByVal srcRow As Long) As Boolean
    Private Function SearchCount(ByVal ws As Worksheet) As Long
    Private Sub PlaceBlocks(ByVal ws As Worksheet, wantKey() As String, ByVal n As Long)
    Private Sub PlaceLots(ByVal ws As Worksheet, wantKey() As String, wantN() As Long, _
    Private Sub Renumber(ByVal ws As Worksheet)
    Private Sub SaveOrder(ByVal ws As Worksheet)
    Private Function LoadOrder(wantKey() As String, wantN() As Long, _
    Private Function BackupSheet() As Worksheet
    Public Sub ClearSay()
    Private Sub Say(ByVal msg As String)
```

### `Sheet1` - 8 lines, 0 routine(s), 0 constant(s) · document module (code-behind) · bound to *Final Calculation Sheet*
* (no procedures at all - only the IDE boilerplate; nothing here can run code)

### `Sheet2` - 18 lines, 2 routine(s), 0 constant(s) · document module (code-behind) · bound to *Live Search*
* event handlers: `Private Sub Worksheet_SelectionChange(ByVal Target As Range)`; `Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)`

### `Sheet3` - 28 lines, 4 routine(s), 0 constant(s) · document module (code-behind) · bound to *Buyer Groups*
* event handlers: `Private Sub Worksheet_Change(ByVal Target As Range)`; `Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)`; `Private Sub Worksheet_Activate()`; `Private Sub Worksheet_SelectionChange(ByVal Target As Range)`

### `Sheet4` - 8 lines, 0 routine(s), 0 constant(s) · document module (code-behind) · bound to *Lists*
* (no procedures at all - only the IDE boilerplate; nothing here can run code)

### `Sheet5` - 8 lines, 0 routine(s), 0 constant(s) · document module (code-behind) · bound to *LS_Backup*
* (no procedures at all - only the IDE boilerplate; nothing here can run code)

### `SortToggle` - 190 lines, 4 routine(s), 8 constant(s) · standard module
```vba
    Private Const LS_SHEET As String = "Live Search"
    Private Const LS_BACKUP_SHEET As String = "LS_Backup"
    Private Const LS_HEADER_ROW As Long = 5
    Private Const LS_FIRST_DATA_ROW As Long = 6
    Private Const LS_LAST_DATA_ROW As Long = 42
    Private Const LS_FIRST_COL As Long = 1    ' A
    Private Const LS_LAST_COL As Long = 35    ' AI
    Private Const LS_STATE_COL As Long = 100  ' CV - unused far column, holds sort state
    Public Sub ToggleLiveSearchSort(ByVal Target As Range)
    Private Function ShouldSwap(arr As Variant, r1 As Long, r2 As Long, colIdx As Long, dir As Long) As Boolean
    Private Sub SwapRows(arr As Variant, r1 As Long, r2 As Long, nCols As Long)
    Public Sub ResetLiveSearchSort()
```

### `ThisWorkbook` - 8 lines, 0 routine(s), 0 constant(s) · document module (code-behind)
* (no procedures at all - only the IDE boilerplate; nothing here can run code)

## 5. What the code is able to do

| capability | severity | hits | where |
|---|---|---|---|
| runs other code / spawns processes | FINDING | 0 |  |
| reads or writes files | FINDING | 0 |  |
| talks to other applications or the network | FINDING | 0 |  |
| auto-runs when the file is opened | FINDING | 0 |  |
| modifies the VBA project itself (code that can rewrite code) | FINDING | 0 |  |
| changes Excel's application state | NOTE | 22 | `BuyerGroupStrip:100`, `BuyerGroupStrip:101`, `BuyerGroupStrip:109`, `BuyerGroupStrip:110`, `BuyerGroupStrip:130`, `BuyerGroupStrip:131`, `BuyerGroupStrip:145`, `BuyerGroupStrip:146`, `BuyerGroupStrip:150`, `BuyerGroupStrip:151`, `SortToggle:42`, `SortToggle:43`, `SortToggle:44`, `SortToggle:89` … |
| can destroy rows or cells | NOTE | 9 | `BuyerGroupStrip:105`, `BuyerGroupStrip:135`, `BuyerGroupStrip:376`, `BuyerGroupStrip:399`, `BuyerGroupStrip:400`, `BuyerGroupStrip:423`, `SortToggle:178`, `SortToggle:179`, `SortToggle:180` |
| sorts or filters a range | NOTE | 0 |  |
| writes into cells | OK | 16 | `BuyerGroupStrip:140`, `BuyerGroupStrip:412`, `BuyerGroupStrip:424`, `BuyerGroupStrip:426`, `BuyerGroupStrip:427`, `BuyerGroupStrip:430`, `BuyerGroupStrip:431`, `BuyerGroupStrip:433`, `SortToggle:71`, `SortToggle:73`, `SortToggle:74`, `SortToggle:75`, `SortToggle:85`, `SortToggle:161` … |
| swallows errors (On Error Resume Next) | NOTE | 4 | `BuyerGroupStrip:441`, `BuyerGroupStrip:470`, `BuyerGroupStrip:482`, `BuyerGroupStrip:487` |
| has a real error handler | OK | 6 | `BuyerGroupStrip:102`, `BuyerGroupStrip:132`, `BuyerGroupStrip:443`, `BuyerGroupStrip:472`, `SortToggle:46`, `SortToggle:156` |
| handles a sheet event | OK | 6 | `Sheet2:9`, `Sheet2:13`, `Sheet3:9`, `Sheet3:15`, `Sheet3:21`, `Sheet3:26` |
| handles a workbook event | OK | 0 |  |

The absence of hits is the finding worth reading: **no** `Shell`, no `CreateObject`, no `Environ`, no `Declare`, no `MSXML`/`WScript`/`ADODB`, no `Kill`/`Open`/`Print #`, no `SaveAs`, no `VBComponents`/`CodeModule` (code that rewrites code), and no `Auto_Open` / `Workbook_Open`. Nothing here reaches outside the workbook and nothing runs on its own.

## 6. What oletools' own scanner says, and what each hit really is

| oletools verdict | keyword | what it caught | the line that triggered it |
|---|---|---|---|
| AutoExec | `Worksheet_Change` | Runs when the file is opened and ActiveX objects trigger eve… | `Sheet3:9` Private Sub Worksheet_Change(ByVal Target As Range) |
| Suspicious | `open` | May open a file… | `BuyerGroupStrip:21` '    1. OnSearchRowChange - opens rows 7:43 as soon as a SEARCH ; `BuyerGroupStrip:87` ' open the strip while any SEARCH box is filled, close it when t; `BuyerGroupStrip:236` Dim r As Long, b As Long, i As Long, t As String, v As Variant,  (+4 more) |
| Suspicious | `put` | May write to a file (if combined with Open)… | `BuyerGroupStrip:113` Say "Buyer Groups: the blocks could not be put back - " & Err.De; `BuyerGroupStrip:136` Say "Buyer Groups: " & n & " buyer block(s) put back into the pl; `BuyerGroupStrip:141` Say "Buyer Groups: " & n & " matching buyer block(s) lifted to t (+3 more) |
| Suspicious | `ChrW` | May attempt to obfuscate specific strings (use option --deob… | `SortToggle:81` h = Replace(Replace(h, " " & ChrW(9650), ""), " " & ChrW(9660), ; `SortToggle:83` h = h & " " & IIf(newDir = 1, ChrW(9650), ChrW(9660)) |
| Suspicious | `Hex Strings` | Hex-encoded strings were detected, may be used to obfuscate… | `Sheet1:2` Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"; `Sheet2:2` Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"; `Sheet3:2` Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}" (+3 more) |

* `detect_autoexec()` hits: 1 · VBA stomping (source stripped so the compiled code runs unseen - a classic malware trick): **not detected**
* `detect_vba_macros()` = True, `detect_xlm_macros()` = False, encrypted/signed for external viewing: False

Every keyword above was traced back to the line that caused it: ordinary words inside comments, status messages and identifier names (`open it while…`, `put them back`, `openBlock`), two legitimate uses of `ChrW` that draw the ▲/▼ arrow on the sort header, and the Office type-library GUIDs in the `VB_Base` attributes the hex-string heuristic trips over. No encoded payload, no obfuscated string, no auto-run.

## 7. State switches: does every one get flipped back?

* `BuyerGroupStrip.UnfloatIfNeeded`: 4 switch(es) - EnableEvents, ScreenUpdating - every switch-off is followed by a switch-on, and the error path is routed through that same cleanup.
* `BuyerGroupStrip.ToggleFloat`: 6 switch(es) - EnableEvents, ScreenUpdating - every switch-off is followed by a switch-on, and the error path is routed through that same cleanup; EnableEvents restored on more than one path (idempotent, harmless), ScreenUpdating restored on more than one path (idempotent, harmless).
* `SortToggle.ToggleLiveSearchSort`: 4 switch(es) - EnableEvents, ScreenUpdating - every switch-off is followed by a switch-on, and the error path is routed through that same cleanup.
* `SortToggle.ResetLiveSearchSort`: 4 switch(es) - EnableEvents, ScreenUpdating - every switch-off is followed by a switch-on, and the error path is routed through that same cleanup.

A balanced toggle pair is only meaningful if the exit paths converge on it - the check above is on the whole routine body, which is exactly how `Resume Done` style cleanup in these modules is written. If Excel ever stops reacting while you type in a SEARCH row, `Application.EnableEvents = True` in the Immediate window restores it without a restart.

## 8. Macro-facing details inside the sheets

| sheet | OLE objects | form controls | hyperlinks | data validations | conditional formats | defined names on this sheet |
|---|---|---|---|---|---|---|
| Final Calculation Sheet | 0 | 0 | 0 | 34 | 37 | 1 |
| Live Search | 0 | 0 | 0 | 37 | 3 | 1 |
| Buyer Groups | 0 | 0 | 0 | 34 | 73 | 0 |
| Lists | 0 | 0 | 0 | 0 | 0 | 0 |
| LS_Backup | 0 | 0 | 0 | 0 | 0 | 0 |

* **no buttons, shapes or controls are wired to macros anywhere.** Each macro is entered from a *cell* - a `Worksheet_Change` on the search row or a `Worksheet_BeforeDoubleClick` on the bar - so there is no caption hiding a `Sub`, nothing to click by accident, and deleting a shape cannot orphan code.

## 9. Cross-checks

* `vba/BuyerGroupStrip.bas` vs module `BuyerGroupStrip`: identical - the copy in the repo is what Excel runs
* `vba/Sheet3.cls` vs module `Sheet3`: identical - the copy in the repo is what Excel runs
* **Final Calculation Sheet** (`Sheet1`) - no procedures at all, only the IDE boilerplate - which is why the search row added there needs no macro and cannot disturb one
* **Lists** (`Sheet4`) - no procedures at all, only the IDE boilerplate
* `Sheet1` talks about no sheet by name - consistent with the mapping in §2
* behaviour check `Sheet2`: it calls `SortToggle`, whose *_SHEET constant names *LS_Backup, Live Search* - **matches** the §2 mapping
* `Sheet2` talks about no sheet by name - consistent with the mapping in §2
* behaviour check `Sheet3`: it calls `BuyerGroupStrip`, whose *_SHEET constant names *Buyer Groups* - **matches** the §2 mapping
* `Sheet3` talks about no sheet by name - consistent with the mapping in §2
* `Sheet4` talks about no sheet by name - consistent with the mapping in §2
* `Sheet5` talks about no sheet by name - consistent with the mapping in §2
* 106 defined names. The three dropdown families - SUG_* (36, Live Search), SUGN_* (34, Buyer Groups), SUGF_* (34, Final Calculation Sheet) - all resolve to dynamic `INDEX()` ranges on the hidden `Lists` sheet; no name points at a macro sheet or an external file
* the FCS search feature is complete without a single line of VBA: **`xl/vbaProject.bin` was verified byte-identical before and after it was added** (`tools/fcs_search_row.py` asserts that on every run), and `SortToggle`/`BuyerGroupStrip` never look at *Final Calculation Sheet*'s rows 1-2, so nothing that already runs can be disturbed

## 10. Findings

* **NOTE** - Live Search runs code on Worksheet_SelectionChange, Worksheet_BeforeDoubleClick: that is by design (the search row and the double-click bar); see §4-§6
* **NOTE** - Buyer Groups runs code on Worksheet_Change, Worksheet_BeforeDoubleClick, Worksheet_Activate, Worksheet_SelectionChange: that is by design (the search row and the double-click bar); see §4-§6
* **NOTE** - the VBA project is not locked for viewing: the full source is readable by every recipient; Tools ▸ VBAProject Properties ▸ Protection if that matters, but remember a locked project also cannot be audited

Count: 0 finding(s), 3 note(s).

## 11. How to read this

* **Nothing runs when the file is opened.** There is no auto-run entry point anywhere; the first line of code executes when a cell changes on *Live Search* or *Buyer Groups*.
* **Nothing leaves the file.** No file, registry, network or COM call exists in the project; the only writes are cell values, row moves and a hidden backup sheet.
* **Destructive work is bracketed by a backup.** Both reorder macros cut whole rows and insert them back, and each one first stashes the untouched order (`LS_Backup`, and `BG_Backup` which is created on first use), so a restore is always available - and a failed restore says so in a message box instead of leaving the sheet half-moved.
* **The search machinery is code-free.** On all three sheets the per-column search, the cascading dropdowns, the yellow hits and the grey misses are formulas, data validation and conditional formatting. They keep working with macros disabled, which is also the fastest way to tell a macro problem from a formula problem: disable macros, and if the sheet behaves, the code is not the culprit.
* **Two things to decide before the workbook travels**: it is unsigned (recipients must enable content), and the project is not locked for viewing (they can also read and edit every line). Both are deliberate for a working file inside the team.
