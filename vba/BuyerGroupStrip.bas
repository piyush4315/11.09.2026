Attribute VB_Name = "BuyerGroupStrip"
Option Explicit

' =====================================================================
'  Buyer Groups - matching lots pinned to the top
'  ---------------------------------------------------------------------
'  Layout of the sheet (row 2 holds the full how-to):
'      row 4          the yellow SEARCH boxes, B4:AI4 (with dropdowns)
'      row 5          column headers
'      row 6          the MATCHED LOTS bar            <- double-click this row
'      rows 7:43      the strip: every lot matching the SEARCH row, hidden
'                     while the SEARCH row is empty
'      row 44         spacer, and the strip's +/- outline control
'      rows 45:94     the buyer blocks (lot rows, then the block's subtotal)
'      rows 95:96     TOTAL and GRAND TOTAL
'      AJ6:AJ42       per-lot match flag (row = Final Calculation Sheet row + 2)
'      AK6:AK42       running rank of the matching lots
'      AL7:AL43       strip row -> source row pointer
'
'  What this module adds
'    1. OnSearchRowChange - opens rows 7:43 as soon as a SEARCH box holds a
'       value and closes them again when the row is cleared.  The strip's
'       contents and the grey-out of the non-matching lots are worksheet
'       formulas, so they need no code at all.
'    2. OnBarDoubleClick  - double-clicking the bar (row 6) lifts every buyer
'       block that holds a matching lot to the top of the grouped table,
'       matched lots first inside each block.  Double-click the bar again and
'       the plain alphabetical order comes back.
'
'  The float only ever moves whole rows (Cut + Insert), so number formats,
'  shading, row groups, row heights and the live links travel with the row and
'  Excel re-points every block sum, lot counter and the TOTAL row's links by
'  itself.  Nothing is retyped here, so no figure can drift.  The untouched
'  order is recorded in the hidden BG_Backup sheet just before the first float.
' =====================================================================

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

' the grouped table as the macro sees it, refreshed by Rescan
Private mBlocks As Long
Private mKey(1 To MAX_BLOCKS) As String               ' the subtotal row's label
Private mFirst(1 To MAX_BLOCKS) As Long               ' first lot row
Private mLast(1 To MAX_BLOCKS) As Long                ' the block's subtotal row
Private mLotCount(1 To MAX_BLOCKS) As Long
Private mLots(1 To MAX_BLOCKS, 1 To MAX_LOTS) As Long ' source rows, top to bottom
Private mMatch(1 To MAX_BLOCKS) As Long               ' how many of them match
Private mWhy As String

' ---------------------------------------------------------------- entry points

Public Sub OnSearchRowChange(ByVal Target As Range)
    If Target Is Nothing Then Exit Sub
    If Target.Worksheet.Name <> BG_SHEET Then Exit Sub
    If Intersect(Target, Target.Worksheet.Rows(BG_SEARCH_ROW)) Is Nothing Then Exit Sub
    SyncStrip Target.Worksheet
    UnfloatIfNeeded Target.Worksheet             ' a new search starts from the plain order again
End Sub

Public Sub OnBarDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)
    If Target.Worksheet.Name <> BG_SHEET Then Exit Sub
    If Target.Row <> BG_BAR_ROW Then Exit Sub
    Cancel = True                                      ' never enter the bar while toggling
    ToggleFloat Target.Worksheet
End Sub

Public Sub SyncStrip(ByVal ws As Worksheet)
    ' open the strip while any SEARCH box is filled, close it when they are clear
    If ws.Name <> BG_SHEET Then Exit Sub
    ws.Rows(BG_STRIP_FIRST & ":" & BG_STRIP_LAST).Hidden = (SearchCount(ws) = 0)
End Sub

Public Sub UnfloatIfNeeded(ByVal ws As Worksheet)
    ' A new search means a new set of matches, so the blocks first go back to the
    ' plain alphabetical order; double-clicking the bar then lifts the new ones.
    ' The state flag is cleared in the same pass, so this runs once per search edit
    ' and typing another character into the SEARCH row costs nothing.
    If ws.Name <> BG_SHEET Then Exit Sub
    If CLng(Val(CStr(ws.Cells(1, STATE_COL).Value))) <> 1 Then Exit Sub
    If Not Prepare(ws) Then Exit Sub
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Failed
    Dim n As Long
    n = RestoreOriginal(ws)
    ws.Cells(1, STATE_COL).ClearContents
    Say "Buyer Groups: new search - the " & n & " buyer block(s) are back in the plain order."
Done:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
Failed:
    Say "Buyer Groups: the blocks could not be put back - " & Err.Description
    Resume Done
End Sub

Public Sub ToggleFloat(ByVal ws As Worksheet)
    If Not Prepare(ws) Then
        Say "Buyer Groups: " & mWhy
        Exit Sub
    End If
    If SearchCount(ws) = 0 Then
        Say "Buyer Groups: type something in the yellow SEARCH row (row 4) first - it decides which lots matter."
        Exit Sub
    End If

    Dim was As Long, n As Long
    was = CLng(Val(CStr(ws.Cells(1, STATE_COL).Value)))

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Failed
    If was = 1 Then
        n = RestoreOriginal(ws)
        ws.Cells(1, STATE_COL).ClearContents
        Say "Buyer Groups: " & n & " buyer block(s) put back into the plain alphabetical order."
    Else
        SaveOrder ws
        n = FloatMatches(ws)
        ws.Cells(1, STATE_COL).Value = 1
        Say "Buyer Groups: " & n & " matching buyer block(s) lifted to the top - double-click row 6 again to put them back."
    End If
Done:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
Failed:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Buyer Groups: the blocks could not be reordered - " & Err.Description & vbCrLf & _
           "Double-click row 6 again to put the order back.", vbExclamation, "Buyer Groups"
    Resume Done
End Sub

' ---------------------------------------------------------------- the float

Private Function FloatMatches(ByVal ws As Worksheet) As Long
    Dim wantKey(1 To MAX_BLOCKS) As String
    Dim wantN(1 To MAX_BLOCKS) As Long
    Dim seq(1 To MAX_BLOCKS, 1 To MAX_LOTS) As Long
    Dim b As Long, i As Long, j As Long, lifted As Long

    ' blocks holding a match first (more matches first, then A-Z), then the rest
    Dim idx(1 To MAX_BLOCKS) As Long
    For b = 1 To mBlocks
        idx(b) = b
    Next b
    Dim x As Long, y As Long, t As Long
    For x = 1 To mBlocks - 1
        For y = 1 To mBlocks - x
            If Better(idx(y + 1), idx(y)) Then
                t = idx(y): idx(y) = idx(y + 1): idx(y + 1) = t
            End If
        Next y
    Next x

    Dim n As Long
    n = mBlocks
    For x = 1 To n
        wantKey(x) = mKey(idx(x))
        If mMatch(idx(x)) > 0 Then lifted = lifted + 1
    Next x
    PlaceBlocks ws, wantKey, n

    ' inside each block, matched lots first; re-read the layout because rows moved
    Prepare ws
    For b = 1 To n
        Prepare ws
        j = FindBlock(wantKey(b))
        If j = 0 Then Err.Raise vbObjectError + 513, "BuyerGroupStrip", "block is missing: " & wantKey(b)
        i = 0
        Dim k As Long
        For k = 1 To mLotCount(j)                          ' the matches, in their own order
            If Matched(ws, mLots(j, k)) Then
                i = i + 1
                seq(b, i) = mLots(j, k)
            End If
        Next k
        For k = 1 To mLotCount(j)                          ' then the lots that do not match
            If Not Matched(ws, mLots(j, k)) Then
                i = i + 1
                seq(b, i) = mLots(j, k)
            End If
        Next k
        wantN(b) = i
    Next b
    PlaceLots ws, wantKey, wantN, seq, n
    Renumber ws
    FloatMatches = lifted
End Function

Private Function RestoreOriginal(ByVal ws As Worksheet) As Long
    Dim wantKey(1 To MAX_BLOCKS) As String
    Dim wantN(1 To MAX_BLOCKS) As Long
    Dim seq(1 To MAX_BLOCKS, 1 To MAX_LOTS) As Long
    Dim n As Long

    n = LoadOrder(wantKey, wantN, seq)
    If n = 0 Then
        Err.Raise vbObjectError + 514, "BuyerGroupStrip", _
                  "no saved order to go back to (the hidden " & BG_BACKUP & " sheet is empty)"
    End If
    PlaceBlocks ws, wantKey, n
    PlaceLots ws, wantKey, wantN, seq, n
    Renumber ws
    RestoreOriginal = n
End Function

' ---------------------------------------------------------------- layout model

Private Sub Rescan(ByVal ws As Worksheet)
    ' Walk the grouped table once: a lot row is any row whose # column holds a
    ' number, the block's subtotal row is the one whose # column holds the sigil.
    Dim r As Long, b As Long, i As Long, t As String, v As Variant, openBlock As Boolean
    mBlocks = 0
    b = 0
    openBlock = False
    For r = BG_TABLE_FIRST To BG_TABLE_LAST
        v = ws.Cells(r, BG_NUM_COL).Value
        t = CStr(v)
        If Len(t) = 0 Then
            ' nothing in the # column: leave the row out of the model
        ElseIf IsNumeric(v) Then
            If Not openBlock Then
                b = b + 1
                If b > MAX_BLOCKS Then Exit Sub
                mFirst(b) = r
                mLast(b) = 0
                mKey(b) = vbNullString
                mLotCount(b) = 0
                mMatch(b) = 0
                openBlock = True
            End If
            i = mLotCount(b) + 1
            If i > MAX_LOTS Then Exit Sub
            mLotCount(b) = i
            mLots(b, i) = SourceRow(ws, r)
            If Matched(ws, mLots(b, i)) Then mMatch(b) = mMatch(b) + 1
        ElseIf AscW(Left$(t, 1)) = SUM_SIGIL Then
            If openBlock Then
                mLast(b) = r
                mKey(b) = Trim$(CStr(ws.Cells(r, BG_LABEL_COL).Value))
                If Len(mKey(b)) = 0 Then mKey(b) = "BLOCK " & b
                openBlock = False
            End If
        End If
    Next r
    mBlocks = b
End Sub

Private Function Prepare(ByVal ws As Worksheet) As Boolean
    Prepare = False
    mWhy = "this is not the Buyer Groups sheet"
    If ws Is Nothing Then Exit Function
    If ws.Name <> BG_SHEET Then Exit Function
    mWhy = "the MATCHED LOTS bar in row " & BG_BAR_ROW & " is missing, so this macro will not touch the sheet"
    If InStr(1, CStr(ws.Cells(BG_BAR_ROW, BG_BUYER_COL).Formula), "$AJ$6:$AJ$42", vbTextCompare) = 0 Then Exit Function
    mWhy = "the strip is not in rows " & BG_STRIP_FIRST & ":" & BG_STRIP_LAST
    If Not ws.Cells(BG_STRIP_FIRST, BG_LOT_COL).HasFormula Then Exit Function
    mWhy = "TOTAL / GRAND TOTAL are not in rows " & (BG_TABLE_LAST + 1) & ":" & (BG_TABLE_LAST + 2)
    If InStr(1, CStr(ws.Cells(BG_TABLE_LAST + 1, BG_LABEL_COL).Value), "TOTAL", vbTextCompare) = 0 Then Exit Function

    Rescan ws
    mWhy = "the buyer blocks could not be read - is a filter or a hidden row in the way?"
    If mBlocks < 1 Then Exit Function
    Dim b As Long, i As Long, key As String
    For b = 1 To mBlocks
        If mLast(b) <= mFirst(b) Then Exit Function              ' a block without its subtotal row
        For i = 1 To mLotCount(b)
            If mLots(b, i) < BG_SRC_FIRST Then Exit Function     ' a row not linked to a source lot
        Next i
        For i = b + 1 To mBlocks
            If mKey(i) = mKey(b) Then
                mWhy = "two blocks share the label '" & mKey(b) & "', so they cannot be told apart"
                Exit Function
            End If
        Next i
    Next b
    Prepare = True
End Function

Private Function FindBlock(ByVal key As String) As Long
    Dim b As Long
    For b = 1 To mBlocks
        If mKey(b) = key Then
            FindBlock = b
            Exit Function
        End If
    Next b
    FindBlock = 0
End Function

Private Function Better(ByVal a As Long, ByVal b As Long) As Boolean
    ' blocks with matches go first, then more matches, then alphabetical
    If (mMatch(a) > 0) <> (mMatch(b) > 0) Then
        Better = (mMatch(a) > 0)
    ElseIf mMatch(a) <> mMatch(b) Then
        Better = (mMatch(a) > mMatch(b))
    Else
        Better = (StrComp(mKey(a), mKey(b), vbTextCompare) < 0)
    End If
End Function

Private Function SourceRow(ByVal ws As Worksheet, ByVal r As Long) As Long
    Dim f As String, p As Long, s As String, i As Long
    f = CStr(ws.Cells(r, BG_LOT_COL).Formula)
    p = InStr(1, f, BG_SRC_MARK, vbTextCompare)
    If p = 0 Then
        SourceRow = -1
        Exit Function
    End If
    s = Mid$(f, p + Len(BG_SRC_MARK))
    i = 1
    Do While Mid$(s, i, 1) Like "#"
        i = i + 1
    Loop
    If i = 1 Then
        SourceRow = -1
    Else
        SourceRow = CLng(Left$(s, i - 1))
    End If
End Function

Private Function Matched(ByVal ws As Worksheet, ByVal srcRow As Long) As Boolean
    Dim v As Variant
    If srcRow < BG_SRC_FIRST Then Exit Function
    v = ws.Cells(srcRow + BG_FLAG_ROW0, BG_FLAG_COL).Value
    Select Case VarType(v)
        Case vbBoolean
            Matched = CBool(v)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency
            Matched = (CDbl(v) <> 0)
        Case Else
            Matched = False
    End Select
End Function

Private Function SearchCount(ByVal ws As Worksheet) As Long
    SearchCount = Application.CountA(ws.Range(ws.Cells(BG_SEARCH_ROW, BG_LOT_COL), _
                                              ws.Cells(BG_SEARCH_ROW, BG_COL_LAST)))
End Function

' ---------------------------------------------------------------- the moves

Private Sub PlaceBlocks(ByVal ws As Worksheet, wantKey() As String, ByVal n As Long)
    Dim slot As Long, t As Long, b As Long
    slot = BG_TABLE_FIRST
    For t = 1 To n
        Rescan ws
        b = FindBlock(wantKey(t))
        If b = 0 Then Err.Raise vbObjectError + 515, "BuyerGroupStrip", "block is missing: " & wantKey(t)
        If mFirst(b) <> slot Then
            ws.Rows(mFirst(b) & ":" & mLast(b)).Cut
            ws.Rows(slot).Insert Shift:=xlShiftDown
            Application.CutCopyMode = False
        End If
        slot = slot + (mLast(b) - mFirst(b) + 1)
    Next t
End Sub

Private Sub PlaceLots(ByVal ws As Worksheet, wantKey() As String, wantN() As Long, _
                      seq() As Long, ByVal n As Long)
    Dim t As Long, i As Long, b As Long, target As Long, cur As Long, j As Long
    For t = 1 To n
        For i = 1 To wantN(t)
            Rescan ws
            b = FindBlock(wantKey(t))
            If b = 0 Then Err.Raise vbObjectError + 516, "BuyerGroupStrip", "block is missing: " & wantKey(t)
            target = mFirst(b) + i - 1
            cur = 0
            For j = 1 To mLotCount(b)
                If mLots(b, j) = seq(t, i) Then cur = mFirst(b) + j - 1
            Next j
            If cur = 0 Then Err.Raise vbObjectError + 517, "BuyerGroupStrip", _
                                  "lot row " & seq(t, i) & " is no longer inside " & wantKey(t)
            If cur > target Then
                ws.Rows(cur).Cut
                ws.Rows(target).Insert Shift:=xlShiftDown
                Application.CutCopyMode = False
            End If
        Next i
    Next t
End Sub

Private Sub Renumber(ByVal ws As Worksheet)
    Dim b As Long, i As Long
    Rescan ws
    For b = 1 To mBlocks
        For i = 1 To mLotCount(b)
            ws.Cells(mFirst(b) + i - 1, BG_NUM_COL).Value = i
        Next i
    Next b
End Sub

' ---------------------------------------------------------------- the backup

Private Sub SaveOrder(ByVal ws As Worksheet)
    Dim bk As Worksheet, b As Long, i As Long, r As Long
    Set bk = BackupSheet()
    Rescan ws
    bk.Cells.Clear
    bk.Cells(1, 1).Value = "Buyer Groups - the untouched block order, written by BuyerGroupStrip " & _
                           "before the first float. Do not edit; delete this sheet to reset the macro."
    bk.Cells(1, 2).Value = Now
    bk.Cells(2, 1).Value = mBlocks
    For b = 1 To mBlocks
        r = 2 + b
        bk.Cells(r, 1).Value = mKey(b)
        bk.Cells(r, 2).Value = mLotCount(b)
        For i = 1 To mLotCount(b)
            bk.Cells(r, 2 + i).Value = mLots(b, i)
        Next i
    Next b
End Sub

Private Function LoadOrder(wantKey() As String, wantN() As Long, _
                           seq() As Long) As Long
    Dim bk As Worksheet, b As Long, i As Long, r As Long, n As Long
    On Error Resume Next
    Set bk = ThisWorkbook.Worksheets(BG_BACKUP)
    On Error GoTo 0
    If bk Is Nothing Then
        LoadOrder = 0
        Exit Function
    End If
    n = Val(CStr(bk.Cells(2, 1).Value))
    If n < 1 Or n > MAX_BLOCKS Then
        LoadOrder = 0
        Exit Function
    End If
    For b = 1 To n
        r = 2 + b
        wantKey(b) = CStr(bk.Cells(r, 1).Value)
        wantN(b) = Val(CStr(bk.Cells(r, 2).Value))
        If Len(wantKey(b)) = 0 Or wantN(b) < 1 Then
            LoadOrder = 0
            Exit Function
        End If
        For i = 1 To wantN(b)
            seq(b, i) = Val(CStr(bk.Cells(r, 2 + i).Value))
        Next i
    Next b
    LoadOrder = n
End Function

Private Function BackupSheet() As Worksheet
    Dim bk As Worksheet
    On Error Resume Next
    Set bk = ThisWorkbook.Worksheets(BG_BACKUP)
    On Error GoTo 0
    If bk Is Nothing Then
        Set bk = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        bk.Name = BG_BACKUP
        bk.Visible = xlSheetHidden
    End If
    Set BackupSheet = bk
End Function

Public Sub ClearSay()
    On Error Resume Next
    If Left$(CStr(Application.StatusBar), 13) = "Buyer Groups:" Then Application.StatusBar = False
End Sub

Private Sub Say(ByVal msg As String)
    On Error Resume Next
    Application.StatusBar = msg
End Sub
