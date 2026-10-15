Attribute VB_Name = "AddInStorage"
Option Explicit

' ============================================================================
'  AddInStorage
'  The ONE module that knows where the add-in keeps its OWN data.
'  Drop this standard module into the add-in once.
'
'  Why this module exists:
'   * The add-in stores its own settings and lists on a worksheet inside
'     THISWORKBOOK (the add-in) -- NOT in ActiveWorkbook (the user's file the
'     macros operate on). In an .xlam that sheet is invisible to the user.
'   * Qualifying add-in storage to ActiveWorkbook is a top-tier add-in bug: it
'     errors (the user's file has no such sheet) or silently hits a same-named
'     sheet the user happens to have open. Everything here uses ThisWorkbook.
'   * Every storage reference is funneled through the accessors below, so a
'     sheet or table rename is a one-line change here, not a hunt across every
'     macro in the project.
'
'  First-time setup (once, in the .xlsm you develop in):
'     Immediate window:  AddInStorage.SetupStorage
'  It creates the storage sheet and both tables if they are missing, and
'  leaves existing ones alone, so running it twice is harmless.
'
'  Usage from any macro:
'     Dim items As Variant, i As Long
'     items = AddInStorage.ItemList()          ' user-editable list
'     If AddInStorage.HasItems(items) Then
'         For i = LBound(items) To UBound(items)
'             Debug.Print items(i)
'         Next i
'     End If
'
'     Debug.Print REPORT_TITLE                 ' a setting, read like a Const
'     AddInStorage.SetSetting "REPORT_TITLE", "Q3 Pack"
'     AddInStorage.SaveStorage                 ' writers do NOT save for you
'
'  SETTINGS (tblSettings: Setting | Value)
'   * Values that would otherwise be `Const` declarations live on the sheet so
'     an end user can change them from the settings form without opening the
'     VBA editor.
'   * They are exposed below as Public Property Get procedures NAMED LIKE THE
'     CONSTANTS THEY REPLACE. A Public Property Get in a standard module has
'     the same global visibility a Public Const had, so existing call sites
'     compile unchanged. Migrating means deleting the old Const lines, nothing
'     else.
'   * Every property falls back to a private DEF_* default. A missing sheet,
'     missing key, blank cell or wrong-typed cell degrades to the built-in
'     value instead of breaking the add-in.
'
'  LIST (tblList: Item)
'   * A single-column, user-editable list (names, codes, departments...).
'     Written with ListObject.Resize, never ListRows.Add -- see WriteTableBlock.
' ============================================================================

' ----------------------------------------------------------------------------
'  EDIT ONLY THIS BLOCK to match the workbook.
' ----------------------------------------------------------------------------
Private Const STORAGE_SHEET       As String = "AddInData"    ' sheet inside ThisWorkbook
Private Const SETTINGS_TABLE      As String = "tblSettings"  ' key | value settings
Private Const SETTINGS_KEY_COLUMN As String = "Setting"      ' column holding the setting NAME
Private Const SETTINGS_VAL_COLUMN As String = "Value"        ' column holding the setting VALUE
Private Const LIST_TABLE          As String = "tblList"      ' user-editable list
Private Const LIST_COLUMN         As String = "Item"         ' its only column

' Excel's hard limit on the contents of a single worksheet cell.
Private Const MAX_CELL_CHARS      As Long = 32767

' ----------------------------------------------------------------------------
'  FALLBACK DEFAULTS -- the safety net, NOT the source of truth.
'  Change a value for real in the settings form; change it here only to move
'  the "if all else fails" floor. Named DEF_* so they do not collide with the
'  Property Get names below. SetupStorage seeds the table with these too.
' ----------------------------------------------------------------------------
Private Const DEF_REPORT_TITLE        As String = "Monthly Report"
Private Const DEF_PERIOD_END          As Date = #12/31/2026#
Private Const DEF_DEFAULT_ROW_HEIGHT  As Double = 15
Private Const DEF_SHOW_TIPS           As Boolean = True

' ----------------------------------------------------------------------------
'  Module-level cache. Reading a worksheet cell is a COM call; settings are
'  read many times per run and change rarely, so they are loaded once into a
'  Dictionary and reused.
'
'  Late-bound (CreateObject) on purpose -- early binding would require every
'  machine running this add-in to carry a reference to Microsoft Scripting
'  Runtime, and a missing reference breaks the whole VBA project, not just
'  this module.
'
'  Anything that writes to the storage sheet MUST call InvalidateCache.
' ----------------------------------------------------------------------------
Private mCache As Object            ' Scripting.Dictionary; Nothing = not loaded

' Last failure from a writer, for the settings form to display. The writers
' return True/False rather than raising, so the caller decides how to report;
' without this text a "could not be written" message has no detail at all.
Private mLastError As String

' True while WriteTableBlock is putting the previous contents back after a
' failed write, so a second failure cannot recurse.
Private mRestoring As Boolean

' Human-readable reason the last writer returned False.
Public Property Get LastError() As String
    LastError = mLastError
End Property

' ============================================================================
'  Sheet / table accessors -- ALWAYS from ThisWorkbook, never ActiveWorkbook.
' ============================================================================

Public Function StorageSheet() As Worksheet
    Set StorageSheet = ThisWorkbook.Worksheets(STORAGE_SHEET)
End Function

Public Function SettingsTable() As ListObject
    Set SettingsTable = StorageSheet().ListObjects(SETTINGS_TABLE)
End Function

Public Function ListTable() As ListObject
    Set ListTable = StorageSheet().ListObjects(LIST_TABLE)
End Function

' True if the storage sheet AND both tables resolve -- call this in a guard at
' the top of anything that writes, or that cannot run on defaults alone.
Public Function StorageReady() As Boolean
    On Error GoTo NotReady
    Dim lo As ListObject
    Set lo = SettingsTable()
    Set lo = ListTable()
    StorageReady = True
    Exit Function
NotReady:
    StorageReady = False
End Function

' ============================================================================
'  List readers -- turn a table column into a VBA array to loop over.
' ============================================================================

' Read one column of an Excel table into a 1-D, 1-based array of its values.
' Blank cells are dropped. Returns a 0-length array if the table is empty --
' callers should guard: If HasItems(arr) Then ...
Public Function TableColumnList(ByVal lo As ListObject, ByVal columnName As String) As Variant
    If lo.DataBodyRange Is Nothing Then
        TableColumnList = Array()
        Exit Function
    End If

    Dim data As Variant
    data = lo.ListColumns(columnName).DataBodyRange.Value

    ' A one-row table hands back a scalar, not a 2-D array. Normalise it so the
    ' loop below has only one shape to deal with.
    If Not IsArray(data) Then
        Dim one(1 To 1, 1 To 1) As Variant
        one(1, 1) = data
        data = one
    End If

    Dim out() As Variant, n As Long, i As Long, k As Long
    n = UBound(data, 1)
    ReDim out(1 To n)
    For i = 1 To n
        If Len(CStr(data(i, 1))) > 0 Then
            k = k + 1
            out(k) = data(i, 1)
        End If
    Next i

    If k = 0 Then
        TableColumnList = Array()
    Else
        ReDim Preserve out(1 To k)
        TableColumnList = out
    End If
End Function

' The user-editable list, ready to loop. Empty array if storage is missing.
Public Function ItemList() As Variant
    On Error GoTo NoItems
    ItemList = TableColumnList(ListTable(), LIST_COLUMN)
    Exit Function
NoItems:
    ItemList = Array()
End Function

' True if an array returned by the list readers has at least one item.
Public Function HasItems(ByVal arr As Variant) As Boolean
    If IsArray(arr) Then HasItems = (UBound(arr) >= LBound(arr))
End Function

' ============================================================================
'  Settings cache
' ============================================================================

' Drop the cached settings so the next read comes off the sheet again.
Public Sub InvalidateCache()
    Set mCache = Nothing
End Sub

' Load tblSettings into the cache. Silent on failure: mCache is left as an
' empty Dictionary so every lookup misses and every caller gets its DEF_*
' fallback. A damaged storage sheet must not stop the add-in from running.
Private Sub LoadCache()
    Set mCache = CreateObject("Scripting.Dictionary")
    mCache.CompareMode = 1                          ' 1 = vbTextCompare

    On Error GoTo Done

    Dim lo As ListObject
    Set lo = SettingsTable()
    If lo.DataBodyRange Is Nothing Then GoTo Done

    ' ONE read per column instead of a COM call per cell.
    '
    ' .Value, NOT .Value2: .Value2 returns a date as its serial number
    ' (12/31/2026 -> 46387). That would not error -- it would quietly hand back
    ' a number where a Date was expected, which shows up as a wrong report
    ' rather than a crash.
    Dim keys As Variant, vals As Variant
    keys = lo.ListColumns(SETTINGS_KEY_COLUMN).DataBodyRange.Value
    vals = lo.ListColumns(SETTINGS_VAL_COLUMN).DataBodyRange.Value

    Dim i As Long, k As String
    If Not IsArray(keys) Then                       ' one-row table -> scalars
        k = Trim$(CStr(keys))
        If Len(k) > 0 Then mCache(k) = vals
        GoTo Done
    End If

    For i = 1 To UBound(keys, 1)
        k = Trim$(CStr(keys(i, 1)))
        If Len(k) > 0 Then mCache(k) = vals(i, 1)   ' last duplicate key wins
    Next i

Done:
    On Error GoTo 0
End Sub

' Raw cached value for a key, or Empty if absent / blank.
Public Function SettingValue(ByVal key As String) As Variant
    If mCache Is Nothing Then LoadCache
    If mCache.Exists(key) Then
        If Len(Trim$(CStr(mCache(key)))) > 0 Then
            SettingValue = mCache(key)
            Exit Function
        End If
    End If
    SettingValue = Empty
End Function

Public Function HasSetting(ByVal key As String) As Boolean
    HasSetting = Not IsEmpty(SettingValue(key))
End Function

' Every setting name, in table order. The settings form builds its list from
' this, so adding a row to tblSettings makes it show up with no code change.
Public Function SettingKeys() As Variant
    On Error GoTo NoKeys
    SettingKeys = TableColumnList(SettingsTable(), SETTINGS_KEY_COLUMN)
    Exit Function
NoKeys:
    SettingKeys = Array()
End Function

' ---- Typed readers. Each returns the fallback on missing / blank / wrong type.
' The IsNumeric / IsDate guards stop a typo on the sheet ("fifteen") from
' raising a type mismatch inside an unrelated macro.

Public Function SettingDouble(ByVal key As String, ByVal fallback As Double) As Double
    Dim v As Variant
    v = SettingValue(key)
    If IsEmpty(v) Then
        SettingDouble = fallback
    ElseIf IsNumeric(v) Then
        SettingDouble = CDbl(v)
    Else
        SettingDouble = fallback
    End If
End Function

Public Function SettingDate(ByVal key As String, ByVal fallback As Date) As Date
    Dim v As Variant
    v = SettingValue(key)
    If IsEmpty(v) Then
        SettingDate = fallback
    ElseIf IsDate(v) Then
        SettingDate = CDate(v)
    ElseIf IsNumeric(v) Then
        ' A date cell formatted as General reads back as a serial number.
        ' Guard the range so a stray 100 does not become 4/9/1900.
        If CDbl(v) >= 1 And CDbl(v) <= 2958465 Then     ' 2958465 = 12/31/9999
            SettingDate = CDate(CDbl(v))
        Else
            SettingDate = fallback
        End If
    Else
        SettingDate = fallback
    End If
End Function

Public Function SettingString(ByVal key As String, ByVal fallback As String) As String
    Dim v As Variant
    v = SettingValue(key)
    If IsEmpty(v) Then
        SettingString = fallback
    Else
        SettingString = CStr(v)
    End If
End Function

' Accepts TRUE/FALSE cells, and the text Yes/No, Y/N, 1/0 a user might type.
Public Function SettingBool(ByVal key As String, ByVal fallback As Boolean) As Boolean
    Dim v As Variant
    v = SettingValue(key)
    If IsEmpty(v) Then
        SettingBool = fallback
    ElseIf VarType(v) = vbBoolean Then
        SettingBool = v
    Else
        Select Case UCase$(Trim$(CStr(v)))
            Case "TRUE", "YES", "Y", "1":   SettingBool = True
            Case "FALSE", "NO", "N", "0":   SettingBool = False
            Case Else:                      SettingBool = fallback
        End Select
    End If
End Function

' ============================================================================
'  The settings themselves. Rename / add to suit your add-in.
'
'  Named like the Const declarations they replace, so no call site needs
'  editing. Read at RUN time, not compile time -- which is the point, but
'  also means a value can change mid-session.
' ============================================================================

Public Property Get REPORT_TITLE() As String
    REPORT_TITLE = SettingString("REPORT_TITLE", DEF_REPORT_TITLE)
End Property

Public Property Get PERIOD_END() As Date
    PERIOD_END = SettingDate("PERIOD_END", DEF_PERIOD_END)
End Property

Public Property Get DEFAULT_ROW_HEIGHT() As Double
    DEFAULT_ROW_HEIGHT = SettingDouble("DEFAULT_ROW_HEIGHT", DEF_DEFAULT_ROW_HEIGHT)
End Property

Public Property Get SHOW_TIPS() As Boolean
    SHOW_TIPS = SettingBool("SHOW_TIPS", DEF_SHOW_TIPS)
End Property

' ============================================================================
'  Writers -- used by the settings form. Nothing else should write here.
'  All of them write into ThisWorkbook and NONE of them save: batch your
'  writes, then call SaveStorage once.
' ============================================================================

' Write one setting's Value cell, matched on the key column. Returns False if
' the key is not in the table -- the caller decides whether that is an error.
Public Function SetSetting(ByVal key As String, ByVal newValue As Variant) As Boolean
    On Error GoTo Failed
    mLastError = ""

    Dim lo As ListObject
    Set lo = SettingsTable()
    If lo.DataBodyRange Is Nothing Then
        mLastError = SETTINGS_TABLE & " has no rows."
        Exit Function
    End If

    Dim keyCol As Range, valCol As Range, i As Long
    Set keyCol = lo.ListColumns(SETTINGS_KEY_COLUMN).DataBodyRange
    Set valCol = lo.ListColumns(SETTINGS_VAL_COLUMN).DataBodyRange

    For i = 1 To keyCol.Rows.Count
        If StrComp(Trim$(CStr(keyCol.Cells(i, 1).Value)), key, vbTextCompare) = 0 Then
            valCol.Cells(i, 1).Value = newValue
            InvalidateCache
            SetSetting = True
            Exit Function
        End If
    Next i

    mLastError = "Setting '" & key & "' was not found in " & SETTINGS_TABLE & "."
    Exit Function

Failed:
    mLastError = "Error " & Err.Number & " writing '" & key & "': " & Err.Description
End Function

' Replace the whole list with the supplied 1-D array (any base; Array() = empty).
Public Function SetItemList(ByVal items As Variant) As Boolean
    On Error GoTo Failed
    mLastError = ""

    Dim n As Long, i As Long, src As Long
    If IsArray(items) Then n = UBound(items) - LBound(items) + 1

    Dim block As Variant
    If n > 0 Then
        Dim buf() As Variant
        ReDim buf(1 To n, 1 To 1)
        src = LBound(items)
        For i = 1 To n
            buf(i, 1) = items(src)
            src = src + 1
        Next i
        block = buf
    End If                                   ' n = 0 -> block stays Empty

    SetItemList = WriteTableBlock(ListTable(), block)
    InvalidateCache
    Exit Function

Failed:
    mLastError = "Error " & Err.Number & " writing the list: " & Err.Description
End Function

' Replace ALL body rows of a table with a 2-D (rows x columns) array, or with
' a single empty row when block is Empty. The single choke point for any
' writer that changes how many rows a table has.
'
' WHY ListObject.Resize AND NOT ListRows.Add / ListRows.Delete:
'   Add and Delete SHIFT CELLS down / up inside the table's own columns. If
'   another table below shares SOME but not all of those columns (wider,
'   narrower or offset), it cannot be moved whole, and Excel refuses with
'   run-time error 1004: "This won't work because it would move cells in a
'   table on your worksheet." Storage sheets crowd tables together, so this
'   bites sooner or later. Resize only redeclares where the table starts and
'   ends -- no cell moves, so neighbours are never at risk.
'
' WHY CHECK FIRST AND KEEP A COPY:
'   The table is cleared before it is rewritten. Refusing an entry AFTER the
'   clear, or failing part-way, would leave the table empty -- and the next
'   successful SaveStorage would save it that way. So every entry is checked
'   before anything is touched, and the previous contents are put back if the
'   write still fails.
Public Function WriteTableBlock(ByVal lo As ListObject, ByVal block As Variant) As Boolean
    On Error GoTo Failed
    mLastError = ""

    Dim n As Long, nCols As Long, r As Long, c As Long
    nCols = lo.ListColumns.Count
    If IsArray(block) Then
        n = UBound(block, 1) - LBound(block, 1) + 1
        If UBound(block, 2) - LBound(block, 2) + 1 <> nCols Then
            mLastError = "Expected " & nCols & " column(s) for " & lo.Name & "."
            Exit Function
        End If
        ' 1. Check everything BEFORE the table is touched.
        For r = LBound(block, 1) To UBound(block, 1)
            For c = LBound(block, 2) To UBound(block, 2)
                If Len(CStr(block(r, c))) > MAX_CELL_CHARS Then
                    mLastError = "Row " & (r - LBound(block, 1) + 1) & " of " & lo.Name & _
                                 " is longer than a cell can hold. Nothing was changed."
                    Exit Function
                End If
            Next c
        Next r
    End If

    ' A ListObject must always keep at least one body row, even an empty one.
    Dim rowsWanted As Long
    rowsWanted = n
    If rowsWanted < 1 Then rowsWanted = 1

    ' A totals row would be counted into the resize and end up as data.
    Dim hadTotals As Boolean
    hadTotals = lo.ShowTotals
    If hadTotals Then lo.ShowTotals = False

    Dim newRange As Range
    Set newRange = lo.Range.Cells(1, 1).Resize(rowsWanted + 1, nCols)

    ' 2. Refuse to grow into a neighbour rather than corrupt the sheet.
    Dim clash As String
    clash = OverlapsAnotherTable(lo, newRange)
    If Len(clash) > 0 Then
        If hadTotals Then lo.ShowTotals = True
        mLastError = lo.Name & " needs " & n & " rows, but growing it to " & _
                     newRange.Address(False, False) & " would run into '" & clash & "'." & _
                     " Move that table further away, or store fewer rows."
        Exit Function
    End If

    ' 3. Keep a copy so a failure below can be undone.
    Dim previous As Variant
    If Not mRestoring Then previous = BodyAs2D(lo)

    ' 4. Blank, resize, write. Blank FIRST: shrinking leaves rows below the new
    '    boundary outside the table, and stale values there look like data.
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    lo.Resize newRange
    If n > 0 Then lo.DataBodyRange.Value = block
    If hadTotals Then lo.ShowTotals = True

    WriteTableBlock = True
    Exit Function

Failed:
    mLastError = "Error " & Err.Number & " writing " & lo.Name & ": " & Err.Description
    If Not mRestoring And IsArray(previous) Then
        Dim saved As String
        saved = mLastError
        mRestoring = True
        WriteTableBlock lo, previous
        mRestoring = False
        mLastError = saved & " The previous contents were put back."
    End If
    If hadTotals Then lo.ShowTotals = True
End Function

' A table's body as a 2-D array ALWAYS, or Empty if it has no body.
' (Range.Value of a single cell is a scalar, not a 1x1 array -- the classic
' trap -- so it is wrapped here once instead of at every caller.)
Private Function BodyAs2D(ByVal lo As ListObject) As Variant
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant
    v = lo.DataBodyRange.Value
    If IsArray(v) Then
        BodyAs2D = v
    Else
        Dim one(1 To 1, 1 To 1) As Variant
        one(1, 1) = v
        BodyAs2D = one
    End If
End Function

' Name of the first OTHER table on the same sheet that target would overlap,
' or "" if the range is clear.
Private Function OverlapsAnotherTable(ByVal lo As ListObject, ByVal target As Range) As String
    Dim other As ListObject
    For Each other In lo.Parent.ListObjects
        If StrComp(other.Name, lo.Name, vbTextCompare) <> 0 Then
            If Not Application.Intersect(target, other.Range) Is Nothing Then
                OverlapsAnotherTable = other.Name
                Exit Function
            End If
        End If
    Next other
End Function

' Persist the add-in to disk.
'
' NOT optional. Writing to a loaded .xlam marks it dirty, but Excel closes
' add-ins at shutdown WITHOUT prompting to save. Skip this and the user's
' settings vanish at the end of the session with no warning at all.
Public Sub SaveStorage()
    On Error Resume Next
    ThisWorkbook.Save
End Sub

' ============================================================================
'  One-time setup
' ============================================================================

' Create the storage sheet and both tables if they are missing. Idempotent:
' anything that already exists is left exactly as it is.
'
' Layout: tblList in column B, tblSettings beside it in D:E, both starting on
' row 2. Side by side (not stacked) so either table can grow downward without
' running into the other.
Public Sub SetupStorage()
    Dim wb As Workbook, ws As Worksheet
    Set wb = ThisWorkbook

    On Error Resume Next
    Set ws = wb.Worksheets(STORAGE_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = STORAGE_SHEET
    End If

    If Not TableExists(ws, LIST_TABLE) Then
        ws.Range("B2").Value = LIST_COLUMN
        ws.Range("B3:B5").Value = Application.Transpose(Array("Alpha", "Beta", "Gamma"))
        ws.ListObjects.Add(xlSrcRange, ws.Range("B2:B5"), , xlYes).Name = LIST_TABLE
    End If

    If Not TableExists(ws, SETTINGS_TABLE) Then
        ws.Range("D2").Value = SETTINGS_KEY_COLUMN
        ws.Range("E2").Value = SETTINGS_VAL_COLUMN
        ws.Range("D3").Value = "REPORT_TITLE":        ws.Range("E3").Value = DEF_REPORT_TITLE
        ws.Range("D4").Value = "PERIOD_END":          ws.Range("E4").Value = DEF_PERIOD_END
        ws.Range("D5").Value = "DEFAULT_ROW_HEIGHT":  ws.Range("E5").Value = DEF_DEFAULT_ROW_HEIGHT
        ws.Range("D6").Value = "SHOW_TIPS":           ws.Range("E6").Value = DEF_SHOW_TIPS
        ws.Range("E4").NumberFormat = "m/d/yyyy"
        ws.ListObjects.Add(xlSrcRange, ws.Range("D2:E6"), , xlYes).Name = SETTINGS_TABLE
    End If

    ws.Columns("B:E").AutoFit
    InvalidateCache
End Sub

Private Function TableExists(ByVal ws As Worksheet, ByVal tableName As String) As Boolean
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next lo
End Function
