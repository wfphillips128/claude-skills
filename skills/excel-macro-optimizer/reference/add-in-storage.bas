Attribute VB_Name = "AddInStorage"
Option Explicit

' ============================================================================
'  AddInStorage
'  The ONE module that knows where the XL Edge add-in keeps its OWN data.
'  Drop this standard module into the .XLAM add-in once.
'
'  Why this module exists:
'   * The add-in stores its own lists, settings, and LAMBDA functions on a
'     HIDDEN "reference" worksheet inside THISWORKBOOK (the add-in) -- NOT in
'     ActiveWorkbook (the user's file the macros operate on).
'   * Qualifying add-in storage to ActiveWorkbook is a top-tier add-in bug: it
'     errors (the user's file has no such sheet) or silently hits a same-named
'     sheet the user happens to have open. Everything here uses ThisWorkbook.
'   * Every storage reference is funneled through the accessors below, so a
'     sheet/table rename -- or the in-progress migration from the EXISTING
'     sheet/table to the NEW one -- is a one-line change here, not a hunt
'     across every macro in the project.
'
'  Usage from any macro:
'     Dim names As Variant, i As Long
'     names = AddInStorage.CompanyList()        ' user-customizable, was hardcoded
'     For i = LBound(names) To UBound(names)
'         Debug.Print names(i)
'     Next i
'
'     Dim lo As ListObject: Set lo = AddInStorage.LambdaTable()
' ============================================================================

' ----------------------------------------------------------------------------
'  EDIT ONLY THIS BLOCK to match the workbook.
' ----------------------------------------------------------------------------
Private Const REFERENCE_SHEET     As String = "shtReference"    ' NEW hidden reference sheet
Private Const LAMBDA_TABLE        As String = "tblAlonzoChurch" ' NEW table of LAMBDA definitions
Private Const COMPANY_TABLE       As String = "tblCompany"      ' list of company/department names
Private Const COMPANY_COLUMN      As Variant = "Company"        ' column header (or 1-based index) to read

' Migration: the EXISTING sheet/table some features still read. Retire when done.
Private Const REFERENCE_SHEET_OLD As String = "Reference"       ' EXISTING reference sheet (legacy)
Private Const LAMBDA_TABLE_OLD    As String = "Functions"       ' EXISTING LAMBDA table (legacy)

' ============================================================================
'  Sheet / table accessors -- ALWAYS from ThisWorkbook, never ActiveWorkbook.
' ============================================================================

' The add-in's NEW hidden reference sheet (lists + LAMBDA table live here).
Public Function ReferenceSheet() As Worksheet
    Set ReferenceSheet = ThisWorkbook.Worksheets(REFERENCE_SHEET)
End Function

' The EXISTING (legacy) reference sheet -- for features not yet migrated.
' Confirm the macro's intent before pointing new code here.
Public Function ReferenceSheetOld() As Worksheet
    Set ReferenceSheetOld = ThisWorkbook.Worksheets(REFERENCE_SHEET_OLD)
End Function

' The NEW table of LAMBDA definitions.
Public Function LambdaTable() As ListObject
    Set LambdaTable = ReferenceSheet().ListObjects(LAMBDA_TABLE)
End Function

' The EXISTING (legacy) LAMBDA table -- for features not yet migrated.
Public Function LambdaTableOld() As ListObject
    Set LambdaTableOld = ReferenceSheetOld().ListObjects(LAMBDA_TABLE_OLD)
End Function

' The user-customizable list of company / department names.
Public Function CompanyTable() As ListObject
    Set CompanyTable = ReferenceSheet().ListObjects(COMPANY_TABLE)
End Function

' ============================================================================
'  List readers -- turn a table column into a VBA array to loop over.
' ============================================================================

' Read one column of an Excel table into a 1-D, 1-based array of its values.
' colRef is the column header (String) or its 1-based position (Long).
' Blank trailing cells are dropped. Returns a 0-length array if the table is
' empty -- callers should guard: If HasItems(arr) Then ...
Public Function TableColumnList(ByVal lo As ListObject, ByVal colRef As Variant) As Variant
    If lo.DataBodyRange Is Nothing Then
        TableColumnList = Array()                  ' no rows -> empty array
        Exit Function
    End If

    Dim lc As ListColumn
    If VarType(colRef) = vbString Then
        Set lc = lo.ListColumns(CStr(colRef))
    Else
        Set lc = lo.ListColumns(CLng(colRef))
    End If

    Dim data As Variant
    data = lc.DataBodyRange.Value2                  ' (rows, 1) 2-D, or scalar if 1 row

    Dim out() As Variant, n As Long, i As Long, k As Long
    If Not IsArray(data) Then                       ' single-row table -> scalar
        If Len(CStr(data)) = 0 Then
            TableColumnList = Array()
        Else
            ReDim out(1 To 1): out(1) = data
            TableColumnList = out
        End If
        Exit Function
    End If

    n = UBound(data, 1)
    ReDim out(1 To n)
    k = 0
    For i = 1 To n
        If Len(CStr(data(i, 1))) > 0 Then           ' skip blanks
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

' The company/department names, ready to loop. Replaces the hardcoded list.
Public Function CompanyList() As Variant
    CompanyList = TableColumnList(CompanyTable(), COMPANY_COLUMN)
End Function

' True if an array returned by the list readers actually has at least one item.
Public Function HasItems(ByVal arr As Variant) As Boolean
    If IsArray(arr) Then HasItems = (UBound(arr) >= LBound(arr))
End Function

' True if the reference sheet can be resolved -- call this in a guard at the
' top of macros that depend on add-in storage.
Public Function StorageReady() As Boolean
    On Error GoTo NotReady
    Dim ws As Worksheet
    Set ws = ReferenceSheet()
    StorageReady = Not ws Is Nothing
    Exit Function
NotReady:
    StorageReady = False
End Function
