Attribute VB_Name = "modRibbon"
Option Explicit

' ============================================================================
'  modRibbon -- every procedure the ribbon XML names, and nothing else.
'
'  Matches templates/customUI14-template.xml. Each callback is a thin shell
'  that hands off to a plain procedure, so the real work can be called (and
'  tested) without a ribbon: from the Immediate window, from another macro, or
'  from Python through Application.Run, which cannot supply an IRibbonControl.
'
'  ---------------------------------------------------------------------------
'  RULES THAT FAIL SILENTLY IF BROKEN
'  ---------------------------------------------------------------------------
'  1. Every callback named in the XML must exist, be Public, live in a
'     STANDARD module, and have exactly the signature the ribbon expects.
'     Missing, Private, misspelled or wrong-signature: the button does
'     nothing, or Excel shows a vague "cannot run the macro" -- never which.
'  2. Callback names resolve by bare name across the whole project. The same
'     Public name in two modules is ambiguous, and the button breaks. Prefix
'     them (Ribbon_*) so they are easy to find and hard to collide.
'  3. An unhandled error in ANY callback can reset the VBA project, which
'     throws away mRibbon below. The get* callbacks therefore swallow errors
'     and return a safe default.
'
'  Signatures (the ones used here):
'     onLoad         Sub X(ribbon As IRibbonUI)
'     button         Sub X(control As IRibbonControl)
'     toggleButton   Sub X(control As IRibbonControl, pressed As Boolean)
'     getLabel / getEnabled / getPressed / getImage / getVisible /
'     getScreentip / getSupertip
'                    Sub X(control As IRibbonControl, ByRef returnedVal)
' ============================================================================

' The ribbon object, captured once at load. The ONLY way VBA can ask Excel to
' re-query a control (call getLabel / getEnabled again) is through this.
Private mRibbon As IRibbonUI

' Position in the list for the Insert Next Item button.
Private mNextIndex As Long

' ============================================================================
'  Ribbon lifecycle
' ============================================================================

' Wired by onLoad="Ribbon_OnLoad" on the <customUI> root element. Without
' that attribute this never runs, mRibbon stays Nothing, and no dynamic label
' or enabled state can ever change after Excel starts. Nothing errors -- the
' controls just freeze.
Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set mRibbon = ribbon
End Sub

' Ask Excel to call the get* callbacks again: one control, or all of them.
'
' mRibbon is lost whenever VBA state resets (an unhandled error, or editing
' code in the VBA editor). The Is Nothing guard degrades that to "the ribbon
' stops refreshing until Excel restarts" instead of error 91 on every click.
Public Sub RefreshRibbon(Optional ByVal controlId As String = "")
    If mRibbon Is Nothing Then Exit Sub
    On Error Resume Next
    If Len(controlId) = 0 Then
        mRibbon.Invalidate
    Else
        mRibbon.InvalidateControl controlId
    End If
End Sub

' True once the ribbon has loaded and has not been lost since. For tests.
Public Function RibbonLoaded() As Boolean
    RibbonLoaded = Not mRibbon Is Nothing
End Function

' ============================================================================
'  Buttons  --  Sub X(control As IRibbonControl)
' ============================================================================

Public Sub Ribbon_InsertNextItem(control As IRibbonControl)
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a cell first.", vbInformation, modAbout.ADDIN_NAME
        Exit Sub
    End If

    If Len(InsertNextItem(ActiveCell)) = 0 Then
        MsgBox "The item list is empty. Add items under Settings.", _
               vbInformation, modAbout.ADDIN_NAME
    End If

    RefreshRibbon "btnInsertItem"             ' its label shows the NEXT item
End Sub

Public Sub Ribbon_ShowSettings(control As IRibbonControl)
    LaunchSettings
End Sub

Public Sub Ribbon_ShowAbout(control As IRibbonControl)
    LaunchAbout
End Sub

Public Sub Ribbon_ReloadSettings(control As IRibbonControl)
    AddInStorage.InvalidateCache
    RefreshRibbon
    MsgBox "Settings reloaded from the add-in's storage sheet.", _
           vbInformation, modAbout.ADDIN_NAME
End Sub

' Emergency reset if a crashed macro left Excel with events off or the
' screen frozen. See AppStateManager.
Public Sub Ribbon_ResetExcelState(control As IRibbonControl)
    AppStateManager.FastModeReset
End Sub

' ============================================================================
'  Toggle button  --  onAction gets a second argument: the new state
' ============================================================================

Public Sub Ribbon_TipsToggle(control As IRibbonControl, pressed As Boolean)
    If AddInStorage.SetSetting("SHOW_TIPS", pressed) Then AddInStorage.SaveStorage
    RefreshRibbon control.ID                  ' re-ask getPressed and getImage
End Sub

Public Sub Ribbon_TipsPressed(control As IRibbonControl, ByRef returnedVal)
    On Error Resume Next
    returnedVal = SHOW_TIPS
End Sub

' getImage may return a built-in icon NAME (an imageMso string) instead of a
' picture -- the cheapest way to swap an icon with state.
Public Sub Ribbon_TipsImage(control As IRibbonControl, ByRef returnedVal)
    On Error Resume Next
    If SHOW_TIPS Then returnedVal = "Info" Else returnedVal = "Help"
End Sub

' ============================================================================
'  Shared get* callbacks -- one procedure, Select Case on control.ID
' ============================================================================

Public Sub Ribbon_GetLabel(control As IRibbonControl, ByRef returnedVal)
    On Error Resume Next
    Select Case control.ID
        Case "btnInsertItem"
            Dim nxt As String
            nxt = PeekNextItem()
            If Len(nxt) = 0 Then
                returnedVal = "Insert Item"
            Else
                returnedVal = "Insert " & nxt
            End If
        Case Else
            returnedVal = control.ID          ' visible clue that a case is missing
    End Select
End Sub

Public Sub Ribbon_GetEnabled(control As IRibbonControl, ByRef returnedVal)
    On Error Resume Next
    returnedVal = True
    Select Case control.ID
        Case "btnInsertItem"
            ' Nothing to write into with no workbook open. Kept current by the
            ' Application events in ThisWorkbook, which call RefreshRibbon.
            returnedVal = Not (ActiveWorkbook Is Nothing)
    End Select
End Sub

' ============================================================================
'  The work itself -- no ribbon types, so callable from anywhere
' ============================================================================

' Write the next list item into target and advance. Returns the text written,
' or "" if the list is empty.
Public Function InsertNextItem(ByVal target As Range) As String
    Dim items As Variant
    items = AddInStorage.ItemList()
    If Not AddInStorage.HasItems(items) Then Exit Function

    Dim n As Long, k As Long
    n = UBound(items) - LBound(items) + 1
    k = mNextIndex Mod n                      ' wraps if the list shrank
    target.Value = items(LBound(items) + k)
    InsertNextItem = CStr(items(LBound(items) + k))
    mNextIndex = k + 1
End Function

' The item the next InsertNextItem will write, without advancing.
Public Function PeekNextItem() As String
    Dim items As Variant
    items = AddInStorage.ItemList()
    If Not AddInStorage.HasItems(items) Then Exit Function
    PeekNextItem = CStr(items(LBound(items) + (mNextIndex Mod (UBound(items) - LBound(items) + 1))))
End Function

' ============================================================================
'  Dialog launchers -- callable from the Immediate window too
' ============================================================================

Public Sub LaunchSettings()
    On Error GoTo Failed

    If Not AddInStorage.StorageReady() Then
        MsgBox "The add-in cannot find its storage sheet, so settings are not " & _
               "available." & vbCrLf & vbCrLf & _
               "It is still running normally on its built-in defaults.", _
               vbExclamation, modAbout.ADDIN_NAME
        Exit Sub
    End If

    ' Settings edits write to the add-in workbook, so events must be alive --
    ' do NOT wrap this in AppStateManager.FastModeOn.
    Dim f As frmSettings
    Set f = New frmSettings                   ' explicit instance, not the default one
    f.Show vbModal

    Dim didSave As Boolean
    didSave = f.Saved
    Unload f
    Set f = Nothing

    If didSave Then
        RefreshRibbon
        Application.StatusBar = modAbout.ADDIN_NAME & " settings saved."
        ' OnTime resolves a bare procedure name against the ACTIVE workbook,
        ' not the add-in, so from an .xlam it cannot find the macro.
        ' Qualifying it with the add-in's own file name makes it resolve.
        Application.OnTime Now + TimeSerial(0, 0, 3), _
            "'" & ThisWorkbook.Name & "'!modRibbon.ClearStatus"
    End If
    Exit Sub

Failed:
    MsgBox "Could not open the settings dialog." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, modAbout.ADDIN_NAME
End Sub

Public Sub LaunchAbout()
    On Error GoTo Failed
    Dim f As frmAbout
    Set f = New frmAbout
    f.Show vbModal
    Unload f
    Set f = Nothing
    Exit Sub

Failed:
    ' The form is the nice presentation, not the only one: if it fails to
    ' build, the name and version still reach the user.
    MsgBox modAbout.VersionLine() & vbCrLf & vbCrLf & _
           "The About window could not be opened." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "About " & modAbout.ADDIN_NAME
End Sub

' Hands the status bar back to Excel after the "saved" message.
Public Sub ClearStatus()
    On Error Resume Next
    Application.StatusBar = False
End Sub
