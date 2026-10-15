Attribute VB_Name = "modSample"
Option Explicit

' Sample for: /opti
' Highlights every negative number in the used range of the active sheet and
' writes a count to the status bar. Correct, and slow on a large sheet: it
' touches every cell through the object model, selects things, and leaves
' screen updating and automatic calculation on.

Public Sub HighlightNegatives()
    Dim c As Range
    Dim n As Long
    For Each c In ActiveSheet.UsedRange
        c.Select
        If IsNumeric(c.Value) And Not IsEmpty(c.Value) Then
            If c.Value < 0 Then
                Selection.Interior.Color = RGB(255, 199, 206)
                Selection.Font.Color = RGB(156, 0, 6)
                n = n + 1
            End If
        End If
    Next c
    Application.StatusBar = n & " negative values highlighted"
End Sub
