Attribute VB_Name = "modAbout"
Option Explicit

' ============================================================================
'  modAbout
'  Identity of the add-in: its name, version and licence text.
'
'  This module holds the FACTS; frmAbout holds the LAYOUT. Two consequences:
'
'   * Bumping the version is a one-line edit HERE. Nothing else in the
'     project hardcodes a version number, so there is no second place to
'     forget. Dialog titles use ADDIN_NAME for the same reason.
'   * The licence text lives in code, not on the storage sheet. It must not
'     be user-editable -- the opposite of everything AddInStorage manages --
'     so keeping it out of the settings plumbing means nobody can blank it
'     out from the settings form.
' ============================================================================

' ----------------------------------------------------------------------------
'  EDIT THIS when you ship a new build.
' ----------------------------------------------------------------------------
Public Const ADDIN_NAME      As String = "My Add-in"
Public Const ADDIN_VERSION   As String = "1.0"
Public Const ADDIN_COPYRIGHT As String = "Copyright (c) 2026 Your Name"

' "My Add-in   Version 1.0" -- one place decides how a version is written.
Public Function VersionLine() As String
    VersionLine = ADDIN_NAME & "   Version " & ADDIN_VERSION
End Function

Public Function VersionString() As String
    VersionString = "Version " & ADDIN_VERSION
End Function

' The MIT licence. Swap in your own licence text if you use a different one.
'
' Built with one assignment per paragraph rather than a single statement over
' many `& _` continuations. VBA allows at most 24 continuations per statement, and
' a licence is exactly the kind of text that grows past a cap and then fails
' to compile at the worst moment. Appending in steps has no such limit.
'
' No hard line breaks inside a paragraph: the dialog's text box wraps to its
' own width, so hard-wrapping here would produce a ragged column.
Public Function LicenseText() As String
    Dim s As String

    s = "MIT License" & vbCrLf & vbCrLf
    s = s & ADDIN_COPYRIGHT & vbCrLf & vbCrLf

    s = s & "Permission is hereby granted, free of charge, to any person " & _
            "obtaining a copy of this software and associated documentation " & _
            "files (the ""Software""), to deal in the Software without " & _
            "restriction, including without limitation the rights to use, " & _
            "copy, modify, merge, publish, distribute, sublicense, and/or sell " & _
            "copies of the Software, and to permit persons to whom the " & _
            "Software is furnished to do so, subject to the following " & _
            "conditions:" & vbCrLf & vbCrLf

    s = s & "The above copyright notice and this permission notice shall be " & _
            "included in all copies or substantial portions of the Software." & _
            vbCrLf & vbCrLf

    s = s & "THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY " & _
            "KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE " & _
            "WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR " & _
            "PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR " & _
            "COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER " & _
            "LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, " & _
            "ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE " & _
            "USE OR OTHER DEALINGS IN THE SOFTWARE."

    LicenseText = s
End Function
