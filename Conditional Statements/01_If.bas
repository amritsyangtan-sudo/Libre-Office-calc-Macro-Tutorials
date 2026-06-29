REM ***** BASIC *****

Option Explicit

Sub IfExample()

    '-------------------------------------------------------
    ' If Statement
    '
    ' - Executes a block of code when a condition is True.
    ' - Used for decision making in programs.
    '
    ' Common Uses:
    '   • Validate user input
    '   • Compare values
    '   • Check conditions
    '   • Control program flow
    '-------------------------------------------------------

    Dim marks As Integer

    marks = 75

    If marks >= 50 Then

        MsgBox "Congratulations! You Passed."

    End If

End Sub
