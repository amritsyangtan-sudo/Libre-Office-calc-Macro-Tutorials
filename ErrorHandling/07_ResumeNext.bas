REM  *****  BASIC  *****

Sub ResumeNextExample()

    '-------------------------------------------------------
    ' Resume Next
    '
    ' - Resume Next is used inside an error handler.
    ' - It skips the statement that caused the error
    '   and continues execution with the next statement.
    ' - Unlike Resume, it does not retry the failed
    '   statement.
    '
    ' Common Uses:
    '   • Skip non-critical errors
    '   • Continue processing records
    '   • Ignore invalid data
    '-------------------------------------------------------

    On Error GoTo ErrorHandler

    Dim number As Integer
    Dim result As Double

    number = 0

    ' This statement causes a division-by-zero error
    result = 100 / number

    ' Execution continues here after Resume Next
    MsgBox "Program continued after the error."

    Exit Sub

ErrorHandler:

    MsgBox "An error occurred."

    ' Skip the statement that caused the error
    Resume Next

End Sub
