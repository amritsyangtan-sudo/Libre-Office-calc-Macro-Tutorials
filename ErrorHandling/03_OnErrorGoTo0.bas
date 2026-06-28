REM  *****  BASIC  *****

Sub OnErrorGoTo0Example()

    '-------------------------------------------------------
    ' On Error GoTo 0
    '
    ' - Disables the current error handler.
    ' - After this statement, any runtime error will stop
    '   the macro and display the normal error message.
    ' - Commonly used after On Error Resume Next to
    '   restore normal error handling.
    '
    ' Common Uses:
    '   • Restore default error handling
    '   • Limit the scope of ignored errors
    '   • Improve debugging
    '-------------------------------------------------------

    Dim number As Integer
    Dim result As Double

    ' Ignore errors
    On Error Resume Next

    number = 0

    ' This error is ignored
    result = 100 / number

    MsgBox "First error ignored."

    ' Restore normal error handling
    On Error GoTo 0

    ' This error is NOT ignored
    result = 100 / number

    MsgBox "This line will not execute."

End Sub
