REM  *****  BASIC  *****

Sub OnErrorResumeNextExample()

    '-------------------------------------------------------
    ' On Error Resume Next
    '
    ' - Enables error handling and continues execution
    '   with the next statement when an error occurs.
    ' - The macro does not stop when a runtime error
    '   is encountered.
    ' - Use this statement with caution because it can
    '   hide programming errors.
    '
    ' Common Uses:
    '   • Optional file operations
    '   • Checking object availability
    '   • Attempting non-critical operations
    '-------------------------------------------------------

    On Error Resume Next

    Dim number As Integer
    Dim result As Double

    number = 0

    ' This causes a division-by-zero error,
    ' but the macro continues running.
    result = 100 / number

    MsgBox "The macro continued after the error."

    MsgBox "Program execution completed."

End Sub
