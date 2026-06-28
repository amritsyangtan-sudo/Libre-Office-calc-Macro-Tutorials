REM  *****  BASIC  *****

Sub OnErrorExample()

    '-------------------------------------------------------
    ' On Error
    '
    ' - Error handling prevents a macro from stopping
    '   unexpectedly when a runtime error occurs.
    ' - On Error GoTo transfers program execution to an
    '   error-handling section when an error occurs.
    ' - This allows the macro to display an error message,
    '   clean up resources, or recover gracefully.
    '
    ' Common Uses:
    '   • Handle missing files
    '   • Handle invalid user input
    '   • Handle division by zero
    '   • Handle UNO API errors
    '-------------------------------------------------------

    On Error GoTo ErrorHandler

    Dim number As Integer
    Dim result As Double

    number = 0

    ' This statement causes a division-by-zero error
    result = 100 / number

    MsgBox "Result: " & result

    Exit Sub

ErrorHandler:

    MsgBox "An unexpected error occurred."

End Sub
