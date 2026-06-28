REM  *****  BASIC  *****

Sub ErrObjectExample()

    '-------------------------------------------------------
    ' Err Object
    '
    ' - The Err object provides information about the
    '   most recent runtime error.
    ' - It is commonly used inside an error handler
    '   to determine what went wrong.
    '
    ' Common Properties:
    '   • Err     - Returns the error number.
    '   • Error$  - Returns the error description.
    '
    ' Common Uses:
    '   • Display error messages
    '   • Log errors
    '   • Debug applications
    '   • Handle different error types
    '-------------------------------------------------------

    On Error GoTo ErrorHandler

    Dim number As Integer
    Dim result As Double

    number = 0

    ' This statement causes a division-by-zero error
    result = 100 / number

    MsgBox result

    Exit Sub

ErrorHandler:

    MsgBox _
        "Error Number: " & Err & Chr(10) & _
        "Error Description: " & Error$

End Sub
