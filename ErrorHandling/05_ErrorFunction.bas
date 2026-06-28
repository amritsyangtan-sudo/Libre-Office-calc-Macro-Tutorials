REM  *****  BASIC  *****

Sub ErrorFunctionExample()

    '-------------------------------------------------------
    ' Error$ Function
    '
    ' - Error$ returns the description of a runtime error.
    ' - It is typically used inside an error handler.
    ' - If an error number is supplied, Error$(number)
    '   returns the corresponding error description.
    '
    ' Common Uses:
    '   • Display readable error messages
    '   • Log error descriptions
    '   • Debug runtime errors
    '-------------------------------------------------------

    On Error GoTo ErrorHandler

    Dim number As Integer
    Dim result As Double

    number = 0

    ' This statement causes a division-by-zero error
    result = 100 / number

    Exit Sub

ErrorHandler:

    MsgBox _
        "Error Number : " & Err & Chr(10) & _
        "Description  : " & Error$

End Sub
