REM  *****  BASIC  *****

Sub ResumeExample()

    '-------------------------------------------------------
    ' Resume Statement
    '
    ' - Resume is used inside an error handler.
    ' - It returns program execution to the statement
    '   that caused the error.
    ' - The error must be corrected before using Resume,
    '   otherwise the same error will occur again.
    '
    ' Common Uses:
    '   • Retry an operation
    '   • Recover from temporary errors
    '   • Correct invalid data before continuing
    '-------------------------------------------------------

    On Error GoTo ErrorHandler

    Dim number As Integer
    Dim result As Double

    number = 0

RetryCalculation:

    ' This statement causes a division-by-zero error
    result = 100 / number

    MsgBox "Result: " & result

    Exit Sub

ErrorHandler:

    MsgBox "Division by zero detected."

    ' Correct the error
    number = 10

    ' Retry the statement that caused the error
    Resume

End Sub
