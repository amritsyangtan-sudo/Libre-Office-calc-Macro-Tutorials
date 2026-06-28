REM  *****  BASIC  *****

Sub ResumeLabelExample()

    '-------------------------------------------------------
    ' Resume <Label>
    '
    ' - Resume <Label> is used inside an error handler.
    ' - Instead of retrying the failed statement, execution
    '   continues from a specified label.
    ' - Useful when you want to skip a failed section and
    '   continue with another part of the program.
    '
    ' Common Uses:
    '   • Skip a failed operation
    '   • Continue processing
    '   • Perform cleanup before exiting
    '-------------------------------------------------------

    On Error GoTo ErrorHandler

    Dim number As Integer
    Dim result As Double

    number = 0

    ' This statement causes a division-by-zero error
    result = 100 / number

    MsgBox "Result: " & result

ContinueExecution:

    MsgBox "Program completed successfully."

    Exit Sub

ErrorHandler:

    MsgBox "An error occurred."

    ' Continue execution from the specified label
    Resume ContinueExecution

End Sub
