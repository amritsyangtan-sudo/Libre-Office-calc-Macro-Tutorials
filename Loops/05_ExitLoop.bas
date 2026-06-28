REM  *****  BASIC  *****

Sub ExitLoopExample()

    '-------------------------------------------------------
    ' Exit Loop
    '
    ' - Exit For immediately exits a For...Next loop.
    ' - Exit Do immediately exits a Do...Loop.
    ' - Useful when a condition has been met and
    '   further looping is unnecessary.
    '
    ' Common Uses:
    '   • Search operations
    '   • Stop processing when found
    '   • Exit on error
    '-------------------------------------------------------

    Dim i As Integer

    For i = 1 To 10

        If i = 5 Then
            Exit For
        End If

        MsgBox "Count: " & i

    Next i

    MsgBox "Loop finished."

End Sub
