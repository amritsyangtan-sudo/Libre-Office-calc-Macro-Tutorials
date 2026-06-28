REM  *****  BASIC  *****

Sub WhileLoopExample()

    '-------------------------------------------------------
    ' While...Wend Loop
    '
    ' - Repeats a block of code while a condition is True.
    ' - The condition is checked before each iteration.
    ' - If the condition is False initially, the loop
    '   does not execute.
    '
    ' Common Uses:
    '   • Count values
    '   • Validate input
    '   • Repeat until a condition changes
    '-------------------------------------------------------

    Dim count As Integer

    count = 1

    While count <= 5

        MsgBox "Count: " & count

        count = count + 1

    Wend

End Sub
