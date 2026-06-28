REM  *****  BASIC  *****

Sub DoWhileExample()

    '-------------------------------------------------------
    ' Do While Loop
    '
    ' - Repeats while a condition is True.
    ' - The condition is checked before each iteration.
    '
    ' Common Uses:
    '   • Read files
    '   • Process user input
    '   • Repeat until a condition changes
    '-------------------------------------------------------

    Dim count As Integer

    count = 1

    Do While count <= 5

        MsgBox "Count: " & count

        count = count + 1

    Loop

End Sub
