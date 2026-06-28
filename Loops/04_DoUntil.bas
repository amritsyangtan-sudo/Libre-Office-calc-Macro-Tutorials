REM  *****  BASIC  *****

Sub DoUntilExample()

    '-------------------------------------------------------
    ' Do Until Loop
    '
    ' - Repeats until a condition becomes True.
    ' - Continues looping while the condition is False.
    '
    ' Common Uses:
    '   • Wait for a condition
    '   • Validate input
    '   • Process data until complete
    '-------------------------------------------------------

    Dim count As Integer

    count = 1

    Do Until count > 5

        MsgBox "Count: " & count

        count = count + 1

    Loop

End Sub
