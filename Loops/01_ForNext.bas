REM  *****  BASIC  *****

Sub ForNextExample()

    '-------------------------------------------------------
    ' For...Next Loop
    '
    ' - Repeats a block of code a specified number
    '   of times.
    ' - Best used when the number of iterations is known.
    '
    ' Common Uses:
    '   • Count numbers
    '   • Process arrays
    '   • Loop through rows or columns
    '-------------------------------------------------------

    Dim i As Integer

    For i = 1 To 5
        MsgBox "Count: " & i
    Next i

End Sub
