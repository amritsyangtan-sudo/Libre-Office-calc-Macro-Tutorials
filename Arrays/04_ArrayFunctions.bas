REM  *****  BASIC  *****

Sub ArrayFunctionsExample()

    '-------------------------------------------------------
    ' Array Functions
    '
    ' - Array functions help create, resize, and
    '   manage arrays.
    '
    ' Common Functions:
    '   • Array
    '   • LBound
    '   • UBound
    '   • ReDim
    '   • Erase
    '-------------------------------------------------------

    Dim colors()
    Dim i As Integer

    colors = Array("Red", "Green", "Blue")

    MsgBox "Lower Bound: " & LBound(colors)
    MsgBox "Upper Bound: " & UBound(colors)

    For i = LBound(colors) To UBound(colors)
        MsgBox colors(i)
    Next i

    Erase colors

    MsgBox "Array erased."

End Sub
