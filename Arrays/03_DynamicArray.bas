REM  *****  BASIC  *****

Sub DynamicArrayExample()

    '-------------------------------------------------------
    ' Dynamic Array
    '
    ' - A dynamic array does not have a fixed size.
    ' - ReDim changes the size of the array during runtime.
    '-------------------------------------------------------

    Dim numbers() As Integer
    Dim i As Integer

    ReDim numbers(4)

    For i = 0 To 4
        numbers(i) = (i + 1) * 10
    Next i

    For i = LBound(numbers) To UBound(numbers)
        MsgBox numbers(i)
    Next i

End Sub
