REM  *****  BASIC  *****

Sub ArrayFunctionsExample()

    '-------------------------------------------------------
    ' Array Functions
    '
    ' - Array functions are used to create, resize,
    '   access, and clear arrays.
    ' - Arrays store multiple values of the same data type
    '   under a single variable name.
    '
    ' Common Array Functions:
    '   • Array
    '   • LBound
    '   • UBound
    '   • ReDim
    '   • Erase
    '-------------------------------------------------------

    Dim fruits()
    Dim numbers(2) As Integer
    Dim i As Integer

    '----------------------------------------
    ' Array()
    '----------------------------------------
    fruits = Array("Apple", "Banana", "Orange")

    MsgBox "First Fruit: " & fruits(0)

    '----------------------------------------
    ' LBound() and UBound()
    '----------------------------------------
    MsgBox "Lower Bound: " & LBound(fruits)
    MsgBox "Upper Bound: " & UBound(fruits)

    '----------------------------------------
    ' ReDim
    '----------------------------------------
    ReDim numbers(4)

    For i = LBound(numbers) To UBound(numbers)
        numbers(i) = (i + 1) * 10
    Next i

    MsgBox "Last Number: " & numbers(UBound(numbers))

    '----------------------------------------
    ' Erase
    '----------------------------------------
    Erase numbers

    MsgBox "Array has been erased."

End Sub
