REM  *****  BASIC  *****

Sub OneDimensionalArrayExample()

    '-------------------------------------------------------
    ' One-Dimensional Array
    '
    ' - A one-dimensional array stores multiple values
    '   in a single row of memory.
    ' - Each value is accessed using an index.
    '
    ' Common Uses:
    '   • Store names
    '   • Store numbers
    '   • Store prices
    '-------------------------------------------------------

    Dim fruits(2) As String

    fruits(0) = "Apple"
    fruits(1) = "Banana"
    fruits(2) = "Orange"

    Dim i As Integer

    For i = LBound(fruits) To UBound(fruits)
        MsgBox fruits(i)
    Next i

End Sub
