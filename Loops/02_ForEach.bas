REM  *****  BASIC  *****

Sub ForEachExample()

    '-------------------------------------------------------
    ' For Each...Next Loop
    '
    ' - Iterates through every element in a collection
    '   or array.
    ' - No index variable is required.
    '
    ' Common Uses:
    '   • Process arrays
    '   • Loop through worksheets
    '   • Loop through UNO collections
    '-------------------------------------------------------

    Dim fruits
    Dim item

    fruits = Array("Apple", "Banana", "Orange")

    For Each item In fruits
        MsgBox item
    Next item

End Sub
