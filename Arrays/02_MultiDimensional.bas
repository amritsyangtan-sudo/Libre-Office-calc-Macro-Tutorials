REM  *****  BASIC  *****

Sub MultiDimensionalArrayExample()

    '-------------------------------------------------------
    ' Multi-Dimensional Array
    '
    ' - A multi-dimensional array stores data in rows
    '   and columns.
    ' - Useful for tables and matrices.
    '-------------------------------------------------------

    Dim students(1,1) As String

    students(0,0) = "John"
    students(0,1) = "IT"

    students(1,0) = "Alice"
    students(1,1) = "HR"

    MsgBox students(0,0)
    MsgBox students(0,1)
    MsgBox students(1,0)
    MsgBox students(1,1)

End Sub
