REM  *****  BASIC  *****

Sub ByValExample()

    '-------------------------------------------------------
    ' ByVal
    '
    ' - ByVal passes a copy of the variable.
    ' - Changes made inside the procedure do not affect
    '   the original variable.
    '-------------------------------------------------------

    Dim number As Integer

    number = 10

    Increase(ByVal number)

    MsgBox "Number: " & number

End Sub

Sub Increase(ByVal value As Integer)

    value = value + 10

    MsgBox "Inside Procedure: " & value

End Sub
