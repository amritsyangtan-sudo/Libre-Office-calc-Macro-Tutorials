REM  *****  BASIC  *****

Sub ByRefExample()

    '-------------------------------------------------------
    ' ByRef
    '
    ' - ByRef passes the original variable.
    ' - Changes made inside the procedure affect the
    '   original variable.
    '-------------------------------------------------------

    Dim number As Integer

    number = 10

    Increase(number)

    MsgBox "Number: " & number

End Sub

Sub Increase(ByRef value As Integer)

    value = value + 10

End Sub
