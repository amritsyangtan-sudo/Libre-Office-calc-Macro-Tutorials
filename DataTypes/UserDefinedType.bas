Type Employee
    Id As Integer
    Name As String
    Salary As Currency
End Type

Sub TypeExample()

    '-------------------------------------------------------
    ' User-Defined Type (Type)
    '
    ' - A Type allows you to group related variables into
    '   a single custom data structure.
    ' - Similar to a record or structure in other languages.
    ' - Useful for organizing related data.
    '-------------------------------------------------------

    Dim emp As Employee

    emp.Id = 1
    emp.Name = "Amrit"
    emp.Salary = 25000.75

    MsgBox _
        "ID: " & emp.Id & Chr(10) & _
        "Name: " & emp.Name & Chr(10) & _
        "Salary: " & emp.Salary

End Sub
