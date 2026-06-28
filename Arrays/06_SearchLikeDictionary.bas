Function GetValue(properties(), key As String) As Variant

    Dim i As Integer

    For i = LBound(properties) To UBound(properties)

        If properties(i).Name = key Then
            GetValue = properties(i).Value
            Exit Function
        End If

    Next i

    GetValue = Empty

End Function

Sub SearchPropertyValue()

    Dim properties(2) As New com.sun.star.beans.PropertyValue

    properties(0).Name = "Name"
    properties(0).Value = "John"

    properties(1).Name = "Age"
    properties(1).Value = 25

    properties(2).Name = "Country"
    properties(2).Value = "Nepal"

    MsgBox GetValue(properties, "Country")

End Sub
