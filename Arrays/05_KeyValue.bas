REM  *****  BASIC  *****

Sub PropertyValueExample()

    '-------------------------------------------------------
    ' PropertyValue (Dictionary-like Object)
    '
    ' - PropertyValue stores data as a Name-Value pair.
    ' - It is commonly used throughout the UNO API to
    '   pass parameters to services and methods.
    ' - Although it is not a true Dictionary, it is
    '   the closest built-in alternative in LibreOffice.
    '
    ' Common Uses:
    '   • Store key-value pairs
    '   • Pass UNO parameters
    '   • Configure services
    '-------------------------------------------------------

    Dim properties(2) As New com.sun.star.beans.PropertyValue
    Dim i As Integer

    '----------------------------------------
    ' Add Name-Value pairs
    '----------------------------------------

    properties(0).Name = "Name"
    properties(0).Value = "John"

    properties(1).Name = "Age"
    properties(1).Value = 25

    properties(2).Name = "Country"
    properties(2).Value = "Nepal"

    '----------------------------------------
    ' Display all values
    '----------------------------------------

    For i = LBound(properties) To UBound(properties)

        MsgBox _
            "Key   : " & properties(i).Name & Chr(10) & _
            "Value : " & properties(i).Value

    Next i

End Sub
