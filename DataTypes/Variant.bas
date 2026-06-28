Sub VariantExample()

    '-------------------------------------------------------
    ' Variant Data Type
    '
    ' - Can store different types of data.
    ' - If no data type is specified, the variable is
    '   automatically treated as a Variant.
    ' - Can hold numbers, text, dates, and Boolean values.
    ' - Useful when the data type is not known in advance.
    ' - Uses more memory than specific data types, so use
    '   Integer, Double, String, etc., whenever possible.
    '-------------------------------------------------------

    ' Declare a Variant variable
    Dim value

    ' Store an Integer
    value = 100
    MsgBox "Integer: " & value

    ' Store a String
    value = "Hello World"
    MsgBox "String: " & value

    ' Store a Boolean
    value = True
    MsgBox "Boolean: " & value

End Sub
