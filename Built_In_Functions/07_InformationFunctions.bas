REM  *****  BASIC  *****

Sub InformationFunctionsExample()

    '-------------------------------------------------------
    ' Information Functions
    '
    ' - Information functions determine the data type
    '   or state of a variable or expression.
    ' - They are useful for validation, debugging,
    '   and writing flexible code.
    '
    ' Common Information Functions:
    '   • IsArray
    '   • IsDate
    '   • IsEmpty
    '   • IsError
    '   • IsMissing
    '   • IsNull
    '   • IsNumeric
    '   • IsObject
    '   • TypeName
    '   • VarType
    '-------------------------------------------------------

    Dim numbers(2) As Integer
    Dim text As String
    Dim obj As Object
    Dim value As Variant

    text = "123"
    value = Empty
    obj = ThisComponent

    '----------------------------------------
    ' IsArray()
    '----------------------------------------
    MsgBox "IsArray: " & IsArray(numbers)

    '----------------------------------------
    ' IsDate()
    '----------------------------------------
    MsgBox "IsDate: " & IsDate("2026-06-28")

    '----------------------------------------
    ' IsEmpty()
    '----------------------------------------
    MsgBox "IsEmpty: " & IsEmpty(value)

    '----------------------------------------
    ' IsNumeric()
    '----------------------------------------
    MsgBox "IsNumeric: " & IsNumeric(text)

    '----------------------------------------
    ' IsObject()
    '----------------------------------------
    MsgBox "IsObject: " & IsObject(obj)

    '----------------------------------------
    ' TypeName()
    '----------------------------------------
    MsgBox "TypeName: " & TypeName(text)

    '----------------------------------------
    ' VarType()
    '----------------------------------------
    MsgBox "VarType: " & VarType(text)

    '----------------------------------------
    ' Notes
    '----------------------------------------
    '
    ' IsNull()    - Checks whether a Variant contains Null.
    ' IsError()   - Checks whether an expression contains
    '               an error value.
    ' IsMissing() - Used only with Optional procedure
    '               parameters.
    '
    ' These functions are typically used in more advanced
    ' programming scenarios.
    '----------------------------------------

End Sub
