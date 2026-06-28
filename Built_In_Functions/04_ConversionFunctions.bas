REM  *****  BASIC  *****

Sub ConversionFunctionsExample()

    '-------------------------------------------------------
    ' Conversion Functions
    '
    ' - Conversion functions convert one data type into
    '   another.
    ' - They are commonly used when working with user
    '   input, files, calculations, and spreadsheets.
    '
    ' Common Conversion Functions:
    '   • CInt
    '   • CLng
    '   • CSng
    '   • CDbl
    '   • CCur
    '   • CStr
    '   • CDate
    '   • CBool
    '   • Val
    '   • Str
    '   • Hex
    '   • Oct
    '-------------------------------------------------------

    Dim textNumber As String
    Dim textDate As String
    Dim number As Double

    textNumber = "123.45"
    textDate = "2026-06-28"
    number = 255

    MsgBox "CInt: " & CInt(textNumber)

    MsgBox "CLng: " & CLng(textNumber)

    MsgBox "CSng: " & CSng(textNumber)

    MsgBox "CDbl: " & CDbl(textNumber)

    MsgBox "CCur: " & CCur(textNumber)

    MsgBox "CStr: " & CStr(100)

    MsgBox "CDate: " & CDate(textDate)

    MsgBox "CBool: " & CBool(1)

    MsgBox "Val: " & Val("123.45ABC")

    MsgBox "Str: " & Str(456)

    MsgBox "Hex: " & Hex(number)

    MsgBox "Oct: " & Oct(number)

End Sub
