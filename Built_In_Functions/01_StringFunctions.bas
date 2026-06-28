REM  *****  BASIC  *****

Sub StringFunctionsExample()

    '-------------------------------------------------------
    ' String Functions
    '
    ' - String functions are used to manipulate text.
    ' - They allow you to measure, extract, search,
    '   compare, and modify strings.
    '
    ' Common String Functions:
    '   • Len
    '   • Left
    '   • Right
    '   • Mid
    '   • Trim
    '   • UCase
    '   • LCase
    '   • InStr
    '   • Replace
    '   • Split
    '-------------------------------------------------------

    Dim text As String
    Dim words()

    text = "  LibreOffice Basic Tutorial  "

    MsgBox "Original: '" & text & "'"

    ' Length of the string
    MsgBox "Len: " & Len(text)

    ' Remove leading and trailing spaces
    MsgBox "Trim: '" & Trim(text) & "'"

    ' First 12 characters
    MsgBox "Left: " & Left(text, 12)

    ' Last 8 characters
    MsgBox "Right: " & Right(text, 8)

    ' Extract characters
    MsgBox "Mid: " & Mid(text, 3, 12)

    ' Convert to uppercase
    MsgBox "UCase: " & UCase(text)

    ' Convert to lowercase
    MsgBox "LCase: " & LCase(text)

    ' Search for text
    MsgBox "InStr: " & InStr(text, "Basic")

    ' Replace text
    MsgBox "Replace: " & Replace(text, "Basic", "Macro")

    ' Split text
    words = Split(Trim(text), " ")

    MsgBox "First Word: " & words(0)

End Sub
