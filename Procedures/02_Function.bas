REM  *****  BASIC  *****

Function AddNumbers(number1 As Integer, number2 As Integer) As Integer

    '-------------------------------------------------------
    ' Function
    '
    ' - A Function is similar to a Sub, but it returns
    '   a value.
    ' - The return value is assigned to the function name.
    '
    ' Common Uses:
    '   • Perform calculations
    '   • Validate data
    '   • Return processed values
    '-------------------------------------------------------

    AddNumbers = number1 + number2

End Function

Sub FunctionExample()

    Dim result As Integer

    result = AddNumbers(10, 20)

    MsgBox "Result: " & result

End Sub
