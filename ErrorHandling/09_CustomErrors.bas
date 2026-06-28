REM  *****  BASIC  *****

Sub CustomErrorsExample()

    '-------------------------------------------------------
    ' Custom Error Handling
    '
    ' - LibreOffice Basic does not support creating
    '   custom runtime errors like some programming
    '   languages.
    ' - Instead, validate your data and handle invalid
    '   conditions manually.
    '
    ' Common Uses:
    '   • Validate user input
    '   • Check business rules
    '   • Verify file existence
    '   • Prevent invalid operations
    '-------------------------------------------------------

    Dim age As Integer

    age = -5

    ' Validate the input
    If age < 0 Then
        MsgBox "Custom Error: Age cannot be negative."
        Exit Sub
    End If

    MsgBox "Age: " & age

End Sub
