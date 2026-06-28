REM  *****  BASIC  *****

Sub InputBoxExample()

    '-------------------------------------------------------
    ' Input Box (InputBox)
    '
    ' - Displays a dialog that prompts the user to enter
    '   a value.
    ' - Returns the text entered by the user.
    ' - Useful for obtaining simple user input without
    '   creating a custom dialog.
    '
    ' Common Uses:
    '   • Enter a name
    '   • Enter an age
    '   • Enter a search value
    '   • Enter a cell reference
    '-------------------------------------------------------

    Dim userName As String

    ' Display the Input Box
    userName = InputBox( _
        "Enter your name:", _
        "User Input", _
        "")

    ' Check whether the user entered a value
    If userName <> "" Then
        MsgBox "Hello, " & userName & "!"
    Else
        MsgBox "No input was entered."
    End If

End Sub
