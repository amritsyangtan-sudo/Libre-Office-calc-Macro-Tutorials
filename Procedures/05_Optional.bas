REM  *****  BASIC  *****

Sub OptionalExample()

    '-------------------------------------------------------
    ' Optional Parameters
    '
    ' - Optional parameters allow arguments to be omitted
    '   when calling a procedure.
    ' - If omitted, a default value is used.
    '-------------------------------------------------------

    DisplayMessage()

    DisplayMessage("Hello, LibreOffice!")

End Sub

Sub DisplayMessage(Optional message As String = "Welcome!")

    MsgBox message

End Sub
