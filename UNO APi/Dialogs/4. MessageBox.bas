REM  *****  BASIC  *****

Sub MessageBoxExample()

    '-------------------------------------------------------
    ' Message Box (MsgBox)
    '
    ' - Displays a message to the user.
    ' - Used to provide information, warnings,
    '   ask questions, or report errors.
    '
    ' Common Buttons:
    '   • OK
    '   • OK / Cancel
    '   • Yes / No
    '   • Yes / No / Cancel
    '   • Retry / Cancel
    '
    ' Common Icons:
    '   • Information
    '   • Warning
    '   • Error
    '   • Question
    '-------------------------------------------------------

    Dim result As Integer

    result = MsgBox( _
        "Do you want to continue?", _
        MB_YESNO + MB_ICONQUESTION, _
        "Confirmation")

    If result = IDYES Then
        MsgBox "You clicked Yes."
    Else
        MsgBox "You clicked No."
    End If

End Sub
