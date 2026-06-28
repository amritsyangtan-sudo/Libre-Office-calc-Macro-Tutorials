REM  *****  BASIC  *****

Sub TextBoxExample()

    '-------------------------------------------------------
    ' Text Box Control
    '
    ' - A Text Box allows the user to enter and edit text.
    ' - It is one of the most commonly used dialog controls.
    '
    ' Common Uses:
    '   • User names
    '   • Email addresses
    '   • Search text
    '   • Comments
    '-------------------------------------------------------

    Dim dialog As Object
    Dim textBox As Object

    ' Load the Standard dialog library
    DialogLibraries.LoadLibrary("Standard")

    ' Create the dialog
    dialog = CreateUnoDialog(DialogLibraries.Standard.Dialog1)

    ' Get the Text Box control
    textBox = dialog.getControl("txtName")

    ' Set the default text
    textBox.Text = "Enter your name"

    ' Display the dialog
    dialog.Execute()

    ' Read the text entered by the user
    MsgBox "You entered: " & textBox.Text

    ' Release the dialog
    dialog.Dispose()

End Sub
