REM  *****  BASIC  *****

Sub LabelExample()

    '-------------------------------------------------------
    ' Label Control
    '
    ' - A Label displays text to the user.
    ' - It is used to provide information or instructions.
    ' - Users cannot edit the text in a Label.
    '
    ' Common Uses:
    '   • Display titles
    '   • Show instructions
    '   • Display messages
    '-------------------------------------------------------

    Dim dialog As Object
    Dim label As Object

    ' Load the Standard dialog library
    DialogLibraries.LoadLibrary("Standard")

    ' Create the dialog
    dialog = CreateUnoDialog(DialogLibraries.Standard.Dialog1)

    ' Get the Label control
    label = dialog.getControl("lblMessage")

    ' Change the label text
    label.Model.Label = "Welcome to LibreOffice Basic!"

    ' Show the dialog
    dialog.Execute()

    ' Release the dialog
    dialog.Dispose()

End Sub
