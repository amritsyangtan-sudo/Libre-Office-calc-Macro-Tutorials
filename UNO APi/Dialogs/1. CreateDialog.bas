REM  *****  BASIC  *****

Sub CreateDialogExample()

    '-------------------------------------------------------
    ' Create Dialog
    '
    ' - A dialog provides a graphical user interface (GUI)
    '   for interacting with the user.
    ' - Dialogs are created using the Dialog Editor in
    '   LibreOffice.
    ' - Once created, they can be loaded and displayed
    '   using a macro.
    '
    ' Steps
    You must first create the dialog:
'Tools → Macros → Organize Dialogs
'Select My Macros → Standard
'Click New
'Name it Dialog1
'Save the dialog
    '-------------------------------------------------------

    Dim dialog As Object

    ' Load the Standard dialog library
    DialogLibraries.LoadLibrary("Standard")

    ' Create the dialog named "Dialog1"
    dialog = CreateUnoDialog(DialogLibraries.Standard.Dialog1)

    ' Display the dialog
    dialog.Execute()

    ' Release the dialog
    dialog.Dispose()

End Sub
