REM  *****  BASIC  *****

Sub FilePickerExample()

    '-------------------------------------------------------
    ' File Picker
    '
    ' - A File Picker allows the user to browse and select
    '   a file from their computer.
    ' - It provides a standard file selection dialog.
    '
    ' Common Uses:
    '   • Open a spreadsheet
    '   • Import a CSV file
    '   • Select a text file
    '   • Choose an image
    '-------------------------------------------------------

    Dim filePicker As Object
    Dim selectedFiles()

    ' Create the File Picker service
    filePicker = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")

    ' Set the dialog title
    filePicker.setTitle("Select a File")

    ' Display the dialog
    If filePicker.execute() = _
        com.sun.star.ui.dialogs.ExecutableDialogResults.OK Then

        ' Get the selected file(s)
        selectedFiles = filePicker.getFiles()

        MsgBox "Selected File:" & Chr(10) & selectedFiles(0)

    Else
        MsgBox "No file selected."
    End If

End Sub
