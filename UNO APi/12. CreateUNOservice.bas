REM  *****  BASIC  *****

Sub CreateUnoServiceExample()

    '-------------------------------------------------------
    ' CreateUnoService
    '
    ' - CreateUnoService() creates an instance of a UNO
    '   service provided by LibreOffice.
    ' - A service provides functionality such as file
    '   access, dialogs, configuration, and more.
    ' - Many LibreOffice APIs begin by creating a service.
    '
    ' Common Services:
    '   • SimpleFileAccess
    '   • FilePicker
    '   • FolderPicker
    '   • Desktop
    '   • PathSettings
    '-------------------------------------------------------

    Dim fileAccess As Object
    Dim desktop As Object
    Dim filePicker As Object

    ' Create the SimpleFileAccess service
    fileAccess = CreateUnoService( _
        "com.sun.star.ucb.SimpleFileAccess")

    ' Create the Desktop service
    desktop = CreateUnoService( _
        "com.sun.star.frame.Desktop")

    ' Create the FilePicker service
    filePicker = CreateUnoService( _
        "com.sun.star.ui.dialogs.FilePicker")

    MsgBox "UNO services created successfully."

End Sub
