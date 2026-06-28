REM  *****  BASIC  *****

Sub CreateFolderExample()

    '-------------------------------------------------------
    ' Create Folder
    '
    ' - Creates a new folder (directory).
    ' - Uses the SimpleFileAccess UNO service.
    ' - Folder paths must be converted to LibreOffice URLs.
    ' - It is recommended to check whether the folder
    '   already exists before creating it.
    '
    ' Common Uses:
    '   • Create backup folders
    '   • Create report directories
    '   • Organize exported files
    '   • Prepare application folders
    '-------------------------------------------------------

    Dim fileAccess As Object
    Dim folderPath As String

    ' Create the SimpleFileAccess service
    fileAccess = CreateUnoService( _
        "com.sun.star.ucb.SimpleFileAccess")

    ' Specify the folder path
    folderPath = ConvertToURL("C:\Temp\Reports")

    ' Check whether the folder already exists
    If Not fileAccess.exists(folderPath) Then

        ' Create the folder
        fileAccess.createFolder(folderPath)

        MsgBox "Folder created successfully."

    Else

        MsgBox "Folder already exists."

    End If

End Sub
