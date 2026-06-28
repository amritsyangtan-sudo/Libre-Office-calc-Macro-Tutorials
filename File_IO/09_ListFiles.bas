REM  *****  BASIC  *****

Sub ListFilesExample()

    '-------------------------------------------------------
    ' List Files
    '
    ' - Lists all files and folders in a directory.
    ' - Uses the SimpleFileAccess UNO service.
    ' - Folder paths must be converted to LibreOffice URLs.
    ' - getFolderContents() returns an array of file URLs.
    '
    ' Common Uses:
    '   • Process multiple files
    '   • Import CSV files
    '   • Display directory contents
    '   • Batch file operations
    '-------------------------------------------------------

    Dim fileAccess As Object
    Dim folderPath As String
    Dim fileList()
    Dim i As Integer

    ' Create the SimpleFileAccess service
    fileAccess = CreateUnoService( _
        "com.sun.star.ucb.SimpleFileAccess")

    ' Specify the folder to list
    folderPath = ConvertToURL("C:\Temp")

    ' Check whether the folder exists
    If Not fileAccess.exists(folderPath) Then
        MsgBox "Folder not found."
        Exit Sub
    End If

    ' Get all files and folders
    fileList = fileAccess.getFolderContents(folderPath, False)

    ' Display each file or folder
    For i = LBound(fileList) To UBound(fileList)
        MsgBox ConvertFromURL(fileList(i))
    Next i

End Sub
