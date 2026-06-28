REM  *****  BASIC  *****

Sub DeleteFileExample()

    '-------------------------------------------------------
    ' Delete File
    '
    ' - Deletes a file from the file system.
    ' - Uses the SimpleFileAccess UNO service.
    ' - The file path must be converted to a LibreOffice URL.
    ' - Always verify that the file exists before deleting it.
    '
    ' Common Uses:
    '   • Delete temporary files
    '   • Remove old reports
    '   • Clean up log files
    '   • Delete exported files
    '-------------------------------------------------------

    Dim fileAccess As Object
    Dim filePath As String

    ' Create the SimpleFileAccess service
    fileAccess = CreateUnoService( _
        "com.sun.star.ucb.SimpleFileAccess")

    ' Specify the file to delete
    filePath = ConvertToURL("C:\Temp\example.txt")

    ' Check whether the file exists
    If fileAccess.exists(filePath) Then

        ' Delete the file
        fileAccess.kill(filePath)

        MsgBox "File deleted successfully."

    Else

        MsgBox "File not found."

    End If

End Sub
