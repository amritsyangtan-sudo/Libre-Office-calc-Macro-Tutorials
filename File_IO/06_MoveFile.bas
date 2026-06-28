REM  *****  BASIC  *****

Sub MoveFileExample()

    '-------------------------------------------------------
    ' Move File
    '
    ' - Moves a file from one location to another.
    ' - Can also be used to rename a file.
    ' - Uses the SimpleFileAccess UNO service.
    ' - File paths must be converted to LibreOffice URLs.
    '
    ' Common Uses:
    '   • Move processed files
    '   • Rename files
    '   • Organize documents
    '   • Archive reports
    '-------------------------------------------------------

    Dim fileAccess As Object
    Dim sourceFile As String
    Dim destinationFile As String

    ' Create the SimpleFileAccess service
    fileAccess = CreateUnoService( _
        "com.sun.star.ucb.SimpleFileAccess")

    ' Source file
    sourceFile = ConvertToURL("C:\Temp\example.txt")

    ' Destination file
    destinationFile = ConvertToURL("C:\Temp\Archive\example.txt")

    ' Move the file
    fileAccess.move(sourceFile, destinationFile)

    MsgBox "File moved successfully."

End Sub
