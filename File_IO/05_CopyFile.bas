REM  *****  BASIC  *****

Sub CopyFileExample()

    '-------------------------------------------------------
    ' Copy File
    '
    ' - Copies a file from one location to another.
    ' - The original file remains unchanged.
    ' - Uses the SimpleFileAccess UNO service.
    ' - File paths must be converted to LibreOffice URLs.
    '
    ' Common Uses:
    '   • Create backup files
    '   • Copy reports
    '   • Duplicate templates
    '   • Archive documents
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
    destinationFile = ConvertToURL("C:\Temp\example_copy.txt")

    ' Copy the file
    fileAccess.copy(sourceFile, destinationFile)

    MsgBox "File copied successfully."

End Sub
