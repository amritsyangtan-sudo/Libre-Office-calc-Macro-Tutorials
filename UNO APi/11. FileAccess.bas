REM  *****  BASIC  *****

Sub FileAccessExample()

    '-------------------------------------------------------
    ' File Access
    '
    ' - File Access allows macros to work with files
    '   and folders on the computer.
    ' - Common operations include checking whether a file
    '   exists, creating folders, copying files,
    '   moving files, and deleting files.
    '
    ' Common Uses:
    '   • Check if a file exists
    '   • Create a folder
    '   • Copy a file
    '   • Move a file
    '   • Delete a file
    '-------------------------------------------------------

    Dim fileAccess As Object
    Dim fileURL As String

    ' Create the SimpleFileAccess service
    fileAccess = CreateUnoService( _
        "com.sun.star.ucb.SimpleFileAccess")

    ' Example file URL
    fileURL = ConvertToURL("C:\Temp\example.txt")

    ' Check whether the file exists
    If fileAccess.exists(fileURL) Then
        MsgBox "File exists."
    Else
        MsgBox "File does not exist."
    End If

End Sub
