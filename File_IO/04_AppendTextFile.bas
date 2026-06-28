REM  *****  BASIC  *****

Sub AppendTextFileExample()

    '-------------------------------------------------------
    ' Append Text File
    '
    ' - Appends (adds) text to the end of an existing file.
    ' - Existing file contents are preserved.
    ' - If the file does not exist, it is created.
    ' - Print # writes text followed by a new line.
    '
    ' Common Uses:
    '   • Write application logs
    '   • Add records to a file
    '   • Store user activity
    '   • Append report data
    '-------------------------------------------------------

    Dim filePath As String
    Dim fileNumber As Integer

    ' Specify the file path
    filePath = "C:\Temp\example.txt"

    ' Get the next available file number
    fileNumber = FreeFile

    ' Open the file for appending
    Open filePath For Append As #fileNumber

    ' Append text to the end of the file
    Print #fileNumber, "This line was appended."
    Print #fileNumber, "Another appended line."

    ' Close the file
    Close #fileNumber

    MsgBox "Text appended successfully."

End Sub
