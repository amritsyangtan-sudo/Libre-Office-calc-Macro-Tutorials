REM  *****  BASIC  *****

Sub WriteTextFileExample()

    '-------------------------------------------------------
    ' Write Text File
    '
    ' - Writes text to a file.
    ' - If the file already exists, it is overwritten.
    ' - If the file does not exist, it is created.
    ' - Print # writes text followed by a new line.
    '
    ' Common Uses:
    '   • Create log files
    '   • Export reports
    '   • Save application data
    '   • Generate text files
    '-------------------------------------------------------

    Dim filePath As String
    Dim fileNumber As Integer

    ' Specify the file path
    filePath = "C:\Temp\example.txt"

    ' Get the next available file number
    fileNumber = FreeFile

    ' Open the file for writing
    Open filePath For Output As #fileNumber

    ' Write text to the file
    Print #fileNumber, "Hello, LibreOffice Basic!"
    Print #fileNumber, "This is the second line."
    Print #fileNumber, "Writing text files is easy."

    ' Close the file
    Close #fileNumber

    MsgBox "Text written successfully."

End Sub
