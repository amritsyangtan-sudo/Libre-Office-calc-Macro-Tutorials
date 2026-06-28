REM  *****  BASIC  *****

Sub OpenFileExample()

    '-------------------------------------------------------
    ' Open File
    '
    ' - Opens an existing text file for reading.
    ' - If the file cannot be opened, an error occurs.
    ' - Always close the file after you finish using it.
    '
    ' Common Uses:
    '   • Read configuration files
    '   • Open log files
    '   • Process text reports
    '   • Load application data
    '-------------------------------------------------------

    Dim filePath As String
    Dim fileNumber As Integer

    ' Specify the file path
    filePath = "C:\Temp\example.txt"

    ' Get the next available file number
    fileNumber = FreeFile

    ' Open the file for reading
    Open filePath For Input As #fileNumber

    MsgBox "File opened successfully."

    ' Close the file
    Close #fileNumber

End Sub
