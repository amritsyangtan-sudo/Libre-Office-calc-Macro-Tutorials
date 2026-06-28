REM  *****  BASIC  *****

Sub ReadTextFileExample()

    '-------------------------------------------------------
    ' Read Text File
    '
    ' - Reads text from a file one line at a time.
    ' - EOF() is used to determine when the end of the
    '   file has been reached.
    ' - Line Input reads an entire line of text.
    '
    ' Common Uses:
    '   • Read configuration files
    '   • Import CSV files
    '   • Process log files
    '   • Read reports
    '-------------------------------------------------------

    Dim filePath As String
    Dim fileNumber As Integer
    Dim lineText As String

    ' Specify the file path
    filePath = "C:\Temp\example.txt"

    ' Get the next available file number
    fileNumber = FreeFile

    ' Open the file for reading
    Open filePath For Input As #fileNumber

    ' Read until the end of the file
    Do While Not EOF(fileNumber)

        ' Read one line
        Line Input #fileNumber, lineText

        ' Display the line
        MsgBox lineText

    Loop

    ' Close the file
    Close #fileNumber

End Sub
