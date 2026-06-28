REM  *****  BASIC  *****

Sub CSVExample()

    '-------------------------------------------------------
    ' CSV (Comma-Separated Values)
    '
    ' - A CSV file stores tabular data as plain text.
    ' - Each line represents a row.
    ' - Values are separated by commas.
    ' - CSV files can be opened by Calc, Excel,
    '   databases, and many other applications.
    '
    ' Common Uses:
    '   • Import spreadsheet data
    '   • Export reports
    '   • Exchange data between applications
    '   • Data migration
    '-------------------------------------------------------

    Dim filePath As String
    Dim fileNumber As Integer
    Dim lineText As String
    Dim values()

    ' Specify the CSV file
    filePath = "C:\Temp\employees.csv"

    ' Get the next available file number
    fileNumber = FreeFile

    ' Open the CSV file
    Open filePath For Input As #fileNumber

    ' Read each record
    Do While Not EOF(fileNumber)

        ' Read one line
        Line Input #fileNumber, lineText

        ' Split the line into fields
        values = Split(lineText, ",")

        ' Display the fields
        MsgBox _
            "ID: " & values(0) & Chr(10) & _
            "Name: " & values(1) & Chr(10) & _
            "Department: " & values(2)

    Loop

    ' Close the file
    Close #fileNumber

End Sub
