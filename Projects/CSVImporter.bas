REM ***** BASIC *****

Function ReadCSV(filePath As String) As Variant

    Dim fileNo As Integer
    Dim line As String
    Dim data()
    Dim row As Integer

    fileNo = FreeFile

    Open filePath For Input As #fileNo

    row = -1

    Do While Not EOF(fileNo)

        Line Input #fileNo, line

        row = row + 1

        ReDim Preserve data(row)

        data(row) = Split(line, ",")

    Loop

    Close #fileNo

    ReadCSV = data

End Function

REM ***** BASIC *****

Sub WriteToSheet(data)

    Dim sheet As Object
    Dim r As Integer
    Dim c As Integer

    sheet = ThisComponent.Sheets(0)

    For r = LBound(data) To UBound(data)

        For c = LBound(data(r)) To UBound(data(r))

            sheet.getCellByPosition(c, r).String = data(r)(c)

        Next c

    Next r

End Sub


REM ***** BASIC *****

Sub CSVImporter()

    Dim csvData

    csvData = ReadCSV("C:\Users\AmritS\Downloads\New folder\data.csv")

    WriteToSheet csvData

    MsgBox "Import completed."

End Sub

