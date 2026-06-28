REM  *****  BASIC  *****

Sub SheetsExample()

    '-------------------------------------------------------
    ' Sheets Collection
    '
    ' - Contains all worksheets in the spreadsheet.
    ' - Access sheets by index or by name.
    '-------------------------------------------------------

    Dim sheets As Object

    sheets = ThisComponent.Sheets

    MsgBox "Number of Sheets: " & sheets.Count

End Sub
