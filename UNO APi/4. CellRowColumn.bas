REM  *****  BASIC  *****

Sub CellRangeRowColumnExample()

    '-------------------------------------------------------
    ' Cells, Ranges, Rows, and Columns
    '
    ' Cell
    ' - Represents a single cell in a worksheet.
    ' - Used to read or write data.
    '
    ' Range
    ' - Represents multiple cells.
    ' - Used to work with a group of cells.
    '
    ' Rows
    ' - Represents all rows in a worksheet.
    ' - Used to change row height, hide rows, etc.
    '
    ' Columns
    ' - Represents all columns in a worksheet.
    ' - Used to change column width, hide columns, etc.
    '-------------------------------------------------------

    Dim sheet As Object
    Dim cell As Object
    Dim range As Object
    Dim rows As Object
    Dim columns As Object

    ' Get the first worksheet
    sheet = ThisComponent.Sheets(0)

    '-------------------------
    ' Cell
    '-------------------------
    cell = sheet.getCellByPosition(0, 0)
    cell.String = "Hello"

    '-------------------------
    ' Range
    '-------------------------
    range = sheet.getCellRangeByName("A2:C4")
    range.String = "LibreOffice"

    '-------------------------
    ' Rows
    '-------------------------
    rows = sheet.Rows
    rows.getByIndex(0).Height = 800

    '-------------------------
    ' Columns
    '-------------------------
    columns = sheet.Columns
    columns.getByIndex(0).Width = 3000

    MsgBox "Cells, Ranges, Rows, and Columns updated successfully."

End Sub
