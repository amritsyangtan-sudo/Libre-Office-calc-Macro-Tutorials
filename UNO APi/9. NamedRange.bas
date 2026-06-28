REM  *****  BASIC  *****

Sub NamedRangesExample()

    '-------------------------------------------------------
    ' Named Ranges
    '
    ' - A named range is a user-defined name assigned to
    '   a cell or group of cells.
    ' - Makes formulas easier to read and maintain.
    ' - Allows you to reference cells by name instead of
    '   cell addresses.
    '
    ' Example:
    '   SalesData  -> Sheet1.A1:A10
    '-------------------------------------------------------

    Dim doc As Object
    Dim namedRanges As Object
    Dim sheet As Object
    Dim address As New com.sun.star.table.CellAddress

    doc = ThisComponent
    sheet = doc.Sheets.getByIndex(0)

    namedRanges = doc.NamedRanges

    address.Sheet = 0
    address.Column = 0
    address.Row = 0

    If Not namedRanges.hasByName("SalesData") Then
        namedRanges.addNewByName( _
            "SalesData", _
            "$Sheet1.$A$1:$A$10", _
            address, _
            0)
    End If

    MsgBox "Named range 'SalesData' created successfully."

End Sub
