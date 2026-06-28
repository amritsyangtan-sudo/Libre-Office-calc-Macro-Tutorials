REM  *****  BASIC  *****

Sub CellFormattingExample()

    '-------------------------------------------------------
    ' Cell Formatting
    '
    ' - Cell formatting changes the appearance of a cell.
    ' - It does not change the actual value stored in the cell.
    ' - Common formatting includes:
    '     • Font Name
    '     • Font Size
    '     • Bold
    '     • Italic
    '     • Font Color
    '     • Cell Background Color
    '     • Horizontal Alignment
    '     • Vertical Alignment
    '-------------------------------------------------------

    Dim sheet As Object
    Dim cell As Object

    ' Get the first worksheet
    sheet = ThisComponent.Sheets(0)

    ' Get cell A1
    cell = sheet.getCellByPosition(0, 0)

    ' Set cell value
    cell.String = "LibreOffice Calc"

    '-------------------------
    ' Font Formatting
    '-------------------------

    ' Font Name
    cell.CharFontName = "Liberation Sans"

    ' Font Size
    cell.CharHeight = 14

    ' Bold
    cell.CharWeight = com.sun.star.awt.FontWeight.BOLD

    ' Italic
    cell.CharPosture = com.sun.star.awt.FontSlant.ITALIC

    ' Font Color (Blue)
    cell.CharColor = RGB(0, 0, 255)

    '-------------------------
    ' Cell Formatting
    '-------------------------

    ' Background Color (Yellow)
    cell.CellBackColor = RGB(255, 255, 0)

    ' Horizontal Alignment (Center)
    cell.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER

    ' Vertical Alignment (Center)
    cell.VertJustify = com.sun.star.table.CellVertJustify.CENTER

    MsgBox "Cell formatting applied successfully."

End Sub
