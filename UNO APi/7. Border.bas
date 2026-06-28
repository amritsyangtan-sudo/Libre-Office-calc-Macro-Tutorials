REM  *****  BASIC  *****

Sub BordersExample()

    '-------------------------------------------------------
    ' Cell Borders
    '
    ' - Borders are lines drawn around a cell or range.
    ' - They improve the appearance and readability of data.
    ' - Borders can be applied to:
    '     • Top
    '     • Bottom
    '     • Left
    '     • Right
    ' - Border properties include:
    '     • Color
    '     • Line Width
    '     • Line Style
    '-------------------------------------------------------

    Dim sheet As Object
    Dim cell As Object
    Dim border As New com.sun.star.table.BorderLine2

    ' Get the first worksheet
    sheet = ThisComponent.Sheets(0)

    ' Get cell A1
    cell = sheet.getCellByPosition(0, 0)

    ' Set cell value
    cell.String = "Border Example"

    '-------------------------
    ' Configure Border
    '-------------------------

    ' Black border
    border.Color = RGB(0, 0, 0)

    ' Border width (100 = approximately 1 pt)
    border.LineWidth = 100

    '-------------------------
    ' Apply Border
    '-------------------------

    cell.TopBorder = border
    cell.BottomBorder = border
    cell.LeftBorder = border
    cell.RightBorder = border

    MsgBox "Borders applied successfully."

End Sub
