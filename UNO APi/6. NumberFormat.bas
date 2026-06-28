REM  *****  BASIC  *****

Sub NumberFormatsExample()

    '-------------------------------------------------------
    ' Number Formats
    '
    ' - Number formatting controls how numeric values are
    '   displayed in a cell.
    ' - It changes only the appearance of the value and
    '   does not modify the actual value stored.
    '
    ' Common Number Formats:
    '   • Number
    '   • Currency
    '   • Date
    '   • Time
    '   • Percentage
    '   • Scientific
    '-------------------------------------------------------

    Dim sheet As Object
    Dim cell As Object
    Dim formats As Object
    Dim locale As New com.sun.star.lang.Locale
    Dim formatKey As Long

    ' Get the first worksheet
    sheet = ThisComponent.Sheets(0)

    ' Get cell A1
    cell = sheet.getCellByPosition(0, 0)

    ' Set a numeric value
    cell.Value = 12345.678

    ' Get the document's number formats
    formats = ThisComponent.NumberFormats

    ' Use the current locale
    locale.Language = "en"
    locale.Country = "US"

    ' Create or retrieve a number format
    formatKey = formats.queryKey("#,##0.00", locale, True)

    If formatKey = -1 Then
        formatKey = formats.addNew("#,##0.00", locale)
    End If

    ' Apply the number format
    cell.NumberFormat = formatKey

    MsgBox "Number format applied successfully."

End Sub
