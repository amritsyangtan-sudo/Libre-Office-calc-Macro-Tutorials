Sub ObjectExample()

    '-------------------------------------------------------
    ' Object Data Type
    '
    ' - Stores a reference to an object.
    ' - Objects represent LibreOffice components such as
    '   documents, sheets, cells, ranges, charts, and more.
    ' - Used to interact with and manipulate LibreOffice
    '   applications through the UNO API.
    ' - Object is one of the most commonly used data types
    '   in LibreOffice Calc macros.
    '-------------------------------------------------------

    ' Declare an Object variable
    Dim sheet As Object

    ' Get the first worksheet in the current spreadsheet
    sheet = ThisComponent.Sheets(0)

    ' Display the worksheet name
    MsgBox "Sheet Name: " & sheet.Name

End Sub
