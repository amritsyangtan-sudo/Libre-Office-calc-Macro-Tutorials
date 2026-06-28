Sub ThisComponentExample()

    '-------------------------------------------------------
    ' ThisComponent
    '
    ' - Refers to the currently active LibreOffice document.
    ' - In Calc, it represents the current spreadsheet.
    ' - It is the starting point for accessing sheets,
    '   cells, ranges, and other document objects.
    '-------------------------------------------------------

    Dim document As Object

    document = ThisComponent

    MsgBox "Document Title: " & document.Title

End Sub
