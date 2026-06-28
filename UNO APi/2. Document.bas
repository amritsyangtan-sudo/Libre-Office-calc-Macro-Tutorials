REM  *****  BASIC  *****

Sub DocumentExample()

    '-------------------------------------------------------
    ' Document
    '
    ' - Represents the current spreadsheet document.
    ' - Provides access to sheets, styles, properties,
    '   and document information.
    '-------------------------------------------------------

    Dim document As Object

    document = ThisComponent

    MsgBox document.getURL()

End Sub
