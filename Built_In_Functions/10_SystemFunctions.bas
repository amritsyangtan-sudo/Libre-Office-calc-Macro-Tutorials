REM  *****  BASIC  *****

Sub SystemFunctionsExample()

    '-------------------------------------------------------
    ' System Functions
    '
    ' - System functions provide access to LibreOffice
    '   services and system-related information.
    ' - They are commonly used when working with the
    '   UNO API.
    '
    ' Common System Functions:
    '   • CreateUnoService
    '   • CreateUnoStruct
    '   • ConvertToURL
    '   • ConvertFromURL
    '-------------------------------------------------------

    Dim fileAccess As Object
    Dim cellAddress As New com.sun.star.table.CellAddress
    Dim systemPath As String
    Dim fileURL As String

    '----------------------------------------
    ' CreateUnoService()
    '----------------------------------------

    fileAccess = CreateUnoService( _
        "com.sun.star.ucb.SimpleFileAccess")

    MsgBox "UNO Service created successfully."

    '----------------------------------------
    ' CreateUnoStruct()
    '----------------------------------------

    cellAddress.Sheet = 0
    cellAddress.Column = 0
    cellAddress.Row = 0

    MsgBox "UNO Struct created successfully."

    '----------------------------------------
    ' ConvertToURL()
    '----------------------------------------

    systemPath = "C:\Temp\example.txt"

    fileURL = ConvertToURL(systemPath)

    MsgBox "File URL:" & Chr(10) & fileURL

    '----------------------------------------
    ' ConvertFromURL()
    '----------------------------------------

    MsgBox "System Path:" & Chr(10) & _
           ConvertFromURL(fileURL)

End Sub
