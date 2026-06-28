REM  *****  BASIC  *****

Sub XMLExample()

    '-------------------------------------------------------
    ' XML (Extensible Markup Language)
    '
    ' - XML is a markup language used to store and
    '   exchange structured data.
    ' - Data is organized using elements (tags).
    ' - LibreOffice can work with XML using the UNO API.
    '
    ' Common Uses:
    '   • Read configuration files
    '   • Import and export structured data
    '   • Exchange data between applications
    '   • Process XML documents
    '-------------------------------------------------------

    Dim filePath As String
    Dim fileNumber As Integer
    Dim xmlText As String
    Dim lineText As String

    ' Specify the XML file
    filePath = "C:\Temp\data.xml"

    ' Get the next available file number
    fileNumber = FreeFile

    ' Open the XML file
    Open filePath For Input As #fileNumber

    ' Read the XML file
    Do While Not EOF(fileNumber)

        Line Input #fileNumber, lineText
        xmlText = xmlText & lineText & Chr(10)

    Loop

    ' Close the file
    Close #fileNumber

    ' Display the XML content
    MsgBox xmlText

End Sub
