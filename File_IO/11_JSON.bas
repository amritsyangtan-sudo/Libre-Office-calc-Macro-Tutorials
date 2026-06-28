REM  *****  BASIC  *****

Sub JSONExample()

    '-------------------------------------------------------
    ' JSON (JavaScript Object Notation)
    '
    ' - JSON is a lightweight data interchange format.
    ' - It is commonly used by web APIs and applications.
    ' - LibreOffice Basic does not provide built-in JSON
    '   parsing or serialization.
    ' - A JSON parser library is required to work with
    '   JSON objects and arrays.
    '
    ' Common Uses:
    '   • Read API responses
    '   • Import application settings
    '   • Exchange data between applications
    '   • Store structured data
    '-------------------------------------------------------

    Dim filePath As String
    Dim fileNumber As Integer
    Dim jsonText As String
    Dim lineText As String

    ' Specify the JSON file
    filePath = "C:\Temp\data.json"

    ' Get the next available file number
    fileNumber = FreeFile

    ' Open the JSON file
    Open filePath For Input As #fileNumber

    ' Read the entire file
    Do While Not EOF(fileNumber)

        Line Input #fileNumber, lineText
        jsonText = jsonText & lineText

    Loop

    ' Close the file
    Close #fileNumber

    ' Display the JSON text
    MsgBox jsonText

End Sub
