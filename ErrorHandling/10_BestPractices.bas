REM  *****  BASIC  *****

Sub BestPracticesExample()

    '-------------------------------------------------------
    ' Error Handling Best Practices
    '
    ' - Write code that prevents errors whenever possible.
    ' - Handle expected errors gracefully.
    ' - Display meaningful error messages.
    ' - Always clean up resources before exiting.
    '
    ' Best Practices:
    '    Use On Error only when necessary.
    '    Validate input before processing.
    '    Close opened files.
    '    Release UNO objects when appropriate.
    '    Use descriptive error messages.
    '    Avoid using On Error Resume Next for an
    '     entire procedure.
    '-------------------------------------------------------

    On Error GoTo ErrorHandler

    Dim filePath As String
    Dim fileNumber As Integer

    filePath = "C:\Temp\example.txt"

    ' Validate the input
    If filePath = "" Then
        MsgBox "File path cannot be empty."
        Exit Sub
    End If

    ' Open the file
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber

    MsgBox "File opened successfully."

Cleanup:

    ' Close the file if it was opened
    Close #fileNumber

    Exit Sub

ErrorHandler:

    MsgBox "An unexpected error occurred."

    ' Always perform cleanup before exiting
    Resume Cleanup

End Sub
