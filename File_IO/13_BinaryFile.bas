REM  *****  BASIC  *****

Sub BinaryFileExample()

    '-------------------------------------------------------
    ' Binary File
    '
    ' - A binary file stores data as raw bytes instead of
    '   plain text.
    ' - Binary files are not intended to be read directly
    '   by humans.
    ' - Used for images, audio, videos, executable files,
    '   and other binary formats.
    '
    ' Common Uses:
    '   • Read image files
    '   • Read executable files
    '   • Process binary data
    '   • Store custom binary formats
    '
    ' Note:
    ' - Binary files are opened using the Binary mode.
    ' - Binary data is read and written using the Get and
    '   Put statements.
    '-------------------------------------------------------

    Dim filePath As String
    Dim fileNumber As Integer
    Dim firstByte As Byte

    ' Specify the binary file
    filePath = "C:\Temp\sample.bin"

    ' Get the next available file number
    fileNumber = FreeFile

    ' Open the file in Binary mode
    Open filePath For Binary As #fileNumber

    ' Read the first byte
    Get #fileNumber, 1, firstByte

    ' Close the file
    Close #fileNumber

    ' Display the byte value
    MsgBox "First Byte: " & firstByte

End Sub


'| Extension | Description          |
'| --------- | -------------------- |
'| `.bin`    | Generic binary file  |
'| `.exe`    | Executable program   |
'| `.dll`    | Dynamic Link Library |
'| `.jpg`    | JPEG image           |
'| `.png`    | PNG image            |
'| `.pdf`    | PDF document         |
'| `.zip`    | Compressed archive   |
