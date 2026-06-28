REM  *****  BASIC  *****

Sub FileFunctionsExample()

    '-------------------------------------------------------
    ' File Functions
    '
    ' - File functions are used to work with files and
    '   directories.
    ' - They allow you to create, open, read, write,
    '   locate, and manage files.
    '
    ' Common File Functions:
    '   • FreeFile
    '   • EOF
    '   • FileLen
    '   • FileDateTime
    '   • Dir
    '   • Kill
    '   • MkDir
    '   • RmDir
    '   • ChDir
    '   • CurDir
    '-------------------------------------------------------

    Dim filePath As String
    Dim fileNumber As Integer

    filePath = "C:\Temp\example.txt"

    '----------------------------------------
    ' FreeFile
    '----------------------------------------
    fileNumber = FreeFile
    MsgBox "Available File Number: " & fileNumber

    '----------------------------------------
    ' Dir
    '----------------------------------------
    If Dir(filePath) <> "" Then
        MsgBox "File exists."
    Else
        MsgBox "File not found."
    End If

    '----------------------------------------
    ' FileLen
    '----------------------------------------
    If Dir(filePath) <> "" Then
        MsgBox "File Size: " & FileLen(filePath) & " bytes"
    End If

    '----------------------------------------
    ' FileDateTime
    '----------------------------------------
    If Dir(filePath) <> "" Then
        MsgBox "Last Modified: " & FileDateTime(filePath)
    End If

    '----------------------------------------
    ' CurDir
    '----------------------------------------
    MsgBox "Current Directory: " & CurDir()

    '----------------------------------------
    ' Note
    '----------------------------------------
    '
    ' MkDir   - Creates a folder.
    ' RmDir   - Removes an empty folder.
    ' Kill    - Deletes a file.
    ' ChDir   - Changes the current directory.
    ' EOF()   - Determines whether the end of an
    '           opened file has been reached.
    '
    ' These functions are typically demonstrated in
    ' dedicated File I/O examples.
    '----------------------------------------

End Sub
