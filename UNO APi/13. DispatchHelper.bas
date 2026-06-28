REM  *****  BASIC  *****

Sub DispatchHelperExample()

    '-------------------------------------------------------
    ' Dispatch Helper
    '
    ' - DispatchHelper executes LibreOffice commands
    '   programmatically.
    ' - It is similar to clicking a menu item or toolbar
    '   button in the LibreOffice user interface.
    ' - Used when a command is available in the UI but
    '   does not have a simple UNO API.
    '
    ' Common Uses:
    '   • Copy
    '   • Paste
    '   • Save Document
    '   • Print
    '   • Undo / Redo
    '   • Sort Data
    '-------------------------------------------------------

    Dim dispatcher As Object
    Dim document As Object
    Dim frame As Object

    ' Get the current document
    document = ThisComponent

    ' Get the document frame
    frame = document.CurrentController.Frame

    ' Create the DispatchHelper service
    dispatcher = CreateUnoService( _
        "com.sun.star.frame.DispatchHelper")

    ' Execute the Save command
    dispatcher.executeDispatch( _
        frame, _
        ".uno:Save", _
        "", _
        0, _
        Array())

    MsgBox "Document saved successfully."

End Sub
