REM  *****  BASIC  *****

Sub EventsExample()

    '-------------------------------------------------------
    ' Events
    '
    ' - An event is an action that occurs while using
    '   LibreOffice.
    ' - When an event occurs, a macro can be executed
    '   automatically.
    '
    ' Common Events:
    '   • Document Open
    '   • Document Save
    '   • Document Close
    '   • Button Click
    '   • Text Changed
    '   • Focus Gained
    '   • Focus Lost
    '-------------------------------------------------------

    MsgBox _
        "This macro demonstrates the concept of events." & Chr(10) & _
        "Assign this macro to a document or dialog event."

End Sub

'How to assign an event
'Document Events
'Open Tools → Customize.
'Select the Events tab.
'Choose an event (for example, Open Document).
'Click Assign Macro.
'Select your macro.
'Click OK.
'Dialog Control Events
'Open Tools → Macros → Organize Dialogs.
'Open your dialog.
'Select a control (such as a button).
'Open the Properties window.
'Go to the Events tab.
'Assign a macro to an event such as Execute action.
