REM ***** BASIC *****

Option Explicit

Sub IfElseExample()

    '-------------------------------------------------------
    ' If...Else Statement
    '
    ' - Executes one block of code if the condition is True.
    ' - Executes another block if the condition is False.
    ' - Used when there are two possible outcomes.
    '
    ' Common Uses:
    '   • Pass or Fail
    '   • Login validation
    '   • Age verification
    '   • File existence checks
    '-------------------------------------------------------

    Dim marks As Integer

    marks = 45

    If marks >= 50 Then

        MsgBox "Congratulations! You Passed."

    Else

        MsgBox "Sorry! You Failed."

    End If

End Sub
