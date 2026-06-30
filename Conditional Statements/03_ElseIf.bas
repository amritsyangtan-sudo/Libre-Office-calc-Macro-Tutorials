REM ***** BASIC *****

Option Explicit

Sub ElseIfExample()

    '-------------------------------------------------------
    ' If...ElseIf...Else Statement
    '
    ' - Evaluates multiple conditions in sequence.
    ' - The first condition that evaluates to True
    '   is executed.
    ' - If none of the conditions are True, the
    '   Else block is executed.
    '
    ' Common Uses:
    '   • Grade calculation
    '   • Salary classification
    '   • Discount calculation
    '   • Menu selection
    '-------------------------------------------------------

    Dim marks As Integer

    marks = 82

    If marks >= 90 Then

        MsgBox "Grade: A+"

    ElseIf marks >= 80 Then

        MsgBox "Grade: A"

    ElseIf marks >= 70 Then

        MsgBox "Grade: B"

    ElseIf marks >= 60 Then

        MsgBox "Grade: C"

    ElseIf marks >= 50 Then

        MsgBox "Grade: D"

    Else

        MsgBox "Grade: F"

    End If

End Sub
