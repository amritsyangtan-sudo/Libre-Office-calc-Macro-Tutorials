REM ***** BASIC *****

Option Explicit

Sub SelectCaseExample()

    '-------------------------------------------------------
    ' Select Case Statement
    '
    ' - Select Case evaluates a single expression
    '   against multiple possible values.
    ' - It is a cleaner alternative to multiple
    '   If...ElseIf statements.
    '
    ' Common Uses:
    '   • Menu selection
    '   • Grade classification
    '   • Day of week
    '   • Status codes
    '-------------------------------------------------------

    Dim grade As String

    grade = "B"

    Select Case grade

        Case "A"
            MsgBox "Excellent"

        Case "B"
            MsgBox "Very Good"

        Case "C"
            MsgBox "Good"

        Case "D"
            MsgBox "Pass"

        Case "F"
            MsgBox "Fail"

        Case Else
            MsgBox "Invalid Grade"

    End Select

End Sub
