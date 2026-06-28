REM  *****  BASIC  *****

Sub MiscellaneousFunctionsExample()

    '-------------------------------------------------------
    ' Miscellaneous Functions
    '
    ' - Miscellaneous functions provide various useful
    '   features that do not belong to a specific category.
    ' - They are commonly used for colors, conditional
    '   expressions, timing, and value selection.
    '
    ' Common Miscellaneous Functions:
    '   • RGB
    '   • Choose
    '   • Switch
    '   • IIf
    '   • Timer
    '-------------------------------------------------------

    Dim startTime As Double
    Dim color As Long
    Dim grade As String
    Dim message As String
    Dim dayName As String

    '----------------------------------------
    ' RGB()
    '----------------------------------------
    color = RGB(255, 0, 0)

    MsgBox "RGB Color Value: " & color

    '----------------------------------------
    ' Choose()
    '----------------------------------------
    dayName = Choose(3, _
        "Sunday", _
        "Monday", _
        "Tuesday", _
        "Wednesday", _
        "Thursday", _
        "Friday", _
        "Saturday")

    MsgBox "Choose: " & dayName

    '----------------------------------------
    ' IIf()
    '----------------------------------------
    grade = IIf(85 >= 50, "Pass", "Fail")

    MsgBox "IIf: " & grade

    '----------------------------------------
    ' Switch()
    '----------------------------------------
    message = Switch( _
        90 >= 80, "Excellent", _
        90 >= 60, "Good", _
        True, "Needs Improvement")

    MsgBox "Switch: " & message

    '----------------------------------------
    ' Timer()
    '----------------------------------------
    startTime = Timer

    Wait 1000

    MsgBox "Elapsed Time: " & _
           Format(Timer - startTime, "0.00") & _
           " seconds"

End Sub
