REM  *****  BASIC  *****

Sub MathFunctionsExample()

    '-------------------------------------------------------
    ' Math Functions
    '
    ' - Math functions perform mathematical calculations.
    ' - They are commonly used in engineering,
    '   finance, statistics, and general programming.
    '
    ' Common Math Functions:
    '   • Abs
    '   • Sgn
    '   • Int
    '   • Fix
    '   • Round
    '   • Sqr
    '   • Sin
    '   • Cos
    '   • Tan
    '   • Atn
    '   • Exp
    '   • Log
    '   • Rnd
    '   • Randomize
    '-------------------------------------------------------

    Dim number As Double

    number = -15.75

    MsgBox "Absolute Value: " & Abs(number)

    MsgBox "Sign: " & Sgn(number)

    MsgBox "Integer Portion (Int): " & Int(number)

    MsgBox "Integer Portion (Fix): " & Fix(number)

    MsgBox "Rounded Value: " & Round(number)

    MsgBox "Square Root of 81: " & Sqr(81)

    MsgBox "Sin(0): " & Sin(0)

    MsgBox "Cos(0): " & Cos(0)

    MsgBox "Tan(0): " & Tan(0)

    MsgBox "Atn(1): " & Atn(1)

    MsgBox "Exp(1): " & Exp(1)

    MsgBox "Log(10): " & Log(10)

    Randomize

    MsgBox "Random Number: " & Rnd()

End Sub
