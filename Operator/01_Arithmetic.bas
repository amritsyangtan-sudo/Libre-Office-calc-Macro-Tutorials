REM ***** BASIC *****

Option Explicit

Sub ArithmeticExample()

    '-------------------------------------------------------
    ' Arithmetic Operators
    '
    ' - Arithmetic operators perform mathematical
    '   calculations on numeric values.
    '
    ' Common Operators:
    '   +   Addition
    '   -   Subtraction
    '   *   Multiplication
    '   /   Division
    '   \   Integer Division
    '   Mod Modulus (Remainder)
    '   ^   Exponentiation
    '-------------------------------------------------------

    Dim number1 As Double
    Dim number2 As Double

    number1 = 20
    number2 = 6

    MsgBox "Number 1 : " & number1 & Chr(10) & _
           "Number 2 : " & number2

    '----------------------------------------
    ' Addition
    '----------------------------------------
    MsgBox "Addition: " & (number1 + number2)

    '----------------------------------------
    ' Subtraction
    '----------------------------------------
    MsgBox "Subtraction: " & (number1 - number2)

    '----------------------------------------
    ' Multiplication
    '----------------------------------------
    MsgBox "Multiplication: " & (number1 * number2)

    '----------------------------------------
    ' Division
    '----------------------------------------
    MsgBox "Division: " & (number1 / number2)

    '----------------------------------------
    ' Integer Division
    '----------------------------------------
    MsgBox "Integer Division: " & (number1 \ number2)

    '----------------------------------------
    ' Modulus
    '----------------------------------------
    MsgBox "Modulus: " & (number1 Mod number2)

    '----------------------------------------
    ' Exponentiation
    '----------------------------------------
    MsgBox "Exponentiation: " & (number1 ^ 2)

End Sub
