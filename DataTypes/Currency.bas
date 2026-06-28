Sub CurrencyExample()

    '-------------------------------------------------------
    ' Currency Data Type
    '
    ' - Stores monetary (currency) values.
    ' - Size: 8 bytes
    ' - Provides high precision for financial calculations.
    ' - Helps avoid many rounding errors that can occur
    '   when using floating-point data types.
    ' - Used for prices, salaries, invoices, taxes,
    '   and other monetary values.
    '-------------------------------------------------------

    ' Declare a Currency variable
    Dim salary As Currency

    ' Assign a monetary value
    salary = 25000.75

    ' Display the value in a message box
    MsgBox "Salary: " & salary

End Sub
