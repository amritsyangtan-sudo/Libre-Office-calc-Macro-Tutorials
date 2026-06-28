Sub LongExample()

    '-------------------------------------------------------
    ' Long Data Type
    '
    ' - Stores large whole numbers (no decimal values).
    ' - Size: 4 bytes
    ' - Range: -2,147,483,648 to 2,147,483,647
    ' - Used when Integer is not large enough, such as
    '   population, account numbers, row counts, etc.
    '-------------------------------------------------------

    ' Declare a Long variable
    Dim population As Long

    ' Assign a large whole number
    population = 8500000

    ' Display the value in a message box
    MsgBox "Population: " & population

End Sub
