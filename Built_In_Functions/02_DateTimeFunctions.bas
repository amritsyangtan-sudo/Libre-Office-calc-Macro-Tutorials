REM  *****  BASIC  *****

Sub DateTimeFunctionsExample()

    '-------------------------------------------------------
    ' Date and Time Functions
    '
    ' - Date and time functions are used to work with
    '   dates and times.
    ' - They allow you to retrieve the current date and
    '   time, extract individual components, and perform
    '   date calculations.
    '
    ' Common Date & Time Functions:
    '   • Date
    '   • Time
    '   • Now
    '   • Year
    '   • Month
    '   • Day
    '   • Hour
    '   • Minute
    '   • Second
    '   • WeekDay
    '   • DateAdd
    '   • DateDiff
    '-------------------------------------------------------

    Dim currentDate As Date
    Dim currentTime As Date
    Dim currentDateTime As Date

    currentDate = Date
    currentTime = Time
    currentDateTime = Now

    MsgBox "Current Date: " & currentDate
    MsgBox "Current Time: " & currentTime
    MsgBox "Current Date & Time: " & currentDateTime

    MsgBox "Year: " & Year(currentDate)
    MsgBox "Month: " & Month(currentDate)
    MsgBox "Day: " & Day(currentDate)

    MsgBox "Hour: " & Hour(currentDateTime)
    MsgBox "Minute: " & Minute(currentDateTime)
    MsgBox "Second: " & Second(currentDateTime)

    MsgBox "Weekday: " & WeekDay(currentDate)

    MsgBox "30 Days Later: " & DateAdd("d", 30, currentDate)

    MsgBox "Days Between: " & _
           DateDiff("d", currentDate, DateAdd("d", 30, currentDate))

End Sub
