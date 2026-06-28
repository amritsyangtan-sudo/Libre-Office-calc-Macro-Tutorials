REM ***** BASIC *****

Type Student

    StudentName As String
    RollNo As Integer

    English As Double
    Math As Double
    Science As Double

End Type

REM ***** BASIC *****

Function Total(student As Student) As Double

    Total = student.English + _
            student.Math + _
            student.Science

End Function


Function Average(student As Student) As Double

    Average = Total(student) / 3

End Function


Function Grade(student As Student) As String

    Dim avg As Double

    avg = Average(student)

    If avg >= 90 Then
        Grade = "A+"

    ElseIf avg >= 80 Then
        Grade = "A"

    ElseIf avg >= 70 Then
        Grade = "B"

    ElseIf avg >= 60 Then
        Grade = "C"

    Else
        Grade = "F"
    End If

End Function

REM ***** BASIC *****

Sub PrintReport(student As Student)

    MsgBox _
        "Student Report" & Chr(10) & Chr(10) & _
        "Name : " & student.StudentName & Chr(10) & _
        "Roll : " & student.RollNo & Chr(10) & Chr(10) & _
        "English : " & student.English & Chr(10) & _
        "Math : " & student.Math & Chr(10) & _
        "Science : " & student.Science & Chr(10) & Chr(10) & _
        "Total : " & Total(student) & Chr(10) & _
        "Average : " & Format(Average(student), "0.00") & Chr(10) & _
        "Grade : " & Grade(student)

End Sub

REM ***** BASIC *****

Sub StudentGradeSystem()

    Dim student As Student

    student.StudentName = InputBox("Student Name")

    student.RollNo = Val(InputBox("Roll Number"))

    student.English = Val(InputBox("English Marks"))

    student.Math = Val(InputBox("Math Marks"))

    student.Science = Val(InputBox("Science Marks"))

    PrintReport student

End Sub
