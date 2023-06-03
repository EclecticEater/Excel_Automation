' Courtesy of ChatGPT

Sub AddCourseName()
    Dim lastRow As Long
    Dim i As Long
   
    lastRow = Cells(Rows.Count, "D").End(xlUp).Row ' Find the last row in column D
   
    ' Iterate through each row starting from row 2
    For i = 2 To lastRow
        Dim currentStudentID As String
        currentStudentID = Cells(i, "F").Value ' Get the student ID for the current row
               
        Dim periodValue As String
        periodValue = Trim(CStr(Cells(i, "D").Value)) ' Get the string representation of the period value and trim any leading/trailing spaces
       
        ' Check if the current row represents period 1 for a student
        If periodValue = "1" Then
            Dim studentCourses As String
            studentCourses = GetStudentCourses(currentStudentID)
            Cells(i, "N").Value = studentCourses ' Place the course names in the new column
        End If

    Next i
End Sub

Function GetStudentCourses(studentID As String) As String
    Dim lastRow As Long
    Dim i As Long
    Dim courses As String
   
    lastRow = Cells(Rows.Count, "F").End(xlUp).Row ' Find the last row in column F
   
    ' Iterate through each row starting from row 2
    For i = 2 To lastRow
        Dim currentStudentID As String
        currentStudentID = Cells(i, "F").Value ' Get the student ID for the current row
       
        ' Check if the current row matches the provided student ID
        If currentStudentID = studentID Then
            Dim classValue As String
            classValue = Cells(i, "K").Value ' Get the value from the Class column
           
            ' Check if the class value contains "ALG", "Math", or "GEOM"
            If InStr(1, classValue, "ALG", vbTextCompare) > 0 Or _
               InStr(1, classValue, "Math", vbTextCompare) > 0 Or _
               InStr(1, classValue, "GEOM", vbTextCompare) > 0 Then
                ' Append the class value to the courses string
                If Len(courses) > 0 Then
                    courses = courses & ", " & classValue
                Else
                    courses = classValue
                End If
            End If
        End If
    Next i
   
    GetStudentCourses = courses ' Return the courses string
End Function

