Sub grade_calculator()

' Check for A worthy grade
If Cells(2, 2).Value >= 90 Then
    Cells(2, 3).Value = "Pass"
    Cells(2, 3).Interior.ColorIndex = 4
    Cells(2, 4).Value = "A"

' Check for B worthy grade
ElseIf Cells(2, 2).Value >= 80 Then
    Cells(2, 3).Value = "Pass"
    Cells(2, 3).Interior.ColorIndex = 4
    Cells(2, 4).Value = "B"

' Check for C worthy grade
ElseIf Cells(2, 2).Value >= 70 Then
    Cells(2, 3).Value = "Warning"
    Cells(2, 3).Interior.ColorIndex = 6
    Cells(2, 4).Value = "C"

' Check for failing grade
Else
    Cells(2, 3).Value = "Fail"
    Cells(2, 3).Interior.ColorIndex = 3
    Cells(2, 4).Value = "F"

End If

End Sub
