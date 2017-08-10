Sub grade_calculator()

' If the value of B2 is greater than or equal to 90..
If Cells(2, 2).Value >= 90 Then
    ' The value of cell C2 is set to "Pass"
    Cells(2, 3).Value = "Pass"
    ' The fill color of cell C2 is set to green
    Cells(2, 3).Interior.ColorIndex = 4
    ' The value of cell D2 is set to "A"
    Cells(2, 4).Value = "A"

' If the value of B2 is greater than or equal to 80..
ElseIf Cells(2, 2).Value >= 80 Then
    ' The value of cell C2 is set to "Pass"
    Cells(2, 3).Value = "Pass"
    ' The fill color of cell C2 is set to green
    Cells(2, 3).Interior.ColorIndex = 4
    ' The value of cell D2 is set to "B"
    Cells(2, 4).Value = "B"

' If the value of B2 is greater than or equal to 70..
ElseIf Cells(2, 2).Value >= 70 Then
    ' The value of cell C2 is set to "Warning"
    Cells(2, 3).Value = "Warning"
    ' The fill color of cell C2 is set to yellow
    Cells(2, 3).Interior.ColorIndex = 6
    ' The value of cell D2 is set to "C"
    Cells(2, 4).Value = "C"

' If all of the previous statements were returned as FALSE...
Else
    ' The value of cell C2 is set to "Fail"
    Cells(2, 3).Value = "Fail"
    ' The fill color of C2 is set to red
    Cells(2, 3).Interior.ColorIndex = 3
    ' The value of cell D2 is set to "F"
    Cells(2, 4).Value = "F"

End If

End Sub
