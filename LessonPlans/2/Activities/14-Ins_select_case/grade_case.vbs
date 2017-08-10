Sub select_case()

Dim i As Integer

For i = 2 To 101

Select Case Cells(i, 1).Value
    Case Is >= 90
       Cells(i, 2).Value = "A"
       Cells(i, 2).Interior.ColorIndex = 10
    Case Is >= 80
        Cells(i, 2).Value = "B"
        Cells(i, 2).Borders.ColorIndex = 10
    Case Is >= 70
       Cells(i, 2) = "C"
       Cells(i, 2).Interior.ColorIndex = 6
    Case Is >= 60
        Cells(i,2) = "D"
        Cells(i, 2).Interior.ColorIndex = 6
        Cells(i, 2).Font.ColorIndex = 3
    Case Else
        Cells(i, 2) = "F"
        Cells(i, 2).Interior.ColorIndex = 3
End Select

Next i

End Sub
