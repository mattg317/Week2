Sub reset():
    
' Store last Grade
    Cells(12, 2) = Cells(2, 2).Value
    Cells(12, 3) = Cells(2, 3).Value
    Cells(12, 4) = Cells(2, 4).Value

' Empty out cells
    Cells(2, 2).Value = ""
    Cells(2, 3).Value = ""
    Cells(2, 3).Interior.ColorIndex = 0
    Cells(2, 4).Value = ""

End Sub