Sub reset():
    
' Store last grade within cells B12, C12, and D12
    Cells(12, 2) = Cells(2, 2).Value
    Cells(12, 3) = Cells(2, 3).Value
    Cells(12, 4) = Cells(2, 4).Value

' Empty out cells B2, C2, and D2 while also removing the fill color from C2
    Cells(2, 2).Value = ""
    Cells(2, 3).Value = ""
    Cells(2, 3).Interior.ColorIndex = 0
    Cells(2, 4).Value = ""

End Sub