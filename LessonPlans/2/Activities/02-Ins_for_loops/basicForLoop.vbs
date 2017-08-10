Sub forLoop()

Dim i As Integer
    For i = 1 To 20
        ' Places a value of 1 in A1 to A20
        Cells(i, 1).Value = 1
        ' Places a value of 1 in A1 to T1
        Cells(1, i).Value = 1
        ' Places increasing values based upon the variable "i" in B2 to B21
        Cells(i + 1, 2).Value = i + 1
    Next i
End Sub
  