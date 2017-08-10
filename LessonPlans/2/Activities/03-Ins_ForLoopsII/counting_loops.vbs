Sub counter()

Dim counter, i As Integer

counter = 0
weight = 0

For i = 2 To 119
    counter = counter + 1
    weight = weight + Cells(i, 3).Value
Next i

Range("G5").Value = counter
Range("G7").Value = weight
Range("G9").Value = weight / counter

End Sub
