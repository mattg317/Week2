Sub counter()

' Set Dimensions
Dim counter
Dim weight As Integer
Dim i As Integer

' Set initial variables
counter = 0
weight = 0

' Loop through our elements
For i = 2 To 119
    ' count elements
    counter = counter + 1
    ' Add the weight of each element
    weight = weight + Cells(i, 3).Value
Next i

' Display results
Range("G5").Value = counter
Range("G7").Value = weight
Range("G9").Value = weight / counter

End Sub
