Sub counter()

' Create the variables for our code
Dim counter As Integer
Dim weight As Integer
Dim i As Integer

' Set the initial values for our variables
counter = 0
weight = 0

' Create a FOR loop to move through our table
For i = 2 To 119
    ' Counts the elements in the table
    counter = counter + 1
    ' Add the weight of each element to find the total weight
    weight = weight + Cells(i, 3).Value
Next i

' Display results in cells G5, G7, and G9
Range("G5").Value = counter
Range("G7").Value = weight
Range("G9").Value = weight / counter

End Sub

