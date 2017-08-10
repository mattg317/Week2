Sub planets()

' Set Dimensions
Dim i As Integer
Dim total As Integer
Dim moons As Integer
Dim mass As Single
Dim earthDiamter As Integer

' Set variables
total = 0
moons = 0
mass = 0
earthDiameter = Range("F5")

For i = 3 To 11

    ' Add to our variables each loop
    total = total + 1
    mass = mass + Range("G" & i)
    moons = moons + Range("M" & i)

    ' BONUS
    ' Take the current planets diameter minus the earths
    Range("H" & i) = Range("F" & i) - earthDiameter
    
Next i

' Print Final Results
Cells(16, 7).Value = total
Cells(17, 7).Value = mass
Cells(18, 7).Value = moons / total

End Sub