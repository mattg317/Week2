Sub planets()

Dim i As Integer
Dim total As Integer
Dim moons As Integer

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
    '
    Range("H" & i) = Range("F" & i) - earthDiameter
    
Next i

' Print Final Results
Cells(16, 7).Value = total
Cells(17, 7).Value = mass
Cells(18, 7).Value = moons / total

End Sub