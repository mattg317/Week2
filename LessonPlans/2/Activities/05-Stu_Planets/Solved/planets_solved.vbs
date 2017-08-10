Sub planets()

' Create the variables for our code
Dim i As Integer
Dim total As Integer
Dim moons As Integer
Dim mass As Single
Dim earthDiamter As Integer

' Set the initial values for our variables
total = 0
moons = 0
mass = 0
earthDiameter = Range("F5")

' Loops through rows 3 to 11. We skip the sun because it is not a planet.
For i = 3 To 11

    ' Counts the total number of planets in our solar system
    total = total + 1

    ' Adds the mass of the planet stored within column G to the total mass
    mass = mass + Range("G" & i)

    ' Adds the number of moons stored within column M to the total number of moons
    moons = moons + Range("M" & i)

    'BONUS: Takes the diameter of the current planet and compares it to that of Earth
    Range("H" & i) = Range("F" & i) - earthDiameter
    
Next i

' Print final results to the table
Cells(16, 7).Value = total
Cells(17, 7).Value = mass
Cells(18, 7).Value = moons / total

End Sub