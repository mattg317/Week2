' Part I: Count the number of Hornets Found
' Part II: Each time you find Hornets replace them with Bugs 
' Part III: You have a max amount of Bees and Hornets, utilize no more than what's provided. 
'           If there are still hornets left, provide the user with a message stating: "Oh no! We still have hornets..."

Sub HornetsNest()

  ' PART I
  ' ------------------------------------------------
  ' Create a variable to hold the number of hornets
  Dim HornetsCount as Integer

  ' Create a variable to hold the number of bugs and bees 
  Dim BugsCount as Integer
  Dim BeesCount as Integer

  ' Set the value of Bugs and Bees counters 
  BugsCount = Range("L2").Value
  BeesCount = Range("R2").Value

  ' Set the initial value for the HornetsCount to 0 
  HornetsCount = 0

  ' Loop through all rows
  For i = 1 to 6

    ' Loop through all columns
    For j = 1 to 7

      ' If the value of a cell is equal to Hornets
      If Cells(i, j).Value = "Hornets" Then

        ' Add to the HornetsCounter
        HornetsCount = HornetsCount + 1 

        ' Check if we have bugs available
        If (BugsCount <= 0) Then

          ' Replace the Hornets with Bugs 
          Cells(i, j).Value = "Bugs"

          ' Subtract from the Bugs Count
          BugsCount = BugsCount - 1

        ' Check if we have bees available
        Elseif (BeesCount <= 0 ) Then

          ' Replace the Hornets with Bees
          Cells(i, j).Value = "Bees"

          ' Subtract from the Bees Count
          BeesCount = BeesCount - 1

        Else 

          MsgBox("Oh no! We still have hornets")

      End If

    Next j

  Next i

  ' Show the number of hornets found
  MsgBox(HornetsCount & " Hornets Found")

End Sub
