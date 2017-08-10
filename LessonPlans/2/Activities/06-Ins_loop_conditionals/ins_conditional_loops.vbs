Sub conditional_loops()

Dim i As Integer

  For i = 1 To 10
  
    ' Check to see if the "i" row is an even number and run the following code if true
    If Cells(i, 1).Value Mod 2 = 0 Then
        ' Places the value "Even Row" in the B column for that row
        Cells(i, 2).Value = "Even Row"
        
    ' Otherwise, run the following code if the "i" row is odd
    Else
        ' Places the value "Odd Row" in the B column for that row
        Cells(i, 2).Value = "Odd Row"
        
    ' Concludes this series of IF/ELSE statements
    End If
    
  Next i

End Sub
