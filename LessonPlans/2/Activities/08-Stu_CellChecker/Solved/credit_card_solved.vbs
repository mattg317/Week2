Sub credit_card()

Dim i As Integer
Dim j As Integer
Dim total As Integer

total = 0
' set a variable to keep track of where to print credit credit card
j = 0

For i = 2 To 101


    ' If we reach a need card type print total amount spent and the
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        ' Print the credit card type
        Range("G" & 2 + j).Value = Cells(i, 1).Value
        
        ' Print total amount of cards being used
        Range("H" & 2 + j).Value = total

        ' Add one to j so we know next CC gets prints on for 3
        j = j + 1
        
        ' Reset total
        total = 0

    
    Else
        ' otherwise we want to keep adding to the total of the current card type
        total = total + 1

     End If
Next i


End Sub
