Sub birth_stone()

    'Sets the values of A1 and B1 and sets their font size to 16
    Cells(1, 1).Value = "Birthday Month"
    Cells(1, 2).Value = "Birthday Stone"
    Cells(1, 1).Font.Size = 16
    Cells(1, 2).Font.Size = 16
    
    'Declares the variable "month" as a string and sets it to the value of A2
    Dim month As String
    month = Cells(2, 1).Value
    
    'Series of conditionals used to determine/set the value of B2 depending upon the value of "month"
    Select Case month
        Case Is = "January"
            Range("B2") = "Garnet"
        Case Is = "February"
            Range("B2") = "Amethyst"
        Case Is = "March"
            Range("B2") = "Aquqamarine"
        Case Is = "April"
            Range("B2") = "Diamond"
        Case Is = "May"
            Range("B2") = "Emerald"
        Case Is = "June"
            Range("B2") = "Pearl Alexandrite"
        Case Is = "July"
            Range("B2") = "Ruby"
        Case Is = "August"
            Range("B2") = "Peridot Sardonyx Spinel"
        Case Is = "September"
            Range("B2") = "Sapphire"
        Case Is = "October"
            Range("B2") = "Tourmaline"
        Case Is = "November"
            Range("B2") = "Tourmaline Opal"
        Case Is = "December"
            Range("B2") = "Tanzanite Zircon Turqoise"
        Case Else
            Range("B2") = "Invalid Month!"
    End Select

End Sub
