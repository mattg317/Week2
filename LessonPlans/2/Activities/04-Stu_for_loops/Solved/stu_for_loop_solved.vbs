Sub loops_and_loops()

    'For loop created to loop through the first 10 rows
    For i = 1 To 10
        'Values in column A will always have a value of "I will eat "
        Cells(i, 1).Value = "I will eat "
        'Values in column B will take the value of the current loop and add 10
        Cells(i, 2).Value = i + 10
        'Values in column C will always have a value of "Chicken Nuggets"
        Cells(i, 3).Value = "Chicken Nuggets"
    Next i

End Sub
