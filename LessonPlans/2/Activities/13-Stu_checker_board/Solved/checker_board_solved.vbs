Sub CheckerBoard()

    ' The first loop is going through each row
    For k = 1 To 8
        ' For even rows, odd numbered columns are black and evens are red
        If k Mod 2 = 0 Then
        ' This nested loop then looks through each column in even rows
            For i = 1 To 8
                ' If the column an odd number, sets the cell's formatting to black
                If i Mod 2 <> 0 Then
                    Cells(i, k).Interior.ColorIndex = 1
                ' If the column is an even number, sets the cell's formatting to red
                Else
                    Cells(i, k).Interior.ColorIndex = 3
                End If
            Next i
        ' For odd rows, odd numbered columns are red and evens are black
        Else
            For i = 1 To 8
                ' if the column is an odd number, sets the cell's formatting to red
                If i Mod 2 <> 0 Then
                    Cells(i, k).Interior.ColorIndex = 3
                ' If the column is an even number, sets the cell's formatting to black
                Else
                    Cells(i, k).Interior.ColorIndex = 1
                End If
            Next i
        End If
    Next k

End Sub