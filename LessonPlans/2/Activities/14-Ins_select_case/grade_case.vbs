Sub select_case()

    Dim i As Integer
    
    'Start a FOR loop that will loop 100 times
    For i = 2 To 101
        'Starts a SELECT CASE conditional that looks at the values in column A one at a time
        Select Case Cells(i, 1).Value
            'If the value contained in column A is greater than or equal to 90, the value in B should be "A"
            Case Is >= 90
               Cells(i, 2).Value = "A"
               Cells(i, 2).Interior.ColorIndex = 10
            'If the value contained in column A is greater than or equal to 80, the value in B should be "B"
            Case Is >= 80
                Cells(i, 2).Value = "B"
                Cells(i, 2).Borders.ColorIndex = 10
            'If the value contained in column A is greater than or equal to 70, the value in B should be "C"
            Case Is >= 70
               Cells(i, 2) = "C"
               Cells(i, 2).Interior.ColorIndex = 6
            'If the value contained in column A is greater than or equal to 60, the value in B should be "D"
            Case Is >= 60
                Cells(i, 2) = "D"
                Cells(i, 2).Interior.ColorIndex = 6
                Cells(i, 2).Font.ColorIndex = 3
            'If none of the previous conditions were met, the value in B should be "F"
            Case Else
                Cells(i, 2) = "F"
                Cells(i, 2).Interior.ColorIndex = 3
        End Select
    Next i

End Sub
