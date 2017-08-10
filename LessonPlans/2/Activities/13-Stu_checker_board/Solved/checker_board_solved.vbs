Sub CheckerBoard()

' Loop Row at a time
' For odd Rows color starts red
' For even color starts black

For k = 1 To 8

    If k Mod 2 = 0 Then
        ' Loop for red starts at columns 1,3,5,7
        For i = 1 To 8
            ' if row is odd print black
            If i Mod 2 <> 0 Then
                Cells(i, k).Interior.ColorIndex = 1
            ' else the row is even print red
            Else
                Cells(i, k).Interior.ColorIndex = 3
            End If

        Next i
    Else
        ' Lopp for black start 2,4,6,8
        For i = 1 To 8

            If i Mod 2 = 0 Then
                Cells(i, k).Interior.ColorIndex = 1
            Else
                Cells(i, k).Interior.ColorIndex = 3
            End If

        Next i
    End If

Next k


End Sub
