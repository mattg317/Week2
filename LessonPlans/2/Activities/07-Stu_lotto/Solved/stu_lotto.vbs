Sub lotto_winner()

first_place = 3957481
second_place = 5865187
third_place = 2817729

' BONUS: runner ups
runner1 = 2275339
runner2 = 5868182
runner3 = 1841402

Cells(1, 6).Value = "Winner - First Name"
Cells(1, 7).Value = "Winner - Last Name"
Cells(1, 8).Value = "Winning Number"

Dim i As Integer

For i = 1 To 1001
    If Cells(i, 3).Value = first_place Then
        MsgBox " Congratulations " + Cells(i, 1).Value
        Cells(2, 5).Value = "First"
        Cells(2, 6).Value = Cells(i, 1).Value
        Cells(2, 7).Value = Cells(i, 2).Value
        Cells(2, 8).Value = first_place

    ElseIf Cells(i, 3).Value = second_place Then
        Cells(3, 5).Value = "Second"
        Cells(3, 6).Value = Cells(i, 1).Value
        Cells(3, 7).Value = Cells(i, 2).Value
        Cells(3, 8).Value = second_place

    ElseIf Cells(i, 3).Value = third_place Then
        Cells(4, 5).Value = "Third"
        Cells(4, 6).Value = Cells(i, 1).Value
        Cells(4, 7).Value = Cells(i, 2).Value
        Cells(4, 8).Value = third_place

    ' BONUS
    ElseIf Cells(i, 3).Value = runner1 Or Cells(i, 3).Value = runner2 Or Cells(i, 3).Value = runner3 Then
        runner_up = Cells(i, 3).Value
        Cells(5, 5).Value = "Runner Up"
        Cells(5, 6).Value = Cells(i, 1).Value
        Cells(5, 7).Value = Cells(i, 2).Value
        Cells(5, 8).Value = runner_up

    End If

Next i

End Sub
