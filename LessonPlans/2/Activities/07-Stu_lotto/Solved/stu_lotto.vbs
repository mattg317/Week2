Sub lotto_winner()

' Set dimensions to Long because of how big the number is
Dim first_place As Long
Dim second_place As Long
Dim third_place As Long
Dim runner1 As Long
Dim runner2 As Long
Dim runner3 As Long

first_place = 3957481
second_place = 5865187
third_place = 2817729

' BONUS: runner ups
runner1 = 2275339
runner2 = 5868182
runner3 = 1841402

Dim i As Integer

' Loop through lotto numbers
For i = 1 To 1001

    ' Check to see if the number matches our winner
    If Cells(i, 3).Value = first_place Then
        MsgBox " Congratulations " + Cells(i, 1).Value
        Cells(2, 6).Value = Cells(i, 1).Value
        Cells(2, 7).Value = Cells(i, 2).Value
        Cells(2, 8).Value = first_place
    ' Check to see if the number matches our winner
    ElseIf Cells(i, 3).Value = second_place Then
        Cells(3, 6).Value = Cells(i, 1).Value
        Cells(3, 7).Value = Cells(i, 2).Value
        Cells(3, 8).Value = second_place

    ' Check to see if the number matches our winner
    ElseIf Cells(i, 3).Value = third_place Then
        Cells(4, 6).Value = Cells(i, 1).Value
        Cells(4, 7).Value = Cells(i, 2).Value
        Cells(4, 8).Value = third_place

    ' BONUS
    ' Check with an Or operator
    ElseIf Cells(i, 3).Value = runner1 Or Cells(i, 3).Value = runner2 Or Cells(i, 3).Value = runner3 Then
        runner_up = Cells(i, 3).Value
        Cells(5, 6).Value = Cells(i, 1).Value
        Cells(5, 7).Value = Cells(i, 2).Value
        Cells(5, 8).Value = runner_up

    End If

Next i

End Sub
