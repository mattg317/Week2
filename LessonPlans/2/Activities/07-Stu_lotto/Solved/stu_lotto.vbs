Sub lotto_winner()

' Make sure to set your variables to Long because of how big the number is
Dim first_place As Long
Dim second_place As Long
Dim third_place As Long
Dim runner1 As Long
Dim runner2 As Long
Dim runner3 As Long
' Set the variable i as an integer since we will be using it to run our FOR loop
Dim i As Integer

' Set the values for first, second, and third place that we will be checking against
first_place = 3957481
second_place = 5865187
third_place = 2817729

' BONUS: Set the initial values runners-up
runner1 = 2275339
runner2 = 5868182
runner3 = 1841402

' Loop through each of the lotto numbers
For i = 1 To 1001

    ' Check to see if the value within the cell in column C matches our first_place value
    If Cells(i, 3).Value = first_place Then
        ' If the value within the cell in column C matches that of our first_place value, run the following code
        MsgBox " Congratulations " + Cells(i, 1).Value
        Cells(2, 6).Value = Cells(i, 1).Value
        Cells(2, 7).Value = Cells(i, 2).Value
        Cells(2, 8).Value = first_place

    ' Check to see if the value within the cell in column C matches our second_place value
    ElseIf Cells(i, 3).Value = second_place Then
        ' If the value within the cell in column C matches that of our second_place value, run the following code
        Cells(3, 6).Value = Cells(i, 1).Value
        Cells(3, 7).Value = Cells(i, 2).Value
        Cells(3, 8).Value = second_place

    ' Check to see if the value within the cell in column C matches our third_place value
    ElseIf Cells(i, 3).Value = third_place Then
        ' If the value within the cell in column C matches that of our third_place value, run the following code
        Cells(4, 6).Value = Cells(i, 1).Value
        Cells(4, 7).Value = Cells(i, 2).Value
        Cells(4, 8).Value = third_place

    ' BONUS: Check for runner ups with an OR operator
    ElseIf Cells(i, 3).Value = runner1 Or Cells(i, 3).Value = runner2 Or Cells(i, 3).Value = runner3 Then
        ' If the value within the cell in column C matches that of any of our runner_up values, run the following code
        runner_up = Cells(i, 3).Value
        Cells(5, 6).Value = Cells(i, 1).Value
        Cells(5, 7).Value = Cells(i, 2).Value
        Cells(5, 8).Value = runner_up

    ' Ends this series of IF/ELSE conditionals
    End If

Next i

End Sub
