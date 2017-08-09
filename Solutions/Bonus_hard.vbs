Sub stock_analysis():

' Set dimensions
Dim total As Double
Dim i As Long
Dim change As Single
Dim j As Integer
Dim start As Long
Dim rowCount As Long
Dim percentChange As Single
Dim days As Integer
Dim dailyChange As Single
Dim averageChange As Single
Dim ws As Worksheet

    For Each ws In Worksheets
    ' Set title row
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"


    ' Set initial values
    j = 0
    total = 0
    change = 0
    start = 2
    dailyChange = 0



    ' get the row number of the last row with data
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To rowCount

        ' If ticker changes then print results
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Stores results in variables
                total = total + ws.Cells(i, 7).Value
                change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                percentChange = Round((change / ws.Cells(start, 3) * 100), 2)
                dailyChange = dailyChange + (ws.Cells(i, 4) - ws.Cells(i, 5))

                ' Average change
                days = (i - start) + 1
                averageChange = dailyChange / days

                ' start of the next stock ticker
                start = i + 1

                ' print the results to a seperate worksheet
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = Round(change, 2)
                ws.Range("K" & 2 + j).Value = "%" & percentChange
                ws.Range("D" & 2 + j).Value = averageChange
                ws.Range("L" & 2 + j).Value = total


            ' colors positives green and negatives red
                Select Case change
                    Case Is > 0
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select


                ' reset variables for new stock ticker
                total = 0
                change = 0
                j = j + 1
                days = 0
                dailyChange = 0

        ' If ticker is still the same add results
            Else
                total = total + ws.Cells(i, 7).Value
                change = change + (ws.Cells(i, 6) - ws.Cells(i, 3))

                ' change in high and low
                dailyChange = dailyChange + (ws.Cells(i, 4) - ws.Cells(i, 5))

            End If

        Next i

        ' take the max and min and place them in a separate part in the worksheet
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))

        ' returns one less because header row not a factor
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)

        ' final ticker symbol for  total, greatest % of increase and decrease, and average
        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)

    Next ws

End Sub
