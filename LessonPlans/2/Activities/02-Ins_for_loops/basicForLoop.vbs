Sub forLoop()

Dim i As Integer

	For i = 1 To 20

	    ' rows
	    Cells(i, 1).Value = 1

	    ' columns
	    Cells(1, i).Value = 1

	    ' variable math
	    Cells(i + 1, 2).Value = i + 1
	   
	Next i

End Sub
  