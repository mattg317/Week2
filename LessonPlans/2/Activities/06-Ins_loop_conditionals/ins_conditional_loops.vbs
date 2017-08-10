Sub conditional_loops()

Dim i As Integer

  For i = 1 To 10
      If Cells(i, 1).Value Mod 2 = 0 Then
        Cells(i, 2).Value = "Even Row"
      Else
        Cells(i,2).Value = "Odd Row"
      End If
  Next i

End Sub
