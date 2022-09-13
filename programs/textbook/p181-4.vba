Sub Main()
  Dim i as Long, sum As Long
  Dim a(9) As Long
  sum = 0
  For i = 0 To 9
    a(i) = i + 1
  Next i
  For i = 1 To 9
    sum = sum + a(i)
  Next i
  Cells(1, 1).Value = sum
End Sub