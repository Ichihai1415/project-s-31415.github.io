Sub main()
  Dim i As Long, count As Long, n As Long
  Dim x As Double, y As Double, pi As Double
  n = 100
  count = 0
  For i = 1 To 10
    x = Rnd
    y = Rnd
    If x * x + y * y <= 1
      count = count + 1
    End If
  Next i
  pi = count / n * 4
  Cells(5, 2).Value = n
  Cells(5, 3).Value = count
  Cells(5, 4).Value = pi
End Sub