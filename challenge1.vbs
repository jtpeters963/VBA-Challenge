Sub stocks()
    Dim ope As Double  ' working opening value
    Dim cls As Double   ' working closing value
    Dim current As String  ' working index name
    Dim totvol As Variant  ' volume sum
    Dim x As Variant       ' Index index
    Dim dateopen As Variant
    Dim dateclosed As Variant
    Dim i As Double        'cell progressor
    workcount = ActiveWorkbook.Worksheets.Count
    For ws = 1 To workcount
      x = 2               'initialize counting variables
      i = 2
      totvol = 0                  'initialize first round of starting variables
      current = Sheets(ws).Range("A2").Value
      ope = Sheets(ws).Range("c2").Value
      dateopen = Sheets(ws).Range("b2").Value
      Do Until IsEmpty(Sheets(ws).Cells(i, 1).Value)                 ' cause I can't remember the row counting script
        If Sheets(ws).Cells(i, 1).Value = Sheets(ws).Cells(i + 1, 1).Value Then  ' check if last day, if not sum volume
          totvol = totvol + Sheets(ws).Cells(i, 7).Value
        Else
          cls = Sheets(ws).Cells(i, 6).Value
          dateclosed = Sheets(ws).Cells(i, 2).Value
          totvol = totvol + Sheets(ws).Cells(i, 7).Value
          Sheets(ws).Cells(x, 9).Value = current
          'Sheets(ws).Cells(x, 13).Value = ope
          'Sheets(ws).Cells(x, 14).Value = cls
          Sheets(ws).Cells(x, 12).Value = totvol
          Sheets(ws).Cells(x, 10).Value = cls - ope
          If ope > 0 Then
            Sheets(ws).Cells(x, 11).Value = (cls - ope) / ope
          End If
          'Sheets(ws).Cells(x, 15).Value = dateopen
          'Sheets(ws).Cells(x, 16).Value = dateclosed
          If Sheets(ws).Cells(x, 10).Value > 0 Then
            Sheets(ws).Cells(x, 10).Interior.Color = RGB(0, 255, 0)
          End If
          If Sheets(ws).Cells(x, 10).Value < 0 Then
            Sheets(ws).Cells(x, 10).Interior.Color = RGB(255, 0, 0)
          End If
          current = Sheets(ws).Cells(i + 1, 1).Value
          ope = Sheets(ws).Cells(i + 1, 4).Value
          dateopen = Sheets(ws).Cells(i + 1, 2).Value
          x = x + 1
          totvol = 0
        End If
        i = i + 1
      Loop
      Sheets(ws).Range("i1") = "Ticker"
      Sheets(ws).Range("j1") = "Yearly Change"
      Sheets(ws).Range("k1") = "Percent Change"
      Sheets(ws).Range("l1") = "Total Stock Volume"
      Sheets(ws).Range("k2:k" & x).Style = "Percent"
    Next ws
End Sub

