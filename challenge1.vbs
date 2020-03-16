Sub stocks()
<<<<<<< HEAD
=======
    Dim current As String  ' working index name
>>>>>>> 72209ec0790b4ef56fcee7d10625a1f0528712df
    workcount = ActiveWorkbook.Worksheets.Count ' workcount counts worksheets
    For ws = 1 To workcount
      x = 2   ' index index initialization
      i = 2     ' active cell initializaion
      j = 2     ' opening date row variable
      Do Until IsEmpty(Sheets(ws).Cells(i, 1).Value) ' cause I can't remember the row counting script
        If Sheets(ws).Cells(i, 1).Value <> Sheets(ws).Cells(i + 1, 1).Value Then
          Sheets(ws).Cells(x, 12).Value = Application.Sum(Sheets(ws).Range("G" & j & ":G" & i)) 'Total volume summation
          Sheets(ws).Cells(x, 9).Value = Sheets(ws).Cells(j, 1).Value   'ticker registration
          Sheets(ws).Cells(x, 10).Value = Sheets(ws).Cells(i, 6).Value - Sheets(ws).Cells(j, 3).Value 'change calc
          If Sheets(ws).Cells(j, 3).Value > 0 Then  ' prevent division by zero
            Sheets(ws).Cells(x, 11).Value = Sheets(ws).Cells(x, 10).Value / Sheets(ws).Cells(j, 3).Value '% chg calc
          End If
          If Sheets(ws).Cells(x, 10).Value > 0 Then                    'cell formating
            Sheets(ws).Cells(x, 10).Interior.Color = RGB(0, 255, 0)
          End If
          If Sheets(ws).Cells(x, 10).Value < 0 Then
            Sheets(ws).Cells(x, 10).Interior.Color = RGB(255, 0, 0)
          End If
          x = x + 1  'progress index variables
          j = i + 1
        End If
        i = i + 1    ' next cell
      Loop
      Sheets(ws).Range("i1") = "Ticker"
      Sheets(ws).Range("j1") = "Yearly Change"
      Sheets(ws).Range("k1") = "Percent Change"
      Sheets(ws).Range("l1") = "Total Stock Volume"
      Sheets(ws).Range("k2:k" & x).Style = "Percent"
    Next ws
End Sub