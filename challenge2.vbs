Sub highlow()
    Dim biginc As Variant
    Dim bigdec As Variant
    Dim bigvol As Variant
    Dim tbiginc As String
    Dim tbigdec As String
    Dim tbigvol As String
    Dim i As Integer
    Dim j As Integer
    workcount = ActiveWorkbook.Worksheets.Count
    For i = 1 To workcount
        MsgBox (Sheets(i).Name)
        biginc = 0
        bigdec = 0
        bigvol = 0
        j = 2
        Do Until IsEmpty(Sheets(i).Cells(j, 9).Value)
            If Sheets(i).Cells(j, 11).Value > biginc Then
                biginc = Sheets(i).Cells(j, 11).Value
                tbiginc = Sheets(i).Cells(j, 9).Value
            End If
            If Sheets(i).Cells(j, 11).Value < bigdec Then
                bigdec = Sheets(i).Cells(j, 11).Value
                tbigdec = Sheets(i).Cells(j, 9).Value
            End If
            If Sheets(i).Cells(j, 12).Value > bigvol Then
                bigvol = Sheets(i).Cells(j, 12).Value
                tbigvol = Sheets(i).Cells(j, 9).Value
            End If
            j = j + 1
        Loop
        Sheets(i).Range("n2").Value = "Greatest % Increase"
        Sheets(i).Range("n3").Value = "Greatest % Decrease"
        Sheets(i).Range("n4").Value = "Greatest Total Volume"
        Sheets(i).Range("o1").Value = "Ticker Symbol"
        Sheets(i).Range("o2").Value = tbiginc
        Sheets(i).Range("o3").Value = tbigdec
        Sheets(i).Range("p2:p3").Style = "Percent"
        Sheets(i).Range("o4").Value = tbigvol
        Sheets(i).Range("p2").Value = biginc
        Sheets(i).Range("p3").Value = bigdec
        Sheets(i).Range("p4").Value = bigvol
    Next i
End Sub