Sub ModerateSolution2015()

    ActiveWorkbook.Worksheets("2015").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("2015").Sort.SortFields.Add2 Key:=Range("A2:A760192"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("2015").Sort.SortFields.Add2 Key:=Range("B2:B760192"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("2015").Sort
        .SetRange Range("A1:G760192")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
 Dim ticker As String
  Dim Total_Stock_Volume, Open_Value, Close_Value  As Double
  Total_Stock_Volume = 0
  Open_Value = 0
  Close_Value = 0

  Dim lastrow, i, Summary_Table_Row As Long
  Summary_Table_Row = 2
    
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row

  Range("I1").Value = "Ticker"
  Range("J1").Value = "Yearly Change"
  Range("K1").Value = "Percent Change"
  Range("L1").Value = "Total Stock Volume"
  
  For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ticker = Cells(i, 1).Value
      
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
      Close_Value = Cells(i, 6).Value
      Range("I" & Summary_Table_Row).Value = ticker
      Range("J" & Summary_Table_Row).Value = Open_Value - Close_Value
        
        If Open_Value <> 0 Then
      Range("K" & Summary_Table_Row).Value = ((Close_Value - Open_Value) / (Abs(Open_Value)))
      Else
      End If
      Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

      Summary_Table_Row = Summary_Table_Row + 1

      Total_Stock_Volume = 0
      Open_Value = 0
      Close_Value = 0
      
    Else
    
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            If Open_Value = 0 Then
               Open_Value = Cells(i, 6).Value
            Else
               Open_Value = Open_Value
            End If

    End If
    
  Next i
  
  For Each J In Range("J2:J" & Cells(Rows.Count, "J").End(xlUp).Row)
 If J.Value > 0 Then
       J.Interior.ColorIndex = 4
    ElseIf J.Value < 0 Then
       J.Interior.ColorIndex = 3
    Else
       J.Interior.ColorIndex = xlNone
    End If
Next J
    Range("O2").Value = "Greatest % increase"
    Range("O3").Value = "Lowest % decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"

    Cells(2, 17).Value = WorksheetFunction.Max(Range("K2:K" & Cells(Rows.Count, 11).End(xlUp).Row))
    Cells(2, 16).Value = WorksheetFunction.Index(Range("I2:K" & Cells(Rows.Count, 9).End(xlUp).Row), WorksheetFunction.Match(Cells(2, 17).Value, Range("K2:K" & Cells(Rows.Count, 12).End(xlUp).Row), 0), 1)
    Range("Q2").NumberFormat = "0.00%"
    Cells(3, 17).Value = WorksheetFunction.Min(Range("K2:K" & Cells(Rows.Count, 11).End(xlUp).Row))
    Cells(3, 16).Value = WorksheetFunction.Index(Range("I2:K" & Cells(Rows.Count, 9).End(xlUp).Row), WorksheetFunction.Match(Cells(3, 17).Value, Range("K2:K" & Cells(Rows.Count, 12).End(xlUp).Row), 0), 1)
    Range("Q3").NumberFormat = "0.00%"
    Cells(4, 17).Value = WorksheetFunction.Max(Range("L2:L" & Cells(Rows.Count, 12).End(xlUp).Row))
    Cells(4, 16).Value = WorksheetFunction.Index(Range("I2:K" & Cells(Rows.Count, 11).End(xlUp).Row), WorksheetFunction.Match(Cells(4, 17).Value, Range("L2:L" & Cells(Rows.Count, 12).End(xlUp).Row), 0), 1)
End Sub


