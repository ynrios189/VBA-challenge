Sub ModerateSolution2016()

    ActiveWorkbook.Worksheets("2016").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("2016").Sort.SortFields.Add2 Key:=Range("A2:A797711"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("2016").Sort.SortFields.Add2 Key:=Range("B2:B797711"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("2016").Sort
        .SetRange Range("A1:G797711")
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

End Sub




