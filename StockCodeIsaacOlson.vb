Sub StockStuff():
    Dim TickerSymbol, NextTicker As String
    Dim StartPrice, EndPrice As Double
    Dim TotalStockVol As Double
    Dim colVal, rowVal, dataRow As Long
    Dim Test As Integer
    
    Dim ws As Worksheet
   ' Dim c, n As Integer
   ' c = ActiveWorkbook.Worksheets.Count
    'For n = 1 To c Step 1
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        
        'putting lables into the first row
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        'getting everything set for the first run
        TickerSymbol = Cells(2, 1).Value
        rowVal = 2
        dataRow = 2
        TotalStockVol = 0
        StartPrice = Cells(2, 3).Value
        Do While (Cells(rowVal, 1).Value <> "")
        
            TotalStockVol = TotalStockVol + Cells(rowVal, 7).Value ' Overflow
            NextTicker = Cells(rowVal + 1, 1).Value
            Test = StrComp(NextTicker, TickerSymbol)
            If (Test <> 0) Then
            'else then
                'set all values for the rows and reintialize all values. working
                Cells(dataRow, 9).Value = TickerSymbol
                Cells(dataRow, 10).Value = -(StartPrice - Cells(rowVal, 6))
                If (StartPrice <> 0) Then
                    Cells(dataRow, 11).Value = Str(Round((Cells(rowVal, 6).Value - StartPrice) / StartPrice * 100, 2)) + "%"
                End If
                Cells(dataRow, 12).Value = TotalStockVol
                
                'reseting variables working
                TotalStockVol = 0
                TickerSymbol = Cells(rowVal + 1, 1).Value
                StartPrice = Cells(rowVal + 1, 3).Value
                
                'coloring working
                If (Cells(dataRow, 10).Value < 0) Then
                    Cells(dataRow, 10).Interior.Color = RGB(255, 0, 0)
                Else
                      Cells(dataRow, 10).Interior.Color = RGB(0, 255, 0)
                End If
                dataRow = dataRow + 1
                
            End If
            
            rowVal = rowVal + 1
         Loop
             
             
        'Challenge 1
        Dim PerMax, PerMin, CheckValue, MaxVolume, CheckVolume As Double
        Dim MaxPTick, MinPTick, VolTick As String
        
        'setting up the area
        Cells(1, 15).Value = "Ticker"
        Cells(1, 16).Value = "Value"
        Cells(2, 14).Value = "Greatest % Increase"
        Cells(3, 14).Value = "Greatest % Decrease"
        Cells(4, 14).Value = "Greatest Total Volume"
        
        'initalize the values
        rowVal = 3
        PerMax = Cells(2, 11).Value
        PerMin = Cells(2, 11).Value
        MaxVolume = Cells(2, 12).Value
        MaxPTick = Cells(2, 9).Value
        VolTick = Cells(2, 9).Value
        MinPTick = Cells(2, 9).Value
        
        Do While (Cells(rowVal, 9).Value <> "")
            CheckValue = Cells(rowVal, 11).Value
            CheckVolume = Cells(rowVal, 12).Value
            'Check Three things and if true update values.
            If (CheckValue > PerMax) Then
                PerMax = CheckValue
                MaxPTick = Cells(rowVal, 9).Value
            ElseIf (CheckValue < PerMin) Then
                PerMin = CheckValue
                MinPTick = Cells(rowVal, 9).Value
            End If
            If (CheckVolume > MaxVolume) Then
                MaxVolume = CheckVolume
                VolTick = Cells(rowVal, 9).Value
            End If
            
            rowVal = rowVal + 1
        Loop
        'writing the values.
        Cells(2, 15).Value = MaxPTick
        Cells(2, 16).Value = Str(PerMax * 100) + "%"
        Cells(3, 15).Value = MinPTick
        Cells(3, 16).Value = Str(PerMin * 100) + "%"
        Cells(4, 15).Value = VolTick
        Cells(4, 16).Value = MaxVolume
        
     Next ws
   ' Next n
End Sub

