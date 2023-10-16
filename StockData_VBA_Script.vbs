Attribute VB_Name = "Module1"
Sub StockAnalysis():

    Dim startTime As Single
    Dim endTime As Single
    
    startTime = Timer
    
For Each ws In Worksheets

    Dim WorksheetName As String
    Dim PercentChange As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecease As Double
    Dim GreatestTotalVolume As Double
    
WorksheetName = ws.Name
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
        TickCount = 2
        j = 2
        
LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To LastRowA
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
            If ws.Cells(TickCount, 10).Value < 0 Then
                ws.Cells(TickCount, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(TickCount, 10).Interior.ColorIndex = 4
            End If
            If ws.Cells(j, 3).Value <> 0 Then
                PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                ws.Cells(TickCount, 11).Value = Format(PercentChange, "Percent")
            Else
                ws.Cells(TickCount, 11).Value = Format(0, "Percent")
            End If
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                TickCount = TickCount + 1
                j = i + 1
            End If
        Next i
            LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
            GreatestIncrease = ws.Cells(2, 11).Value
            GreatestDecrease = ws.Cells(2, 11).Value
            GreatestTotalVolume = ws.Cells(2, 12).Value
              
        For i = 2 To LastRowI
            If ws.Cells(i, 12).Value > GreatestTotalVolume Then
            GreatestTotalVolume = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        Else
            GreatestTotalVolume = GreatestTotalVolume
        End If
            If ws.Cells(i, 11).Value > GreatestIncrease Then
            GreatestIncrease = ws.Cells(i, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        Else
            GreatestIncrease = GreatestIncrease
        End If
            If ws.Cells(i, 11).Value < GreatestDecrease Then
            GreatestDecrease = ws.Cells(i, 11).Value
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        Else
            GreatestDecrease = GreatestDecrease
        End If
            ws.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatestTotalVolume, "Scientific")
            Next i
            
     endTime = Timer
     MsgBox "This script ran in " & (endTime - startTime) & " seconds"
    
    Worksheets(WorksheetName).Columns("A:Z").AutoFit
        
    Next ws
    
End Sub
