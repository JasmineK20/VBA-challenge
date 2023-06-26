Attribute VB_Name = "Module1"

Sub MultiYearStock()
    Dim ws As Worksheet
    Dim i As Double
    Dim j As Double
    Dim ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim Volume As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim StartRow As Double
    Dim EndRow As Double
    
    For Each ws In Worksheets
        ws.Activate
        LastRow = Cells(Rows.Count, "A").End(xlUp).Row
        Cells(1, "i").Value = "Ticker"
        Cells(1, "j").Value = "Yearly Change"
        Cells(1, "k").Value = "Percent Change"
        Cells(1, "l").Value = "Stock Volume"
        
        For i = 2 To LastRow
        
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                OpenPrice = Cells(i, 3).Value
                StartRow = Cells(i, 1).Row
                ticker = Cells(i, 1).Value
                Cells(Rows.Count, "I").End(xlUp).Offset(1, 0).Activate
                ActiveCell.Value = ticker
            End If
            
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                ClosePrice = Cells(i, 6).Value
                EndRow = Cells(i, 1).Row
                Volume = Application.WorksheetFunction.Sum(Range(Cells(StartRow, "G"), Cells(EndRow, "G")))
                YearlyChange = ClosePrice - OpenPrice
                PercentChange = (ClosePrice - OpenPrice) / OpenPrice
                
                Cells(Rows.Count, "J").End(xlUp).Offset(1, 0).Activate
                ActiveCell.Value = Format(YearlyChange, "#.00")
                Cells(Rows.Count, "K").End(xlUp).Offset(1, 0).Activate
                ActiveCell.Value = PercentChange
                Cells(Rows.Count, "L").End(xlUp).Offset(1, 0).Activate
                ActiveCell.Value = Volume
            End If
            
        Next i
        
        Range("i:l").EntireColumn.AutoFit
        LastRow2 = Cells(Rows.Count, "I").End(xlUp).Row
        
        For j = 2 To LastRow2
            If Cells(j, "j").Value > 0 Then
                Cells(j, "j").Interior.Color = vbGreen
            Else
                Cells(j, "j").Interior.Color = vbRed
            End If
        Next j
        
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
        Range("N2").Value = "Greatest Percent Increase"
        Range("N3").Value = "Greatest Percent Decrease"
        Range("N4").Value = "Greatest Total Volume"
        Range("P2").Value = Application.WorksheetFunction.Max(Range(Cells(2, "K"), Cells(LastRow2, "K")))
        Range("P3").Value = Application.WorksheetFunction.Min(Range(Cells(2, "K"), Cells(LastRow2, "K")))
        Range("P4").Value = Application.WorksheetFunction.Max(Range(Cells(2, "L"), Cells(LastRow2, "L")))
        
        HighIncrease = Range("P2").Value
        Range("K:K").Find(HighIncrease).Activate
        ticker = ActiveCell.Offset(0, -2).Value
        Range("O2").Value = ticker
        
        HighDecrease = Range("P3").Value
        Range("K:K").Find(HighDecrease).Activate
        ticker = ActiveCell.Offset(0, -2).Value
        Range("O3").Value = ticker
        
        HighVolume = Range("P4").Value
        Range("L:L").Find(HighVolume).Activate
        ticker = ActiveCell.Offset(0, -3).Value
        Range("O4").Value = ticker
        
        Range("K:K").NumberFormat = "0.00%"
        Range("P2").NumberFormat = "0.00%"
        Range("P3").NumberFormat = "0.00%"
        Range("n:p").EntireColumn.AutoFit
        
    Next ws
End Sub
