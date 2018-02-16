Attribute VB_Name = "Module1"
Sub loop_through_all_worksheets()
    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Set starting_ws = ActiveSheet

    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        Call CalculateTotalVolume
    Next

    starting_ws.Activate

End Sub

Sub CalculateTotalVolume()
    Dim Volume As Double
    Dim i As Long
    Dim j As Long
    Dim TotalRecord As Long
    Dim Ticker_Counter As Integer
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim GPctIncrease As Double
    Dim GPctDecrease As Double
    Dim GTotalVol As Double
 
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    TotalRecord = Cells(Rows.Count, 1).End(xlUp).Row
    Ticker_Counter = 1
    
    Open_Price = Cells(2, 3).Value
     j = 2
    For i = 2 To TotalRecord
        Open_Price = Cells(j, 3).Value
            
        'Calculate total volume for each stock
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            Volume = Volume + Cells(i, 7).Value
        Else 'calculte data and save them to excel sheet
            Volume = Volume + Cells(i, 7).Value
            Close_Price = Cells(i, 6).Value
            YearlyChange = Close_Price - Open_Price
            If Open_Price = 0# Then
                PercentChange = 0#
            Else
                PercentChange = YearlyChange / Open_Price
            End If
            Cells(Ticker_Counter + 1, 9).Value = Cells(i, 1).Value
            Cells(Ticker_Counter + 1, 10).Value = YearlyChange
            If YearlyChange > 0 Then
                Cells(Ticker_Counter + 1, 10).Interior.ColorIndex = 10
            ElseIf YearlyChange < 0 Then
                Cells(Ticker_Counter + 1, 10).Interior.ColorIndex = 3
            End If
            Cells(Ticker_Counter + 1, 11).Value = PercentChange
            Cells(Ticker_Counter + 1, 11).NumberFormat = "0.00%"
            Cells(Ticker_Counter + 1, 12).Value = Volume
            Ticker_Counter = Ticker_Counter + 1
          
        
            Volume = 0
            j = i + 1
        End If
    Next i
    
     
    'Find the greastest percentage increase
        GPctIncrese = Cells(2, 11).Value
        For i = 3 To Ticker_Counter
    
            If Cells(i, 11).Value > GPctIncrease Then
                GPctIncrease = Cells(i, 11).Value
                ticker = Cells(i, 9).Value
            End If
        Next i
        
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("P2").Value = ticker
        Range("Q2").Value = GPctIncrease
        Range("Q2").NumberFormat = "0.00%"
        
        'Find the greastest percentage decrease
        GPctDecrease = Cells(2, 11).Value
        For i = 3 To Ticker_Counter
            If Cells(i, 11).Value < GPctDecrease Then
                GPctDecrease = Cells(i, 11).Value
                ticker = Cells(i, 9).Value
            End If
        Next i
      
        
        Range("O3").Value = "Greatest % Decrease"
        Range("P3").Value = ticker
        Range("Q3").Value = GPctDecrease
        Range("Q3").NumberFormat = "0.00%"
        
        'Find the greastest total volume
        GTotalVol = Cells(2, 12).Value
        For i = 3 To Ticker_Counter
            If Cells(i, 12).Value > GTotalVol Then
                GTotalVol = Cells(i, 12).Value
                ticker = Cells(i, 9).Value
            End If
        Next i
      
        
        Range("O4").Value = "Greatest Total Volume"
        Range("P4").Value = ticker
        Range("Q4").Value = GTotalVol

End Sub

