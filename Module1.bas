Attribute VB_Name = "Module1"
Sub StocksVs2()
    
    'Set variables for making summary table based on year
    Dim CurrDate As Double
    Dim CurrYear As Integer
    Dim LastDate As Double
    Dim LastYear As Integer
    Dim NextDate As Double
    Dim NextYear As Integer
    
    Dim CurrTicker As String
    Dim LastTicker As String
    Dim OpenPrice As Currency
    Dim ClosePrice As Currency
    
    Dim SummaryTableRow As Integer
    Dim SummaryTickerHighPercent As String
    Dim SummaryTickerLowPercent As String
    Dim SummaryTickerHighTotal As String
    Dim LastOverallMax As Double
    Dim LastOverallMin As Double
    Dim LastHighestVolTotal As Variant
    
    'Set column indexes
    Const TickerColIndex As Integer = 1
    Const DateColIndex As Integer = 2
    Const OpenColIndex As Integer = 3
    Const CloseColIndex As Integer = 6
    Const StockVolColIndex As Integer = 7
    Const STYear As Integer = 9
    Const STTickerColIndex As Integer = 10
    Const STPercentChangeColIndex As Integer = 12
    Const TotalStockVolIndex As Integer = 13
    
    'Assign values to variables outside of For Loop
    SummaryTableRow = 2
    LastOverallMax = 0
    LastOverallMin = 0
    LastHighestVolTotal = 0
    
    For Each ws In Worksheets
    
    
    
        'Set variables for inside the For Loop
        Dim lastrow As Long
        Dim YearlyChange As Currency
        Dim PercentDiff As Double
        Dim StockTotal As Variant
        Dim CurrStockVol As Variant
        Dim LastStockVol As Variant
        Dim Max As Long
        
        
        'Assign some values
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        StockTotal = 0
        
        LastTicker = ""
        LastYear = 0
        
        
        For Row = 2 To lastrow
            CurrDate = ws.Cells(Row, DateColIndex).Value
            CurrYear = Left(CurrDate, 4)
            CurrTicker = ws.Cells(Row, TickerColIndex).Value
            If LastYear & LastTicker <> CurrYear & CurrTicker Then
                'Add current year to summary table
                Range("I" & SummaryTableRow).Value = CurrYear
                'add ticker to summary table
                Range("J" & SummaryTableRow).Value = CurrTicker
                'Store the open price value
                OpenPrice = ws.Cells(Row, OpenColIndex).Value
                'Store the StockTotal
                StockTotal = ws.Cells(Row, StockVolColIndex).Value
            Else
                CurrStockVol = ws.Cells(Row, StockVolColIndex).Value
                StockTotal = CDec(StockTotal) + CDec(CurrStockVol)
            End If
            LastTicker = CurrTicker
            LastYear = CurrYear
            NextDate = ws.Cells(Row + 1, DateColIndex).Value
            NextYear = Left(NextDate, 4)
            NextTicker = ws.Cells(Row + 1, TickerColIndex).Value
            
            If CurrYear & CurrTicker <> NextYear & NextTicker Then
                'Add the last stock volume to the total
                LastStockVol = ws.Cells(Row, StockVolColIndex).Value
                StockTotal = CDec(StockTotal) + CDec(LastStockVol)
                
                'Puts the total volume into the summary table
                Range("M" & SummaryTableRow).Value = StockTotal
                
                'Store the close price value
                ClosePrice = ws.Cells(Row, CloseColIndex).Value
                
                'Calculate the yearly price change and put in summary table
                YearlyChange = ClosePrice - OpenPrice
                Range("K" & SummaryTableRow).Value = YearlyChange
                
                'Calcualte the percent difference and put in summary table
                If OpenPrice <> 0 Then
                    PercentDiff = (((YearlyChange) / OpenPrice) * 100)
                    Range("L" & SummaryTableRow).Value = PercentDiff
                ElseIf OpenPrice = 0 Then
                    Range("L" & SummaryTableRow).Value = "Not Defined"
                End If
                
                'Format colors of cells
                If YearlyChange > 0 Then
                    Range("K" & SummaryTableRow).Interior.ColorIndex = 4 '4 is green
                ElseIf YearlyChange < 0 Then
                    Range("K" & SummaryTableRow).Interior.ColorIndex = 3 '3 is red
                End If

                'Add final stock total to summary table at this point
                Range("M" & SummaryTableRow).Value = StockTotal
                
                'Now that this ticker and year are done add a new row to the summary table
                SummaryTableRow = SummaryTableRow + 1
            End If
        Next Row
    Next ws
    
    'Bonus - Find the ticker with the highest percent change, lowest percent change, and greatest volume
    For Row = 2 To (SummaryTableRow - 1)
            'Bonus highest percent change
            If IsNumeric(Cells(Row, STPercentChangeColIndex).Value) And LastOverallMax < Cells(Row, STPercentChangeColIndex).Value Then
                 LastOverallMax = Cells(Row, STPercentChangeColIndex).Value
                 SummaryTickerHighPercent = Cells(Row, STTickerColIndex).Value
            End If
            
             
    
            'Bonus lowest percent change
            If IsNumeric(Cells(Row, STPercentChangeColIndex).Value) And LastOverallMin > Cells(Row, STPercentChangeColIndex).Value Then
                LastOverallMin = Cells(Row, STPercentChangeColIndex).Value
                SummaryTickerLowPercent = Cells(Row, STTickerColIndex).Value
            End If

            
  
            'Bonus highest volume
            If CDec(Cells(Row, TotalStockVolIndex).Value) > LastHighestVolTotal Then
                LastHighestVolTotal = Cells(Row, TotalStockVolIndex).Value
                SummaryTickerHighTotal = Cells(Row, STTickerColIndex).Value
            End If

     Next Row
     
     'Put the values in the bonus table
    Cells(2, 16).Value = SummaryTickerHighPercent
    Cells(2, 17).Value = LastOverallMax
    Cells(3, 16).Value = SummaryTickerLowPercent
    Cells(3, 17).Value = LastOverallMin
    Cells(4, 16).Value = SummaryTickerHighTotal
    Cells(4, 17).Value = LastHighestVolTotal
End Sub







   
   
   
   
   
   
   
   
   
   

