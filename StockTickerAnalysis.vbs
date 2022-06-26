Attribute VB_Name = "Module1"
' MODULE 2- VB CHALLENGE - PRADIP RAODEO 6/20/22
Sub StockDataAnalysis():
' Set up variable of datatype worksheet
Dim v_Sheet As Worksheet
Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Stock Volume"
v_ticker = ""
v_TotalStockVolume = 0
v_summaryTableRow = 2

v_YearOpenPrice = 0
v_YearClosePrice = 0
v_TickerRowCount = 0
v_PercentChange = 0
v_ticker_row = 0

For Each v_Sheet In ThisWorkbook.Worksheets
    v_SheetName = v_Sheet.Name
    v_Sheet.Range("J1").Value = "Ticker"
    v_Sheet.Range("K1").Value = "Yearly Change"
    v_Sheet.Range("L1").Value = "Percent Change"
    v_Sheet.Range("M1").Value = "Total Stock Volume"
    'v_Sheet.Range("N1").Value = "Yr Open Price"
    'v_Sheet.Range("O1").Value = "Yr Close Price"
    'Find last row in sheet
    v_lastRow = v_Sheet.Cells(Rows.Count, 1).End(xlUp).Row

    For v_row = 2 To v_lastRow
     
        
        If v_Sheet.Cells(v_row + 1, 1).Value <> v_Sheet.Cells(v_row, 1).Value Then
            ' Stock ticker changes above
            v_ticker = v_Sheet.Cells(v_row, 1).Value
            v_Sheet.Cells(v_summaryTableRow, 10).Value = v_ticker
            'v_Sheet.Range("J" & v_summaryTableRow).Value = v_ticker
            
            ' make sure you add volume from last row
            v_TotalStockVolume = v_TotalStockVolume + v_Sheet.Cells(v_row, 7).Value
            v_Sheet.Cells(v_summaryTableRow, 13).Value = v_TotalStockVolume
            v_Sheet.Cells(v_summaryTableRow, 13).Style = "Comma"
            'as ticker changes last value for ticker is closed price
            v_YearClosePrice = v_Sheet.Cells(v_row, 6).Value
            
            'v_Sheet.Cells(v_summaryTableRow, 14).Value = v_YearOpenPrice
            'v_Sheet.Cells(v_summaryTableRow, 15).Value = v_YearClosePrice
            
            'Print yearly change for Ticker
            v_Sheet.Cells(v_summaryTableRow, 11).Value = v_YearClosePrice - v_YearOpenPrice
            v_Sheet.Cells(v_summaryTableRow, 11).Style = "Currency"
            
            If v_YearClosePrice - v_YearOpenPrice < 0 Then
                v_Sheet.Cells(v_summaryTableRow, 11).Interior.ColorIndex = 3 ' Red
            ElseIf v_YearClosePrice - v_YearOpenPrice > 0 Then
                v_Sheet.Cells(v_summaryTableRow, 11).Interior.ColorIndex = 4 'GREEN
            Else
             ' NONE
            End If
            v_PercentChange = (v_Sheet.Cells(v_summaryTableRow, 11).Value) * (100 / v_YearOpenPrice)
            v_Sheet.Cells(v_summaryTableRow, 12).Value = v_PercentChange
            
            v_summaryTableRow = v_summaryTableRow + 1
            v_TotalStockVolume = 0
            v_YearOpenPrice = 0
            v_YearClosePrice = 0
            v_PercentChange = 0
        Else
            If v_YearOpenPrice > 0 Then
                ' Do nothing
            Else
                v_YearOpenPrice = v_Sheet.Cells(v_row, 3).Value
            End If
            
            v_TotalStockVolume = v_TotalStockVolume + v_Sheet.Cells(v_row, 7).Value
        End If
'        If v_row = v_lastRow Then
'            MsgBox (v_SheetName)
'        End If
        
    
    Next v_row
        v_summaryTableRow = 2
        
        'BONUS QUESTIONS
        v_Sheet.Cells(3, 15).Value = "Greatest % increase"
        v_Sheet.Cells(3, 17).Value = WorksheetFunction.Max(v_Sheet.Range("K2", "K" & v_lastRow))
        
        v_ticker_row = 0
        v_ticker_row = WorksheetFunction.Match(v_Sheet.Cells(3, 17).Value, v_Sheet.Range("K2", "K" & v_lastRow), 0) + 1
        v_Sheet.Cells(3, 16).Value = v_Sheet.Range("J" & Int(v_ticker_row)).Value
        
        v_Sheet.Cells(4, 15).Value = "Greatest % decrease"
        v_Sheet.Cells(4, 17).Value = WorksheetFunction.Min(v_Sheet.Range("K2", "K" & v_lastRow))
        
        
        v_ticker_row = 0
        v_ticker_row = WorksheetFunction.Match(v_Sheet.Cells(4, 17).Value, v_Sheet.Range("K2", "K" & v_lastRow), 0) + 1
        v_Sheet.Cells(4, 16).Value = v_Sheet.Range("J" & Int(v_ticker_row)).Value
        
        v_Sheet.Cells(5, 15).Value = "Greatest Total Volume"
        v_Sheet.Cells(5, 17).Value = WorksheetFunction.Max(v_Sheet.Range("M2", "M" & v_lastRow))
        v_Sheet.Cells(5, 17).Style = "Comma"


        v_ticker_row = 0
        v_ticker_row = WorksheetFunction.Match(v_Sheet.Cells(5, 17).Value, v_Sheet.Range("M2", "M" & v_lastRow), 0) + 1
        v_Sheet.Cells(5, 16).Value = v_Sheet.Range("J" & Int(v_ticker_row)).Value

Next v_Sheet

End Sub
