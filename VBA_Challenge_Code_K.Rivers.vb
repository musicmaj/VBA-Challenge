Sub alphabetws()

    ' Write Leaderboard Columns
    ActiveSheet.Range("I1").Value = "Ticker"
    ActiveSheet.Range("J1").Value = "Quarterly Change"
    ActiveSheet.Range("K1").Value = "Percent Change"
    ActiveSheet.Range("L1").Value = "Volume"
    
    ActiveSheet.Range("O2").Value = "Greatest % Increase"
    ActiveSheet.Range("O3").Value = "Greatest % Decrease"
    ActiveSheet.Range("O4").Value = "Greatest Total Volume"
    ActiveSheet.Range("P1").Value = "Ticker"
    ActiveSheet.Range("P2").Value = "Value"
    
    
    
    ' variables
    Dim ticker As String
    Dim summary_row As Integer
    Dim ticker_open As Double
    Dim ticker_close As Double
    Dim quarter_change As Double
    Dim percent_change As Double
    Dim total_volume As LongLong
    Dim last_row As Long
    
    Dim max_change As Double
    Dim max_change_ticker As String
    Dim min_change As Double
    Dim min_change_ticker As String
    Dim max_volume As LongLong
    Dim max_volume_ticker As String
    
    Dim i As Long ' row number
    
    ' Initialize variables
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
    
    i = 2
    summary_row = 2
    ticker_open = ActiveSheet.Cells(2, 3).Value
    total_volume = 0
    
    max_change = 0
    min_change = 0
    max_volume = 0
    
    
    ' Loop
    For i = 2 To lastRow
        If ActiveSheet.Cells(i + 1, 1).Value <> ActiveSheet.Cells(i, 1).Value Then
            ticker = ActiveSheet.Cells(i, 1).Value
            
            ' Display ticker label
            ActiveSheet.Cells(summary_row, 9).Value = ticker
            
            ' Get quarterly change
            ticker_close = ActiveSheet.Cells(i, 6).Value
            quarter_change = ticker_close - ticker_open
            ActiveSheet.Cells(summary_row, 10).Value = quarter_change
            
            ' We need to set Red for loss and Green for gain
            ' And White for no change
            If (quarter_change > 0) Then
                ActiveSheet.Cells(summary_row, 10).Interior.ColorIndex = 4
            ElseIf (quarter_change < 0) Then
                ActiveSheet.Cells(summary_row, 10).Interior.ColorIndex = 3
            Else
                ActiveSheet.Cells(summary_row, 10).Interior.ColorIndex = 2
            End If
            
            ' Get percentage change from the opening value
            ' Opening price can be 0.00, so we need to catch that error
            If (ticker_open <> 0) Then
                price_change = quarter_change / ticker_open
            Else
                price_change = 0
            End If
            ActiveSheet.Cells(summary_row, 11).NumberFormat = "0.00%"
            ActiveSheet.Cells(summary_row, 11).Value = price_change
        
            ' Display total stock volume
            ActiveSheet.Cells(summary_row, 12).Value = total_volume
            
            ' Determine min/max change and highest total volume
            If (price_change < min_change) Then
                min_change = price_change
                min_change_ticker = ticker
                ActiveSheet.Cells(3, 16).Value = min_change_ticker
                ActiveSheet.Cells(3, 17).Value = min_change
                ActiveSheet.Cells(3, 17).NumberFormat = "0.00%"
            ElseIf (price_change > max_change) Then
                max_change = price_change
                max_change_ticker = ticker
                ActiveSheet.Cells(2, 16).Value = max_change_ticker
                ActiveSheet.Cells(2, 17).Value = max_change
                ActiveSheet.Cells(2, 17).NumberFormat = "0.00%"
            End If
            
            If (total_volume > max_volume) Then
                max_volume = total_volume
                max_volume_ticker = ticker
                ActiveSheet.Cells(4, 16).Value = max_volume_ticker
                ActiveSheet.Cells(4, 17).Value = max_volume
            End If
            
            ' After getting the quarterly change, reset values for next company
            total_volume = 0
            ticker_open = ActiveSheet.Cells(i + 1, 3).Value
            
            
            
            summary_row = summary_row + 1
        Else
            total_volume = total_volume + ActiveSheet.Cells(i, 7).Value
        End If
    Next i
    
End Sub

Sub forEachWs()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ActiveWorkbook.Worksheets(ws.Name).Activate
        Call alphabetws
        
        ActiveSheet.Columns("I:Q").AutoFit
    Next ws
End Sub

