Sub Stockdata()

Dim ws As Worksheet


For Each ws In Worksheets
            
            Dim worksheetname As String
            worksheetname = ws.Name
            ' Find last row in worksheet with data
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            ' Add ticker and total stock volume columns
            ws.Cells(1, 10).Value = "Ticker"
            ws.Cells(1, 11).Value = "Total Stock Volume"

            Dim ticker_name As String
            Dim total_stock_volume As Double
            total_stock_volume = 0
    
            Dim summary_table_row As Integer
            summary_table_row = 2
            ' for loop to run data to last row
            For i = 2 To LastRow
    
            ' move to next ticker symbol that is different from the prior
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ticker_name = Cells(i, 1).Value
            ' calculate total stock volume for each ticker
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
            
            ' display the ticker symbol to coincide with the total stock volume
            Range("J" & summary_table_row).Value = ticker_name
            Range("K" & summary_table_row).Value = total_stock_volume

            summary_table_row = summary_table_row + 1
            total_stock_volume = 0
            
            Else: total_stock_volume = total_stock_volume + Cells(i, 7).Value


            End If
            Next i
            
        
Next ws


End Sub



