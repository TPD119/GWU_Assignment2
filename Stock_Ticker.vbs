Attribute VB_Name = "Module1"
Sub stock_checker()
  Dim column As Integer
  Dim lastrow As Integer
  Dim ticker_row As Integer
  Dim ticker_total As Double
  Dim first_value As Double
  Dim last_value As Double
  Dim annual_change As Double
  Dim percent_change As Double
  
  Dim greatest_total As Double
  Dim greatest_increase As Double
  Dim greatest_decrease As Double
  
        For Each ws In Worksheets
          column = 1
          ticker_row = 2
          'lastrow = 1000
          last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
          ' Loop through rows in the column
          first_value = ws.Cells(2, 3).Value
          ws.Cells(1, 9).Value = "Ticker"
          ws.Cells(1, 10).Value = "Yearly Change"
          ws.Cells(1, 11).Value = "Percent Change"
          ws.Cells(1, 12).Value = "Total Stock Volume"
          ws.Cells(1, 16).Value = "Ticker"
          ws.Cells(1, 17).Value = "Total"
          ws.Cells(2, 15).Value = "Greatest Percent Increase"
          ws.Cells(3, 15).Value = "Greatest Percent Decrease"
          ws.Cells(4, 15).Value = "Greatest Total Increase"
          
          For i = 2 To last_row
            
            ' Searches for when the value of the next cell is different than that of the current cell
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                ws.Cells(ticker_row, 9).Value = ws.Cells(i, column).Value
                ticker_total = ticker_total + ws.Cells(i, 7).Value
                
                ws.Cells(ticker_row, 12).Value = ticker_total
                last_value = ws.Cells(i, 6).Value
                
                annual_change = last_value - first_value
                
                ws.Cells(ticker_row, 10).Value = annual_change
                If first_value = 0 Then
                    percent_change = 1
                Else
                    percent_change = annual_change / first_value
                End If
                
                ws.Cells(ticker_row, 11).Value = percent_change
                ws.Cells(ticker_row, 11).NumberFormat = "0.00%"
                
                If annual_change > 0 Then
                    ws.Cells(ticker_row, 10).Interior.ColorIndex = 4
                ElseIf annual_change < 0 Then
                    ws.Cells(ticker_row, 10).Interior.ColorIndex = 3
                End If
                
                first_value = ws.Cells(i + 1, 3).Value
                
                ticker_row = ticker_row + 1
                
                If ticker_total > greatest_total Then
                    greatest_total = ticker_total
                    greatest_ticker = ws.Cells(i, column).Value
                End If
                
                If percent_change > greatest_increase Then
                    greatest_increase = percent_change
                    increase_ticker = ws.Cells(i, column).Value
                End If
                
                If percent_change < greatest_decrease Then
                    greatest_decrease = percent_change
                    decrease_ticker = ws.Cells(i, column).Value
                End If
                
                ticker_total = 0
                
            Else
                ticker_total = ticker_total + ws.Cells(i, 7).Value
            End If
            
          Next i
          ws.Cells(4, 16).Value = greatest_ticker
          ws.Cells(4, 17).Value = greatest_total
          
          ws.Cells(2, 17).Value = greatest_increase
          ws.Cells(2, 17).NumberFormat = "0.00%"
          ws.Cells(2, 16).Value = increase_ticker
          
          ws.Cells(3, 17).Value = greatest_decrease
          ws.Cells(3, 17).NumberFormat = "0.00%"
          ws.Cells(3, 16).Value = decrease_ticker
          
          greatest_total = 0
          greatest_increase = 0
          greatest_decrease = 0
        Next ws
  
End Sub
