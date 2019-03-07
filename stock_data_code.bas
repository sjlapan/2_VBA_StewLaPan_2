Attribute VB_Name = "Module1"
Sub stocksummary()

For Each ws In Worksheets

' Label columns to store ticker value, yearly change, percent change, and total volume in

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Vol"

' Create variable to store ticker value

Dim ticker_value As String

' Create and initialize a variable to store total volume

Dim total_vol As LongLong

total_vol = 0

' Keep track of summary table row

Dim summary_table_row As Integer

' Keep track of open and close entries for a ticker value

Dim year_open As Double
Dim year_close As Double

' Create variables to store the line value of the year open and year close
Dim open_line As Long
Dim close_line As Long

' Initialize year_open and the open/close line trackers before the loop

year_open = ws.Range("C2").Value
open_line = 2
close_line = 2
summary_table_row = 2

' Find the last row of the spreadsheet

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Format percent change cells
ws.Range("K:K").NumberFormat = "0.00%"

' Loop through all ticker values

For i = 2 To LastRow

    ' Compile if the cell value in ticker column has changed
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'set ticker value to current cell
        ticker_value = ws.Cells(i, 1)
        
        ' Add final value to total volume
        total_vol = total_vol + ws.Cells(i, 7).Value
        
        'store the close value for the ticker value
        year_close = ws.Cells(i, 6).Value
        close_line = close_line + 1
        
        ' Print ticker value and total volume to summary table
        ws.Range("I" & summary_table_row).Value = ticker_value
        ws.Range("L" & summary_table_row).Value = total_vol
        
        'Calculate yearly change
        ws.Range("J" & summary_table_row).Value = (year_close - year_open)
        
        ' Format cells to more clearly show growth or loss
        If ws.Range("J" & summary_table_row).Value > 0 Then
            ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
        Else
            ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
        End If
        
        'Calculate percent change, providing exception for zero values
        If year_open = 0 Then
            For j = open_line To close_line
                If Cells(j, 3) <> 0 Then
                    year_open = Cells(j, 3)
                    ws.Range("K" & summary_table_row).Value = ((year_close - year_open) / year_open)
                    Exit For
                Else
                    ws.Range("K" & summary_table_row).Value = 0
                End If
            Next j
        Else
            ws.Range("K" & summary_table_row).Value = ((year_close - year_open) / year_open)
        End If
        
        ' Set the year open for the new ticker value
        year_open = ws.Cells(i + 1, 3).Value
        open_line = i + 1
        
        'Reset total volume
        total_vol = 0
    
        'Add to summary table line
        summary_table_row = summary_table_row + 1
    Else
        total_vol = total_vol + ws.Cells(i, 7).Value
        close_line = close_line + 1
    End If
Next i

Next ws

End Sub

