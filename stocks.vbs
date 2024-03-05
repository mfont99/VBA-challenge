Attribute VB_Name = "Module2"
Sub stocks()

Dim ws As Worksheet

Dim ticker_count As Integer     'Variables
Dim current_open As Double
Dim current_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim volume_change As Double
Dim total_volume As Double
Dim last_row_data As Double
Dim last_row_calcs As Double
Dim greatest_gain As Double
Dim greatest_loss As Double
Dim greatest_volume As Double
Dim gg_count As String
Dim gl_count As String
Dim gv_count As String

For Each ws In ThisWorkbook.Worksheets        'Loop through each worksheet

    ticker_count = 2
    current_open = ws.Cells(2, 3).Value
    last_row_data = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To last_row_data

        volume_change = ws.Cells(i, 7).Value        'Running total of current ticker's volume
        total_volume = total_volume + volume_change
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ws.Cells(ticker_count, 9).Value = ws.Cells(i, 1).Value      'Set tickers under Ticker header
            
            current_close = ws.Cells(i, 6).Value        'Set current close price
            
            yearly_change = (current_close - current_open)      'Calculate and set yearly change column
            ws.Cells(ticker_count, 10).Value = yearly_change
            
            percent_change = (current_close - current_open) / (current_open)    'Calculate and set percentage change
            ws.Cells(ticker_count, 11).Value = percent_change
                
            ws.Cells(ticker_count, 12).Value = total_volume     'Set total volume
            
            total_volume = 0                            'Reset and increment for next ticker
            current_open = ws.Cells(i + 1, 3).Value
            ticker_count = ticker_count + 1
        
        End If
    
    Next i
    
    last_row_calcs = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
    greatest_gain = ws.Cells(2, 11).Value
    
    For j = 2 To last_row_calcs
        
        ws.Cells(j, 11).NumberFormat = "0.00%"      'Format column "K" to be percentages
        
        If greatest_gain < ws.Cells(j + 1, 11).Value Then       'Determine greatest percentage gain
            greatest_gain = ws.Cells(j + 1, 11).Value
            gg_count = ws.Cells(j + 1, 9).Value
        End If
        
        If greatest_loss > ws.Cells(j + 1, 11).Value Then       'Determine greatest percentage loss
            greatest_loss = ws.Cells(j + 1, 11).Value
            gl_count = ws.Cells(j + 1, 9).Value
        End If
        
        If greatest_volume < ws.Cells(j + 1, 12).Value Then     'Determine greatest total volume
            greatest_volume = ws.Cells(j + 1, 12).Value
            gv_count = ws.Cells(j + 1, 9).Value
        End If
        
        If ws.Cells(j, 10).Value >= 0 Then                      'Set background colors in column "J"
            ws.Cells(j, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 3
        End If
        
    Next j
    
    ws.Cells(2, 17).Value = greatest_gain       'Set values for greatest % increase, greatest % decrease, and greatest total volume
    ws.Cells(3, 17).Value = greatest_loss
    ws.Cells(4, 17).Value = greatest_volume
    
    ws.Cells(2, 16).Value = gg_count            'Set the corresponding tickers
    ws.Cells(3, 16).Value = gl_count
    ws.Cells(4, 16).Value = gv_count
    
    'Formatting
    
    ws.Range("I1").Value = "Ticker"                     'Add headers/labels
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ws.Cells(2, 17).NumberFormat = "0.00%"              'Format Q2 and Q3 to be percentages
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ws.Range("I1").EntireColumn.AutoFit                 'Adjust column widths
    ws.Range("J1").EntireColumn.AutoFit
    ws.Range("K1").EntireColumn.AutoFit
    ws.Range("L1").EntireColumn.AutoFit
    ws.Range("O1").EntireColumn.AutoFit
    ws.Range("P1").EntireColumn.AutoFit
    ws.Range("Q1").EntireColumn.AutoFit
    
    
Next ws

End Sub













