Attribute VB_Name = "Module1"
Sub stockticker()
'initiate worksheet loop
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Range("A1").Value = "<ticker>"
        'ws.Range("A1").Select
    
        'declare and define necessary variables
        Dim stock_name As String
        stock_name = Cells(2, 1).Value
        Dim stock_volume As Double
        stock_volume = 0
        Dim table_row As Double
        table_row = 2
        Dim Starting_value As Double
        Starting_value = Cells(2, 3).Value
        Dim Closing_value As Double
        Dim Yearly_change As Double
        Dim Greatest_pct_increase As Double
        Dim Greatest_pct_decrease As Double
        Dim Greatest_total_volume As Double
        Greatest_pct_increase = 0
        Greatest_pct_decrease = 0
        Greatest_total_volume = 0
        
        Dim biggest_gain_stock As String
        Dim biggest_loss_stock As String
        Dim biggest_volume_stock As String
        
        
        'set headers for summary table
        ws.Range("L1").Value = "Ticker"
        ws.Range("M1").Value = "Yearly Change"
        ws.Range("N1").Value = "Percentage Change"
        ws.Range("O1").Value = "Total Stock Volume"
        ws.Range("R2").Value = "Greatest % Increase"
        ws.Range("R3").Value = "Greatest % Decrease"
        ws.Range("R4").Value = "Greatest Total Volume"
        ws.Range("S1").Value = "Ticker"
        ws.Range("T1").Value = "Value"
            
        
        'track output location
        Dim Summary_table_row As Integer
        Summary_table_row = 2
        
        'count rows
        Dim last_row As Double
        last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
        'MsgBox last_row
        
        'loop through all rows of stock worksheet
        For Row = 2 To last_row
        
            'check if stock name matches the stock name for the following row
            If ws.Cells(Row + 1, 1).Value <> stock_name Then
            
                
                'update stock volume total
                stock_volume = stock_volume + ws.Cells(Row, 7).Value
                
                'record closing value
                Closing_value = ws.Range("F" & Row).Value
                
                'output values to summary table
                ws.Range("O" & Summary_table_row).Value = stock_volume
                ws.Range("L" & Summary_table_row).Value = stock_name
                ws.Range("M" & Summary_table_row).Value = Closing_value - Starting_value
                ws.Range("N" & Summary_table_row).NumberFormat = "0.00%"
                ws.Range("N" & Summary_table_row).Value = ((Closing_value / Starting_value) - 1)
                
                Yearly_change = ws.Range("N" & Summary_table_row).Value
                
                'format color for yearly change
                If ws.Range("M" & Summary_table_row).Value < 0 Then
                    ws.Range("M" & Summary_table_row).Interior.ColorIndex = 3
                ElseIf ws.Range("M" & Summary_table_row).Value > 0 Then
                    ws.Range("M" & Summary_table_row).Interior.ColorIndex = 4
                End If
                
                'format color for percent change
                
                If ws.Range("N" & Summary_table_row).Value < 0 Then
                    ws.Range("N" & Summary_table_row).Interior.ColorIndex = 3
                ElseIf ws.Range("N" & Summary_table_row).Value > 0 Then
                    ws.Range("N" & Summary_table_row).Interior.ColorIndex = 4
                End If
                
                'look for greatest volume
                If stock_volume > Greatest_total_volume Then
                    Greatest_total_volume = stock_volume
                    biggest_volume_stock = stock_name
                End If
                
                'look for biggest gainer and loser
                If Yearly_change > Greatest_pct_increase Then
                    Greatest_pct_increase = Yearly_change
                    biggest_gain_stock = stock_name
                ElseIf Yearly_change < Greatest_pct_decrease Then
                    Greatest_pct_decrease = Yearly_change
                    biggest_loss_stock = stock_name
                End If
                
                
                
                
                'reset stock volume, starting value and closing value
                stock_volume = 0
                Starting_value = ws.Cells(Row + 1, 3).Value
                stock_name = ws.Cells(Row + 1, 1).Value
                Summary_table_row = Summary_table_row + 1
                
                
                
            Else
                'add stock volume of current row to running total
                stock_volume = stock_volume + ws.Cells(Row, 7).Value
                
                
            
            
            End If
        Next Row
    
        'Format superlatives
        ws.Range("T2").NumberFormat = "0.00%"
        ws.Range("T3").NumberFormat = "0.00%"
        
        
        'print out superlatives
        ws.Range("R2").Value = "Greatest % Increase"
        ws.Range("R3").Value = "Greatest % Decrease"
        ws.Range("R4").Value = "Greatest Total Volume"
        ws.Range("S1").Value = "Ticker"
        ws.Range("T1").Value = "Value"
        ws.Range("S2").Value = biggest_gain_stock
        ws.Range("T2").Value = Greatest_pct_increase
        ws.Range("S3").Value = biggest_loss_stock
        ws.Range("T3").Value = Greatest_pct_decrease
        ws.Range("S4").Value = biggest_volume_stock
        ws.Range("T4").Value = Greatest_total_volume
        
        
        'MsgBox (biggest_gain_stock)
    Next ws
       
         'Autofit cells
        ActiveSheet.UsedRange.EntireColumn.AutoFit
    
    
    End Sub
