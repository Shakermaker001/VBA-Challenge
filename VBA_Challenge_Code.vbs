VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Analysis()

 ' Declare Current as a worksheet object variable.
         Dim ws As Worksheet

         ' Loop through all of the worksheets in the active workbook.
         For Each ws In Worksheets






Dim ticker As String '3-digit abbreviation
Dim yearly_change As Double ' value change (+,-)
Dim percent_chnage As Double ' (%)
Dim total_stock_volume As LongLong 'total volume traded
Dim analysis_table_row As Integer ' location of looped data
Dim open_price_row As Double

last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row ' pings last row in for loop

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("O2").Value = "Greatest Percentage Increase"
ws.Range("O3").Value = "Greatest Percentage Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

yearly_change = 0
percennt_change = 1
total_stock_volume = 0
analysis_table_row = 2
open_price_row = 2

    For i = 2 To last_row
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ticker = ws.Cells(i, 1).Value
            
            yearly_change = ws.Cells(i, 6).Value - ws.Cells(open_price_row, 3).Value
            
            percent_change = yearly_change / ws.Cells(open_price_row, 3).Value
            
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
            ws.Range("i" & analysis_table_row).Value = ticker
            
            ws.Range("J" & analysis_table_row).Value = yearly_change
            
            ws.Range("K" & analysis_table_row).Value = percent_change
        
            ws.Range("K" & analysis_table_row).NumberFormat = "0.00%"
            
            ws.Range("L" & analysis_table_row).Value = total_stock_volume
            
            If ws.Range("J" & analysis_table_row).Value > 0 Then
             ws.Range("J" & analysis_table_row).Interior.ColorIndex = 4
             ElseIf ws.Range("J" & analysis_table_row).Value < 0 Then
              ws.Range("J" & analysis_table_row).Interior.ColorIndex = 3
              End If
              
              
             
            
            analysis_table_row = analysis_table_row + 1
            
            total_stock_volume = 0
            
            ticker = 0
            
            open_price_row = i + 1
            
        Else
        
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        
            
        
        End If
    
    Next i
    
    ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & last_row))
    
            ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & last_row))
            ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & last_row))
    max_increase_index = WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K" & last_row), 0)
     max_decrease_index = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K" & last_row), 0)
      max_volume_index = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L" & last_row), 0)

    
   ws.Range("P2").Value = ws.Cells(max_increase_index + 1, 9).Value
    ws.Range("P3").Value = ws.Cells(max_decrease_index + 1, 9).Value
     ws.Range("P4").Value = ws.Cells(max_volume_index + 1, 9).Value

Next ws

End Sub
