Attribute VB_Name = "Module1"
Sub DataRetrieval():

Dim ws As Worksheet
For Each ws In Worksheets

ws.Range("J1").Value = "Ticker Symbol"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Ticker Yearly Change"
ws.Range("M1").Value = "Total Stock Volume"
ws.Range("P2").Value = "Ticker"
ws.Range("Q2").Value = "Value"
ws.Range("O3").Value = "Greatest % Increase"
ws.Range("O4").Value = "Greatest % Decrease"
ws.Range("O5").Value = "Greatest Total Volume"


Dim ticker As String
Dim total_stock_volume As Double
total_stock_volume = 0
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim high_price As Double
high_price = 0
Dim low_price As Double
low_price = 0
Dim percentage_change As Double
percentage_change = 0
Dim yearly_change As Double
yearly_change = 0

Dim greatest_percentage_increase As Double
Dim greatest_increase_ticker As String
Dim greatest_percentage_decrease As Double
Dim greatest_decrease_ticker As String
Dim greatest_total_volume As Double
Dim greatest_volume_ticker As String



Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
  
  Dim Lastrow As Long
  Dim i As Integer
  
  Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  For i = 2 To Lastrow


    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      ticker = ws.Cells(i, 1).Value
      close_price = ws.Cells(i, 6).Value
      open_price = ws.Cells(i, 3).Value
      yearly_change = (close_price - open_price)
      
      
      
      If yearly_change = 0 Then
        percentage_change = yearly_change
    ElseIf yearly_change <> 0 Then
        percentage_change = yearly_change / open_price
        
    End If
      total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
      ws.Range("J" & Summary_Table_Row).Value = ticker
      ws.Range("L" & Summary_Table_Row).Value = yearly_change
      ws.Range("K" & Summary_Table_Row).Value = percentage_change
      
If yearly_change > 0 Then
        ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
ElseIf yearly_change < 0 Then
        ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
        
        End If
        
If percentage_change > 0 Then
    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
ElseIf percentage_change < 0 Then
    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3

        End If
        
      ws.Range("M" & Summary_Table_Row).Value = total_stock_volume
      ws.Range("K" & Summary_Table_Row).Value = percentage_change
      Summary_Table_Row = Summary_Table_Row + 1
      total_stock_volume = 0
    
    Else
      total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                    

End If

    Next i
  
greatest_percentage_increase = Application.WorksheetFunction.Max(ws.Range("K:K"))
ws.Range("Q3").Value = greatest_percentage_increase

greatest_percentage_decrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
ws.Range("Q4").Value = greatest_percentage_decrease

greatest_total_volume = Application.WorksheetFunction.Max(ws.Range("M:M"))
ws.Range("Q5").Value = greatest_total_volume

ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q4").NumberFormat = "0.00%"


      For i = 2 To Lastrow
      
If ws.Cells(i, 11).Value = greatest_percentage_increase Then
    greatest_increase_ticker = ws.Cells(i, 10).Value
    
ElseIf ws.Cells(i, 11) = greatest_percentage_decrease Then
    greatest_decrease_ticker = ws.Cells(i, 10).Value
  
End If

    Next i
    
ws.Range("P3").Value = greatest_increase_ticker
ws.Range("P4").Value = greatest_decrease_ticker


For i = 2 To Lastrow

If ws.Cells(i, 13).Value = greatest_total_volume Then
    greatest_volume_ticker = ws.Cells(i, 10).Value
End If
Next i
ws.Range("P5").Value = greatest_volume_ticker

ws.Range("J:M").Columns.AutoFit
ws.Range("O:Q").Columns.AutoFit

    Next ws
    End Sub
    
    


        
           
