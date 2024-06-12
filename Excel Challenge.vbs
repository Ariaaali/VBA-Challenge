Attribute VB_Name = "Module2"
Sub Stockprice()

Dim ws As Worksheet
Dim name_ticker As String
Dim row_name_ticker As Integer
Dim row_percentage_change As Integer
Dim row_quarterly_change As Integer

Dim RowCount As Long
Dim First_Row As Long

Dim opening_price As Double
Dim closing_price As Double
Dim Quarterly_Change As Double
Dim Percentage_change As Double
Dim Total_Stock_Volume As Double
Dim total_opening_price As Double
Dim total_closing_price As Double

For Each ws In Worksheets
row_name_ticker = 2
row_quartly_change = 2
row_total_stock_volume = 2
row_percentage_change = 2

total_opening_price = 0
total_closing_price = 0

ws.Cells(1, 9).Value = "<ticker>"
ws.Cells(1, 10).Value = "Quarterly change"
ws.Cells(1, 11).Value = "Percentage Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To RowCount

opening_price = ws.Cells(i, 3).Value

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

name_ticker = ws.Cells(i, 1).Value

ws.Range("I" & row_name_ticker).Value = name_ticker

closing_price = ws.Cells(i, 6).Value

Quarterly_Change = closing_price - opening_price

Percentage_change = (closing_price - opening_price) / opening_price

ws.Range("J" & row_quartly_change).Value = Quarterly_Change
  If Quarterly_Change >= 0 Then
    ws.Range("J" & row_quartly_change).Interior.ColorIndex = 4
   Else
   
    ws.Range("J" & row_quartly_change).Interior.ColorIndex = 3
    
    End If
    
ws.Range("J" & row_quartly_change).Style = "Currency"

ws.Range("K" & row_percentage_change).Value = Percentage_change

ws.Range("K" & row_percentage_change).NumberFormat = "0.00%"

row_name_ticker = row_name_ticker + 1
row_quartly_change = row_quartly_change + 1
row_total_stock_volume = row_total_stock_volume + 1
row_percentage_change = row_percentage_change + 1


Total_Stock_Volume = 0



Else

Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

ws.Range("L" & row_total_stock_volume).Value = Total_Stock_Volume

 
 End If
  
  Next i
  
     Next ws

End Sub
  


