Sub MarketAnalysis():
    Dim tickername As String
    Dim tickervolume As Double
    Dim summary_ticker_row As Integer
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim lastrow As Long
    Dim lastrow_summary_table As Long
    Dim i As Long
    
' Set appropriate header

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

' Initialize Variable
tickervolume = 0
summary_ticker_row = 2
open_price = Cells(2, 3).Value

' Count total number of rows in the first column
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' Loop through by ticker name
For i = 2 To lastrow
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' Set ticker name
tickername = Cells(i, 1).Value

' Add volume of trade
tickervolume = tickervolume + Cells(i, 7).Value

' Print the ticker name in the summary table
Range("I" & summary_ticker_row).Value = tickername

' Print the trade volume for each ticker in the summary table
Range("L" & summary_ticker_row).Value = tickervolume

' Find Yearly Change
close_price = Cells(i, 6).Value
yearly_change = (close_price - open_price)

' Print yearly change in the summary table
Range("J" & summary_ticker_row).Value = yearly_change

' Find Percent Change
If open_price = 0 Then
percent_change = 0

Else
percent_change = yearly_change / open_price

End If
' Print Percent Change in the summary table
Range("K" & summary_ticker_row).Value = percent_change
Columns("K").NumberFormat = "0.00%"

'Reset the row counter. Add one to the summary_ticker_row
summary_ticker_row = summary_ticker_row + 1

'Reset volume of trade to zero
tickervolume = 0

'Reset the opening price
open_price = Cells(i + 1, 3)
            
Else
              
'Add the volume of trade
tickervolume = tickervolume + Cells(i, 7).Value

            
End If
        
Next i
'Conditional formatting that will highlight positive change in green and negative change in red
'First find the last row of the summary table

lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
'Color code yearly change
For i = 2 To lastrow_summary_table
If Cells(i, 10).Value > 0 Then
Cells(i, 10).Interior.ColorIndex = 10

Else
Cells(i, 10).Interior.ColorIndex = 3
End If

Next i

'Highlight the stock price changes
'First label the cells according to the sample .png provided in the assignment

        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"

'Determine the max and min values in column "Percent Change" and just max in column "Total Stock Volume"
'collect the ticker name, and the corresponding values for the percent change and total volume of trade for that ticker
    
For i = 2 To lastrow_summary_table
'Find the maximum percent change
If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
Cells(2, 16).Value = Cells(i, 9).Value
Cells(2, 17).Value = Cells(i, 11).Value
Cells(2, 17).NumberFormat = "0.00%"

'Find the minimum percent change
ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
Cells(3, 16).Value = Cells(i, 9).Value
Cells(3, 17).Value = Cells(i, 11).Value
Cells(3, 17).NumberFormat = "0.00%"
            
'Find the maximum volume of trade
ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
Cells(4, 16).Value = Cells(i, 9).Value
Cells(4, 17).Value = Cells(i, 12).Value
            
End If
        
Next i
        
End Sub
