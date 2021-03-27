Attribute VB_Name = "Module1"
Option Explicit

'Create a script that will loop through all the stocks for one year and output the following information.
'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.
'You should also have conditional formatting that will highlight positive change in green and negative change in red.
'The result should look as follows.


'list ticker symbol from column A

Sub StockAnalysis()

'variable for stock name
Dim ws As Worksheet
  
For Each ws In Worksheets
  Dim Ticker_Name As String


  ' Set an initial variable for holding
  Dim Ticker_Volume As Double
  Ticker_Volume = 0

  ' Keep track of the location for each ticke in the summary table
  Dim Summary_Row As Long
  Summary_Row = 2
  
  'variable for lastrow in the data
  Dim Lastrow As Long
  
  'varibale for i
  Dim i As Long

  'variable for date
   Dim OpenDate_BegYear As Long
   Dim CloseDate_EndYear As Long
   Dim OpenPrice_BegYear As Double
   Dim ClosePrice_EndYear As Double
   Dim Price_Diff As Double
   Dim Percent_Diff As Double

  
   OpenPrice_BegYear = 0
   ClosePrice_EndYear = 0
   Price_Diff = 0
   Percent_Diff = 0
   
   Dim GreatestInc As Double
   Dim GreatestDec As Double
   Dim GreatestTotal As Double
   Dim GreatestIncTickerName As String
   Dim GreatestDecTickerName As String
   Dim GreatestTotalTickerName As String

   GreatestInc = 0
   GreatestDec = 0
   GreatestTotal = 0
   
  OpenPrice_BegYear = ws.Cells(2, 3).Value
  
  Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
 
  ws.Range("J1") = "Ticker"
  ws.Range("K1") = "PriceDiff"
  ws.Range("L1") = "%Diff"
  ws.Range("M1") = "TickerVolume"
  
  ws.Range("O2") = "Greatest % Inc."
  ws.Range("O3") = "Greatest % Dec."
  ws.Range("O4") = "Greatest Total Vol."
  ws.Range("P1") = "Ticker"
  ws.Range("Q1") = "Value"
  
  ' Loop through all Tickers
  For i = 2 To Lastrow
  
    ' Check if still same ticker, if not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' set the stock names
      Ticker_Name = ws.Cells(i, 1).Value
    
      'get closing price and price diff, initial opening price referenced above
      ClosePrice_EndYear = ws.Cells(i, 6).Value
      Price_Diff = ClosePrice_EndYear - OpenPrice_BegYear
      
      If OpenPrice_BegYear <> "0" Then
      Percent_Diff = (Price_Diff / OpenPrice_BegYear)
      Else: Percent_Diff = "0"
      End If
    
      ' Add the stock volume
      Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
      
      
      ' Print the stock names in the Summary Table
      ws.Range("J" & Summary_Row).Value = Ticker_Name

      ' Print the stock volume to the Summary Table
      ws.Range("M" & Summary_Row).Value = Ticker_Volume
      
       ' Print the price & percent diff to the Summary Table
      ws.Range("K" & Summary_Row).Value = Price_Diff
      ws.Range("L" & Summary_Row).Value = Percent_Diff
      

     
      
      ' Add one to the summary table row
      Summary_Row = Summary_Row + 1
      OpenPrice_BegYear = ws.Cells(i + 1, 3).Value
      
      'get the max % inc, min % dec and max total vol
      If Percent_Diff > GreatestInc Then
      GreatestInc = Percent_Diff
      GreatestIncTickerName = Ticker_Name
      ElseIf Percent_Diff < GreatestDec Then
      GreatestDec = Percent_Diff
      GreatestDecTickerName = Ticker_Name
      End If

      If Ticker_Volume > GreatestTotal Then
      GreatestTotal = Ticker_Volume
      GreatestTotalTickerName = Ticker_Name
      End If
      
        ws.Range("P2").Value = GreatestIncTickerName
        ws.Range("Q2").Value = GreatestInc
        ws.Range("P3").Value = GreatestDecTickerName
        ws.Range("Q3").Value = GreatestDec
        ws.Range("P4").Value = GreatestTotalTickerName
        ws.Range("Q4").Value = GreatestTotal
 
      ' Reset the Totals
      ClosePrice_EndYear = 0
      Price_Diff = 0
      Percent_Diff = 0
      Ticker_Volume = 0

    Else
    
      ' Add to the Ticker Total
      Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
     
     

End If

Next i

Call format




Next ws

End Sub


 
Sub format()

Dim ws As Worksheet
  
For Each ws In Worksheets
 Dim i As Double
 Dim LastTickerRow As Double
 
 LastTickerRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
 
    
    For i = 2 To LastTickerRow
    If ws.Cells(i, 11).Value > "0" Then
    ws.Cells(i, 11).Interior.ColorIndex = 4
    ElseIf ws.Cells(i, 11).Value <= "0" Then
    ws.Cells(i, 11).Interior.ColorIndex = 3
    End If
    
    ws.Cells(i, 12).NumberFormat = "0.00%"
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).NumberFormat = "000,000"
    Next i
    
Next ws
    
End Sub



