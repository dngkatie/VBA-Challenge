Sub stockdata()


'declare all the variables
Dim Ticker_Name As String
Dim lastrow As Long
  ' Keep track of the ticker
  Dim Summary_Table_Row As Integer
  
  Summary_Table_Row = 2




'setting the label for each value
Cells(1, 10).Value = "Ticker"
Cells(1, 13).Value = "Total Stock Volume"
Cells(1, 11).Value = "Yearly Changes"
Cells(1, 12).Value = "Percent Changes"
'count the row in the file
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
   
    ' Set the sticker name
      Ticker_Name = Cells(i, 1).Value
 ' Print the sticker in the Summary Table
      Range("j" & Summary_Table_Row).Value = Ticker_Name

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
    End If
Next i
 
 
 
 
 '--------------------------------------------------------------
 
 
 
 'getting value for total value
  ' Initialize variables for starting price and ending price
 
   Dim startingPrice As Single
   Dim endingPrice As Single
   'counting the row in the stickers comlumn
   Dim ticker_row As Long
   ticker_row = Range("j2").End(xlDown).Row
 'declare the sticker as array with range
   Dim tickers() As String
   ReDim tickers(ticker_row)
 'declare and initialize the value for total volume
   Dim total_volume As Long
   totalVolume = 0
   
 'getting the value from table into the array
   For N = 0 To ticker_row
   tickers(N) = Cells(N + 2, 10).Value
   Next N

'main code for getting the value of volume,starting and ending price

 For C = 0 To ticker_row - 2
 ticker = tickers(C)
  For j = 2 To lastrow
           ' Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 7).Value

           End If
           ' get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 3).Value

           End If

           ' get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       ' Output data for current ticker

       Cells(2 + C, 13).Value = totalVolume
       Cells(2 + C, 11).Value = endingPrice - startingPrice
       Cells(2 + C, 12).Value = (endingPrice - startingPrice) / startingPrice
       
       Cells(2 + C, 12).Value = FormatPercent(Cells(2 + C, 12))
       Cells(2 + C, 11).Value = Format(Cells(2 + C, 11), "#.00")
       
       
       
totalVolume = 0
      If Cells(2 + C, 11).Value >= 0 Then
        Cells(2 + C, 11).Interior.ColorIndex = 4
        Else
        Cells(2 + C, 11).Interior.ColorIndex = 3
        End If
    
        
Next C

'--------------------------------------------------'

'Bonus'
Dim Max As Double
Dim Min As Double
Dim MaxVolume As Double
Dim ws As Worksheet
Dim b As Long
Dim Maxcell, Mincell, Volcell As Long

For Each ws In ThisWorkbook.Worksheets

' finding the greatest % increase

Max = ws.Cells(2, 12).Value

For b = 2 To ticker_row
If ws.Cells(b, 12).Value > Max Then
Max = ws.Cells(b, 12).Value
Maxcell = b
End If
Next b

ws.Range("r2").Value = Max
ws.Range("Q2").Value = ws.Range("j" & Maxcell).Value

'find the greatest % decrease
Min = ws.Cells(2, 12).Value
For k = 2 To ticker_row
If ws.Cells(k, 12).Value < Min Then
Min = ws.Cells(k, 12).Value
Mincell = k
End If
Next k

ws.Range("r3").Value = Min
ws.Range("Q3").Value = ws.Range("j" & Mincell).Value

'finding the greatest total volume

MaxVolume = ws.Cells(2, 13).Value

For t = 2 To ticker_row
If ws.Cells(t, 13).Value > MaxVolume Then
MaxVolume = ws.Cells(t, 13).Value
Volcell = t
End If
Next t

ws.Range("r4").Value = MaxVolume
ws.Range("Q4").Value = ws.Range("j" & Volcell).Value

'setting the name for all the value

ws.Range("P2").Value = "Greatest % Increase"
ws.Range("P3").Value = "Greatest % Decrease"
ws.Range("P4").Value = "Greatest Total Volume"
ws.Range("Q1").Value = "Ticker"
ws.Range("R1").Value = "Value"




Next ws


End Sub