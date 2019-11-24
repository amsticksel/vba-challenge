Sub StockMarket3()
   For Each ws In Worksheets
   WorksheetName = ws.Name
   
  ' Variable to hold the ticker name, closing value, and opening value
  Dim TickerName As String
    Dim CloVal As Double
    Dim OpVal As Double
  ' Variable to hold the total volume and set it as 0 to start
  Dim TotalVolume As Double
  TotalVolume = 0
  
  'Variable to hold the percent change and store it as a %
  Dim PerChange As Double
    PerChange = 0
   
'create a summary table for them all
  Dim Summary_Table_Row As Long
  Summary_Table_Row = 2

'set the headers of the summary table
ws.Range("J1").Value = "Ticker"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"
ws.Range("M1").Value = "Total"

'find out when the last row happens
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  For i = 2 To LastRow
  
' Check if we are in a new ticker symbol, if it is not...
If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
'capture opening value and ticker name
OpVal = ws.Cells(i, 3).Value
TickerName = ws.Cells(i, 1).Value
'add to the total
TotalVolume = TotalVolume + CLng(ws.Cells(i, 7).Value)

'Now Check if we are on the last row of our ticker symbol, if it is
ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' Add to the Ticker Total
TotalVolume = TotalVolume + CLng(ws.Cells(i, 7).Value)
'Capture closing value
CloVal = ws.Cells(i, 6).Value

' Now we can print the StockMarket Summary in the Summary Table
'year change will be closing value - opening value
YrChange = CloVal - OpVal
'per change is closing value-opening value / opening value
    If OpVal = 0 Then
    PerChange = 0
    Else
    PerChange = ((CloVal - OpVal) / OpVal)
    End If
'Column J will hold the ticker name
ws.Range("J" & Summary_Table_Row).Value = TickerName
'Column K will hold the year change amount
ws.Range("K" & Summary_Table_Row).Value = YrChange
'With a nested if we can set the style of this column based on if it is negative or positive value

    If ws.Range("K" & Summary_Table_Row).Value >= 0 Then
    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
    Else
    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
    End If
    

'Column L has the percent change, which we format as a percentage
    ws.Range("L" & Summary_Table_Row).Value = PerChange
    ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
'Finally column M has the total volume
    ws.Range("M" & Summary_Table_Row).Value = TotalVolume
    
' Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
      
'Reset total volume
TotalVolume = 0

' Now, If the cell immediately following a row is the same brand...
    Else
      ' We just want to add to the Total
TotalVolume = TotalVolume + CLng(ws.Cells(i, 7).Value)

    End If
'Now we reset everything
YrChange = 0
PerChange = 0
      
Next i
'Bonuts Part, define some ranges
Dim rng As Range
Dim rng2 As Range
Dim dblMin As Double
Dim dblMax As Double
Dim totalMax As Double
'greatest percent increase
'greatest perent decrease
'greatest total volume
'setting the table up
ws.Range("O2").Value = "Greatest Percent Decrease"
ws.Range("O3").Value = "Greatest Percent Incresae"
ws.Range("O4").Value = "Largest Total"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Amount"

'Set range from which to determine smallest value
Set rng = ws.Range("L:L")
Set rng2 = ws.Range("M:M")

'Worksheet function MIN returns the smallest value in a range
'I learned and adapted this code from https://stackoverflow.com/questions/37049762/finding-minimum-and-maximum-values-from-range-of-values
dblMin = Application.WorksheetFunction.Min(rng)
dblMax = Application.WorksheetFunction.Max(rng)
totalMax = Application.WorksheetFunction.Max(rng2)
'finding the last row in our summary table
   LastRow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
   For ii = 2 To LastRow2
'look for minimum value and write it to the cell set up previously
If ws.Cells(ii, 12).Value = dblMin Then
ws.Range("Q2").Value = ws.Cells(ii, 12)
ws.Range("Q2").NumberFormat = "0.00%"
'get the ticker name
ws.Range("P2").Value = ws.Cells(ii, 10)
Else
End If
'another if statement for maximum percentage
If ws.Cells(ii, 12).Value = dblMax Then
ws.Range("Q3").Value = ws.Cells(ii, 12)
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("P3").Value = ws.Cells(ii, 10)
Else
End If
'and one more for the total
If ws.Cells(ii, 13).Value = totalMax Then
ws.Range("Q4").Value = ws.Cells(ii, 13)
ws.Range("P4").Value = ws.Cells(ii, 10)

Else
End If
'loop again
    Next ii
 
'loop to the next worksheet

    Next ws

End Sub

