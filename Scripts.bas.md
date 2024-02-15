Sub TotalStockVol()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    
' Set an initial variable for holding the ticker name
    Dim Ticker As String
' Set initial variables for holding the total stock volume
    Dim Total_Stock_Volume As LongLong
    
'Set Total Stock Volume to 0
    Total_Stock_Volume = 0
'Set headers for the stats table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Long
        Summary_Table_Row = 2

'Find the last row
  Dim lastRow As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
' Loop through all ticker records
    For i = 2 To lastRow

' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' Set the Ticker
    Ticker = ws.Cells(i, 1).Value

 ' Add to the total stock volume
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

' Add the Ticker in the Summary Table
    ws.Range("I" & Summary_Table_Row).Value = Ticker
    
' Add the Total Stock Volume to the Summary Table
    ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
    
' Add one to the summary table row, so different tickers will be added to the row one by one.
    Summary_Table_Row = Summary_Table_Row + 1
'
'      ' Reset the Total Stock Volume
    Total_Stock_Volume = 0
'
' If the cell immediately following a row is the same brand...
    Else

' Add to the Total Stock Volume
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

    End If
    Next i
    Next ws
    
End Sub

Sub Yearlychange()

Dim YearEndValue   As Double
Dim YearStartValue  As Double
Dim lastRow         As Long
Dim ws              As Worksheet
Dim i               As Long
Dim Ticker          As String
Dim LastMonthDate   As String
Dim startMonthDate  As String
Dim Summarylastrow  As Long
Dim TargetRow       As Long


For Each ws In ThisWorkbook.Worksheets

lastRow = ws.Range("I" & Rows.Count).End(xlUp).Row
Summarylastrow = ws.Range("A" & Rows.Count).End(xlUp).Row

'Year start value would start from row 2

Ticker = ws.Range("I2").Value
YearStartValue = ws.Range("C2").Value
TargetRow = 2

ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percentage Change"

For i = 2 To Summarylastrow
    ' Check if we are still within the same ticker, if it is not, the last row of the same ticker would be i and the year end value would be in the same row F column.
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        YearEndValue = ws.Range("F" & i).Value
         ws.Range("J" & TargetRow).Value = YearEndValue - YearStartValue
        
        If YearStartValue = 0 Then
            ws.Range("K" & TargetRow).Value = 0
        Else
            ws.Range("K" & TargetRow).Value = ws.Range("J" & TargetRow).Value / YearStartValue
        End If
        ' Format the numbers of the yearly changes. If positive, then green; if negative, then red
        If ws.Range("J" & TargetRow).Value > 0 Then
          ws.Range("J" & TargetRow).Interior.ColorIndex = 4
        Else
          ws.Range("J" & TargetRow).Interior.ColorIndex = 3
        End If
        
        YearStartValue = ws.Range("C" & i + 1).Value
        TargetRow = TargetRow + 1
    End If
Next

ws.Range("K2: K" & lastRow).NumberFormat = "0.00%"

Next ws

End Sub

Sub GreatestValue()

Dim MaxIncrease As Double
Dim MaxDecrease As Double
Dim i   As Long
Dim lastRow     As Long
Dim ticker1      As String
Dim ticker2      As String
Dim ticker3      As String
Dim MaxTotalVol     As LongLong
Dim ws              As Worksheet

  For Each ws In ThisWorkbook.Worksheets
  
lastRow = ws.Range("I" & Rows.Count).End(xlUp).Row
MaxIncrease = ws.Range("K2").Value
MaxDecrease = ws.Range("K2").Value
MaxTotalVol = ws.Range("L2").Value
For i = 2 To lastRow

'Loop to find the max value by comparing each value row by row.

    If ws.Range("K" & i).Value > MaxIncrease Then
        MaxIncrease = ws.Range("K" & i).Value
        ticker1 = ws.Range("I" & i).Value
        
    End If
    If ws.Range("K" & i).Value < MaxDecrease Then
        MaxDecrease = ws.Range("K" & i).Value
        ticker2 = ws.Range("I" & i).Value
        
    End If
    If ws.Range("L" & i).Value > MaxTotalVol Then
        MaxTotalVol = ws.Range("L" & i).Value
        ticker3 = ws.Range("I" & i).Value
    End If

Next

'Create the second summary table with column headers and row names.

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Stock Volume"


ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

ws.Range("P2").Value = ticker1
ws.Range("Q2").Value = MaxIncrease
ws.Range("Q2").NumberFormat = "0.00%"

ws.Range("P3").Value = ticker2
ws.Range("Q3").Value = MaxDecrease
ws.Range("Q3").NumberFormat = "0.00%"

ws.Range("P4").Value = ticker3
ws.Range("Q4").Value = MaxTotalVol
ws.Range("Q4").NumberFormat = "0"

Next ws

End Sub

'Add sub to run all subs at once

Sub Consolidate()

Call TotalStockVol
Call Yearlychange
Call GreatestValue
MsgBox "Process Completed"
End Sub




