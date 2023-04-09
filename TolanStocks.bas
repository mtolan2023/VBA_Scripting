Attribute VB_Name = "Module1"
Sub Multi_Year_Stocks()

'Analyze each sheet in Worksheet and find total annual volume for each Ticker symbol, find annual change (first open to last close) and format cells.

For Each ws In Worksheets

' part 1 format and find first last and volume

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

' volume needs to be string because totals are too large for long

Dim op As Double
Dim cl As Double
Dim change As Double
Dim percentchg As Double
Dim Ticker As String
Dim Volume As String
Dim sumrow As Integer

Volume = 0
sumrow = 2
    
' count total rows


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

 ' ticker volume and change
 
    
    For i = 2 To lastrow
        
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
           Volume = Volume + ws.Cells(i, 7).Value
            ws.Range("I" & sumrow).Value = Ticker
            ws.Range("L" & sumrow).Value = Volume
            'this is looking at last value in each ticker so cl is found here
            
            cl = ws.Cells(i, 6).Value
            change = cl - op
            
            ws.Range("J" & sumrow).Value = change
            
            'conditional formatting
            
                If ws.Range("J" & sumrow).Value > 0 Then
                    ws.Range("J" & sumrow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & sumrow).Interior.ColorIndex = 3
                End If
            percentchg = (change / op)
            ws.Range("K" & sumrow).Value = percentchg
            
            ' change column formatting to percentage
            
            ws.Range("K" & sumrow).NumberFormat = "0.00%"
            
        
            sumrow = sumrow + 1
            change = 0
            Volume = 0
            
            ' looking for first instance of each ticker
            
        ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            op = ws.Cells(i, 3).Value
            Volume = Volume + ws.Cells(i, 7).Value
            
        Else
            Volume = Volume + ws.Cells(i, 7).Value
            
        End If
        
        
    Next i
    
    
    
'  Second part. Find max Increase decrease and volume for each sheet

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"


Dim tickerrow As Long
Dim changerng As Range
Dim volrng As Range
Dim decrease As Double
Dim increase As Double
Dim maxvolume As String
Dim minticker As String
Dim maxticker As String
Dim volticker As String

maxvolume = 0

' find last value in new table

tickerrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

' set ranges for increase, decrease and max volume

Set changerng = ws.Range("K2:K" & tickerrow)
Set volrang = ws.Range("L2:L" & tickerrow)

' using worksheet functions to find min and max in new ranges

decrease = Application.WorksheetFunction.Min(changerng)
increase = Application.WorksheetFunction.Max(changerng)
maxvolume = Application.WorksheetFunction.Max(volrang)

' looking for values that match max and mins and populating values and ticker in columns p and q

For i = 2 To tickerrow
    If ws.Cells(i, 11).Value = decrease Then
        minticker = ws.Cells(i, 9).Value
        ws.Range("P3").Value = minticker
        ws.Range("Q3").Value = decrease
    ElseIf ws.Cells(i, 11).Value = increase Then
        maxticker = ws.Cells(i, 9).Value
        ws.Range("P2").Value = maxticker
        ws.Range("Q2").Value = increase
    End If
    If ws.Cells(i, 12).Value = maxvolume Then
        volticker = ws.Cells(i, 9).Value
        ws.Range("P4").Value = volticker
        ws.Range("Q4").Value = maxvolume
    End If
    
Next i

' adding percentage formatting to q cells and adjusting cells size to fit

    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Columns("A:Q").AutoFit
    
Next ws


End Sub


