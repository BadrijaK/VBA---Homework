VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockDataAnalysis():

'Loop over multiple worksheets
For Each ws In Worksheets

'Worksheet name
WorksheetName = ws.Name

'Defining variables

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Declaring variables
Dim i As Long
Dim j As Long
Dim tickerCount As Long
Dim endrow As Long
Dim percentChange As Double
Dim green As Double
Dim red As Double

'Assigning the intergers
tickerCount = 2
red = 3
green = 4

j = 2

'Navigate to last row
endrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
  'Summarising and looping variables
For i = 2 To endrow
            
   'Calculating the ticker symbol and yearly change variables
  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
        ws.Cells(tickerCount, 9).Value = ws.Cells(i, 1).Value
                
        ws.Cells(tickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
    'Colour formatting
  If ws.Cells(tickerCount, 10).Value < 0 Then

        ws.Cells(tickerCount, 10).Interior.ColorIndex = red
                
        Else
                
                    
        ws.Cells(tickerCount, 10).Interior.ColorIndex = green
                
        End If
      'Calculating percent change and formulating percent symbol
  If ws.Cells(j, 3).Value <> 0 Then

         p = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
         pe = p / ws.Cells(j, 3).Value
                    
        ws.Cells(tickerCount, 11).Value = Format(pe, "Percent")
                    
        Else
                    
        ws.Cells(tickerCount, 11).Value = Format(0, "Percent")
                    
        End If

       'Total stock volume
   ws.Cells(tickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
    
        tickerCount = tickerCount + 1
        j = i + 1
            
            End If
            
            Next i
            
                
        
            
    Next ws
    
        
                
End Sub


