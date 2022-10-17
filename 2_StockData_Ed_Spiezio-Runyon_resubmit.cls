VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Multi_Tabs()
    'Run macro across multiple worksheet tabs within same workbook (source: https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html)
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        'Run TickerStocks subroutine
        Call TickerStocks
     Next
    Application.ScreenUpdating = True
End Sub

Sub TickerStocks()

'Set initial variables
Dim Ticker As String
Dim OpenVal As Double
Dim CloseVal As Double
Dim Percent_Change As Double
Dim Stock_Volume As Double
Dim LastRow As Long
Dim LastRowK As Long
Dim LastRowL As Long
Dim MaxResult As Double
Dim MinResult As Double
Dim Volume_Result As Double
Dim PctInc As String
Dim PctDec As String
Dim VolInc As String

'Set loop length to include last row of column (source: https://www.everything-excel.com/loop-to-last-row)
LastRow = Range("a" & Rows.Count).End(xlUp).Row

'Set initial variable for holding totals per ticker symbol
Dim Ticker_Total As Single
Ticker_Total = 0

'Track location for each ticker symbol and results in table1
Dim Table_Row As Integer
Table_Row = 2

'Add column and row titles
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"

'Set column formats (Source: https://stackoverflow.com/questions/36933414/excel-vba-round-of-decimal-for-2-points-or-to-entire-sheet)
Range("j1").EntireColumn.NumberFormat = "#0.00"
Range("k1").EntireColumn.NumberFormat = "#0.00%"
Range("L1").EntireColumn.NumberFormat = "#0,000"

'Set values to 0
OpenVal = 0
CloseVal = 0

'Loop through ticker symbols
For r = 2 To LastRow

    'Check to see if ticker symbol in list changes
    If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
    
        'Set the ticker symbol
        Ticker = Cells(r, 1).Value
        
        'Determine closing value for the year
        CloseVal = Cells(r, 6).Value
        
        'Calculate yearly price change and add to the ticker total
        Ticker_Total = CloseVal - OpenVal
        
        'Calculate the annual percentage change
        Percent_Change = Ticker_Total / Abs(OpenVal)
        
        'Calculate total stock volume
        Stock_Volume = Stock_Volume + Cells(r, 7).Value
        
        'List the total stock volume in table1
        Range("L" & Table_Row).Value = Stock_Volume
        
        'List the ticker symbols in table1
        Range("i" & Table_Row).Value = Ticker
        
        'List the total ticker amounts in table1
        Range("j" & Table_Row).Value = Ticker_Total
        
        'List the percent change in table1
        Range("k" & Table_Row).Value = Percent_Change
        
        'Determine if annual Ticker_Total change is positive
        If Ticker_Total >= 0 Then
        
            'If Ticker_Total is positive, color green
            Range("j" & Table_Row).Interior.ColorIndex = 4
            
        Else
        
            'If Ticker_Total is negative, color red
            Range("j" & Table_Row).Interior.ColorIndex = 3
            
        End If
        
        'Add one to table1
        Table_Row = Table_Row + 1
        
        'Reset values for the next round
        Ticker_Total = 0
        OpenVal = 0
        CloseVal = 0
        Stock_Volume = 0
        Percent_Change = 0
        
    'If the cell following a row is the same ticker symbol
    Else
    
        'Determine if opening value for the year has been populated
        If OpenVal = 0 Then
        
            'Determine opening value for the year
            OpenVal = Cells(r, 3).Value
            
        'Otherwise
        End If
            
        'Calculate total stock volume
        Stock_Volume = Stock_Volume + Cells(r, 7).Value
        
    End If
    
Next r

'PERCENT INCREASE/DECREASE AND VOLUME CALCULATIONS

'Find and list greatest % increase in Percent_Change
'Set loop length to include last row of column (source: https://www.everything-excel.com/loop-to-last-row)
LastRowK = Range("k" & Rows.Count).End(xlUp).Row

'Set initial max
MaxResult = Cells(2, 11)

'Loop through percent change
For k = 2 To LastRowK
        
    'If stock volume is greater than those above...
    If Cells(k + 1, 11).Value >= MaxResult Then
    
        'Replace saved MaxResult and associated ticker symbol
        MaxResult = Cells(k + 1, 11).Value
        PctInc = Cells(k + 1, 9).Value
        
    End If
    
Next k

'Print results to table
Cells(2, 16).Value = MaxResult
Cells(2, 15).Value = PctInc

'Format percentage of result to 2 decimal points (Source: https://excelchamps.com/vba/functions/formatpercent/)
Range("p2").Value = FormatPercent(Range("p2"))

'Find and list greatest % decrease in Percent_Change
'Set loop length to include last row of column (source: https://www.everything-excel.com/loop-to-last-row)
LastRowK = Range("k" & Rows.Count).End(xlUp).Row

'Set initial min
MinResult = Cells(2, 11)

'Loop through percent change
For k = 2 To LastRowK
    
    'If stock volume is greater than those above...
    If Cells(k + 1, 11).Value <= MinResult Then
    
        'Replace saved MinResult and associated ticker symbol
        MinResult = Cells(k + 1, 11).Value
        PctDec = Cells(k + 1, 9).Value
        
    End If
    
Next k

'Print results to table
Cells(3, 16).Value = MinResult
Cells(3, 15).Value = PctDec

'Format percentage of result to 2 decimal points (Source: https://excelchamps.com/vba/functions/formatpercent/)
Range("p3").Value = FormatPercent(Range("p3"))


'Find and list greatest total volume
'Set loop length to include last row of column (source: https://www.everything-excel.com/loop-to-last-row)
LastRowL = Range("L" & Rows.Count).End(xlUp).Row

'Set initial volume max
Volume_Result = Cells(2, 12)

'Loop through stock volume
For L = 2 To LastRowL

    'If stock volume is greater than those above...
    If Cells(L + 1, 12).Value >= Volume_Result Then
    
        'Replace saved stock volume result and associated ticker symbol
        Volume_Result = Cells(L + 1, 12).Value
        VolInc = Cells(L + 1, 9).Value
        
    End If
    
Next L

'Print results to table
Cells(4, 16).Value = Volume_Result
Cells(4, 15).Value = VolInc

'Format volume result to include commas (Source: https://stackoverflow.com/questions/45733877/numberformat-with-0-decimals-but-thousands-separator)
Range("p4").NumberFormat = "#,##0"

End Sub

