VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Dim Percent_Change As Integer
Dim Stock_Volume As Currency
Dim KRange As Range
Dim LastRow As Long
Dim LastRowK As Long
Dim LastRowL As Long
Dim MaxResult As Integer
Dim MinResult As Integer
Dim Volume_Result As Currency
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

'Loop through ticker symbols
For r = 2 To LastRow
    'Check to see if ticker symbol in list changes
    If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
        'Set the ticker symbol
        Ticker = Cells(r, 1).Value
        'Add to the ticker total
        Ticker_Total = Ticker_Total + Cells(r, 6).Value - Cells(r, 3).Value
        'List the ticker symbols in table1
        Range("i" & Table_Row).Value = Ticker
        'List the total ticker amounts in table1
        Range("j" & Table_Row).Value = Ticker_Total
        'List the percent change in table1
        Range("k" & Table_Row).Value = Percent_Change
        'Calculate total stock volume
        Stock_Volume = Stock_Volume + Cells(r, 7).Value
        'List the total stock volume in table1
        Range("L" & Table_Row).Value = Stock_Volume
        'Reset the Stock_Volume
        Stock_Volume = 0
        'Reset the Percent_Change
        Percent_Change = 0
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
        'Reset the Ticker_Total
        Ticker_Total = 0
    'If the cell following a row is the same ticker symbol
    Else
        'Calculate yearly price change and add to the ticker total
        Ticker_Total = Ticker_Total + Cells(r, 6).Value - Cells(r, 3).Value
        'Calculate the annual percentage change
        Percent_Change = Ticker_Total / Cells(r, 3).Value * 100
        'Calculate total stock volume
        Stock_Volume = Stock_Volume + Cells(r, 7).Value
        'List the total stock volume in table1
        Range("L" & Table_Row).Value = Stock_Volume
    End If
    Next r

'Find and list greatest % increase in Percent_Change
'Set loop length to include last row of column (source: https://www.everything-excel.com/loop-to-last-row)
LastRowK = Range("k" & Rows.Count).End(xlUp).Row
'Set initial max
MaxResult = Cells(2, 11)
For k = 2 To LastRowK
    If Cells(k + 1, 11).Value >= MaxResult Then
        MaxResult = Cells(k + 1, 11).Value
        PctInc = Cells(k + 1, 9).Value
    End If
Next k
Cells(2, 16).Value = MaxResult
Cells(2, 15).Value = PctInc

'Find and list greatest % decrease in Percent_Change
'Set loop length to include last row of column (source: https://www.everything-excel.com/loop-to-last-row)
LastRowK = Range("k" & Rows.Count).End(xlUp).Row
'Set initial min
MinResult = Cells(2, 11)
For k = 2 To LastRowK
    If Cells(k + 1, 11).Value <= MinResult Then
        MinResult = Cells(k + 1, 11).Value
        PctDec = Cells(k + 1, 9).Value
    End If
Next k
Cells(3, 16).Value = MinResult
Cells(3, 15).Value = PctDec

'Find and list greatest total volume
'Set loop length to include last row of column (source: https://www.everything-excel.com/loop-to-last-row)
LastRowL = Range("L" & Rows.Count).End(xlUp).Row
'Set initial max
Volume_Result = Cells(2, 12)
For L = 2 To LastRowL
    If Cells(L + 1, 12).Value >= Volume_Result Then
        Volume_Result = Cells(L + 1, 12).Value
        VolInc = Cells(L + 1, 9).Value
    End If
Next L
Cells(4, 16).Value = Volume_Result
Cells(4, 15).Value = VolInc

End Sub


