Sub BetaStock()

For Each ws In Worksheets
Dim WorksheetName As String
WorksheetName = ws.Name

    Sheets(ws.Name).Select
    

'Dim ws As Worksheet
Dim Ticker As String
Dim YearlyChng As Double
Dim StockPercentChng As Double
Dim StockVolume As Double
Dim LastRow As Double
Dim Summary_table_row As Double
Dim stockOpen As Double
Dim stockClose As Double
Dim stockChange As Double
Dim tickerRow As Long

Summary_table_row = 2
tickerCount = 0


'Draw Summary Table ---------------------------------
'CLEAR SUMMARY COLUMNS ------------------------------------------
Columns("J:Q").Select
Selection.Clear

'SUMMARY tABLE HEADER
Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percentage Change"
Range("M1").Value = "Total Volume"

Set ws = ActiveSheet
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
tickerRow = 2

For i = 2 To LastRow

    If Cells(i + 1, 1) <> Cells(i, 1) Then
        ' ticker value
        Ticker = Cells(i, 1).Value
        
        'Print Summary Table
        ' Print the Ticker in the Summary Table
        Range("J" & Summary_table_row).Value = Ticker
        
        
        'Calculate values -------------------------------------------------
        stockOpen = ws.Cells(tickerRow, 3).Value
        stockClose = ws.Cells(i, 6).Value
        stockChange = stockClose - stockOpen
        StockVolume = StockVolume + ws.Cells(i, 7).Value
        
        If stockOpen = 0 Then
        StockPercentChng = 0
        Else
        StockPercentChng = stockChange / stockOpen
        End If
        'print stock yearly change
        Range("K" & Summary_table_row).Value = stockChange
        'print Percentage change
        
        If StockPercentChng <= 0 Then
        Range("K" & Summary_table_row).Interior.ColorIndex = 3
        Else
        Range("K" & Summary_table_row).Interior.ColorIndex = 4
        End If
        
        Range("L" & Summary_table_row).Style = "Percent"
        Range("L" & Summary_table_row).Value = StockPercentChng
        'Print Stock Volume
        
        'Add style by value
        
        
        Range("M" & Summary_table_row).Value = StockVolume
        
        
        Summary_table_row = Summary_table_row + 1
    
    End If


Next i

'----------------------------------------------

Dim ThisLastRow
'last row of greatest %

    ThisLastRow = Cells(Rows.Count, 10).End(xlUp).Row
    
    'Limits table --------------------------------
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
        
'iterate through summary table
    
    For x = 2 To ThisLastRow
    
    If Cells(x, 12).Value > Cells(x + 1, 12).Value Then
        
        Ticker_Gt_Inc = Cells(x, 10).Value
        Vol_Gt_Inc = Cells(x, 11).Value
    
    End If
    
    '
    If Cells(x, 12).Value < Cells(x + 1, 12).Value Then
        
        Ticker_Gt_Dec = Cells(x, 10).Value
        Vol_Gt_Dec = Cells(x, 11).Value
    
    End If
    
    
    If Cells(x, 13).Value > Cells(x + 1, 13).Value Then
        
        Ticker_Gt_Total_Vol = Cells(x, 10).Value
        Vol_Gt_Total = Cells(x, 13).Value
    
    End If
    
    'Print results
    Cells(2, 16).Value = Ticker_Gt_Inc
    Cells(2, 17).Value = Vol_Gt_Inc
    Cells(2, 17).Style = "Percent"
    Cells(3, 16).Value = Ticker_Gt_Dec
    Cells(3, 17).Value = Vol_Gt_Dec
    Cells(3, 17).Style = "Percent"
    Cells(4, 16).Value = Ticker_Gt_Total_Vol
    Cells(4, 17).Value = Vol_Gt_Total
    
    Next x

 Next ws


End Sub
