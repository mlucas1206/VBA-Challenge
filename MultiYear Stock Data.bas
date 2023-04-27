Attribute VB_Name = "Module1"
Sub Stocks()
'speedup code running
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False



'Worksheet Loop
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This is the subprocess for ticker column
Dim lastrow As Long

'To get the last row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox lastrow

'Put header
ws.Cells(1, 9).Value = "Ticker"

'Copy Values
ws.Range("I2:I" & lastrow).Value = ws.Range("A2:A" & lastrow).Value
'Remove Duplicates
ws.Range("I2:I" & lastrow).RemoveDuplicates Columns:=1, Header:=xlNo

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Rest of code
'Variables: Earliest row, Latest row earliest date, last date, last unique ticker
Dim first, last, firstrow, latestrow, tickrow As Long


'Variable for greatest things
Dim rng As Range

'Variable for ticker name
Dim tick As String

'Variable for totals and difference
Dim total, diff As Double

'Put headers
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"


'For loop every row to base on the tickers in the ticker column
tickrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

For a = 2 To tickrow
    'Refresh variables
    first = 0
    last = 0
    tick = ws.Cells(a, 9).Value
    total = 0
    
    'Every tick scans the whole table. Will use r for scanned row
    For r = 2 To lastrow
        'Conditional grab date if ticker matches
        If tick = ws.Cells(r, 1) Then
        
            'if first row for specific ticker, assign date
            If first = 0 And last = 0 Then
                first = ws.Cells(r, 2).Value
                last = first
                firstrow = r
                latestrow = r
                
            End If
        
            'Every row scan adjusts the variables' values
            'latest date check
            If ws.Cells(r, 2).Value > last Then
                last = ws.Cells(r, 2).Value
                latestrow = r
            'Earliest date check
            ElseIf ws.Cells(r, 2).Value < first Then
                first = ws.Cells(r, 2).Value
                firstrow = r
            End If
            
            'Add to total
            total = total + ws.Cells(r, 7).Value
            
        End If
       
    Next r
    
    'Place total into column L
    ws.Cells(a, 12).Value = total
    ws.Cells(a, 12).NumberFormat = "0"
        
    'Add formula for getting difference
    diff = ws.Cells(latestrow, 6).Value - ws.Cells(firstrow, 3).Value
    ws.Cells(a, 10).Value = diff
        
    'Condition for bg color
    If diff > 0 Then
        ws.Cells(a, 10).Interior.ColorIndex = 4
    ElseIf diff < 0 Then
        ws.Cells(a, 10).Interior.ColorIndex = 3
    Else
        ws.Cells(a, 10).Interior.ColorIndex = xlNone
    End If
        
    'Put percentage on column K
    ws.Cells(a, 11).Value = WorksheetFunction.Round(diff / ws.Cells(firstrow, 3), 4)
    ws.Cells(a, 11).NumberFormat = "0.0000"

Next a

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Greatest things
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'change to decimal format
ws.Cells(2, 17).NumberFormat = "0.0000"
ws.Cells(3, 17).NumberFormat = "0.0000"

ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K2:K" & tickrow))
ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K2:K" & tickrow))
ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L2:L" & tickrow))

    'find ticker
    'https://software-solutions-online.com/excel-vba-find-find-value-ws.Range-ws.Cells-vba/
    Set rng = ws.Range("K:K").Find(ws.Cells(2, 17).Value)
    ws.Cells(2, 16).Value = ws.Cells(rng.Row, 9)

    Set rng = ws.Range("K:K").Find(ws.Cells(3, 17).Value)
    ws.Cells(3, 16).Value = ws.Cells(rng.Row, 9)

    Set rng = ws.Range("L:L").Find(ws.Cells(4, 17).Value)
    ws.Cells(4, 16).Value = ws.Cells(rng.Row, 9)

    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Range("K2:K" & tickrow).NumberFormat = "0.00%"


'Autofit Columns
'From https://excelchamps.com/vba/autofit/#AutoFit_a_Column
ws.Range("J1:R1").EntireColumn.AutoFit

Next ws

'Re-enable settings
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

End Sub
