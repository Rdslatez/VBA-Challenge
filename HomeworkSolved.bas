Attribute VB_Name = "Module1"
Sub Stocks()
Dim openstart As Double
Dim incr As Double
Dim voltick As String
Dim decrtick As String
Dim incrtick As String
Dim decr As Double
Dim maxvol As Double
Dim closeend As Double
Dim change As Double
Dim perc As Double
Dim vol As Double
Dim tick As String
Dim length As Long
Dim side As Long
Dim ws As Worksheet
For Each ws In Worksheets
side = 2
maxvol = 0
decr = 0
incr = 0
length = Range("A1", Range("A1").End(xlDown)).Rows.Count
vol = 0
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

    For i = 2 To length
    
        tick = ws.Cells(i, 1).Value
        
        
        vol = vol + Cells(i, 7).Value
        
        
        If ws.Cells(i + 1, 1).Value <> tick Then
        closeend = ws.Cells(i, 6).Value
        change = closeend - openstart
        ws.Cells(side, 9).Value = tick
        ws.Cells(side, 10).NumberFormat = "0.00"
        ws.Cells(side, 10).Value = change
        ws.Columns("I").ColumnWidth = Len(Cells(1, 9)) * 2
        ws.Columns("J").ColumnWidth = Len(Cells(1, 10))
        ws.Columns("K").ColumnWidth = Len(Cells(1, 11))
        ws.Columns("L").ColumnWidth = Len(Cells(1, 12))
    
        
        If change > 0 Then
            ws.Cells(side, 10).Interior.ColorIndex = 4
        ElseIf change < 0 Then
            ws.Cells(side, 10).Interior.ColorIndex = 3
        End If
        perc = change / openstart
        ws.Cells(side, 11).Value = perc
        ws.Cells(side, 11).NumberFormat = "0.00%"
        ws.Cells(side, 12).Value = vol
        side = side + 1
            If perc > incr Then
            incr = perc
            incrtick = tick
            End If
            If perc < decr Then
            decr = perc
            decrtick = tick
            End If
            If maxvol < vol Then
            maxvol = vol
            voltick = tick
            End If
        vol = CLng(0)
        End If
        If ws.Cells(i - 1, 1).Value <> tick Then
        openstart = ws.Cells(i, 3).Value
        End If
        
        If (i + 1) > length Then
            ws.Cells(1, 16).Value = "Ticker"
            ws.Columns("P").ColumnWidth = Len(Cells(1, 16))
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Increase:"
            
            ws.Columns("Q").ColumnWidth = Len(Cells(2, 15))
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(2, 16).Value = incrtick
            ws.Cells(2, 17).Value = incr
            ws.Cells(3, 15).Value = "Greatest % Decrease:"
            ws.Cells(3, 16).Value = decrtick
            ws.Cells(3, 17).Value = decr
            ws.Cells(3, 17).NumberFormat = "0.00%"
            ws.Cells(4, 15).Value = "Greatest Total Volume:"
            ws.Columns("O").ColumnWidth = Len(Cells(4, 15))
            ws.Cells(4, 16).Value = voltick
            ws.Cells(4, 17).NumberFormat = "0"
            ws.Cells(4, 17).Value = maxvol
            
        End If
        
    
    Next i
    
Next ws



End Sub
