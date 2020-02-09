Attribute VB_Name = "Module1"
Sub stock()

Dim year, ticker As String
Dim rownum As Long
Dim total_vol, start_yr, end_yr As Double
Dim tbl_row As Integer
Dim ws As Worksheet
Dim max_id As Range


For Each ws In ThisWorkbook.Worksheets

    'rownum = Worksheets(year).UsedRange.Rows.Count
    rownum = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'set initial values and headers for the summary table
    start_yr = ws.Cells(2, 3).Value
    total_vol = 0
    ws.Range("I1:L1") = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    tbl_row = 2
    ws.Range("I2") = rownum

    For i = 2 To rownum
        If ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then
            total_vol = total_vol + Cells(i, 7)
        Else
            end_yr = ws.Cells(i, 6).Value
            ticker_name = ws.Cells(i, 1).Value
            total_diff = end_yr - start_yr
        
            If start_yr = 0 Then
                percent_change = "N/A"
            Else
                percent_change = total_diff / start_yr
            End If
        
            'write values into table
            ws.Range("I" & tbl_row).Value = ticker_name
            ws.Range("J" & tbl_row) = total_diff
            ws.Range("K" & tbl_row) = percent_change
            ws.Range("K" & tbl_row).NumberFormat = "0.00%"
            ws.Range("L" & tbl_row) = total_vol
                
            If ws.Range("J" & tbl_row) < 0 Then
                ws.Range("J" & tbl_row).Interior.Color = vbRed
            ElseIf ws.Range("J" & tbl_row) > 0 Then
                ws.Range("J" & tbl_row).Interior.Color = vbGreen
            End If
        
            tbl_row = tbl_row + 1
            total_vol = 0
            start_yr = Cells(i + 1, 3).Value
        End If
    
    Next i

'challenge

    'ws.Range("O1:P1") = Array("Ticker", "Value")
    'ws.Range("N2") = "Greatest % Increase"
    'ws.Range("N3") = "Greatest % Decrease"
    'ws.Range("N4") = "Greatest Total Volume"
    'ws.Range("P2") = Application.WorksheetFunction.Max(ws.Range("K:K"))
    'ws.Range("P3") = Application.WorksheetFunction.Min(ws.Range("K:K"))
    'ws.Range("P4") = Application.WorksheetFunction.Max(ws.Range("L:L"))
    'ws.Range("P2:P4").NumberFormat = "0.00%"
    'Set max_id = Range("K:K").Find(what:=Range("P2").Value, LookIn:=-4163, LookAt:=xlWhole)

    'ws.Range("O2").Value = max_id.Offset(0, -2).Value
Next ws

End Sub



