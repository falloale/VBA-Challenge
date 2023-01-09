
Sub MODULE_2()
'define everything
    Dim sheet As Worksheet
    Dim ticker As String
    Dim vol As Integer
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim Summary_Table_Row As Integer

'this prevents my overflow error
    On Error Resume Next

'run through each worksheet
    For Each sheet In ThisWorkbook.Worksheets

'set headers
    sheet.Cells(1, 9).Value = "Ticker"
    sheet.Cells(1, 10).Value = "Yearly Change"
    sheet.Cells(1, 11).Value = "Percent Change"
    sheet.Cells(1, 12).Value = "Total Stock Volume"

'setup integers for loop
    Summary_Table_Row = 2

'loop
    For i = 2 To sheet.UsedRange.Rows.Count

If sheet.Cells(i + 1, 1).Value <> sheet.Cells(i, 1).Value Then

'find all the values
    ticker = sheet.Cells(i, 1).Value
    vol = sheet.Cells(i, 7).Value
    year_open = sheet.Cells(i, 3).Value
    year_close = sheet.Cells(i, 6).Value
    yearly_change = year_close - year_open
    percent_change = (year_close - year_open) / year_close
    
'insert values into summary
    sheet.Cells(Summary_Table_Row, 9).Value = ticker
    sheet.Cells(Summary_Table_Row, 10).Value = yearly_change
    sheet.Cells(Summary_Table_Row, 11).Value = percent_change
    sheet.Cells(Summary_Table_Row, 12).Value = vol

    Summary_Table_Row = Summary_Table_Row + 1
    vol = 0
End If
    'finish loop
Next i

sheet.Columns("K").NumberFormat = "0.00%"

'format columns colors
    Dim rg As Range
    Dim a As Double
    Dim z As Double
    
    
    
    Set rg = ws.Range("J2", Range("J2").End(xlDown))
    c = rg.Cells.Count
    For a = 1 To z
        
        Set color_cell = rg(a)
        Select Case color_cell

        Case Is >= 0
    With color_cell
    .Interior.ColorIndex = 4
    .Font.ColorIndex = 1
    
End With

    Case Is < 0
    With color_cell
    .Interior.ColorIndex = 3
    .Font.ColorIndex = 1
    
End With

End Select

Next a

'move to next worksheet
Next sheet

End Sub
    
