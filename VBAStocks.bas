Attribute VB_Name = "Module1"
Sub tickertotaler_moderate()
Dim ws As Worksheet
On Error Resume Next

For Each ws In Worksheets
ws.Cells(1, 8).Value = "Ticker"
ws.Cells(1, 9).Value = "Yearly Change"
ws.Cells(1, 10).Value = "Percent Change"
ws.Cells(1, 11).Value = "Total Stock Volume"

    'Setup the intergers for loop
    Summary_Table_Row = 2
        Row = ActiveSheet.UsedRange.Rows.Count
    
        'Find the data
        For i = 2 To ws.UsedRange.Rows.Count
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                 ticker = ws.Cells(i, 1).Value
                   vol = ws.Cells(i, 7).Value
                
                 year_open = ws.Cells(i, 3).Value
                    year_close = ws.Cells(i, 6).Value
        
                 yearly_change = year_close - year_open
                  percent_change = year_close / year_open
                
                   'Find the Summary
                  ws.Cells(Summary_Table_Row, 8).Value = ticker
                   ws.Cells(Summary_Table_Row, 9).Value = yearly_change
                    ws.Cells(Summary_Table_Row, 10).Value = percent_change
                    ws.Cells(Summary_Table_Row, 11).Value = vol
                Summary_Table_Row = Summary_Table_Row + 1
                
                vol = 0
        
             End If
        Next i
        
'Change the percentage formatting
ws.Columns("J").NumberFormat = "0.00%"


Next
End Sub
