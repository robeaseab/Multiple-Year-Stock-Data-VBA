Option Explicit

Sub stocks_hard()
Dim ws As Worksheet


For Each ws In Worksheets
'title rows
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Volume Total"
    
    Dim voltotal As Currency
        voltotal = 0
    Dim totalrow As Integer
        totalrow = 2
    
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim i As Variant
      
    For i = 2 To lastrow
        ws.Cells(i, 12).NumberFormat = "$#,##0.00_)"
        ws.Cells(i, 7).NumberFormat = "$#,##0.00_)"
        
        'creater ticker and total volumn columns
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(totalrow, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(totalrow, 12).Value = voltotal
            totalrow = totalrow + 1
            voltotal = 0
        ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            voltotal = voltotal + ws.Cells(i, 7).Value
        End If
    Next i
    
    Dim yr_open As Double
    Dim yr_close As Double
    Dim yr_change As Double
    
    
    yr_open = ws.Cells(2, 3).Value
    yr_close = 0
    yr_change = 0
    totalrow = 2
      
    For i = 2 To lastrow
        ws.Cells(i, 10).NumberFormat = "#,##0.000000000_)"
        ws.Cells(i, 11).NumberFormat = "0.00%"
        
        'year change and year % change columns
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And yr_open <> 0 Then
            yr_close = ws.Cells(i, 6).Value
            yr_change = yr_close - yr_open
            ws.Cells(totalrow, 10).Value = yr_change
            ws.Cells(totalrow, 11).Value = (yr_change / yr_open)

            totalrow = totalrow + 1
            yr_open = 0
            yr_close = 0
            yr_change = 0
            
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And yr_open = 0 Then
            'same as above but leave out percent to account for if year open is 0
            yr_close = Cells(i, 6).Value
            yr_change = yr_close - yr_open
            ws.Cells(totalrow, 10).Value = yr_change
            ws.Cells(totalrow, 11).Value = "Undefined"
            
            totalrow = totalrow + 1
            yr_open = 0
            yr_close = 0
            yr_change = 0
        
        ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value And ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            yr_open = ws.Cells(i, 3).Value
            
        End If
    Next i
    
 'color column I based on value
 Dim lastrow_col_I As Long
    lastrow_col_I = ws.Cells(Rows.Count, 9).End(xlUp).Row
 For i = 2 To lastrow_col_I
    If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    Else: ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
    
 Next i
 
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(4, 17).NumberFormat = "$#,##0.00_)"


Dim ticker As String
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_volume As Currency

greatest_increase = 0
greatest_decrease = 0
greatest_volume = 0
    

    
 For i = 2 To lastrow_col_I
 greatest_increase = WorksheetFunction.Max(Columns(11))
greatest_decrease = WorksheetFunction.Min(Columns(11))
greatest_volume = WorksheetFunction.Max(Columns(12))
    If ws.Cells(i, 11).Value = greatest_increase Then
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
    End If
    
    If ws.Cells(i, 11).Value = greatest_decrease Then
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
    End If
    
    If Cells(i, 12).Value = greatest_volume Then
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
    End If
      
 Next i
 
    
Next ws

End Sub




